"""Fetch recent emails from QQ IMAP and extract PDF attachments.

Filtering: an email is considered a candidate "weekly report" if any of:
  - its subject mentions 周报 / weekly / report
  - its body mentions 周报 / weekly / report
  - it has a PDF attachment whose filename hints at a person name
  - its sender matches one of the known student emails

We default to permissive: ANY PDF attachment in recent emails becomes a
candidate, and the matcher decides whether it actually belongs to a student.
"""
from __future__ import annotations

import email
import imaplib
import logging
import os
import re
import socket
import time
from datetime import datetime, timedelta
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import List, Optional, Set

from .matcher import PdfCandidate

logger = logging.getLogger("wra")

_IMAP_TIMEOUT_SECONDS = 60
_FETCH_MAX_ATTEMPTS = 3
_FETCH_RETRY_SLEEP_SECONDS = 2


def _close_imap(M: Optional[imaplib.IMAP4_SSL]) -> None:
    """Best-effort close/logout for an IMAP connection."""
    if M is None:
        return

    try:
        M.close()
    except Exception:
        pass

    try:
        M.logout()
    except Exception:
        pass


def _connect_imap(
    host: str,
    port: int,
    user: str,
    password: str,
) -> imaplib.IMAP4_SSL:
    """Create a fresh IMAP connection and select INBOX."""
    logger.info(f"Connecting to IMAP {host}:{port} as {user}")

    M = imaplib.IMAP4_SSL(host, port)
    M.login(user, password)

    typ, _ = M.select("INBOX", readonly=True)
    if typ != "OK":
        _close_imap(M)
        raise imaplib.IMAP4.error("Failed to select INBOX")

    return M


def _fetch_raw_message_with_retries(
    M: Optional[imaplib.IMAP4_SSL],
    host: str,
    port: int,
    user: str,
    password: str,
    msg_uid: bytes,
    max_attempts: int = _FETCH_MAX_ATTEMPTS,
) -> tuple[Optional[imaplib.IMAP4_SSL], Optional[bytes]]:
    """Fetch one email with retries.

    If the current IMAP connection times out or becomes unusable, close it,
    reconnect, and retry this message. If this message still fails after all
    retries, skip only this message and let later messages continue.
    """
    last_error: Optional[Exception] = None

    for attempt in range(1, max_attempts + 1):
        try:
            if M is None:
                M = _connect_imap(host, port, user, password)

            typ, msg_data = M.uid("FETCH", msg_uid, "(RFC822)")
            if typ != "OK" or not msg_data or not msg_data[0]:
                logger.warning(
                    f"Fetch message {msg_uid!r} returned no data "
                    f"(attempt {attempt}/{max_attempts})"
                )
                return M, None

            raw = msg_data[0][1]
            if not isinstance(raw, (bytes, bytearray)):
                logger.warning(
                    f"Fetch message {msg_uid!r} returned non-bytes payload "
                    f"(attempt {attempt}/{max_attempts})"
                )
                return M, None

            return M, bytes(raw)

        except (
            socket.timeout,
            TimeoutError,
            OSError,
            imaplib.IMAP4.abort,
            imaplib.IMAP4.error,
        ) as e:
            last_error = e
            logger.warning(
                f"Failed to fetch message {msg_uid!r} "
                f"(attempt {attempt}/{max_attempts}): {e}"
            )

            # Important: after timeout, this connection may be poisoned.
            # Do not reuse it for the next message.
            _close_imap(M)
            M = None

            if attempt < max_attempts:
                time.sleep(_FETCH_RETRY_SLEEP_SECONDS)

    logger.warning(
        f"Skipped message {msg_uid!r} after {max_attempts} failed fetch attempt(s): "
        f"{last_error}"
    )
    return M, None


def _decode_str(s) -> str:
    """Decode an RFC2047 / bytes header string to a clean str."""
    if s is None:
        return ""
    if isinstance(s, bytes):
        try:
            return s.decode("utf-8", errors="replace")
        except Exception:
            return s.decode("latin-1", errors="replace")
    parts = decode_header(s)
    out = []
    for text, charset in parts:
        if isinstance(text, bytes):
            try:
                out.append(text.decode(charset or "utf-8", errors="replace"))
            except (LookupError, TypeError):
                out.append(text.decode("utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def _safe_filename(name: str) -> str:
    """Make a filesystem-safe filename out of an attachment name."""
    name = name.strip()
    # Replace path separators and other risky characters
    name = re.sub(r"[\\/]", "_", name)
    name = re.sub(r"[\x00-\x1f]", "", name)
    if not name:
        name = "attachment.pdf"
    return name


# Heuristic keywords for body/subject filtering (we still keep PDFs that don't
# match any of these — matching gets the final say.)
_HINT_KEYWORDS = ("周报", "weekly", "weekly report", "report", "周汇报", "周总结")


def fetch_email_pdf_attachments(
    host: str,
    port: int,
    user: str,
    password: str,
    save_dir: str,
    since_dt: datetime,
    until_dt: Optional[datetime] = None,
    student_emails: Optional[Set[str]] = None,
) -> List[PdfCandidate]:
    """Connect via IMAP-SSL, search recent emails, save PDF attachments to save_dir.

    Returns a list of PdfCandidate(source='email', ...) for each PDF saved.

    since_dt / until_dt are *naive or aware* datetimes — IMAP SEARCH uses date
    granularity (day-level), so we widen the IMAP search by 1 day on each side
    and then filter by the email's actual Date header.
    """
    save_path = Path(save_dir)
    save_path.mkdir(parents=True, exist_ok=True)
    student_emails = student_emails or set()

    candidates: List[PdfCandidate] = []
    M: Optional[imaplib.IMAP4_SSL] = None

    try:
        # Socket timeout to avoid long hangs on bad networks/credentials.
        # A timed-out connection will be discarded and rebuilt per message fetch.
        socket.setdefaulttimeout(_IMAP_TIMEOUT_SECONDS)

        M = _connect_imap(host, port, user, password)

        # Build IMAP SINCE date (one day earlier to be safe)
        search_since = (since_dt - timedelta(days=1)).strftime("%d-%b-%Y")
        # IMAP SEARCH: SINCE returns messages with internal date >= the given date
        typ, data = M.uid("SEARCH", None, "SINCE", search_since)
        if typ != "OK" or not data or not data[0]:
            logger.info(f"No messages found since {search_since}")
            return []

        msg_ids = data[0].split()
        logger.info(f"IMAP returned {len(msg_ids)} message UID(s) since {search_since}")

        for msg_id in msg_ids:
            M, raw = _fetch_raw_message_with_retries(
                M=M,
                host=host,
                port=port,
                user=user,
                password=password,
                msg_uid=msg_id,
            )

            if raw is None:
                continue

            try:
                msg = email.message_from_bytes(raw)
            except Exception as e:
                logger.warning(f"Failed to parse message {msg_id!r}: {e}")
                continue

            # Filter by Date header against [since_dt, until_dt]
            try:
                date_hdr = msg.get("Date")
                if date_hdr:
                    msg_dt = parsedate_to_datetime(date_hdr)
                else:
                    msg_dt = None
            except Exception:
                msg_dt = None

            if msg_dt is not None:
                # Make naive for comparison if needed
                msg_naive = msg_dt.replace(tzinfo=None) if msg_dt.tzinfo else msg_dt
                since_naive = since_dt.replace(tzinfo=None) if since_dt.tzinfo else since_dt
                if msg_naive < since_naive:
                    continue
                if until_dt is not None:
                    until_naive = until_dt.replace(tzinfo=None) if until_dt.tzinfo else until_dt
                    if msg_naive > until_naive:
                        continue

            subject = _decode_str(msg.get("Subject", ""))
            from_hdr = _decode_str(msg.get("From", ""))
            sender_email_addr = parseaddr(from_hdr)[1].lower() if from_hdr else ""

            # Quick body extraction (only used as a hint)
            body_text = _extract_text_body(msg).lower()
            subj_lower = subject.lower()
            has_hint = any(kw in subj_lower or kw in body_text for kw in _HINT_KEYWORDS)
            sender_is_known_student = sender_email_addr in student_emails

            # Iterate parts to find PDF attachments
            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                disp = (part.get("Content-Disposition") or "").lower()
                ctype = (part.get_content_type() or "").lower()

                fname = part.get_filename()
                fname_decoded = _decode_str(fname) if fname else ""

                # Accept if:
                #  - explicitly attached AND filename ends with .pdf, OR
                #  - content-type is application/pdf
                is_pdf = (
                    ctype == "application/pdf"
                    or (fname_decoded.lower().endswith(".pdf"))
                )
                if not is_pdf:
                    continue

                # Heuristic decision to keep this PDF as a candidate.
                # We keep ALL PDFs from recent emails — matcher decides ownership.
                # The hints/sender are stored for the matcher to use.

                payload = part.get_payload(decode=True)
                if not payload:
                    continue

                # Save to disk
                safe_name = _safe_filename(fname_decoded or "attachment.pdf")
                if not safe_name.lower().endswith(".pdf"):
                    safe_name += ".pdf"

                # Avoid filename collision
                target = save_path / f"{msg_id.decode() if isinstance(msg_id, bytes) else msg_id}_{safe_name}"
                try:
                    target.write_bytes(payload)
                except OSError as e:
                    logger.warning(f"Failed to save attachment {safe_name}: {e}")
                    continue

                # mtime: use the message's date if available, else now
                mtime = (msg_dt.timestamp() if msg_dt else datetime.now().timestamp())
                try:
                    os.utime(target, (mtime, mtime))
                except OSError:
                    pass

                logger.info(
                    f"Saved email attachment: {target.name} "
                    f"(from={sender_email_addr}, subj='{subject[:40]}', "
                    f"hint={has_hint}, known_student={sender_is_known_student})"
                )

                candidates.append(PdfCandidate(
                    path=str(target.resolve()),
                    source="email",
                    mtime=mtime,
                    sender_email=sender_email_addr or None,
                    original_name=fname_decoded or safe_name,
                ))
    except imaplib.IMAP4.error as e:
        logger.error(f"IMAP error: {e}")
    except Exception as e:
        logger.error(f"Unexpected error during IMAP fetch: {e}")
    finally:
        _close_imap(M)

    logger.info(f"Total PDFs saved from email: {len(candidates)}")
    return candidates


def _extract_text_body(msg: email.message.Message) -> str:
    """Best-effort plain-text body extraction (returns empty string on failure)."""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_maintype() != "text":
                    continue
                if part.get_content_subtype() != "plain":
                    continue
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    try:
                        return payload.decode(charset, errors="replace")
                    except (LookupError, TypeError):
                        return payload.decode("utf-8", errors="replace")
            # fall back to html
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        try:
                            return payload.decode(charset, errors="replace")
                        except (LookupError, TypeError):
                            return payload.decode("utf-8", errors="replace")
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                try:
                    return payload.decode(charset, errors="replace")
                except (LookupError, TypeError):
                    return payload.decode("utf-8", errors="replace")
    except Exception:
        pass
    return ""
