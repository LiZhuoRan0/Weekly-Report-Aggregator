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
from datetime import datetime, timedelta
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import List, Optional, Set

from .matcher import PdfCandidate

logger = logging.getLogger("wra")


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
        logger.info(f"Connecting to IMAP {host}:{port} as {user}")
        # 30s socket timeout to avoid long hangs on bad networks/credentials
        import socket
        socket.setdefaulttimeout(30)
        M = imaplib.IMAP4_SSL(host, port)
        M.login(user, password)
        # QQ's selectable mailbox name for inbox is "INBOX". For 已发送 etc., names differ.
        typ, _ = M.select("INBOX", readonly=True)
        if typ != "OK":
            logger.error("Failed to select INBOX")
            return []

        # Build IMAP SINCE date (one day earlier to be safe)
        search_since = (since_dt - timedelta(days=1)).strftime("%d-%b-%Y")
        # IMAP SEARCH: SINCE returns messages with internal date >= the given date
        typ, data = M.search(None, "SINCE", search_since)
        if typ != "OK" or not data or not data[0]:
            logger.info(f"No messages found since {search_since}")
            return []

        msg_ids = data[0].split()
        logger.info(f"IMAP returned {len(msg_ids)} message id(s) since {search_since}")

        for msg_id in msg_ids:
            try:
                typ, msg_data = M.fetch(msg_id, "(RFC822)")
                if typ != "OK" or not msg_data or not msg_data[0]:
                    continue
                raw = msg_data[0][1]
                if not isinstance(raw, (bytes, bytearray)):
                    continue
                msg = email.message_from_bytes(raw)
            except Exception as e:
                logger.warning(f"Failed to fetch/parse message {msg_id}: {e}")
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
        if M is not None:
            try:
                M.close()
            except Exception:
                pass
            try:
                M.logout()
            except Exception:
                pass

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
