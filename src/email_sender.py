"""SMTP email sender. Sends one or more emails, each with one merged-PDF attachment."""
import logging
import smtplib
import ssl
import time
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from pathlib import Path
from typing import List, Tuple

logger = logging.getLogger("wra")

_DEFAULT_SMTP_MAX_ATTEMPTS = 3
_DEFAULT_SMTP_RETRY_DELAYS = (5, 15)


def _is_transient_smtp_error(error: BaseException) -> bool:
    """Return True only for SMTP/network failures that are worth retrying."""
    if isinstance(error, smtplib.SMTPAuthenticationError):
        return False
    if isinstance(error, smtplib.SMTPRecipientsRefused):
        return False
    if isinstance(error, smtplib.SMTPResponseException):
        # SMTP 4xx responses are temporary; 5xx responses are permanent.
        return 400 <= error.smtp_code < 500
    return isinstance(
        error,
        (
            ssl.SSLError,
            OSError,
            smtplib.SMTPServerDisconnected,
            smtplib.SMTPConnectError,
            smtplib.SMTPHeloError,
        ),
    )


def format_recipient_header(recipients: List[Tuple[str, str]]) -> str:
    """Build the RFC-compliant To header with UTF-8 display names."""
    return ", ".join(formataddr((name, email)) for name, email in recipients)


def send_emails_with_attachments(
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str,
    sender_display_name: str,
    recipients: List[Tuple[str, str]],
    subject_base: str,
    body_text: str,
    attachments: List[str],
    dry_run: bool = False,
    max_attempts: int = _DEFAULT_SMTP_MAX_ATTEMPTS,
    retry_delays: Tuple[int, ...] = _DEFAULT_SMTP_RETRY_DELAYS,
) -> bool:
    """Send the merged report. If multiple attachments, send them as separate emails.

    - subject_base: base subject; if multiple attachments, ' (i/n)' is appended
    - body_text: included verbatim in every email; appended with part info if split
    - attachments: list of file paths. Each is sent in its OWN email (one per attachment).
    - Transient SMTP/TLS failures are retried with a fresh connection. The same
      Message-ID is reused across attempts to help recipients' servers deduplicate.

    Returns True iff every email was sent (or, in dry_run, would have been sent) successfully.
    """
    if not recipients:
        logger.error("No recipients provided.")
        return False
    if not attachments:
        logger.error("No attachments to send.")
        return False

    n = len(attachments)
    all_ok = True
    max_attempts = max(1, int(max_attempts))

    for i, att_path in enumerate(attachments, start=1):
        if n > 1:
            subject = f"{subject_base} ({i}/{n})"
            extra_body = f"\n\n（本邮件附件为分卷 {i}/{n}，共 {n} 个附件分多封发送）"
        else:
            subject = subject_base
            extra_body = ""

        msg = EmailMessage()
        msg["From"] = formataddr((sender_display_name, smtp_user))
        to_header = format_recipient_header(recipients)
        msg["To"] = to_header
        msg["Subject"] = subject
        message_domain = smtp_user.rsplit("@", 1)[-1] if "@" in smtp_user else None
        msg["Message-ID"] = make_msgid(domain=message_domain)
        msg.set_content(body_text + extra_body)

        # Attach the PDF
        try:
            data = Path(att_path).read_bytes()
        except OSError as e:
            logger.error(f"Failed to read attachment {att_path}: {e}")
            all_ok = False
            continue
        att_name = Path(att_path).name
        msg.add_attachment(
            data,
            maintype="application",
            subtype="pdf",
            filename=att_name,
        )

        size_mb = len(data) / 1024 / 1024
        logger.info(f"Prepared email {i}/{n}: subject='{subject}', "
                    f"to={recipients}, to_header={to_header!r}, "
                    f"attachment='{att_name}' ({size_mb:.2f} MB)")

        if dry_run:
            logger.info(f"[DRY RUN] would send email {i}/{n} (skipped)")
            continue

        sent = False
        for attempt in range(1, max_attempts + 1):
            try:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(
                    smtp_host,
                    smtp_port,
                    context=context,
                    timeout=60,
                ) as server:
                    server.login(smtp_user, smtp_password)
                    refused = server.send_message(msg)

                if refused:
                    logger.error(
                        f"Failed to send email {i}/{n}: recipients refused: {refused}"
                    )
                    all_ok = False
                    break

                logger.info(f"Sent email {i}/{n} successfully.")
                sent = True
                break
            except (smtplib.SMTPException, OSError) as e:
                transient = _is_transient_smtp_error(e)
                if not transient or attempt >= max_attempts:
                    logger.error(
                        f"Failed to send email {i}/{n} "
                        f"(attempt {attempt}/{max_attempts}): {e}"
                    )
                    all_ok = False
                    break

                delay_index = min(attempt - 1, len(retry_delays) - 1)
                delay = retry_delays[delay_index] if retry_delays else 0
                logger.warning(
                    f"Temporary SMTP failure for email {i}/{n} "
                    f"(attempt {attempt}/{max_attempts}): {e}. "
                    f"Reconnecting and retrying in {delay}s."
                )
                if delay > 0:
                    time.sleep(delay)

        if not sent:
            all_ok = False

    return all_ok


def build_email_subject_and_body(
    target_time: str,
    submitted_names: List[str],
    not_submitted_names: List[str],
) -> Tuple[str, str]:
    """Build the email subject and body following the user's required format.

    Subject: 周报汇总 - YYYY年M月D日   (parsed from target_time YYYY_MM_DD_HH_MM)
    Body:
        已交周报：xxx，yyy
        未交周报：zzz，www
    """
    parts = target_time.split("_")
    if len(parts) >= 3:
        try:
            y = int(parts[0]); m = int(parts[1]); d = int(parts[2])
            subject = f"周报汇总 - {y}年{m}月{d}日"
        except ValueError:
            subject = f"周报汇总 - {target_time}"
    else:
        subject = f"周报汇总 - {target_time}"

    submitted_str = "，".join(submitted_names) if submitted_names else "（无）"
    not_submitted_str = "，".join(not_submitted_names) if not_submitted_names else "（无）"

    body = (
        f"已交周报：{submitted_str}\n"
        f"未交周报：{not_submitted_str}\n"
    )
    return subject, body
