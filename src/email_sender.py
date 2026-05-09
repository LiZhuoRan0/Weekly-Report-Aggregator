"""SMTP email sender. Sends one or more emails, each with one merged-PDF attachment."""
import logging
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from typing import List, Optional, Tuple

logger = logging.getLogger("wra")


def send_emails_with_attachments(
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str,
    sender_display_name: str,
    recipients: List[str],
    subject_base: str,
    body_text: str,
    attachments: List[str],
    dry_run: bool = False,
) -> bool:
    """Send the merged report. If multiple attachments, send them as separate emails.

    - subject_base: base subject; if multiple attachments, ' (i/n)' is appended
    - body_text: included verbatim in every email; appended with part info if split
    - attachments: list of file paths. Each is sent in its OWN email (one per attachment).

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

    for i, att_path in enumerate(attachments, start=1):
        if n > 1:
            subject = f"{subject_base} ({i}/{n})"
            extra_body = f"\n\n（本邮件附件为分卷 {i}/{n}，共 {n} 个附件分多封发送）"
        else:
            subject = subject_base
            extra_body = ""

        msg = EmailMessage()
        msg["From"] = formataddr((sender_display_name, smtp_user))
        msg["To"] = ", ".join(recipients)
        msg["Subject"] = subject
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
                    f"to={recipients}, attachment='{att_name}' ({size_mb:.2f} MB)")

        if dry_run:
            logger.info(f"[DRY RUN] would send email {i}/{n} (skipped)")
            continue

        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=60) as server:
                server.login(smtp_user, smtp_password)
                server.send_message(msg)
            logger.info(f"Sent email {i}/{n} successfully.")
        except (smtplib.SMTPException, OSError) as e:
            logger.error(f"Failed to send email {i}/{n}: {e}")
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
