import smtplib

from src.email_sender import format_recipient_header, send_emails_with_attachments


def test_recipient_header_contains_encoded_real_names() -> None:
    header = format_recipient_header([
        ("\u5434\u94ed\u6656", "minghuiwu1999@foxmail.com"),
        ("\u6c5f\u97e9\u9a8f", "1758082886@qq.com"),
    ])

    assert "minghuiwu1999@foxmail.com" in header
    assert "1758082886@qq.com" in header
    assert "=?utf-8?" in header.lower()


def test_smtp_transient_failure_is_retried_with_same_message_id(
    monkeypatch, tmp_path
) -> None:
    attachment = tmp_path / "report.pdf"
    attachment.write_bytes(b"%PDF-1.4\n%%EOF\n")
    attempts = 0
    message_ids = []

    class FakeSMTP:
        def __init__(self, *args, **kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc_value, traceback):
            return False

        def login(self, user, password):
            pass

        def send_message(self, message):
            nonlocal attempts
            attempts += 1
            message_ids.append(message["Message-ID"])
            if attempts == 1:
                raise smtplib.SMTPServerDisconnected(
                    "EOF occurred in violation of protocol"
                )
            return {}

    monkeypatch.setattr("src.email_sender.smtplib.SMTP_SSL", FakeSMTP)

    sent = send_emails_with_attachments(
        smtp_host="smtp.example.com",
        smtp_port=465,
        smtp_user="sender@example.com",
        smtp_password="password",
        sender_display_name="Sender",
        recipients=[("Recipient", "recipient@example.com")],
        subject_base="Weekly report",
        body_text="Body",
        attachments=[str(attachment)],
        max_attempts=3,
        retry_delays=(0, 0),
    )

    assert sent is True
    assert attempts == 2
    assert len(message_ids) == 2
    assert message_ids[0] == message_ids[1]


def test_smtp_authentication_failure_is_not_retried(
    monkeypatch, tmp_path
) -> None:
    attachment = tmp_path / "report.pdf"
    attachment.write_bytes(b"%PDF-1.4\n%%EOF\n")
    attempts = 0

    class FakeSMTP:
        def __init__(self, *args, **kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc_value, traceback):
            return False

        def login(self, user, password):
            nonlocal attempts
            attempts += 1
            raise smtplib.SMTPAuthenticationError(535, b"Authentication failed")

        def send_message(self, message):
            raise AssertionError("send_message must not be called")

    monkeypatch.setattr("src.email_sender.smtplib.SMTP_SSL", FakeSMTP)

    sent = send_emails_with_attachments(
        smtp_host="smtp.example.com",
        smtp_port=465,
        smtp_user="sender@example.com",
        smtp_password="wrong-password",
        sender_display_name="Sender",
        recipients=[("Recipient", "recipient@example.com")],
        subject_base="Weekly report",
        body_text="Body",
        attachments=[str(attachment)],
        max_attempts=3,
        retry_delays=(0, 0),
    )

    assert sent is False
    assert attempts == 1
