import imaplib
from datetime import datetime

import pytest

from src.email_fetcher import (
    EmailFetchError,
    _connect_imap_with_retries,
    fetch_email_pdf_attachments,
)


def test_imap_connection_retries_transient_failure(monkeypatch) -> None:
    attempts = 0
    expected_connection = object()

    def fake_connect(host, port, user, password):
        nonlocal attempts
        attempts += 1
        if attempts < 3:
            raise OSError("EOF occurred in violation of protocol")
        return expected_connection

    monkeypatch.setattr("src.email_fetcher._connect_imap", fake_connect)

    connection = _connect_imap_with_retries(
        "imap.example.com",
        993,
        "user@example.com",
        "password",
        retry_delays=(0, 0),
    )

    assert connection is expected_connection
    assert attempts == 3


def test_imap_connection_raises_after_all_attempts(monkeypatch) -> None:
    attempts = 0

    def fake_connect(host, port, user, password):
        nonlocal attempts
        attempts += 1
        raise OSError("TLS connection reset")

    monkeypatch.setattr("src.email_fetcher._connect_imap", fake_connect)

    with pytest.raises(EmailFetchError, match="Unable to connect"):
        _connect_imap_with_retries(
            "imap.example.com",
            993,
            "user@example.com",
            "password",
            retry_delays=(0, 0),
        )

    assert attempts == 3


def test_imap_authentication_failure_is_not_retried(monkeypatch) -> None:
    attempts = 0

    def fake_connect(host, port, user, password):
        nonlocal attempts
        attempts += 1
        raise imaplib.IMAP4.error("AUTHENTICATIONFAILED invalid credentials")

    monkeypatch.setattr("src.email_fetcher._connect_imap", fake_connect)

    with pytest.raises(EmailFetchError, match="Unable to connect"):
        _connect_imap_with_retries(
            "imap.example.com",
            993,
            "user@example.com",
            "wrong-password",
            retry_delays=(0, 0),
        )

    assert attempts == 1


def test_email_fetch_failure_is_propagated(monkeypatch, tmp_path) -> None:
    def fake_connect(*args, **kwargs):
        raise EmailFetchError("IMAP unavailable")

    monkeypatch.setattr(
        "src.email_fetcher._connect_imap_with_retries",
        fake_connect,
    )

    with pytest.raises(EmailFetchError, match="IMAP unavailable"):
        fetch_email_pdf_attachments(
            host="imap.example.com",
            port=993,
            user="user@example.com",
            password="password",
            save_dir=str(tmp_path),
            since_dt=datetime(2026, 11, 1),
            until_dt=datetime(2026, 11, 2),
        )
