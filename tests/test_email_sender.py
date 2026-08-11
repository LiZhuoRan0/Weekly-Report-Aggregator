from src.email_sender import format_recipient_header


def test_recipient_header_contains_encoded_real_names() -> None:
    header = format_recipient_header([
        ("\u5434\u94ed\u6656", "minghuiwu1999@foxmail.com"),
        ("\u6c5f\u97e9\u9a8f", "1758082886@qq.com"),
    ])

    assert "minghuiwu1999@foxmail.com" in header
    assert "1758082886@qq.com" in header
    assert "=?utf-8?" in header.lower()
