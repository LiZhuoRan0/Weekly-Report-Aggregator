from pathlib import Path

import pytest

from src.config import Student, load_target_emails


def test_target_recipient_keeps_configured_name(tmp_path: Path) -> None:
    recipients_file = tmp_path / "TargetEmail.txt"
    recipients_file.write_text(
        "\u7504\u7a0b <2406069629@qq.com>\n"
        "External Name <external@example.com>\n",
        encoding="utf-8",
    )
    students = [
        Student(
            chinese_name="\u5f20\u4e09",
            pinyin_aliases=["zhang san"],
            emails=["2406069629@qq.com"],
        )
    ]

    assert load_target_emails(str(recipients_file), students) == [
        ("\u7504\u7a0b", "2406069629@qq.com"),
        ("External Name", "external@example.com"),
    ]


@pytest.mark.parametrize("line", [
    "2406069629 <2406069629@qq.com>",
    "minghuiwu1999 <minghuiwu1999@foxmail.com>",
])
def test_target_recipient_rejects_suspicious_display_name(
    tmp_path: Path, line: str
) -> None:
    recipients_file = tmp_path / "TargetEmail.txt"
    recipients_file.write_text(line + "\n", encoding="utf-8")

    with pytest.raises(ValueError, match="Suspicious display name"):
        load_target_emails(str(recipients_file))
