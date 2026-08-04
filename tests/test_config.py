from pathlib import Path

from src.config import Student, load_target_emails


def test_target_recipient_uses_student_name_for_matching_email(tmp_path: Path) -> None:
    recipients_file = tmp_path / "TargetEmail.txt"
    recipients_file.write_text(
        "2406069629 <2406069629@qq.com>\n"
        "External Name <external@example.com>\n",
        encoding="utf-8",
    )
    students = [
        Student(
            chinese_name="张三",
            pinyin_aliases=["zhang san"],
            emails=["2406069629@qq.com"],
        )
    ]

    assert load_target_emails(str(recipients_file), students) == [
        ("张三", "2406069629@qq.com"),
        ("External Name", "external@example.com"),
    ]


def test_target_recipient_without_students_keeps_configured_name(tmp_path: Path) -> None:
    recipients_file = tmp_path / "TargetEmail.txt"
    recipients_file.write_text(
        "2406069629 <2406069629@qq.com>\n", encoding="utf-8"
    )

    assert load_target_emails(str(recipients_file)) == [
        ("2406069629", "2406069629@qq.com")
    ]
