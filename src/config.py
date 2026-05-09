"""Configuration and input file loaders."""
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List


@dataclass
class Student:
    """A single student with Chinese name and one or more email addresses."""
    chinese_name: str
    emails: List[str] = field(default_factory=list)


@dataclass
class Config:
    imap_host: str
    imap_port: int
    imap_user: str
    imap_password: str

    smtp_host: str
    smtp_port: int
    smtp_user: str
    smtp_password: str

    file_path: str
    target_time: str  # "YYYY_MM_DD_HH_MM"
    lookback_days: int
    max_attachment_size_mb: int
    sender_display_name: str


def load_config(path: str = "config.json") -> Config:
    """Load config.json from disk."""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(
            f"Config file not found: {path}. Copy config.json.example to config.json and fill in your credentials."
        )
    with p.open("r", encoding="utf-8") as f:
        data = json.load(f)

    return Config(
        imap_host=data["imap"]["host"],
        imap_port=int(data["imap"]["port"]),
        imap_user=data["imap"]["user"],
        imap_password=data["imap"]["password"],
        smtp_host=data["smtp"]["host"],
        smtp_port=int(data["smtp"]["port"]),
        smtp_user=data["smtp"]["user"],
        smtp_password=data["smtp"]["password"],
        file_path=data["FilePath"],
        target_time=data["TargetTime"],
        lookback_days=int(data.get("lookback_days", 3)),
        max_attachment_size_mb=int(data.get("max_attachment_size_mb", 20)),
        sender_display_name=data.get("sender_display_name", data["smtp"]["user"]),
    )


# Splits on commas (中/英), semicolons (中/英), whitespace
_DELIM_RE = re.compile(r"[\s,，;；]+")
_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

_INLINE_COMMENT_RE = re.compile(
    r"(#|//|备注[:：]|注释[:：]|comment[:：]).*$",
    re.IGNORECASE,
)

def _strip_comment(line: str) -> str:
    """Remove inline annotations/comments from one line."""
    return _INLINE_COMMENT_RE.sub("", line).strip()

def load_students(path: str) -> List[Student]:
    """Parse students.txt.

    Each non-empty line: <Chinese name> <email1> [<email2> ...]
    Delimiters: whitespace / comma (, ，) / semicolon (; ；) — any combination.
    Order is preserved (used as the merged-PDF ordering).
    """
    students: List[Student] = []
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Students file not found: {path}")

    with p.open("r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()

            if not line or line.startswith("#") or line.startswith("//"):
                continue

            line = _strip_comment(line)

            if not line:
                continue

            tokens = [t for t in _DELIM_RE.split(line) if t]
            if not tokens:
                continue
            # First token = Chinese name; rest = emails
            name = tokens[0]
            emails = [m.group(0).lower() for m in _EMAIL_RE.finditer(line)]
            if not emails:
                # Allow students without email (won't help for email matching)
                # but still keep them in the ordering.
                pass
            students.append(Student(chinese_name=name, emails=emails))
    return students


def load_target_emails(path: str) -> List[str]:
    """Read recipient emails from TargetEmail.txt.

    Supported examples:

        advisor1@example.com
        advisor2@example.com  # Prof. Wang
        advisor3@example.com  // TA
        advisor4@example.com 备注：课程老师
        advisor5@example.com, advisor6@example.com  # multiple recipients

    Anything after #, //, 备注:, 注释:, or comment: is treated as annotation.
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Target email file not found: {path}")

    out: List[str] = []
    seen = set()

    with p.open("r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()

            if not line:
                continue

            # Full-line comments
            if line.startswith("#") or line.startswith("//"):
                continue

            # Remove inline annotations
            line = _strip_comment(line)

            if not line:
                continue

            # Extract valid emails only.
            # This is safer than checking if "@" is in a token.
            for match in _EMAIL_RE.finditer(line):
                email = match.group(0).lower()
                if email not in seen:
                    out.append(email)
                    seen.add(email)

    return out
