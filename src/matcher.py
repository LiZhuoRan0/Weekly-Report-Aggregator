"""Match PDF files to students using filename, PDF content, and (optionally) sender email.

Match flow per PDF:
  1. Score every student against this PDF using:
     - filename contains pinyin variant       (strong signal)
     - filename contains Chinese name         (strong)
     - PDF content contains Chinese name      (strong)
     - PDF content contains pinyin variant    (medium)
     - sender email matches student's email   (medium, only for email PDFs)
  2. Pick the student with the highest score (above a threshold).
  3. If two students tie, leave unmatched and log it.

Each student then gets *one* PDF chosen from candidates:
  - Local FilePath beats email
  - Among multiple in same source: take the latest mtime
"""
from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .config import Student
from .pdf_utils import extract_text_first_pages, is_valid_pdf
from .pinyin_utils import get_pinyin_variants

logger = logging.getLogger("wra")


@dataclass
class PdfCandidate:
    """A candidate PDF that may belong to some student."""
    path: str                       # filesystem path (local FilePath or saved attachment)
    source: str                     # 'local' or 'email'
    mtime: float                    # seconds since epoch — for picking 'latest'
    sender_email: Optional[str] = None  # only set when source == 'email'
    original_name: str = ""         # original filename (for matching/reporting)


@dataclass
class MatchResult:
    """Result of matching one PDF."""
    pdf: PdfCandidate
    matched_student: Optional[Student]
    score: int
    reasons: List[str] = field(default_factory=list)


# ---- Scoring weights ----
W_FILENAME_PINYIN  = 50
W_FILENAME_CHINESE = 60
W_CONTENT_CHINESE  = 40
W_CONTENT_PINYIN   = 25
W_SENDER_EMAIL     = 20  # auxiliary

MIN_SCORE_TO_MATCH = 40   # below this, leave unmatched


def _score_student_for_pdf(
    student: Student,
    variants: List[str],
    filename_norm: str,
    content_norm: str,
    sender_email: Optional[str],
) -> Tuple[int, List[str]]:
    """Score how well this student matches this PDF. Returns (score, reasons)."""
    score = 0
    reasons: List[str] = []

    # 1. Chinese name in filename / content
    cn = student.chinese_name
    if cn and cn in filename_norm:
        score += W_FILENAME_CHINESE
        reasons.append(f"filename contains Chinese name '{cn}'")
    if cn and cn in content_norm:
        score += W_CONTENT_CHINESE
        reasons.append(f"PDF content contains Chinese name '{cn}'")

    # 2. Pinyin variant in filename / content
    matched_filename_variant = None
    matched_content_variant = None
    for v in variants:
        if len(v) < 4:  # avoid spurious short matches
            continue
        if matched_filename_variant is None and v in filename_norm:
            matched_filename_variant = v
        if matched_content_variant is None and v in content_norm:
            matched_content_variant = v
        if matched_filename_variant and matched_content_variant:
            break
    if matched_filename_variant:
        score += W_FILENAME_PINYIN
        reasons.append(f"filename contains pinyin '{matched_filename_variant}'")
    if matched_content_variant:
        score += W_CONTENT_PINYIN
        reasons.append(f"PDF content contains pinyin '{matched_content_variant}'")

    # 3. Sender email (auxiliary)
    if sender_email and student.emails:
        s = sender_email.lower()
        if s in student.emails:
            score += W_SENDER_EMAIL
            reasons.append(f"sender '{sender_email}' is in student's emails")

    return score, reasons


def match_pdfs(
    candidates: List[PdfCandidate],
    students: List[Student],
) -> Tuple[List[MatchResult], Dict[str, str]]:
    """Match every candidate PDF to at most one student.

    Returns:
      - list of MatchResult (one per input candidate, in input order)
      - dict mapping student.chinese_name -> chosen PDF path (one per submitting student)

    Selection rules for the per-student dict:
      - If both 'local' and 'email' PDFs exist for a student, prefer 'local'.
      - Within the same source, prefer the latest mtime.
    """
    # Pre-compute pinyin variants for each student (case-insensitive lowercase).
    student_variants: Dict[str, List[str]] = {
        s.chinese_name: get_pinyin_variants(s.chinese_name) for s in students
    }

    results: List[MatchResult] = []

    for cand in candidates:
        if not is_valid_pdf(cand.path):
            results.append(MatchResult(pdf=cand, matched_student=None, score=0,
                                        reasons=["unreadable PDF"]))
            continue

        filename_norm = cand.original_name.lower()
        # Also include the actual file basename in case it differs
        filename_norm += " " + Path(cand.path).name.lower()

        content = extract_text_first_pages(cand.path, max_pages=2)
        # Lowercase + collapse whitespace for matching
        content_norm = content.lower().replace("\n", " ").replace("\r", " ")
        # Keep Chinese chars as-is in content_norm (lowercase doesn't affect them).

        # Score each student
        best_score = -1
        best_students: List[Student] = []
        best_reasons: List[str] = []
        for s in students:
            score, reasons = _score_student_for_pdf(
                s,
                student_variants[s.chinese_name],
                filename_norm,
                content_norm,
                cand.sender_email,
            )
            if score > best_score:
                best_score = score
                best_students = [s]
                best_reasons = reasons
            elif score == best_score and score > 0:
                best_students.append(s)

        if best_score < MIN_SCORE_TO_MATCH or not best_students:
            results.append(MatchResult(
                pdf=cand, matched_student=None, score=best_score,
                reasons=["no student exceeded match threshold"],
            ))
            continue

        if len(best_students) > 1:
            tied = ", ".join(s.chinese_name for s in best_students)
            results.append(MatchResult(
                pdf=cand, matched_student=None, score=best_score,
                reasons=[f"tied between: {tied}; leaving unmatched"],
            ))
            logger.warning(f"PDF '{cand.path}' tied between students: {tied}")
            continue

        results.append(MatchResult(
            pdf=cand, matched_student=best_students[0], score=best_score,
            reasons=best_reasons,
        ))

    # Now choose one PDF per student per source-priority + latest mtime.
    per_student_local: Dict[str, MatchResult] = {}
    per_student_email: Dict[str, MatchResult] = {}
    for r in results:
        if r.matched_student is None:
            continue
        bucket = per_student_local if r.pdf.source == "local" else per_student_email
        name = r.matched_student.chinese_name
        if name not in bucket or r.pdf.mtime > bucket[name].pdf.mtime:
            bucket[name] = r

    chosen: Dict[str, str] = {}
    for s in students:
        name = s.chinese_name
        if name in per_student_local:
            chosen[name] = per_student_local[name].pdf.path
        elif name in per_student_email:
            chosen[name] = per_student_email[name].pdf.path

    return results, chosen


def write_match_report(
    results: List[MatchResult],
    chosen: Dict[str, str],
    students: List[Student],
    report_path: str,
) -> None:
    """Write a human-readable matching report."""
    Path(report_path).parent.mkdir(parents=True, exist_ok=True)
    lines: List[str] = []
    lines.append("=" * 70)
    lines.append("Weekly Report Matching Report")
    lines.append("=" * 70)
    lines.append("")

    lines.append(f"Total candidate PDFs scanned: {len(results)}")
    matched = [r for r in results if r.matched_student is not None]
    unmatched = [r for r in results if r.matched_student is None]
    lines.append(f"  Successfully matched:        {len(matched)}")
    lines.append(f"  Unmatched / skipped:         {len(unmatched)}")
    lines.append(f"  Students with chosen PDF:    {len(chosen)} / {len(students)}")
    lines.append("")

    lines.append("-" * 70)
    lines.append("Per-PDF match details")
    lines.append("-" * 70)
    for r in results:
        who = r.matched_student.chinese_name if r.matched_student else "<UNMATCHED>"
        lines.append(f"\n[{r.pdf.source}] {r.pdf.path}")
        lines.append(f"  original_name = {r.pdf.original_name}")
        if r.pdf.sender_email:
            lines.append(f"  sender        = {r.pdf.sender_email}")
        lines.append(f"  matched_to    = {who}  (score={r.score})")
        for reason in r.reasons:
            lines.append(f"    - {reason}")

    lines.append("")
    lines.append("-" * 70)
    lines.append("Per-student final selection")
    lines.append("-" * 70)
    for s in students:
        if s.chinese_name in chosen:
            src = "local" if any(
                r.matched_student and r.matched_student.chinese_name == s.chinese_name
                and r.pdf.path == chosen[s.chinese_name] and r.pdf.source == "local"
                for r in results
            ) else "email"
            lines.append(f"  ✓ {s.chinese_name:<10} <- [{src}] {chosen[s.chinese_name]}")
        else:
            lines.append(f"  ✗ {s.chinese_name:<10} (no PDF found)")

    Path(report_path).write_text("\n".join(lines), encoding="utf-8")
    logger.info(f"Match report written to {report_path}")


def scan_local_pdfs(file_path: str) -> List[PdfCandidate]:
    """Recursively find all .pdf files under file_path and return candidates."""
    base = Path(file_path)
    if not base.exists():
        logger.error(f"FilePath does not exist: {file_path}")
        return []
    if not base.is_dir():
        logger.error(f"FilePath is not a directory: {file_path}")
        return []
    out: List[PdfCandidate] = []
    for p in base.rglob("*.pdf"):
        if not p.is_file():
            continue
        try:
            mtime = p.stat().st_mtime
        except OSError:
            mtime = 0.0
        out.append(PdfCandidate(
            path=str(p.resolve()),
            source="local",
            mtime=mtime,
            original_name=p.name,
        ))
    # Also pick up .PDF (case insensitive on case-sensitive filesystems)
    for p in base.rglob("*.PDF"):
        if not p.is_file():
            continue
        try:
            mtime = p.stat().st_mtime
        except OSError:
            mtime = 0.0
        out.append(PdfCandidate(
            path=str(p.resolve()),
            source="local",
            mtime=mtime,
            original_name=p.name,
        ))
    # Deduplicate by absolute path
    seen = set()
    deduped = []
    for c in out:
        if c.path in seen:
            continue
        seen.add(c.path)
        deduped.append(c)
    logger.info(f"Found {len(deduped)} local PDF(s) under {file_path}")
    return deduped
