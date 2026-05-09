#!/usr/bin/env python3
"""Weekly Report Aggregator — main entry point.

Workflow:
  1. Parse args, load config / students.txt / TargetEmail.txt
  2. Sleep until TargetTime (Beijing)
  3. Scan local FilePath for PDFs
  4. Fetch recent (lookback_days before TargetTime) emails for PDF attachments
  5. Match every PDF to a student
  6. Pick one PDF per student (local > email; latest mtime within source)
  7. Merge into a single PDF with bookmarks
  8. If file size > max_attachment_size_mb, split into multiple parts
  9. Send via SMTP to TargetEmail recipients
  10. Write match report and logs

Usage:
  python main.py
  python main.py --dry-run                # match + merge but never send
  python main.py --no-wait                # ignore TargetTime, run immediately
  python main.py --config myconfig.json --students students.txt --target-email TargetEmail.txt
"""
from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import List

from src.config import load_config, load_students, load_target_emails
from src.email_fetcher import fetch_email_pdf_attachments
from src.email_sender import build_email_subject_and_body, send_emails_with_attachments
from src.logger import setup_logger
from src.matcher import (
    match_pdfs,
    scan_local_pdfs,
    write_match_report,
)
from src.pdf_utils import split_merged_pdf_by_size
from src.scheduler import (
    compute_email_window,
    now_bjt,
    parse_target_time,
    sleep_until,
)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Weekly Report Aggregator — merge weekly reports and email them.",
    )
    p.add_argument("--config", default="config.json", help="Path to config.json")
    p.add_argument("--students", default="students.txt", help="Path to students.txt")
    p.add_argument(
        "--target-email", default="TargetEmail.txt",
        help="Path to TargetEmail.txt (recipients, one per line)",
    )
    p.add_argument(
        "--dry-run", action="store_true",
        help="Run the full pipeline but do not send any email. "
             "Generated PDFs and match report are written for review.",
    )
    p.add_argument(
        "--no-wait", action="store_true",
        help="Do not wait until TargetTime; execute immediately. "
             "(TargetTime is still used to name the output file and email.)",
    )
    p.add_argument(
        "--output-dir", default="output",
        help="Directory for merged PDFs and match reports.",
    )
    p.add_argument(
        "--keep-temp", action="store_true",
        help="Keep the temp directory of fetched email attachments (for debugging).",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()

    logger = setup_logger("logs")
    logger.info("=" * 60)
    logger.info("Weekly Report Aggregator starting up")
    logger.info("=" * 60)

    # ----- Load inputs -----
    try:
        cfg = load_config(args.config)
    except FileNotFoundError as e:
        logger.error(str(e)); return 1
    except (KeyError, ValueError) as e:
        logger.error(f"Invalid config.json: {e}"); return 1

    try:
        students = load_students(args.students)
    except FileNotFoundError as e:
        logger.error(str(e)); return 1
    if not students:
        logger.error("No students found in students.txt"); return 1
    logger.info(f"Loaded {len(students)} students from {args.students}")

    try:
        recipients = load_target_emails(args.target_email)
    except FileNotFoundError as e:
        logger.error(str(e)); return 1
    if not recipients:
        logger.error("No recipients found in TargetEmail.txt"); return 1
    logger.info(f"Loaded {len(recipients)} recipient(s) from {args.target_email}")

    try:
        target_dt = parse_target_time(cfg.target_time)
    except ValueError as e:
        logger.error(str(e)); return 1
    logger.info(f"TargetTime parsed as: {target_dt.isoformat()} (Beijing)")

    # ----- Wait until TargetTime (mode A) -----
    if args.no_wait:
        logger.info("--no-wait set: skipping wait-until-TargetTime.")
    else:
        if target_dt <= now_bjt():
            logger.warning(
                f"TargetTime {target_dt.isoformat()} is already in the past. "
                f"Executing immediately."
            )
        else:
            sleep_until(target_dt)

    # ----- Scan local PDFs -----
    local_candidates = scan_local_pdfs(cfg.file_path)

    # ----- Fetch email attachments -----
    since_dt = compute_email_window(target_dt, cfg.lookback_days)
    logger.info(
        f"Email search window: {since_dt.isoformat()} → {target_dt.isoformat()} (Beijing)"
    )
    student_emails_set = set()
    for s in students:
        for e in s.emails:
            student_emails_set.add(e.lower())

    tmp_dir = tempfile.mkdtemp(prefix="wra_email_")
    logger.info(f"Email attachments will be saved to: {tmp_dir}")
    email_candidates = fetch_email_pdf_attachments(
        host=cfg.imap_host,
        port=cfg.imap_port,
        user=cfg.imap_user,
        password=cfg.imap_password,
        save_dir=tmp_dir,
        since_dt=since_dt,
        until_dt=target_dt,
        student_emails=student_emails_set,
    )

    all_candidates = local_candidates + email_candidates
    logger.info(
        f"Total PDF candidates: {len(all_candidates)} "
        f"({len(local_candidates)} local + {len(email_candidates)} email)"
    )

    # ----- Match -----
    results, chosen = match_pdfs(all_candidates, students)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    report_path = output_dir / f"match_report_{cfg.target_time}.txt"
    write_match_report(results, chosen, students, str(report_path))

    # ----- Build merge order (preserve students.txt order; skip non-submitters) -----
    merge_items: List = []
    submitted_names: List[str] = []
    not_submitted_names: List[str] = []
    for s in students:
        if s.chinese_name in chosen:
            merge_items.append((s.chinese_name, chosen[s.chinese_name]))
            submitted_names.append(s.chinese_name)
        else:
            not_submitted_names.append(s.chinese_name)

    logger.info(f"Submitted: {len(submitted_names)} | Not submitted: {len(not_submitted_names)}")
    logger.info(f"Submitted names: {submitted_names}")
    logger.info(f"Not submitted names: {not_submitted_names}")

    if not merge_items:
        logger.error("No students submitted — nothing to merge. Exiting.")
        if not args.keep_temp:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return 2

    # ----- Merge (with chunking if oversized) -----
    base_name = f"WeeklyReport_{cfg.target_time}"
    max_bytes = cfg.max_attachment_size_mb * 1024 * 1024
    merged_paths = split_merged_pdf_by_size(
        items=merge_items,
        output_dir=str(output_dir),
        base_name=base_name,
        max_size_bytes=max_bytes,
    )
    if not merged_paths:
        logger.error("Failed to produce any merged PDF.")
        if not args.keep_temp:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return 3

    # ----- Compose and send email -----
    subject, body = build_email_subject_and_body(
        target_time=cfg.target_time,
        submitted_names=submitted_names,
        not_submitted_names=not_submitted_names,
    )
    logger.info(f"Email subject: {subject}")
    logger.info(f"Email body:\n{body}")

    send_ok = send_emails_with_attachments(
        smtp_host=cfg.smtp_host,
        smtp_port=cfg.smtp_port,
        smtp_user=cfg.smtp_user,
        smtp_password=cfg.smtp_password,
        sender_display_name=cfg.sender_display_name,
        recipients=recipients,
        subject_base=subject,
        body_text=body,
        attachments=merged_paths,
        dry_run=args.dry_run,
    )

    # ----- Cleanup -----
    if args.keep_temp:
        logger.info(f"Keeping temp dir: {tmp_dir}")
    else:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    if args.dry_run:
        logger.info("Dry run complete. No email was actually sent.")
        return 0

    if send_ok:
        logger.info("All emails sent successfully.")
        return 0
    else:
        logger.error("Some emails failed to send. See log above.")
        return 4


if __name__ == "__main__":
    sys.exit(main())
