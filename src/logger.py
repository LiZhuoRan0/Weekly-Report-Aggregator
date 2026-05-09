"""Logging configuration."""
import logging
import sys
from datetime import datetime
from pathlib import Path


def setup_logger(log_dir: str = "logs", level: int = logging.INFO) -> logging.Logger:
    """Set up the root logger with file + console handlers.

    Returns the configured logger. Log file path: logs/run_{YYYYmmdd_HHMMSS}.log
    """
    log_dir_path = Path(log_dir)
    log_dir_path.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir_path / f"run_{timestamp}.log"

    logger = logging.getLogger("wra")
    logger.setLevel(level)
    # Avoid duplicate handlers if called twice
    if logger.handlers:
        return logger

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    logger.info(f"Logger initialised. Log file: {log_file}")
    return logger


def get_logger() -> logging.Logger:
    """Get the configured logger (call setup_logger first)."""
    return logging.getLogger("wra")
