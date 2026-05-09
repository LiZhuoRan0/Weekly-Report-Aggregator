"""Scheduler — wait until TargetTime (Beijing time) before executing."""
import logging
import time
from datetime import datetime, timedelta, timezone

# Python 3.9+ has zoneinfo. Fall back to fixed UTC+8 if not available.
try:
    from zoneinfo import ZoneInfo
    BJT = ZoneInfo("Asia/Shanghai")
    _USES_ZONEINFO = True
except Exception:  # pragma: no cover
    BJT = timezone(timedelta(hours=8))
    _USES_ZONEINFO = False

logger = logging.getLogger("wra")


def parse_target_time(target_time: str) -> datetime:
    """Parse a 'YYYY_MM_DD_HH_MM' string as Beijing-time and return a tz-aware datetime."""
    try:
        dt_naive = datetime.strptime(target_time, "%Y_%m_%d_%H_%M")
    except ValueError as e:
        raise ValueError(
            f"Invalid TargetTime '{target_time}'. Expected format YYYY_MM_DD_HH_MM."
        ) from e
    return dt_naive.replace(tzinfo=BJT)


def now_bjt() -> datetime:
    """Return current time in Beijing time."""
    return datetime.now(tz=BJT)


def sleep_until(target_dt: datetime) -> None:
    """Block until `target_dt` (must be tz-aware). Logs progress periodically."""
    if target_dt.tzinfo is None:
        target_dt = target_dt.replace(tzinfo=BJT)

    while True:
        now = now_bjt()
        delta = (target_dt - now).total_seconds()
        if delta <= 0:
            logger.info(f"Reached TargetTime: {target_dt.isoformat()}")
            return
        # Choose a sleep step. For long waits, sleep up to 5 minutes at a time
        # so the process remains responsive.
        if delta > 300:
            step = 300
        elif delta > 30:
            step = 30
        else:
            step = max(1, int(delta))

        logger.info(
            f"Waiting until {target_dt.isoformat()} (Beijing). "
            f"Now={now.isoformat()}, remaining≈{int(delta)}s. Sleeping {step}s."
        )
        time.sleep(step)


def compute_email_window(target_dt: datetime, lookback_days: int) -> datetime:
    """Compute the 'since' datetime for IMAP search:

    Per the spec: 'most recent 3 days' is measured from TargetTime, looking BACK.
    So since = target - lookback_days.
    """
    return target_dt - timedelta(days=lookback_days)
