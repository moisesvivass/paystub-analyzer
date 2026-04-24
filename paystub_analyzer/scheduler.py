"""
Background scheduler for automated paystub processing.

Usage:
    python main.py --schedule
    python -m paystub_analyzer.scheduler

Config (.env):
    SCHEDULE_HOUR       = 9            # 24h format, default 9
    SCHEDULE_MINUTE     = 0            # default 0
    SCHEDULE_TIMEZONE   = America/Toronto
    SCHEDULE_MAX_EMAILS = 50
"""
import os
import sys
from pathlib import Path
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from paystub_analyzer.config import (
    SCHEDULE_DAY_OF_WEEK, SCHEDULE_HOUR, SCHEDULE_MINUTE,
    SCHEDULE_TIMEZONE, SCHEDULE_MAX_EMAILS,
)
from paystub_analyzer.database import start_run, finish_run, fail_run
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

_LOCK_FILE = ".scheduler.lock"


def _acquire_lock() -> object | None:
    """Return an open file handle if the lock is acquired, else None.

    Uses exclusive file creation (mode 'x') which is atomic on all platforms.
    """
    try:
        lock = open(_LOCK_FILE, "x")
        lock.write(str(os.getpid()))
        lock.flush()
        return lock
    except FileExistsError:
        return None


def _release_lock(lock: object) -> None:
    lock.close()
    Path(_LOCK_FILE).unlink(missing_ok=True)


def _scheduled_job() -> None:
    """Single scheduled run — idempotent, lock-protected."""
    lock = _acquire_lock()
    if lock is None:
        logger.warning("Scheduler: previous run still in progress — skipping this tick")
        return

    run_id = start_run("scheduled")
    logger.info(f"Scheduler: starting run #{run_id}")
    try:
        from main import run_pipeline
        stats = run_pipeline(mode="update", limit=SCHEDULE_MAX_EMAILS)
        finish_run(run_id, stats["found"], stats["processed"], stats["failed"])
        logger.info(f"Scheduler: run #{run_id} completed — {stats}")
    except Exception as e:
        fail_run(run_id, str(e))
        logger.error(f"Scheduler: run #{run_id} failed — {e}", exc_info=True)
    finally:
        _release_lock(lock)


def start_scheduler() -> BackgroundScheduler:
    """Create and start the APScheduler background scheduler."""
    scheduler = BackgroundScheduler(timezone=SCHEDULE_TIMEZONE)
    scheduler.add_job(
        _scheduled_job,
        trigger=CronTrigger(
            day_of_week=SCHEDULE_DAY_OF_WEEK,
            hour=SCHEDULE_HOUR,
            minute=SCHEDULE_MINUTE,
            timezone=SCHEDULE_TIMEZONE,
        ),
        id="paystub_processor",
        replace_existing=True,
        misfire_grace_time=3600,
    )
    scheduler.start()
    logger.info(
        f"Scheduler started — runs every {SCHEDULE_DAY_OF_WEEK} at "
        f"{SCHEDULE_HOUR:02d}:{SCHEDULE_MINUTE:02d} ({SCHEDULE_TIMEZONE}), "
        f"max {SCHEDULE_MAX_EMAILS} emails/run"
    )
    return scheduler


if __name__ == "__main__":
    from paystub_analyzer.logger import setup_logging
    from paystub_analyzer.config import LOG_FILE
    from paystub_analyzer.database import init_db
    import time

    setup_logging(LOG_FILE)
    init_db()

    scheduler = start_scheduler()
    try:
        while True:
            time.sleep(60)
    except KeyboardInterrupt:
        scheduler.shutdown()
        logger.info("Scheduler stopped")
        sys.exit(0)
