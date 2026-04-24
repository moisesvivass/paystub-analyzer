"""
Email tracking backed by SQLite (processed_emails table).
Keeps a thin JSON fallback for migration of legacy processed_ids.json.
"""
import json
import os
from paystub_analyzer.database import get_all_processed_email_ids, mark_email_processed
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

_LEGACY_FILE = "processed_ids.json"
_migration_done = False


def _migrate_legacy_tracker() -> None:
    """One-time import of processed_ids.json into the DB."""
    global _migration_done
    if _migration_done or not os.path.exists(_LEGACY_FILE):
        _migration_done = True
        return
    with open(_LEGACY_FILE) as f:
        ids: list[str] = json.load(f)
    for email_id in ids:
        mark_email_processed(email_id)
    os.rename(_LEGACY_FILE, _LEGACY_FILE + ".migrated")
    _migration_done = True
    logger.info(f"Migrated {len(ids)} IDs from legacy tracker to DB")


def load_processed_ids() -> set[str]:
    """Return all processed email IDs from the DB."""
    _migrate_legacy_tracker()
    return get_all_processed_email_ids()


def filter_new_messages(messages: list[dict], processed_ids: set[str]) -> list[dict]:
    """Return only messages whose ID is not yet processed."""
    new = [m for m in messages if m["id"] not in processed_ids]
    logger.info(f"{len(new)} new emails out of {len(messages)} total")
    return new


def mark_processed(email_id: str) -> None:
    """Persist a processed email ID to the DB."""
    mark_email_processed(email_id)
