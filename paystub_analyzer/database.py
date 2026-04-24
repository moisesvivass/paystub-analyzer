import sqlite3
from contextlib import contextmanager
from typing import Generator
from paystub_analyzer.config import DB_FILE
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)


@contextmanager
def get_connection() -> Generator[sqlite3.Connection, None, None]:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    """Create all tables if they don't exist."""
    with get_connection() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS paystubs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company TEXT NOT NULL,
                pay_period_start TEXT,
                pay_period_end TEXT,
                gross_pay REAL,
                net_pay REAL,
                federal_tax REAL,
                provincial_tax REAL,
                cpp REAL,
                ei REAL,
                vacation_pay REAL,
                hours_worked REAL,
                created_at TEXT DEFAULT (datetime('now')),
                UNIQUE(company, pay_period_end)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS processed_emails (
                email_id TEXT PRIMARY KEY,
                processed_at TEXT DEFAULT (datetime('now'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS run_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at TEXT NOT NULL,
                finished_at TEXT,
                status TEXT NOT NULL DEFAULT 'running',
                mode TEXT,
                emails_found INTEGER DEFAULT 0,
                emails_processed INTEGER DEFAULT 0,
                emails_failed INTEGER DEFAULT 0,
                error_message TEXT
            )
        """)
    logger.info("Database initialized")


def insert_paystub(data: dict) -> bool:
    """Insert a paystub. Returns True if inserted, False if duplicate."""
    with get_connection() as conn:
        try:
            conn.execute("""
                INSERT INTO paystubs (
                    company, pay_period_start, pay_period_end,
                    gross_pay, net_pay, federal_tax, provincial_tax,
                    cpp, ei, vacation_pay, hours_worked
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                str(data.get("company", ""))[:255],
                data.get("pay_period_start"),
                data.get("pay_period_end"),
                float(data.get("gross_pay") or 0),
                float(data.get("net_pay") or 0),
                float(data.get("federal_tax") or 0),
                float(data.get("provincial_tax") or 0),
                float(data.get("cpp") or 0),
                float(data.get("ei") or 0),
                float(data.get("vacation_pay") or 0),
                float(data.get("hours_worked") or 0) if data.get("hours_worked") is not None else None,
            ))
            logger.info(f"Saved to DB — {data.get('company')} {data.get('pay_period_end')}")
            return True
        except sqlite3.IntegrityError:
            logger.info(f"Already in DB — {data.get('company')} {data.get('pay_period_end')}")
            return False


def is_email_processed(email_id: str) -> bool:
    """Check if a Gmail message ID was already processed."""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT 1 FROM processed_emails WHERE email_id = ?", (email_id,)
        ).fetchone()
        return row is not None


def mark_email_processed(email_id: str) -> None:
    """Record a Gmail message ID as processed."""
    with get_connection() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO processed_emails (email_id) VALUES (?)", (email_id,)
        )


def get_all_processed_email_ids() -> set[str]:
    """Return all processed email IDs from DB."""
    with get_connection() as conn:
        rows = conn.execute("SELECT email_id FROM processed_emails").fetchall()
        return {row["email_id"] for row in rows}


def get_all_paystubs() -> list[sqlite3.Row]:
    """Return all paystubs ordered by pay period start."""
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM paystubs ORDER BY pay_period_start"
        ).fetchall()


def start_run(mode: str) -> int:
    """Log a new scheduler run. Returns the run ID."""
    with get_connection() as conn:
        cursor = conn.execute(
            "INSERT INTO run_history (started_at, mode, status) VALUES (datetime('now'), ?, 'running')",
            (mode,)
        )
        return cursor.lastrowid


def finish_run(run_id: int, emails_found: int, processed: int, failed: int) -> None:
    """Mark a run as completed."""
    with get_connection() as conn:
        conn.execute("""
            UPDATE run_history
            SET finished_at = datetime('now'), status = 'completed',
                emails_found = ?, emails_processed = ?, emails_failed = ?
            WHERE id = ?
        """, (emails_found, processed, failed, run_id))


def fail_run(run_id: int, error: str) -> None:
    """Mark a run as failed with an error message."""
    with get_connection() as conn:
        conn.execute("""
            UPDATE run_history
            SET finished_at = datetime('now'), status = 'failed', error_message = ?
            WHERE id = ?
        """, (error[:2000], run_id))
