import sqlite3
import os
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

DB_FILE = "paystubs.db"

def get_connection():
    return sqlite3.connect(DB_FILE)

def init_db():
    """Create the paystubs table if it doesn't exist."""
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
        logger.info("✅ Database initialized")

def insert_paystub(data: dict):
    """Insert a paystub record — skip if already exists."""
    with get_connection() as conn:
        try:
            conn.execute("""
                INSERT INTO paystubs (
                    company, pay_period_start, pay_period_end,
                    gross_pay, net_pay, federal_tax, provincial_tax,
                    cpp, ei, vacation_pay, hours_worked
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                data.get("company"),
                data.get("pay_period_start"),
                data.get("pay_period_end"),
                data.get("gross_pay"),
                data.get("net_pay"),
                data.get("federal_tax"),
                data.get("provincial_tax"),
                data.get("cpp"),
                data.get("ei"),
                data.get("vacation_pay"),
                data.get("hours_worked"),
            ))
            logger.info(f"💾 Saved to DB — {data.get('company')} {data.get('pay_period_end')}")
        except sqlite3.IntegrityError:
            logger.info(f"⏭️ Already in DB — {data.get('company')} {data.get('pay_period_end')}")

def get_all_paystubs() -> list:
    """Return all paystubs from the database."""
    with get_connection() as conn:
        cursor = conn.execute("SELECT * FROM paystubs ORDER BY pay_period_start")
        return cursor.fetchall()