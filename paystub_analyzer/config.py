import os
from dotenv import load_dotenv

load_dotenv()

CREDENTIALS_FILE: str = os.getenv("CREDENTIALS_FILE", "client_secret.json")
PDF_PASSWORD: str = os.getenv("PDF_PASSWORD", "")
OUTPUT_EXCEL: str = os.getenv("OUTPUT_EXCEL", "paystubs.xlsx")
EMAIL_QUERY: str = os.getenv("EMAIL_QUERY", "")
DB_FILE: str = os.getenv("DB_FILE", "paystubs.db")
LOG_FILE: str = os.getenv("LOG_FILE", "process.log")

# Scheduler config
SCHEDULE_HOUR: int = int(os.getenv("SCHEDULE_HOUR", "9"))
SCHEDULE_MINUTE: int = int(os.getenv("SCHEDULE_MINUTE", "0"))
SCHEDULE_TIMEZONE: str = os.getenv("SCHEDULE_TIMEZONE", "America/Toronto")
SCHEDULE_MAX_EMAILS: int = int(os.getenv("SCHEDULE_MAX_EMAILS", "50"))

SCOPES: list[str] = ["https://www.googleapis.com/auth/gmail.readonly"]


def validate_config() -> list[str]:
    """Return a list of missing required config keys."""
    import os as _os
    missing = []
    if not _os.path.exists(CREDENTIALS_FILE):
        missing.append(f"CREDENTIALS_FILE (file not found: {CREDENTIALS_FILE})")
    if not PDF_PASSWORD:
        missing.append("PDF_PASSWORD")
    if not EMAIL_QUERY:
        missing.append("EMAIL_QUERY")
    return missing
