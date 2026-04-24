import argparse
import sys
from paystub_analyzer.logger import setup_logging, get_logger
from paystub_analyzer.config import LOG_FILE, validate_config

setup_logging(LOG_FILE)
logger = get_logger(__name__)


def run_pipeline(mode: str = "update", limit: int = 50) -> dict:
    """Execute the paystub processing pipeline. Returns run stats."""
    from paystub_analyzer.gmail_client import authenticate_gmail, get_paystub_emails, download_pdf, rate_limited_sleep
    from paystub_analyzer.pdf_processor import extract_text_from_pdf
    from paystub_analyzer.claude_extractor import extract_data_with_claude
    from paystub_analyzer.excel_report import create_excel
    from paystub_analyzer.tracker import load_processed_ids, filter_new_messages, mark_processed
    from paystub_analyzer.database import init_db, insert_paystub

    init_db()

    missing = validate_config()
    if missing:
        raise ValueError(f"Missing required config keys: {', '.join(missing)}")

    logger.info(f"Starting pipeline — mode={mode} limit={limit}")

    service = authenticate_gmail()
    messages = get_paystub_emails(service)
    processed_ids = load_processed_ids()

    if mode == "update":
        messages = filter_new_messages(messages, processed_ids)

    messages = messages[:limit]

    if not messages:
        logger.info("No new emails to process")
        return {"found": 0, "processed": 0, "failed": 0}

    all_data: list[dict] = []
    processed = 0
    failed = 0

    for i, msg in enumerate(messages):
        rate_limited_sleep(i)
        logger.info(f"Processing {i + 1}/{len(messages)}...")
        try:
            pdf_bytes = download_pdf(service, msg["id"])
            if not pdf_bytes:
                logger.warning(f"No PDF in email {i + 1} — skipping")
                failed += 1
                continue
            text = extract_text_from_pdf(pdf_bytes)
            data = extract_data_with_claude(text)
            insert_paystub(data)
            mark_processed(msg["id"])
            all_data.append(data)
            processed += 1
        except ValueError as e:
            logger.error(f"Validation error on email {i + 1}: {e}")
            failed += 1
        except RuntimeError as e:
            logger.error(f"API failure on email {i + 1} after retries: {e}")
            failed += 1
        except Exception as e:
            logger.error(f"Unexpected error on email {i + 1}: {e}", exc_info=True)
            failed += 1

    if all_data:
        logger.info("Generating Excel report...")
        create_excel(all_data)

    stats = {"found": len(messages), "processed": processed, "failed": failed}
    logger.info(f"Done — {stats}")
    return stats


def main() -> None:
    parser = argparse.ArgumentParser(description="Paystub Analyzer")
    parser.add_argument(
        "--mode", choices=["full", "update"], default="update",
        help="full = all emails | update = new emails only (default)"
    )
    parser.add_argument(
        "--limit", type=int, default=50,
        help="Max emails to process per run (default: 50)"
    )
    parser.add_argument(
        "--schedule", action="store_true",
        help="Start the background scheduler (runs on cron, does not exit)"
    )
    args = parser.parse_args()

    if args.schedule:
        from paystub_analyzer.scheduler import start_scheduler
        logger.info("Starting scheduler mode — press Ctrl+C to stop")
        scheduler = start_scheduler()
        try:
            import time
            while True:
                time.sleep(60)
        except KeyboardInterrupt:
            scheduler.shutdown()
            logger.info("Scheduler stopped")
        return

    try:
        run_pipeline(mode=args.mode, limit=args.limit)
    except Exception as e:
        logger.error(f"Pipeline failed: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
