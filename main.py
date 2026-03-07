import argparse
from paystub_analyzer.gmail_client import authenticate_gmail, get_paystub_emails, download_pdf
from paystub_analyzer.pdf_processor import extract_text_from_pdf
from paystub_analyzer.claude_extractor import extract_data_with_claude
from paystub_analyzer.excel_report import create_excel
from paystub_analyzer.tracker import load_processed_ids, save_processed_ids, filter_new_messages
from paystub_analyzer.database import init_db, insert_paystub
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

def main():
    # ── CLI arguments ──────────────────────────────────────────────────────────
    parser = argparse.ArgumentParser(description="Paystub Analyzer")
    parser.add_argument(
        "--mode",
        choices=["full", "update"],
        default="full",
        help="full = all emails | update = new emails only"
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=5,
        help="Max number of emails to process (default: 5)"
    )
    args = parser.parse_args()

    logger.info(f"🚀 Starting Paystub Analyzer — mode={args.mode} limit={args.limit}")

    # ── Initialize database ────────────────────────────────────────────────────
    init_db()

    # ── Gmail ──────────────────────────────────────────────────────────────────
    logger.info("Connecting to Gmail...")
    service = authenticate_gmail()
    messages = get_paystub_emails(service)

    # ── Tracker ────────────────────────────────────────────────────────────────
    processed_ids = load_processed_ids()

    if args.mode == "update":
        messages = filter_new_messages(messages, processed_ids)

    messages = messages[:args.limit]

    if not messages:
        logger.info("✅ No new emails to process!")
        return

    # ── Process ────────────────────────────────────────────────────────────────
    all_data = []
    for i, msg in enumerate(messages):
        logger.info(f"Processing {i+1}/{len(messages)}...")
        pdf_bytes = download_pdf(service, msg['id'])
        if pdf_bytes:
            try:
                text = extract_text_from_pdf(pdf_bytes)
                data = extract_data_with_claude(text)
                all_data.append(data)
                insert_paystub(data)                    # ← save to SQLite
                processed_ids.add(msg['id'])
            except Exception as e:
                logger.error(f"❌ Skipping paystub {i+1}: {e}")
        else:
            logger.warning(f"❌ No PDF found in email {i+1}")

    # ── Save & Report ──────────────────────────────────────────────────────────
    save_processed_ids(processed_ids)
    logger.info("Generating Excel report...")
    create_excel(all_data)
    logger.info("✅ Done!")

if __name__ == "__main__":
    main()