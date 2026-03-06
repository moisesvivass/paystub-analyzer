from paystub_analyzer.gmail_client import authenticate_gmail, get_paystub_emails, download_pdf
from paystub_analyzer.pdf_processor import extract_text_from_pdf
from paystub_analyzer.claude_extractor import extract_data_with_claude
from paystub_analyzer.excel_report import create_excel
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

def main():
    logger.info("🚀 Starting Paystub Analyzer...")

    logger.info("Connecting to Gmail...")
    service = authenticate_gmail()

    messages = get_paystub_emails(service)

    all_data = []
    for i, msg in enumerate(messages):
        logger.info(f"Processing {i+1}/{len(messages)}...")
        pdf_bytes = download_pdf(service, msg['id'])
        if pdf_bytes:
            try:
                text = extract_text_from_pdf(pdf_bytes)
                data = extract_data_with_claude(text)
                all_data.append(data)
            except Exception as e:
                logger.error(f"❌ Skipping paystub {i+1}: {e}")
        else:
            logger.warning(f"❌ No PDF found in email {i+1}")

    logger.info("Generating Excel report...")
    create_excel(all_data)
    logger.info("✅ Done!")

if __name__ == "__main__":
    main()