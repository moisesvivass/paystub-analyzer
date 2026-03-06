import io
import PyPDF2
from paystub_analyzer.config import PDF_PASSWORD
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    if pdf_bytes is None:
        raise ValueError("No PDF bytes received — attachment may be missing")
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        if reader.is_encrypted:
            reader.decrypt(PDF_PASSWORD)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        logger.info(f"PDF text extracted — {len(text)} characters")
        return text
    except Exception as e:
        logger.error(f"Failed to extract PDF text: {e}")
        raise