import io
import PyPDF2
from PyPDF2.errors import PdfReadError
from paystub_analyzer.config import PDF_PASSWORD
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """Extract plain text from a PDF, decrypting with PDF_PASSWORD if needed."""
    if pdf_bytes is None:
        raise ValueError("No PDF bytes received — attachment may be missing")
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        if reader.is_encrypted:
            result = reader.decrypt(PDF_PASSWORD)
            if result == 0:
                raise ValueError("PDF decryption failed — wrong password")
        text = "".join(page.extract_text() or "" for page in reader.pages)
        if not text.strip():
            raise ValueError("PDF parsed successfully but no text was extracted")
        logger.info(f"PDF text extracted — {len(text)} characters")
        return text
    except PdfReadError as e:
        logger.error(f"PDF is corrupted or not a valid PDF: {e}")
        raise ValueError(f"Invalid PDF: {e}") from e
