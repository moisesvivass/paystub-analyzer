import pytest
from paystub_analyzer.pdf_processor import extract_text_from_pdf

def test_extract_text_none_input():
    """Should raise ValueError if no bytes are provided."""
    with pytest.raises(ValueError):
        extract_text_from_pdf(None)

def test_extract_text_invalid_bytes():
    """Should raise an exception if bytes are not a valid PDF."""
    with pytest.raises(Exception):
        extract_text_from_pdf(b"this is not a pdf")