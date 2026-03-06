import os
from dotenv import load_dotenv

load_dotenv()

CREDENTIALS_FILE = os.getenv("CREDENTIALS_FILE")
PDF_PASSWORD = os.getenv("PDF_PASSWORD")
OUTPUT_EXCEL = os.getenv("OUTPUT_EXCEL")
EMAIL_QUERY = os.getenv("EMAIL_QUERY")
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']