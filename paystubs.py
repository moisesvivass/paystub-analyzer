import os
import base64
import json
import logging
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import PyPDF2
import openpyxl
from anthropic import Anthropic

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s — %(levelname)s — %(message)s',
    handlers=[
        logging.FileHandler('process.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

anthropic_client = Anthropic()

CREDENTIALS_FILE = os.getenv("CREDENTIALS_FILE")
PDF_PASSWORD = os.getenv("PDF_PASSWORD")
OUTPUT_EXCEL = os.getenv("OUTPUT_EXCEL")

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def get_paystub_emails(service):
    results = service.users().messages().list(
        userId='me',
        q='subject:"Ares Infrastructure Incorporated Pay Stub" OR subject:"Bar-Quip Construction Limited (SB) Pay Stub"'
    ).execute()
    return results.get('messages', [])

def download_pdf(service, message_id):
    message = service.users().messages().get(userId='me', id=message_id, format='full').execute()

    def find_pdf_in_parts(parts):
        for part in parts:
            if part.get('mimeType') == 'application/pdf' or part.get('filename', '').endswith('.pdf'):
                data = part['body'].get('data')
                att_id = part['body'].get('attachmentId')
                if data:
                    return base64.urlsafe_b64decode(data)
                elif att_id:
                    att = service.users().messages().attachments().get(
                        userId='me', messageId=message_id, id=att_id
                    ).execute()
                    return base64.urlsafe_b64decode(att['data'])
            if 'parts' in part:
                result = find_pdf_in_parts(part['parts'])
                if result:
                    return result
        return None

    parts = message['payload'].get('parts', [])
    return find_pdf_in_parts(parts)

def extract_text_from_pdf(pdf_bytes, password):
    import io
    reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    if reader.is_encrypted:
        reader.decrypt(password)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

def extract_data_with_claude(text):
    response = anthropic_client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": f"""Extract the following data from this paystub and return ONLY a JSON object:
            - pay_period_start (YYYY-MM-DD)
            - pay_period_end (YYYY-MM-DD)
            - gross_pay (number)
            - net_pay (number)
            - federal_tax (number)
            - provincial_tax (number)
            - cpp (number)
            - ei (number)
            - vacation_pay (number)
            - hours_worked (number)
            - company (text, extract the company name from the paystub)
            
            Paystub text:
            {text}
            
            Return ONLY valid JSON, no explanation."""
        }]
    )
    raw = response.content[0].text
    clean = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(clean)

def create_excel(data_list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Paystubs 2025"

    headers = ["Company", "Pay Period Start", "Pay Period End", "Gross Pay",
               "Net Pay", "Federal Tax", "Provincial Tax",
               "CPP", "EI", "Vacation Pay", "Hours Worked"]
    ws.append(headers)

    for data in sorted(data_list, key=lambda x: x.get('pay_period_start', '')):
        ws.append([
            data.get('company'),
            data.get('pay_period_start'),
            data.get('pay_period_end'),
            data.get('gross_pay'),
            data.get('net_pay'),
            data.get('federal_tax'),
            data.get('provincial_tax'),
            data.get('cpp'),
            data.get('ei'),
            data.get('vacation_pay'),
            data.get('hours_worked')
        ])

    wb.save(OUTPUT_EXCEL)
    logger.info(f"Excel saved at: {OUTPUT_EXCEL}")

def main():
    logger.info("Connecting to Gmail...")
    service = authenticate_gmail()

    logger.info("Searching for paystubs...")
    messages = get_paystub_emails(service)
    logger.info(f"Found {len(messages)} emails")

    all_data = []
    for i, msg in enumerate(messages):
        logger.info(f"Processing {i+1}/{len(messages)}...")
        pdf_bytes = download_pdf(service, msg['id'])
        if pdf_bytes:
            try:
                text = extract_text_from_pdf(pdf_bytes, PDF_PASSWORD)
                data = extract_data_with_claude(text)
                all_data.append(data)
                logger.info(f"✅ {data.get('pay_period_start')} — Gross: ${data.get('gross_pay')}")
            except Exception as e:
                logger.error(f"❌ Error processing paystub {i+1}: {e}")
        else:
            logger.warning(f"❌ No PDF found in email {i+1}")

    logger.info("Generating Excel...")
    create_excel(all_data)
    logger.info("Done! 🎉")

if __name__ == "__main__":
    main()