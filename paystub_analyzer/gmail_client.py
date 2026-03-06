import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from paystub_analyzer.config import CREDENTIALS_FILE, SCOPES, EMAIL_QUERY
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

def authenticate_gmail():
    creds = None
    if __import__('os').path.exists('token.json'):
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
    logger.info("Searching Gmail for paystub emails...")
    results = service.users().messages().list(
        userId='me',
        q=EMAIL_QUERY
    ).execute()
    messages = results.get('messages', [])
    logger.info(f"Found {len(messages)} emails")
    return messages

def download_pdf(service, message_id):
    message = service.users().messages().get(
        userId='me', id=message_id, format='full'
    ).execute()

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