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
    """
    Download the first PDF attachment from a Gmail message.

    Handles three structures Gmail may return:
      1. Multipart message  → payload.parts[] tree (recursive)
      2. Single-part message → payload.body.data directly
      3. Attachment too large to inline → payload.body.attachmentId
    """
    message = service.users().messages().get(
        userId='me', id=message_id, format='full'
    ).execute()

    payload = message['payload']
    logger.debug(f"[{message_id}] mimeType={payload.get('mimeType')} "
                 f"has_parts={'parts' in payload}")

    # ── helpers ────────────────────────────────────────────────────────────────

    def decode_body(body, label=""):
        """Decode inline base64 data or fetch a remote attachment."""
        data = body.get('data')
        att_id = body.get('attachmentId')
        if data:
            logger.debug(f"  Decoding inline data ({label})")
            return base64.urlsafe_b64decode(data)
        if att_id:
            logger.debug(f"  Fetching attachment id={att_id} ({label})")
            att = service.users().messages().attachments().get(
                userId='me', messageId=message_id, id=att_id
            ).execute()
            return base64.urlsafe_b64decode(att['data'])
        return None

    def is_pdf_part(part):
        mime = part.get('mimeType', '')
        name = part.get('filename', '')
        return mime == 'application/pdf' or name.lower().endswith('.pdf')

    def find_pdf_in_parts(parts):
        """Recursively walk a parts tree looking for a PDF."""
        for part in parts:
            logger.debug(f"  Inspecting part: mimeType={part.get('mimeType')} "
                         f"filename={part.get('filename', '')!r}")
            if is_pdf_part(part):
                result = decode_body(part['body'], label=part.get('filename', 'pdf'))
                if result:
                    return result
            # Recurse into nested multipart/* parts
            if 'parts' in part:
                result = find_pdf_in_parts(part['parts'])
                if result:
                    return result
        return None

    # ── case 1: multipart message ──────────────────────────────────────────────
    if 'parts' in payload:
        result = find_pdf_in_parts(payload['parts'])
        if result:
            logger.info(f"[{message_id}] PDF found in multipart tree ({len(result):,} bytes)")
            return result

    # ── case 2 & 3: single-part — check payload body directly ─────────────────
    if is_pdf_part(payload):
        result = decode_body(payload.get('body', {}), label="payload body")
        if result:
            logger.info(f"[{message_id}] PDF found in payload body ({len(result):,} bytes)")
            return result

    # ── nothing found ──────────────────────────────────────────────────────────
    logger.warning(f"[{message_id}] No PDF attachment found — skipping.")
    return None