import base64
import time
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from paystub_analyzer.config import CREDENTIALS_FILE, SCOPES, EMAIL_QUERY
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)


def authenticate_gmail() -> Resource:
    """Authenticate and return an authorized Gmail API service."""
    import os
    creds: Credentials | None = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)


def get_paystub_emails(service: Resource, max_results: int = 500) -> list[dict]:
    """Fetch all paystub emails matching EMAIL_QUERY with full pagination."""
    if not EMAIL_QUERY:
        raise ValueError("EMAIL_QUERY is not set in .env")

    logger.info(f"Searching Gmail for paystub emails (query: {EMAIL_QUERY!r})...")
    all_messages: list[dict] = []
    next_page_token: str | None = None

    while True:
        try:
            params: dict = {"userId": "me", "q": EMAIL_QUERY, "maxResults": 500}
            if next_page_token:
                params["pageToken"] = next_page_token

            results = service.users().messages().list(**params).execute()
            batch = results.get("messages", [])
            all_messages.extend(batch)
            next_page_token = results.get("nextPageToken")

            if not next_page_token or len(all_messages) >= max_results:
                break

        except HttpError as e:
            logger.error(f"Gmail API error while listing messages: {e}")
            raise

    logger.info(f"Found {len(all_messages)} emails")
    return all_messages[:max_results]


def download_pdf(service: Resource, message_id: str) -> bytes | None:
    """
    Download the first PDF attachment from a Gmail message.

    Handles three structures Gmail may return:
      1. Multipart message  → payload.parts[] tree (recursive)
      2. Single-part message → payload.body.data directly
      3. Attachment too large to inline → payload.body.attachmentId
    """
    try:
        message = service.users().messages().get(
            userId="me", id=message_id, format="full"
        ).execute()
    except HttpError as e:
        logger.error(f"[{message_id}] Gmail API error fetching message: {e}")
        raise

    payload = message["payload"]
    logger.debug(f"[{message_id}] mimeType={payload.get('mimeType')} "
                 f"has_parts={'parts' in payload}")

    def decode_body(body: dict, label: str = "") -> bytes | None:
        data = body.get("data")
        att_id = body.get("attachmentId")
        if data:
            return base64.urlsafe_b64decode(data)
        if att_id:
            logger.debug(f"  Fetching attachment id={att_id} ({label})")
            try:
                att = service.users().messages().attachments().get(
                    userId="me", messageId=message_id, id=att_id
                ).execute()
                return base64.urlsafe_b64decode(att["data"])
            except HttpError as e:
                logger.error(f"  Failed to fetch attachment {att_id}: {e}")
                return None
        return None

    def is_pdf_part(part: dict) -> bool:
        mime = part.get("mimeType", "")
        name = part.get("filename", "")
        return mime == "application/pdf" or name.lower().endswith(".pdf")

    def find_pdf_in_parts(parts: list[dict]) -> bytes | None:
        for part in parts:
            if is_pdf_part(part):
                result = decode_body(part["body"], label=part.get("filename", "pdf"))
                if result:
                    return result
            if "parts" in part:
                result = find_pdf_in_parts(part["parts"])
                if result:
                    return result
        return None

    if "parts" in payload:
        result = find_pdf_in_parts(payload["parts"])
        if result:
            logger.info(f"[{message_id}] PDF found in multipart tree ({len(result):,} bytes)")
            return result

    if is_pdf_part(payload):
        result = decode_body(payload.get("body", {}), label="payload body")
        if result:
            logger.info(f"[{message_id}] PDF found in payload body ({len(result):,} bytes)")
            return result

    logger.warning(f"[{message_id}] No PDF attachment found — skipping.")
    return None


def rate_limited_sleep(index: int, delay: float = 1.0) -> None:
    """Sleep between API calls to avoid hitting rate limits."""
    if index > 0:
        time.sleep(delay)
