import json
import os
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

TRACKER_FILE = "processed_ids.json"

def load_processed_ids():
    """Lee los IDs ya procesados del archivo JSON."""
    if not os.path.exists(TRACKER_FILE):
        return set()
    with open(TRACKER_FILE, 'r') as f:
        data = json.load(f)
    return set(data)

def save_processed_ids(ids: set):
    """Guarda los IDs procesados en el archivo JSON."""
    with open(TRACKER_FILE, 'w') as f:
        json.dump(list(ids), f, indent=2)
    logger.info(f"💾 Saved {len(ids)} processed IDs to {TRACKER_FILE}")

def filter_new_messages(messages, processed_ids):
    """Filtra solo los emails que no han sido procesados."""
    new = [m for m in messages if m['id'] not in processed_ids]
    logger.info(f"📬 {len(new)} new emails out of {len(messages)} total")
    return new