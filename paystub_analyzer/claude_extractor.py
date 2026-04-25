import json
import time
import anthropic
from anthropic import APIStatusError, APIConnectionError, APITimeoutError
from paystub_analyzer.config import ANTHROPIC_API_KEY
from paystub_analyzer.models import PaystubData
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
_MODEL = "claude-haiku-4-5-20251001"
_MAX_RETRIES = 3
_RETRY_DELAY = 2.0


def extract_data_with_claude(text: str) -> dict:
    """Extract structured payroll data from paystub text using Claude AI.

    Retries up to _MAX_RETRIES times on transient API errors.
    Raises ValueError for JSON/validation failures (not retried).
    """
    last_error: Exception | None = None

    for attempt in range(1, _MAX_RETRIES + 1):
        try:
            response = _client.messages.create(
                model=_MODEL,
                max_tokens=1024,
                messages=[{
                    "role": "user",
                    "content": (
                        "Extract the following fields from this paystub and return ONLY a JSON object:\n"
                        "- pay_period_start (YYYY-MM-DD)\n"
                        "- pay_period_end (YYYY-MM-DD)\n"
                        "- gross_pay (number)\n"
                        "- net_pay (number)\n"
                        "- federal_tax (number)\n"
                        "- provincial_tax (number)\n"
                        "- cpp (number)\n"
                        "- ei (number)\n"
                        "- vacation_pay (number)\n"
                        "- hours_worked (number or null)\n"
                        "- company (text)\n\n"
                        f"Paystub:\n{text}\n\n"
                        "Return ONLY valid JSON, no explanation."
                    )
                }]
            )
            raw: str = response.content[0].text
            clean = raw.replace("```json", "").replace("```", "").strip()

            try:
                data = json.loads(clean)
            except json.JSONDecodeError as e:
                logger.error(f"Claude returned invalid JSON: {e} — raw (200 chars): {raw[:200]!r}")
                raise ValueError(f"Claude response could not be parsed as JSON: {e}") from e

            paystub = PaystubData(**data)
            paystub.validate_math()
            logger.info(f"Data extracted — {paystub.company} — {paystub.pay_period_end}")
            return paystub.model_dump()

        except (APIStatusError, APIConnectionError, APITimeoutError) as e:
            last_error = e
            logger.warning(f"Claude API error (attempt {attempt}/{_MAX_RETRIES}): {e}")
            if attempt < _MAX_RETRIES:
                time.sleep(_RETRY_DELAY * (2 ** (attempt - 1)))

        except (ValueError, anthropic.BadRequestError):
            raise

    raise RuntimeError(f"Claude API failed after {_MAX_RETRIES} attempts") from last_error
