import json
from anthropic import Anthropic
from paystub_analyzer.models import PaystubData
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

anthropic_client = Anthropic()

def extract_data_with_claude(text: str) -> dict:
    try:
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
        try:
            data = json.loads(clean)
        except json.JSONDecodeError as e:
            logger.error(f"❌ Claude returned invalid JSON: {e} — raw response (first 200 chars): {raw[:200]!r}")
            raise ValueError(f"Claude response could not be parsed as JSON: {e}") from e

        # ── Pydantic validation ────────────────────────────────────────────────
        paystub = PaystubData(**data)
        paystub.validate_math()

        logger.info(f"✅ Data extracted — {paystub.company} — {paystub.pay_period_start}")
        return paystub.model_dump()

    except ValueError:
        raise
    except Exception as e:
        logger.error(f"❌ Claude extraction failed: {e}", exc_info=True)
        raise