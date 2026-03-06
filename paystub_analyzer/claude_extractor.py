import json
import os
from dotenv import load_dotenv
from anthropic import Anthropic
from paystub_analyzer.logger import get_logger

load_dotenv()

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
        data = json.loads(clean)
        logger.info(f"✅ Data extracted — {data.get('company')} — {data.get('pay_period_start')}")
        return data
    except Exception as e:
        logger.error(f"❌ Claude extraction failed: {e}")
        raise