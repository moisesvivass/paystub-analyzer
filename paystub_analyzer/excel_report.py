import openpyxl
from paystub_analyzer.config import OUTPUT_EXCEL
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

def create_excel(data_list: list) -> None:
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Paystubs"

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
        logger.info(f"✅ Excel saved at: {OUTPUT_EXCEL}")

    except Exception as e:
        logger.error(f"❌ Failed to create Excel: {e}")
        raise