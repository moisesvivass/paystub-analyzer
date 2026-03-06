import os
import openpyxl
from paystub_analyzer.config import OUTPUT_EXCEL
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

HEADERS = ["Company", "Pay Period Start", "Pay Period End", "Gross Pay",
           "Net Pay", "Federal Tax", "Provincial Tax",
           "CPP", "EI", "Vacation Pay", "Hours Worked"]

def load_existing_data() -> list:
    """Lee los datos ya existentes en el Excel."""
    if not os.path.exists(OUTPUT_EXCEL):
        return []
    wb = openpyxl.load_workbook(OUTPUT_EXCEL)
    ws = wb["Paystubs"]
    existing = []
    for row in ws.iter_rows(min_row=2, values_only=True):  # skip header
        if any(row):  # skip empty rows
            existing.append(row)
    logger.info(f"📂 Loaded {len(existing)} existing records from Excel")
    return existing

def deduplicate(existing: list, new_data: list) -> list:
    """Evita duplicados comparando Company + Pay Period End."""
    existing_keys = set()
    for row in existing:
        key = (row[0], row[2])  # Company + Pay Period End
        existing_keys.add(key)

    added = 0
    for data in new_data:
        key = (data.get('company'), data.get('pay_period_end'))
        if key not in existing_keys:
            existing.append([
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
            existing_keys.add(key)
            added += 1

    logger.info(f"✅ {added} new records added — {len(existing)} total")
    return existing

def create_excel(data_list: list) -> None:
    try:
        # Cargar datos existentes y agregar nuevos sin duplicados
        existing = load_existing_data()
        all_data = deduplicate(existing, data_list)

        # Ordenar por fecha
        all_data.sort(key=lambda x: x[1] or '')

        # Escribir Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Paystubs"
        ws.append(HEADERS)
        for row in all_data:
            ws.append(list(row))

        wb.save(OUTPUT_EXCEL)
        logger.info(f"💾 Excel saved at: {OUTPUT_EXCEL}")

    except Exception as e:
        logger.error(f"❌ Failed to create Excel: {e}")
        raise