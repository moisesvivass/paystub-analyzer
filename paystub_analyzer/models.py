from pydantic import BaseModel, field_validator
from typing import Optional
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)


class PaystubData(BaseModel):
    company: str
    pay_period_start: str
    pay_period_end: str
    gross_pay: float
    net_pay: float
    federal_tax: float = 0.0
    provincial_tax: float = 0.0
    cpp: float = 0.0
    ei: float = 0.0
    vacation_pay: float = 0.0
    hours_worked: Optional[float] = None

    @field_validator("gross_pay", "net_pay", "federal_tax",
                     "provincial_tax", "cpp", "ei", "vacation_pay", mode="before")
    @classmethod
    def parse_numeric(cls, v):
        """Convert strings like '$2,134.90' to float if needed."""
        if isinstance(v, str):
            return float(v.replace("$", "").replace(",", "").strip())
        return v

    def validate_math(self):
        """Warn if Gross Pay - Deductions does not roughly equal Net Pay."""
        deductions = self.federal_tax + self.provincial_tax + self.cpp + self.ei
        expected_net = self.gross_pay - deductions
        difference = abs(expected_net - self.net_pay)
        if difference > 5.0:
            logger.warning(
                f"⚠️ Math check failed for {self.company} {self.pay_period_end} — "
                f"Expected net: ${expected_net:.2f} | Actual net: ${self.net_pay:.2f} | "
                f"Difference: ${difference:.2f}"
            )