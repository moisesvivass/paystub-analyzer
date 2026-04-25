from typing import Optional
from pydantic import BaseModel, field_validator, model_validator
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


def _sanitize_text(value: str) -> str:
    """Prevent formula injection in Excel by stripping leading formula chars."""
    value = value.strip()
    while value and value[0] in _FORMULA_PREFIXES:
        value = value[1:].strip()
    return value[:500]


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

    @field_validator("company", "pay_period_start", "pay_period_end", mode="before")
    @classmethod
    def sanitize_strings(cls, v: object) -> str:
        if not isinstance(v, str):
            raise ValueError(f"Expected string, got {type(v).__name__}")
        return _sanitize_text(v)

    @field_validator(
        "gross_pay", "net_pay", "federal_tax",
        "provincial_tax", "cpp", "ei", "vacation_pay",
        mode="before"
    )
    @classmethod
    def parse_numeric(cls, v: object) -> float:
        """Convert strings like '$2,134.90' to float if needed."""
        if v is None:
            return 0.0
        if isinstance(v, str):
            cleaned = v.replace("$", "").replace(",", "").strip()
            return float(cleaned) if cleaned else 0.0
        if isinstance(v, (int, float)):
            return float(v)
        raise ValueError(f"Cannot convert {v!r} to float")

    @field_validator("gross_pay", "net_pay", mode="after")
    @classmethod
    def must_be_non_negative(cls, v: float) -> float:
        if v < 0:
            raise ValueError(f"Pay value cannot be negative: {v}")
        return v

    @model_validator(mode="after")
    def validate_pay_period_order(self) -> "PaystubData":
        if self.pay_period_start and self.pay_period_end:
            if self.pay_period_start > self.pay_period_end:
                raise ValueError(
                    f"pay_period_start ({self.pay_period_start}) is after "
                    f"pay_period_end ({self.pay_period_end})"
                )
        return self

    def validate_math(self) -> None:
        """Warn if Gross Pay minus deductions doesn't roughly equal Net Pay."""
        deductions = self.federal_tax + self.provincial_tax + self.cpp + self.ei
        expected_net = self.gross_pay - deductions
        difference = abs(expected_net - self.net_pay)
        if difference > 5.0:
            logger.warning(
                f"Math check — {self.company} {self.pay_period_end}: "
                f"expected net ${expected_net:.2f}, actual ${self.net_pay:.2f}, "
                f"diff ${difference:.2f}"
            )
