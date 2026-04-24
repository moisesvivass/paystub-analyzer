import pytest
from pydantic import ValidationError
from paystub_analyzer.models import PaystubData, _sanitize_text


BASE = {
    "company": "Test Corp",
    "pay_period_start": "2025-01-01",
    "pay_period_end": "2025-01-15",
    "gross_pay": 2000.0,
    "net_pay": 1500.0,
    "federal_tax": 250.0,
    "provincial_tax": 150.0,
    "cpp": 80.0,
    "ei": 20.0,
    "vacation_pay": 80.0,
    "hours_worked": 80.0,
}


def test_valid_paystub():
    p = PaystubData(**BASE)
    assert p.company == "Test Corp"
    assert p.gross_pay == 2000.0


def test_parse_numeric_string():
    data = {**BASE, "gross_pay": "$2,134.90", "net_pay": "1,500.00"}
    p = PaystubData(**data)
    assert p.gross_pay == 2134.90
    assert p.net_pay == 1500.00


def test_formula_injection_stripped():
    data = {**BASE, "company": "=CMD|'calc'!A1"}
    p = PaystubData(**data)
    assert not p.company.startswith("=")


def test_formula_injection_plus():
    data = {**BASE, "company": "+evil"}
    p = PaystubData(**data)
    assert not p.company.startswith("+")


def test_negative_gross_pay_rejected():
    with pytest.raises(ValidationError):
        PaystubData(**{**BASE, "gross_pay": -100.0})


def test_pay_period_order_validation():
    with pytest.raises(ValidationError):
        PaystubData(**{**BASE, "pay_period_start": "2025-01-31", "pay_period_end": "2025-01-01"})


def test_sanitize_text_strips_formula_chars():
    assert _sanitize_text("=evil") == "evil"
    assert _sanitize_text("+cmd") == "cmd"
    assert _sanitize_text("normal text") == "normal text"


def test_hours_worked_optional():
    data = {k: v for k, v in BASE.items() if k != "hours_worked"}
    p = PaystubData(**data)
    assert p.hours_worked is None


def test_validate_math_warns_on_large_difference(caplog):
    import logging
    data = {**BASE, "gross_pay": 2000.0, "net_pay": 100.0,
            "federal_tax": 0.0, "provincial_tax": 0.0, "cpp": 0.0, "ei": 0.0}
    p = PaystubData(**data)
    with caplog.at_level(logging.WARNING):
        p.validate_math()
    assert "Math check" in caplog.text
