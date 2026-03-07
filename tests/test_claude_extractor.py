from unittest.mock import MagicMock, patch
from paystub_analyzer.claude_extractor import extract_data_with_claude

FAKE_RESPONSE = """{
    "company": "Test Company Ltd",
    "pay_period_start": "2025-01-01",
    "pay_period_end": "2025-01-07",
    "gross_pay": 2000.00,
    "net_pay": 1500.00,
    "federal_tax": 200.00,
    "provincial_tax": 100.00,
    "cpp": 80.00,
    "ei": 33.00,
    "vacation_pay": 80.00,
    "hours_worked": 40
}"""

@patch("paystub_analyzer.claude_extractor.anthropic_client")
def test_extract_data_success(mock_client):
    """Should extract and validate data without calling the real API."""
    mock_message = MagicMock()
    mock_message.content[0].text = FAKE_RESPONSE
    mock_client.messages.create.return_value = mock_message

    result = extract_data_with_claude("fake paystub text")

    assert result["company"] == "Test Company Ltd"
    assert result["gross_pay"] == 2000.00
    assert result["net_pay"] == 1500.00

@patch("paystub_analyzer.claude_extractor.anthropic_client")
def test_extract_data_invalid_json(mock_client):
    """Should raise an exception if Claude returns invalid JSON."""
    mock_message = MagicMock()
    mock_message.content[0].text = "this is not json"
    mock_client.messages.create.return_value = mock_message

    import pytest
    with pytest.raises(Exception):
        extract_data_with_claude("fake paystub text")