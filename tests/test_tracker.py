import pytest
from unittest.mock import patch
from paystub_analyzer.tracker import filter_new_messages, mark_processed, load_processed_ids


def test_filter_new_messages_all_new():
    messages = [{"id": "abc"}, {"id": "def"}]
    result = filter_new_messages(messages, set())
    assert len(result) == 2


def test_filter_new_messages_none_new():
    messages = [{"id": "abc"}, {"id": "def"}]
    result = filter_new_messages(messages, {"abc", "def"})
    assert len(result) == 0


def test_filter_new_messages_some_new():
    messages = [{"id": "abc"}, {"id": "def"}, {"id": "xyz"}]
    result = filter_new_messages(messages, {"abc"})
    assert len(result) == 2


def test_mark_and_load_processed_ids(tmp_path, monkeypatch):
    """Marking an ID persists it through load_processed_ids (DB-backed)."""
    db_path = str(tmp_path / "test.db")
    monkeypatch.setenv("DB_FILE", db_path)

    import importlib
    import paystub_analyzer.config as cfg
    importlib.reload(cfg)
    import paystub_analyzer.database as db
    importlib.reload(db)
    db.init_db()

    import paystub_analyzer.tracker as tracker
    importlib.reload(tracker)

    tracker.mark_processed("email_001")
    tracker.mark_processed("email_002")

    ids = tracker.load_processed_ids()
    assert "email_001" in ids
    assert "email_002" in ids
