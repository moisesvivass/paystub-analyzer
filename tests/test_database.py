import pytest
import importlib


@pytest.fixture
def db(tmp_path, monkeypatch):
    monkeypatch.setenv("DB_FILE", str(tmp_path / "test.db"))
    import paystub_analyzer.config as cfg
    importlib.reload(cfg)
    import paystub_analyzer.database as database
    importlib.reload(database)
    database.init_db()
    return database


SAMPLE = {
    "company": "Acme Corp",
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


def test_insert_paystub_success(db):
    assert db.insert_paystub(SAMPLE) is True


def test_insert_paystub_duplicate_returns_false(db):
    db.insert_paystub(SAMPLE)
    assert db.insert_paystub(SAMPLE) is False


def test_get_all_paystubs_returns_inserted(db):
    db.insert_paystub(SAMPLE)
    rows = db.get_all_paystubs()
    assert len(rows) == 1
    assert rows[0]["company"] == "Acme Corp"


def test_is_email_processed(db):
    assert db.is_email_processed("msg_001") is False
    db.mark_email_processed("msg_001")
    assert db.is_email_processed("msg_001") is True


def test_get_all_processed_email_ids(db):
    db.mark_email_processed("id_a")
    db.mark_email_processed("id_b")
    ids = db.get_all_processed_email_ids()
    assert {"id_a", "id_b"}.issubset(ids)


def test_run_history_complete_lifecycle(db):
    run_id = db.start_run("test")
    assert run_id is not None
    db.finish_run(run_id, emails_found=10, processed=8, failed=2)
    with db.get_connection() as conn:
        row = conn.execute("SELECT * FROM run_history WHERE id = ?", (run_id,)).fetchone()
    assert row["status"] == "completed"
    assert row["emails_found"] == 10


def test_run_history_fail(db):
    run_id = db.start_run("test")
    db.fail_run(run_id, "Something broke")
    with db.get_connection() as conn:
        row = conn.execute("SELECT * FROM run_history WHERE id = ?", (run_id,)).fetchone()
    assert row["status"] == "failed"
    assert "Something broke" in row["error_message"]
