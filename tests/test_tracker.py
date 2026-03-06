from paystub_analyzer.tracker import filter_new_messages, load_processed_ids, save_processed_ids

def test_filter_new_messages_all_new():
    """All messages should be returned if no IDs have been processed."""
    messages = [{"id": "abc"}, {"id": "def"}]
    result = filter_new_messages(messages, set())
    assert len(result) == 2

def test_filter_new_messages_none_new():
    """No messages should be returned if all IDs were already processed."""
    messages = [{"id": "abc"}, {"id": "def"}]
    result = filter_new_messages(messages, {"abc", "def"})
    assert len(result) == 0

def test_filter_new_messages_some_new():
    """Only unprocessed messages should be returned."""
    messages = [{"id": "abc"}, {"id": "def"}, {"id": "xyz"}]
    result = filter_new_messages(messages, {"abc"})
    assert len(result) == 2

def test_save_and_load_processed_ids(tmp_path, monkeypatch):
    """Saved IDs should match loaded IDs."""
    monkeypatch.chdir(tmp_path)
    ids = {"abc123", "def456"}
    save_processed_ids(ids)
    loaded = load_processed_ids()
    assert loaded == ids