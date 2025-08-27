import os
import time
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.cache import load_json_cache, save_json_cache

def test_cache_roundtrip(tmp_path):
    path = tmp_path / "c.json"
    save_json_cache(str(path), {"a": 1})
    data = load_json_cache(str(path), 60)
    assert data == {"a": 1}

def test_cache_expiry(tmp_path):
    path = tmp_path / "c.json"
    save_json_cache(str(path), {"a": 1})
    old = time.time() - 600
    os.utime(path, (old, old))
    assert load_json_cache(str(path), 60) is None
