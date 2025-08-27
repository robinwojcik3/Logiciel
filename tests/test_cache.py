import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from utils.cache import load_json_cache, save_json_cache


def test_cache_roundtrip(tmp_path):
    data = {"a": 1}
    p = tmp_path / "cache.json"
    save_json_cache(str(p), data)
    loaded = load_json_cache(str(p), max_age=5)
    assert loaded == data


def test_cache_expired(tmp_path):
    data = {"a": 1}
    p = tmp_path / "cache.json"
    save_json_cache(str(p), data)
    # Simulate old file
    os.utime(p, (time.time() - 10, time.time() - 10))
    loaded = load_json_cache(str(p), max_age=5)
    assert loaded is None
