import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from utils.cache import load_json_cache, save_json_cache


def test_save_and_load_cache(tmp_path):
    path = tmp_path / "cache.json"
    data = {"hello": "world"}
    save_json_cache(path, data)
    assert load_json_cache(path, 60) == data


def test_cache_expiry(tmp_path):
    path = tmp_path / "cache.json"
    save_json_cache(path, {"a": 1})
    old = time.time() - 1000
    os.utime(path, (old, old))
    assert load_json_cache(path, 10) is None
