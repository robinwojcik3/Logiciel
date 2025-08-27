import os
import time
from utils.cache import load_json_cache, save_json_cache


def test_cache_roundtrip(tmp_path):
    path = tmp_path / "cache.json"
    data = {"a": 1}
    save_json_cache(path, data)
    assert load_json_cache(path, 60) == data


def test_cache_expired(tmp_path):
    path = tmp_path / "cache.json"
    data = {"a": 1}
    save_json_cache(path, data)
    old = time.time() - 120
    os.utime(path, (old, old))
    assert load_json_cache(path, 60) is None
