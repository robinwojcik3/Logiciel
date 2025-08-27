import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from utils.cache import load_json_cache, save_json_cache


def test_cache_roundtrip(tmp_path):
    path = tmp_path / "cache.json"
    data = {"x": 1}
    save_json_cache(path, data)
    assert load_json_cache(path, 5) == data

    old = time.time() - 10
    os.utime(path, (old, old))
    assert load_json_cache(path, 5) is None
