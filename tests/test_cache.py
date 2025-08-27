import os
import time

from utils.cache import load_json_cache, save_json_cache


def test_cache_save_and_load(tmp_path):
    data = {"a": 1}
    path = tmp_path / "test.json"
    save_json_cache(str(path), data)
    assert load_json_cache(str(path), 60) == data
    old = time.time() - 120
    os.utime(path, (old, old))
    assert load_json_cache(str(path), 60) is None
