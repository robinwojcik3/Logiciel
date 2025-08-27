import json
import os
import sys
import tempfile
import time

sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from utils.cache import load_json_cache, save_json_cache


def test_cache_roundtrip():
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "data.json")
        data = {"a": 1}
        save_json_cache(path, data)
        assert load_json_cache(path, 60) == data


def test_cache_expired():
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "data.json")
        with open(path, "w", encoding="utf-8") as fh:
            json.dump({"a": 1}, fh)
        os.utime(path, (time.time() - 120, time.time() - 120))
        assert load_json_cache(path, 60) is None
