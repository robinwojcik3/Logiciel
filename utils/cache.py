import json
import os
import time
from typing import Any

def load_json_cache(path: str, max_age: int):
    try:
        if os.path.isfile(path) and (time.time() - os.path.getmtime(path) < max_age):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return None

def save_json_cache(path: str, data: Any) -> None:
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception:
        pass
