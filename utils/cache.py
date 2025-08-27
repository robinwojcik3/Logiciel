import json
import os
import time
from typing import Any, Optional


def load_json_cache(path: str, max_age: int) -> Optional[Any]:
    """Load JSON data from *path* if file exists and is younger than *max_age* seconds."""
    try:
        if os.path.isfile(path) and (time.time() - os.path.getmtime(path) < max_age):
            with open(path, "r", encoding="utf-8") as fh:
                return json.load(fh)
    except Exception:
        pass
    return None


def save_json_cache(path: str, data: Any) -> None:
    """Save *data* as JSON to *path*, creating parent directories if needed."""
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False)
    except Exception:
        pass
