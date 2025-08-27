import os

def share_available(path: str) -> bool:
    try:
        return os.path.exists(path)
    except Exception:
        return False
