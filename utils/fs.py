import os

def share_available(path: str) -> bool:
    """Return True if *path* exists, False otherwise without raising."""
    try:
        return os.path.exists(path)
    except Exception:
        return False
