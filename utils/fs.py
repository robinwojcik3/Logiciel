import os

def share_available(path: str) -> bool:
    """Quickly test if a network share or path is reachable."""
    try:
        return os.path.exists(path)
    except Exception:
        return False
