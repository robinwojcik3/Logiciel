import os


def share_available(path: str) -> bool:
    """Teste rapidement la disponibilité d'un partage réseau."""
    try:
        return os.path.exists(path)
    except Exception:
        return False
