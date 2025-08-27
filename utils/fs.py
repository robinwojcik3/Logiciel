"""Fonctions utilitaires liées au système de fichiers."""

from __future__ import annotations

import os


def share_available(path: str) -> bool:
    """Test rapide de disponibilité d'un partage réseau."""
    try:
        return os.path.exists(path)
    except Exception:
        return False
