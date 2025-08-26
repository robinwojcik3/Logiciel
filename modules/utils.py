"""Fonctions utilitaires communes au projet."""

import datetime
import json
import os
from typing import List


def log_with_time(msg: str) -> None:
    """Affiche un message préfixé par l'heure courante."""
    now = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] {msg}")


def normalize_name(s: str) -> str:
    """Normalise une chaîne en supprimant les espaces doubles et en la mettant en minuscule."""
    s2 = s.replace("\u00A0", " ").replace("\u202F", " ")
    while "  " in s2:
        s2 = s2.replace("  ", " ")
    return s2.strip().lower()


def to_long_unc(path: str) -> str:
    """Convertit un chemin en notation UNC longue pour éviter les limites Windows."""
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path


def chunk_even(lst: List[str], k: int) -> List[List[str]]:
    """Découpe une liste en *k* morceaux de taille aussi égale que possible."""
    if not lst:
        return []
    k = max(1, min(k, len(lst)))
    base = len(lst) // k
    extra = len(lst) % k
    out, start = [], 0
    for i in range(k):
        size = base + (1 if i < extra else 0)
        out.append(lst[start:start + size])
        start += size
    return out


PREFS_PATH = os.path.join(os.path.expanduser("~"), "ExportCartesContexteEco.config.json")


def load_prefs() -> dict:
    """Charge les préférences depuis un fichier JSON."""
    if os.path.isfile(PREFS_PATH):
        try:
            with open(PREFS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_prefs(data: dict) -> None:
    """Sauvegarde les préférences dans un fichier JSON."""
    try:
        with open(PREFS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
