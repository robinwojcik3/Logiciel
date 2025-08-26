import datetime
import json
import os
from typing import List

# =========================
# Utils communs
# =========================

def log_with_time(msg: str) -> None:
    """Afficher un message horodaté dans la console."""
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")

def normalize_name(s: str) -> str:
    """Normaliser une chaîne (espaces, casse)."""
    s2 = s.replace("\u00A0", " ").replace("\u202F", " ")
    while "  " in s2:
        s2 = s2.replace("  ", " ")
    return s2.strip().lower()

def to_long_unc(path: str) -> str:
    """Préfixer correctement les chemins UNC."""
    if path.startswith("\\\\?\\"): return path
    if path.startswith("\\\\"):   return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path

def chunk_even(lst: List[str], k: int) -> List[List[str]]:
    """Répartir une liste en blocs de taille similaire."""
    if not lst: return []
    k = max(1, min(k, len(lst)))
    base = len(lst) // k
    extra = len(lst) % k
    out, start = [], 0
    for i in range(k):
        size = base + (1 if i < extra else 0)
        out.append(lst[start:start+size])
        start += size
    return out

def load_prefs(path: str) -> dict:
    """Charger un fichier de préférences JSON."""
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_prefs(path: str, data: dict) -> None:
    """Sauvegarder un dictionnaire de préférences au format JSON."""
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
