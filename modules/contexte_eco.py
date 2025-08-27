import os
import threading
import tkinter as tk
from tkinter import ttk
from typing import List

from utils.cache import load_json_cache, save_json_cache
from utils.fs import share_available

BASE_SHARE = r"\\192.168.1.240\commun\PARTAGE"  # fictif pour l'exemple
SUBPATH = r"Espace_RWO\CARTO ROBIN"


def discover_projects() -> List[str]:
    """Scan le partage réseau pour récupérer la liste des projets."""
    base = os.path.join(BASE_SHARE, SUBPATH)
    if not share_available(base):
        return []
    try:
        with os.scandir(base) as it:
            return [entry.name for entry in it if entry.is_dir()]
    except Exception:
        return []


def discover_projects_cached(max_age: int = 300) -> List[str]:
    cache_path = os.path.join(os.path.expanduser("~"), ".app_cache", "projects.json")
    cached = load_json_cache(cache_path, max_age)
    if cached is not None:
        return cached
    projs = discover_projects()
    save_json_cache(cache_path, projs)
    return projs


class ContexteEcoTab(ttk.Frame):
    """Onglet chargé à la demande qui scanne les projets en arrière-plan."""

    def __init__(self, parent):
        super().__init__(parent)
        self.status_var = tk.StringVar(value="En attente")
        self._list_container = None
        self._build_header_light()
        self.after(200, self._populate_projects_async)

    # --- UI léger ---
    def _build_header_light(self):
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="Actualiser", command=self._populate_projects_async).pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.status_var).pack(side=tk.LEFT, padx=10)

    # --- Construction différée du conteneur lourd ---
    def _build_list_container(self):
        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, pady=10)
        listbox = tk.Listbox(frame)
        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.config(yscrollcommand=scroll.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        return listbox

    # --- Peuplement asynchrone ---
    def _populate_projects_async(self):
        self._set_status("Chargement des projets…")

        def work():
            projs = discover_projects_cached()
            self.after(0, lambda: self._populate_projects_ui(projs))

        threading.Thread(target=work, daemon=True).start()

    def _populate_projects_ui(self, projs: List[str]):
        if self._list_container is None:
            self._list_container = self._build_list_container()
        self._list_container.delete(0, tk.END)
        for p in projs:
            self._list_container.insert(tk.END, p)
        self._set_status(f"{len(projs)} projets")

    def _set_status(self, msg: str):
        self.status_var.set(msg)


# --- Exemples de fonctions avec imports lourds différés ---
def export_word(*args, **kwargs):
    from docx import Document  # import différé
    return Document()


def run_web_capture(*args, **kwargs):
    from selenium import webdriver  # import différé
    return webdriver


def load_geodata(*args, **kwargs):
    import geopandas as gpd  # import différé
    return gpd
