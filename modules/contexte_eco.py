import os
import threading
import tkinter as tk
from tkinter import ttk
from typing import List

from utils.cache import load_json_cache, save_json_cache
from utils.fs import share_available


BASE_SCAN = os.path.dirname(__file__)


def discover_projects():
    """Scanne les projets dans le dossier BASE_SCAN."""
    if not share_available(BASE_SCAN):
        return []
    try:
        return [d for d in os.listdir(BASE_SCAN)
                if os.path.isdir(os.path.join(BASE_SCAN, d))]
    except Exception:
        return []


def discover_projects_cached(max_age: int = 300):
    """Découverte avec cache JSON de 5 min."""
    cache = os.path.join(os.path.expanduser("~"), ".app_cache", "projects.json")
    cached = load_json_cache(cache, max_age)
    if cached is not None:
        return cached
    projs = discover_projects()
    save_json_cache(cache, projs)
    return projs


class ContexteEcoTab(ttk.Frame):
    """Onglet « Contexte éco » avec scan réseau asynchrone."""

    def __init__(self, parent):
        super().__init__(parent)
        self._build_header_light()
        self._list_container = None
        self.after(200, self._populate_projects_async)

    # --- Construction légère ---
    def _build_header_light(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=4)
        ttk.Button(bar, text="Actualiser", command=self._populate_projects_async).pack(side="left")
        self._status = ttk.Label(bar, text="Prêt")
        self._status.pack(side="left", padx=8)

    def _build_list_container(self):
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True)
        lb = tk.Listbox(frame)
        lb.pack(fill="both", expand=True)
        return lb

    # --- Chargement asynchrone ---
    def _populate_projects_async(self) -> None:
        self._set_status("Chargement des projets…")

        def work():
            projs = discover_projects_cached()
            self.after(0, lambda: self._populate_projects_ui(projs))

        threading.Thread(target=work, daemon=True).start()

    def _populate_projects_ui(self, projs: List[str]) -> None:
        if self._list_container is None:
            self._list_container = self._build_list_container()
        self._list_container.delete(0, tk.END)
        for p in projs:
            self._list_container.insert(tk.END, p)
        self._set_status(f"{len(projs)} projets")

    def _set_status(self, msg: str) -> None:
        self._status.config(text=msg)

    # --- Fonctions lourdes importées à la demande ---
    def export_word(self):
        from docx import Document
        return Document()

    def run_web_capture(self):
        from selenium import webdriver
        return webdriver.Chrome()

    def load_geodata(self):
        import geopandas as gpd
        return gpd.GeoDataFrame()

    def load_heif(self, path: str):
        import pillow_heif
        return pillow_heif.read_heif(path)
