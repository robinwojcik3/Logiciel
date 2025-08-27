"""Onglet « Contexte éco » avec chargement asynchrone des projets."""

from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import ttk
from typing import List

from utils.cache import load_json_cache, save_json_cache
from utils.fs import share_available

BASE_SHARE = r"\\192.168.1.240\commun\PARTAGE"
SUBPATH = r"Espace_RWO\CARTO ROBIN"


def get_base_share() -> str:
    return os.path.join(BASE_SHARE, SUBPATH)


def discover_projects() -> List[str]:
    base = get_base_share()
    if not share_available(base):
        return []
    try:
        return [f for f in os.listdir(base) if f.lower().endswith(".qgz")]
    except Exception:
        return []


def discover_projects_cached(max_age: int = 300) -> List[str]:
    cache = os.path.join(os.path.expanduser("~"), ".app_cache", "projects.json")
    cached = load_json_cache(cache, max_age)
    if cached is not None:
        return cached
    projs = discover_projects()
    save_json_cache(cache, projs)
    return projs


class ContexteEcoTab(ttk.Frame):
    def __init__(self, parent: tk.Misc):
        super().__init__(parent)
        self._status = tk.StringVar(value="Prêt")
        self._build_header_light()
        self._list_container = None
        self.after(200, self._populate_projects_async)

    # --- Construction légère -------------------------------------------------
    def _build_header_light(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill="x")
        ttk.Label(bar, textvariable=self._status).pack(side="left", padx=5, pady=5)
        ttk.Button(bar, text="Actualiser", command=self._populate_projects_async).pack(
            side="right", padx=5, pady=5
        )

    def _build_list_container(self) -> ttk.Frame:
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True)
        self._listbox = tk.Listbox(frame)
        self._listbox.pack(fill="both", expand=True)
        return frame

    # --- Gestion de l'état ---------------------------------------------------
    def _set_status(self, msg: str) -> None:
        self._status.set(msg)

    def _fill_list(self, _container: ttk.Frame, projs: List[str]) -> None:
        self._listbox.delete(0, tk.END)
        for p in projs:
            self._listbox.insert(tk.END, p)

    # --- Chargement asynchrone ----------------------------------------------
    def _populate_projects_async(self) -> None:
        self._set_status("Chargement des projets...")

        def work() -> None:
            projs = discover_projects_cached()
            self.after(0, lambda: self._populate_projects_ui(projs))

        threading.Thread(target=work, daemon=True).start()

    def _populate_projects_ui(self, projs: List[str]) -> None:
        if self._list_container is None:
            self._list_container = self._build_list_container()
        self._fill_list(self._list_container, projs)
        self._set_status(f"{len(projs)} projets")
