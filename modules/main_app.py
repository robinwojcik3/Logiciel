"""Fenêtre principale avec chargement différé des onglets."""

from __future__ import annotations

import os
import time
import tkinter as tk
from tkinter import ttk

from .contexte_eco_tab import ContexteEcoTab
from .remonter_tab import RemonterTab
from .plantnet_tab import PlantNetTab

LAZY_TABS = os.getenv("LAZY_TABS", "1") != "0"


def show_splash(root: tk.Tk):
    """Affiche une fenêtre de démarrage optionnelle."""
    if os.getenv("APP_SPLASH", "0") == "1":
        win = tk.Toplevel(root)
        win.overrideredirect(True)
        tk.Label(win, text="Initialisation…").pack(padx=20, pady=20)
        win.update_idletasks()
        x = root.winfo_screenwidth() // 2 - win.winfo_reqwidth() // 2
        y = root.winfo_screenheight() // 2 - win.winfo_reqheight() // 2
        win.geometry(f"+{x}+{y}")
        return win
    return None


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("App")
        self._t0 = time.perf_counter()

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        self._tabs = [
            ("Contexte éco", lambda parent: ContexteEcoTab(parent)),
            ("Remonter le temps & Bassin versant", lambda parent: RemonterTab(parent)),
            ("Pl@ntNet", lambda parent: PlantNetTab(parent)),
        ]
        self._loaded = set()

        if LAZY_TABS:
            for title, _ in self._tabs:
                self.nb.add(ttk.Frame(self.nb), text=title)
            self.nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)
            self.after(0, lambda: self.nb.select(0))
        else:
            for title, factory in self._tabs:
                frame = factory(self.nb)
                self.nb.add(frame, text=title)
                self._loaded.add(self.nb.index("end") - 1)
            self.after(0, self._log_ready)

        self._splash = show_splash(self)
        self.after(0, self._select_first)

    # --- Gestion des onglets --------------------------------------------------
    def _select_first(self) -> None:
        if LAZY_TABS:
            self.nb.select(0)
        self._log_ready()

    def _on_tab_changed(self, _event) -> None:
        idx = self.nb.index("current")
        if idx in self._loaded:
            return
        title, factory = self._tabs[idx]
        placeholder = self.nb.nametowidget(self.nb.select())
        real = factory(self.nb)
        self.nb.forget(idx)
        self.nb.insert(idx, real, text=title)
        self.nb.select(idx)
        self._loaded.add(idx)

    def _log_ready(self) -> None:
        delay = time.perf_counter() - self._t0
        print(f"UI prête en {delay:.3f}s")
        if self._splash is not None:
            self._splash.destroy()
            self._splash = None


def launch() -> None:
    app = MainApp()
    app.mainloop()
