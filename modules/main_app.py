import os
import time
import tkinter as tk
from tkinter import ttk

from .contexte_eco import ContexteEcoTab

LAZY_TABS = os.getenv("LAZY_TABS", "1") != "0"


def show_splash(root):
    if os.getenv("APP_SPLASH", "0") == "1":
        win = tk.Toplevel(root)
        win.overrideredirect(True)
        tk.Label(win, text="Initialisation…").pack(padx=20, pady=20)
        win.update_idletasks()
        x = (root.winfo_screenwidth() - win.winfo_reqwidth()) // 2
        y = (root.winfo_screenheight() - win.winfo_reqheight()) // 2
        win.geometry(f"+{x}+{y}")
        return win
    return None


class MainApp(tk.Tk):
    def __init__(self, t0: float):
        super().__init__()
        self.t0 = t0
        self.splash = None
        self.title("App")
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        self._tabs = [
            ("Contexte éco", self._make_contexte_eco),
            ("Remonter le temps & Bassin versant", self._make_rlt),
            ("Pl@ntNet", self._make_plantnet),
        ]
        self._loaded = set()

        if LAZY_TABS:
            for title, _ in self._tabs:
                self.nb.add(ttk.Frame(self.nb), text=title)
            self.nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)
            self.after(0, lambda: self.nb.select(0))
        else:
            for title, factory in self._tabs:
                self.nb.add(factory(self.nb), text=title)

        self.after(0, self._log_startup)

    # --- Tab factories ---
    def _make_contexte_eco(self, parent):
        return ContexteEcoTab(parent)

    def _make_rlt(self, parent):
        f = ttk.Frame(parent)
        ttk.Label(f, text="Remonter le temps & Bassin versant").pack(padx=20, pady=20)
        return f

    def _make_plantnet(self, parent):
        f = ttk.Frame(parent)
        ttk.Label(f, text="Pl@ntNet").pack(padx=20, pady=20)
        return f

    # --- Lazy loading ---
    def _on_tab_changed(self, _event):
        idx = self.nb.index("current")
        if idx in self._loaded:
            return
        self._load_tab(idx)

    def _load_tab(self, idx: int):
        title, factory = self._tabs[idx]
        placeholder = self.nb.nametowidget(self.nb.select())
        real = factory(self.nb)
        self.nb.forget(idx)
        self.nb.insert(idx, real, text=title)
        self.nb.select(idx)
        self._loaded.add(idx)
        if idx == 0:
            t2 = time.time()
            print(f"Premier onglet prêt en {t2 - self.t0:.3f}s")
            if self.splash:
                self.splash.destroy()

    def _log_startup(self):
        t1 = time.time()
        print(f"Fenêtre affichée en {t1 - self.t0:.3f}s")

    # --- Splash ---
    def set_splash(self, splash):
        self.splash = splash


def launch():
    t0 = time.time()
    app = MainApp(t0)
    splash = show_splash(app)
    if splash:
        app.set_splash(splash)
    app.mainloop()


if __name__ == "__main__":
    launch()
