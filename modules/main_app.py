import os
import time
import tkinter as tk
from tkinter import ttk

from .contexte_eco import ContexteEcoTab

LAZY_TABS = os.getenv("LAZY_TABS", "1") != "0"


def show_splash(root: tk.Tk):
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
    def __init__(self, start_time: float):
        super().__init__()
        self.start_time = start_time
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
            for idx, (title, factory) in enumerate(self._tabs):
                self.nb.add(factory(self.nb), text=title)
                self._loaded.add(idx)

        self.splash = show_splash(self)
        self.update_idletasks()
        print(f"Fenêtre affichée en {time.time() - self.start_time:.2f}s")

    # ----- Gestion d'onglets -----
    def _on_tab_changed(self, _event):
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
        if idx == 0:
            print(f"Premier onglet chargé en {time.time() - self.start_time:.2f}s")
            if self.splash is not None:
                self.splash.destroy()
                self.splash = None

    # ----- Fabriques d'onglets -----
    def _make_contexte_eco(self, parent):
        return ContexteEcoTab(parent)

    def _make_rlt(self, parent):
        from docx import Document  # import lourd
        _ = Document  # éviter l'avertissement 'unused'
        frame = ttk.Frame(parent)
        ttk.Label(frame, text="Remonter le temps & Bassin versant").pack(padx=10, pady=10)
        return frame

    def _make_plantnet(self, parent):
        import pillow_heif  # import lourd
        pillow_heif.register_heif_opener()
        frame = ttk.Frame(parent)
        ttk.Label(frame, text="Pl@ntNet").pack(padx=10, pady=10)
        return frame


def launch() -> None:
    start_time = time.time()
    app = MainApp(start_time)
    app.mainloop()
