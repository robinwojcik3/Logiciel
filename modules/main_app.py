import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont

from .style_helper import StyleHelper
from .utils import load_prefs, save_prefs
from .contexte_eco import ContexteEcoTab
from .remonter_le_temps import RemonterLeTempsTab
from .plantnet import PlantNetTab
class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Contexte éco — Outils")
        self.root.geometry("1060x760"); self.root.minsize(900, 640)

        self.prefs = load_prefs()
        self.style_helper = StyleHelper(root, self.prefs)
        self.theme_var = tk.StringVar(value=self.prefs.get("theme", "light"))
        self.style_helper.apply(self.theme_var.get())

        # Header global + bouton thème
        top = ttk.Frame(root, style="Header.TFrame", padding=(12, 8))
        top.pack(fill=tk.X)
        ttk.Label(top, text="Contexte éco — Suite d’outils", style="Card.TLabel",
                  font=tkfont.Font(family="Segoe UI", size=16, weight="bold")).pack(side=tk.LEFT)
        btn_theme = ttk.Button(top, text="Changer de thème", command=self._toggle_theme)
        btn_theme.pack(side=tk.RIGHT)

        # Notebook
        nb = ttk.Notebook(root)
        nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        self.tab_ctx   = ContexteEcoTab(nb, self.style_helper, self.prefs)
        self.tab_rlt   = RemonterLeTempsTab(nb, self.style_helper, self.prefs)
        self.tab_plant = PlantNetTab(nb, self.style_helper, self.prefs)

        nb.add(self.tab_ctx, text="Contexte éco")
        nb.add(self.tab_rlt, text="Remonter le temps & Bassin versant")
        nb.add(self.tab_plant, text="Pl@ntNet")

        # Raccourcis utiles
        root.bind("<Control-1>", lambda _e: nb.select(0))
        root.bind("<Control-2>", lambda _e: nb.select(1))
        root.bind("<Control-3>", lambda _e: nb.select(2))

        # Sauvegarde prefs à la fermeture
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _toggle_theme(self):
        themes = ["light", "dark", "funky"]
        current = self.theme_var.get()
        new_theme = themes[(themes.index(current) + 1) % len(themes)]
        self.theme_var.set(new_theme)
        self.prefs["theme"] = new_theme
        save_prefs(self.prefs)
        self.style_helper.apply(new_theme)

    def _on_close(self):
        try:
            save_prefs(self.prefs)
        finally:
            self.root.destroy()

