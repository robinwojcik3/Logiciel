"""Onglet fictif pour l'identification Pl@ntNet."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class PlantNetTab(ttk.Frame):
    def __init__(self, parent: tk.Misc):
        super().__init__(parent)
        ttk.Label(self, text="Pl@ntNet").pack(padx=20, pady=20)
