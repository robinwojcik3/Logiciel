"""Onglet fictif pour « Remonter le temps & Bassin versant »."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class RemonterTab(ttk.Frame):
    def __init__(self, parent: tk.Misc):
        super().__init__(parent)
        ttk.Label(self, text="Remonter le temps & Bassin versant").pack(padx=20, pady=20)
