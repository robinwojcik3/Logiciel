import tkinter as tk
from tkinter import ttk


class StyleHelper:
    def __init__(self, master, prefs: dict):
        self.master = master
        self.prefs = prefs
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

    def apply(self, theme: str = "light"):
        if theme == "light":
            bg, fg, card_bg, accent, border, subfg = "#F6F7F9", "#111827", "#FFFFFF", "#2563EB", "#E5E7EB", "#6B7280"
            active_accent = "#1D4ED8"
        elif theme == "dark":
            bg, fg, card_bg, accent, border, subfg = "#0F172A", "#E5E7EB", "#111827", "#3B82F6", "#1F2937", "#9CA3AF"
            active_accent = "#2563EB"
        else:  # thème funky
            bg, fg, card_bg, accent, border, subfg = "#1E1E2F", "#FCEFF9", "#27293D", "#FF47A1", "#37394D", "#FFE66D"
            active_accent = "#E0007D"

        self.master.configure(bg=bg)
        s = self.style
        s.configure(".", background=bg, foreground=fg, fieldbackground=card_bg, bordercolor=border)
        s.configure("Card.TFrame", background=card_bg, bordercolor=border, relief="solid", borderwidth=1)
        s.configure("Header.TFrame", background=card_bg, bordercolor=border, relief="flat")
        s.configure("TLabel", background=bg, foreground=fg)
        s.configure("Card.TLabel", background=card_bg, foreground=fg)
        s.configure("Subtle.TLabel", background=bg, foreground=subfg)
        s.configure("Tooltip.TLabel", background="#111827", foreground="#F9FAFB")
        s.configure("Accent.TButton", padding=10, background=accent, foreground="#FFFFFF")
        s.map("Accent.TButton", background=[("active", active_accent)], foreground=[("active", "#FFFFFF")])
        s.configure("Card.TCheckbutton", background=card_bg, foreground=fg)
        s.configure("Card.TRadiobutton", background=card_bg, foreground=fg)
        s.configure("Card.TEntry", fieldbackground=card_bg)
        s.configure("Status.TLabel", background=card_bg, foreground=subfg)
        s.configure("TProgressbar", troughcolor=border)
