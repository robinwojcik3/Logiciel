import os
import json
import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont
from typing import List

PREFS_PATH = os.path.join(os.path.expanduser("~"), "ExportCartesContexteEco.config.json")
OUT_IMG = r"C:\Users\utilisateur\Mon Drive\1 - Bota & Travail\+++++++++  BOTA  +++++++++\---------------------- 3) BDD\PYTHON\2) Contexte éco\OUTPUT"

def log_with_time(msg: str) -> None:
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")


def normalize_name(s: str) -> str:
    s2 = s.replace("\u00A0", " ").replace("\u202F", " ")
    while "  " in s2:
        s2 = s2.replace("  ", " ")
        
    return s2.strip().lower()


def to_long_unc(path: str) -> str:
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path


def chunk_even(lst: List[str], k: int) -> List[List[str]]:
    if not lst:
        return []
    k = max(1, min(k, len(lst)))
    base = len(lst) // k
    extra = len(lst) % k
    out, start = [], 0
    for i in range(k):
        size = base + (1 if i < extra else 0)
        out.append(lst[start:start + size])
        start += size
    return out


def load_prefs() -> dict:
    if os.path.isfile(PREFS_PATH):
        try:
            with open(PREFS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_prefs(d: dict) -> None:
    try:
        with open(PREFS_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, s):
        self.widget.config(state='normal')
        self.widget.insert(tk.END, s)
        self.widget.see(tk.END)
        self.widget.config(state='disabled')
        self.widget.update_idletasks()

    def flush(self):
        pass


class ToolTip:
    def __init__(self, widget, text: str, delay: int = 600):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tipwindow = None
        self.id = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)

    def _schedule(self, _=None):
        self._cancel()
        self.id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def _show(self):
        if self.tipwindow:
            return
        x = self.widget.winfo_rootx() + 15
        y = self.widget.winfo_rooty() + 25
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        ttk.Label(tw, text=self.text, style="Tooltip.TLabel", padding=(8, 4)).pack()

    def _hide(self, _=None):
        self._cancel()
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None


class StyleHelper:
    def __init__(self, root: tk.Tk, prefs: dict):
        self.root = root
        self.prefs = prefs
        self.style = ttk.Style(root)

    def apply(self, theme: str) -> None:
        if theme == "dark":
            bg = "#2b2b2b"
            fg = "#ffffff"
        elif theme == "funky":
            bg = "#ffd700"
            fg = "#000000"
        else:
            bg = "#f0f0f0"
            fg = "#000000"

        self.style.configure("TFrame", background=bg)
        self.style.configure("TLabel", background=bg, foreground=fg)
        self.style.configure("TButton", background=bg)
        self.style.configure("Tooltip.TLabel", background=fg, foreground=bg)
