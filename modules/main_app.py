#!/usr/bin/env "C:/Program Files/QGIS 3.40.3/apps/Python312/python.exe"
# -*- coding: utf-8 -*-
r"""
Application à onglets (ttk.Notebook) :

Onglet « Contexte éco » :
    - Export des mises en page QGIS en PNG.
    - Identification des zonages (ID Contexte éco) avec tampon configurable.
    - Sélection commune des shapefiles ZE/AE et console partagée.
    - Boutons « Remonter le temps », « Ouvrir Google Maps » et « Bassin versant »
      utilisant le centroïde de la zone d'étude.

Onglet « Identification Pl@ntNet » :
    - Reconnaissance de plantes via l'API Pl@ntNet.

Pré-requis Python : qgis (environnement QGIS), selenium, pillow, python-docx,
openpyxl (non utilisé ici), chromedriver dans PATH.
"""

import os
import sys
import re
import json
import time
import shutil
import tempfile
import datetime
import threading
import urllib.request
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont
from typing import List, Optional, Tuple
from concurrent.futures import ProcessPoolExecutor, as_completed
import requests
from io import BytesIO
import pillow_heif
import zipfile
import traceback

# ==== Imports supplémentaires pour l'onglet Contexte éco ====
import geopandas as gpd

# Import du scraper Wikipédia
from .wikipedia_scraper import DEP, run_wikipedia_scrape

# Import du worker QGIS externalisé
from .export_worker import worker_run


# ==== Imports spécifiques onglet 2 (gardés en tête de fichier comme le script source) ====
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

from PIL import Image

# Enregistrer le décodeur HEIF
pillow_heif.register_heif_opener()

# =========================
# Paramètres globaux
# =========================
# Contexte éco — Export Cartes
DPI_DEFAULT        = 300
N_WORKERS_DEFAULT  = max(1, min((os.cpu_count() or 2) - 1, 6))
MARGIN_FAC_DEFAULT = 1.15
OVERWRITE_DEFAULT  = False

LAYER_AE_NAME = "Aire d'étude élargie"
LAYER_ZE_NAME = "Zone d'étude"

BASE_SHARE = r"\\192.168.1.240\commun\PARTAGE"
SUBPATH    = r"Espace_RWO\CARTO ROBIN"

OUT_IMG    = r"C:\Users\utilisateur\Mon Drive\1 - Bota & Travail\+++++++++  BOTA  +++++++++\---------------------- 3) BDD\PYTHON\2) Contexte éco\OUTPUT"

# Dossier par défaut pour la sélection des shapefiles (onglet 1)
DEFAULT_SHAPE_DIR = r"C:\Users\utilisateur\Mon Drive\1 - Bota & Travail\+++++++++  BOTA  +++++++++\---------------------- 2) CARTO terrain"

# QGIS
QGIS_ROOT = r"C:\Program Files\QGIS 3.40.3"
QGIS_APP  = os.path.join(QGIS_ROOT, "apps", "qgis")
PY_VER    = "Python312"

# Préférences
PREFS_PATH = os.path.join(os.path.expanduser("~"), "ExportCartesContexteEco.config.json")

# Constantes « Remonter le temps » et « Bassin versant » (issues du script source)
LAYERS = [
    ("Aujourd’hui",   "10"),
    ("2000-2005",     "18"),
    ("1965-1980",     "20"),
    ("1950-1965",     "19"),
]
URL = ("https://remonterletemps.ign.fr/comparer/?lon={lon}&lat={lat}"
       "&z=17&layer1={layer}&layer2=19&mode=dub1")
WAIT_TILES_DEFAULT = 1.5
IMG_WIDTH = Cm(12.5 * 0.8)
WORD_FILENAME = "Comparaison_temporelle_Paysage.docx"
OUTPUT_DIR_RLT = os.path.join(OUT_IMG, "Remonter le temps")
COMMENT_TEMPLATE = (
    "Rédige un commentaire synthétique de l'évolution de l'occupation du sol observée "
    "sur les images aériennes de la zone d'étude, aux différentes dates indiquées "
    "(1950–1965, 1965–1980, 2000–2005, aujourd’hui). Concentre-toi sur les grandes "
    "dynamiques d'aménagement (urbanisation, artificialisation, évolution des milieux "
    "ouverts ou boisés), en identifiant les principales transformations visibles. "
    "Fais ta réponse en un seul court paragraphe. Intègre les éléments de contexte "
    "historique et territorial propres à la commune de {commune} pour interpréter ces évolutions."
)

# Onglet 3 — Identification Pl@ntNet
API_KEY = "2b10vfT6MvFC2lcAzqG1ZMKO"  # Votre clé API Pl@ntNet
PROJECT = "all"
API_URL = f"https://my-api.plantnet.org/v2/identify/{PROJECT}?api-key={API_KEY}"

# =========================
# Utils communs
# =========================
def log_with_time(msg: str) -> None:
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")

def normalize_name(s: str) -> str:
    s2 = s.replace("\u00A0", " ").replace("\u202F", " ")
    while "  " in s2:
        s2 = s2.replace("  ", " ")
    return s2.strip().lower()

def to_long_unc(path: str) -> str:
    if path.startswith("\\\\?\\"): return path
    if path.startswith("\\\\"):   return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path

def chunk_even(lst: List[str], k: int) -> List[List[str]]:
    if not lst: return []
    k = max(1, min(k, len(lst)))
    base = len(lst) // k
    extra = len(lst) % k
    out, start = [], 0
    for i in range(k):
        size = base + (1 if i < extra else 0)
        out.append(lst[start:start+size])
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
    def __init__(self, widget): self.widget = widget
    def write(self, s):
        self.widget.config(state='normal')
        self.widget.insert(tk.END, s)
        self.widget.see(tk.END)
        self.widget.config(state='disabled')
        self.widget.update_idletasks()
    def flush(self): pass

class ToolTip:
    def __init__(self, widget, text: str, delay: int = 600):
        self.widget = widget; self.text = text; self.delay = delay
        self.tipwindow = None; self.id = None
        widget.bind("<Enter>", self._schedule); widget.bind("<Leave>", self._hide)
    def _schedule(self, _=None):
        self._cancel(); self.id = self.widget.after(self.delay, self._show)
    def _cancel(self):
        if self.id: self.widget.after_cancel(self.id); self.id = None
    def _show(self):
        if self.tipwindow: return
        x = self.widget.winfo_rootx() + 15; y = self.widget.winfo_rooty() + 25
        self.tipwindow = tw = tk.Toplevel(self.widget); tw.wm_overrideredirect(True); tw.wm_geometry(f"+{x}+{y}")
        ttk.Label(tw, text=self.text, style="Tooltip.TLabel", padding=(8, 4)).pack()
    def _hide(self, _=None):
        self._cancel()
        if self.tipwindow: self.tipwindow.destroy(); self.tipwindow = None

# =========================
# Fonctions Pl@ntNet
# =========================
def resize_image(image_path, max_size=(800, 800), quality=70):
    """
    Redimensionne et compresse une image.

    :param image_path: Chemin de l'image à traiter.
    :param max_size: Tuple indiquant la taille maximale (largeur, hauteur).
    :param quality: Qualité de compression (1-100).
    :return: BytesIO de l'image traitée ou None en cas d'erreur.
    """
    try:
        with Image.open(image_path) as img:
            img.thumbnail(max_size)
            buffer = BytesIO()
            img.save(buffer, format='JPEG', quality=quality)
            buffer.seek(0)
            return buffer
    except Exception as e:
        print(f"Erreur lors du redimensionnement de l'image : {e}")
        return None

def identify_plant(image_path, organ):
    """
    Envoie une image à l'API Pl@ntNet pour identification.

    :param image_path: Chemin de l'image à envoyer.
    :param organ: Type d'organe de la plante (par exemple, 'flower').
    :return: Nom scientifique de la plante identifiée ou None.
    """
    print(f"Envoi de l'image à l'API : {image_path}")
    try:
        resized_image = resize_image(image_path)
        if not resized_image:
            print(f"Échec du redimensionnement de l'image : {image_path}")
            return None

        files = {
            'images': (os.path.basename(image_path), resized_image, 'image/jpeg')
        }
        data = {
            'organs': organ
        }

        response = requests.post(API_URL, files=files, data=data)

        print(f"Réponse de l'API : {response.status_code}")
        if response.status_code == 200:
            json_result = response.json()
            try:
                species = json_result['results'][0]['species']['scientificNameWithoutAuthor']
                print(f"Plante identifiée : {species}")
                return species
            except (KeyError, IndexError):
                print(f"Aucun résultat trouvé pour l'image : {image_path}")
                return None
        else:
            print(f"Erreur API : {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Exception lors de l'identification de la plante : {e}")
        return None

def copy_and_rename_file(file_path, dest_folder, new_name, count):
    """
    Copie et renomme un fichier dans le dossier de destination.

    :param file_path: Chemin du fichier original.
    :param dest_folder: Dossier de destination.
    :param new_name: Nom scientifique de la plante.
    :param count: Compteur pour différencier les fichiers portant le même nom.
    """
    ext = os.path.splitext(file_path)[1]
    if count == 1:
        new_file_name = f"{new_name} @plantnet{ext}"
    else:
        new_file_name = f"{new_name} @plantnet({count}){ext}"
    new_path = os.path.join(dest_folder, new_file_name)
    try:
        shutil.copy(file_path, new_path)
        print(f"Fichier copié et renommé : {file_path} -> {new_path}")
    except Exception as e:
        print(f"Erreur lors de la copie du fichier : {e}")


def discover_projects() -> List[str]:
    partage = BASE_SHARE
    base_dir: Optional[str] = None
    try:
        if os.path.isdir(partage):
            for e in os.listdir(partage):
                if normalize_name(e) == normalize_name("Espace_RWO"):
                    base_espace = os.path.join(partage, e)
                    for s in os.listdir(base_espace):
                        if normalize_name(s) == normalize_name("CARTO ROBIN"):
                            base_dir = os.path.join(base_espace, s); break
                    break
    except Exception as e:
        log_with_time(f"Accès PARTAGE impossible via listdir: {e}")

    if not base_dir or not os.path.isdir(base_dir):
        base_dir = os.path.join(BASE_SHARE, SUBPATH)

    for d in (base_dir, to_long_unc(base_dir)):
        try:
            if not os.path.isdir(d): continue
            files = os.listdir(d)
            qgz = [f for f in files if f.lower().endswith(".qgz")]
            qgz = [f for f in qgz if normalize_name(f).startswith(normalize_name("Contexte éco -"))]
            return [os.path.join(d, f) for f in sorted(qgz)]
        except Exception:
            continue
    return []

# =========================
# Fonctions IGN (onglet 2) — identiques au script source
# =========================
def dms_to_dd(text: str) -> float:
    pat = r"(\d{1,3})[°d]\s*(\d{1,2})['m]\s*([\d\.]+)[\"s]?\s*([NSEW])"
    alt = r"(\d{1,3})\s+(\d{1,2})\s+([\d\.]+)\s*([NSEW])"
    m = re.search(pat, text, re.I) or re.search(alt, text, re.I)
    if not m:
        raise ValueError(f"Format DMS invalide : {text}")
    deg, mn, sc, hemi = m.groups()
    dd = float(deg) + float(mn)/60 + float(sc)/3600
    return -dd if hemi.upper() in ("S", "W") else dd

def dd_to_dms(lat: float, lon: float) -> str:
    """Convertit des coordonnées décimales en DMS (degrés, minutes, secondes)."""
    def _convert(value: float, positive: str, negative: str) -> str:
        hemi = positive if value >= 0 else negative
        value = abs(value)
        deg = int(value)
        minutes_full = (value - deg) * 60
        minutes = int(minutes_full)
        seconds = (minutes_full - minutes) * 60
        return f"{deg}°{minutes:02d}'{seconds:04.1f}\"{hemi}"

    return f"{_convert(lat, 'N', 'S')} {_convert(lon, 'E', 'W')}"

def add_hyperlink(paragraph, url: str, text: str, italic: bool = True):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    from docx.oxml import OxmlElement, ns
    fld_simple = OxmlElement('w:hyperlink')
    fld_simple.set(ns.qn('r:id'), r_id)

    run = OxmlElement('w:r')
    r_pr = OxmlElement('w:rPr')
    if italic:
        i = OxmlElement('w:i')
        r_pr.append(i)
    u = OxmlElement('w:u')
    u.set(ns.qn('w:val'), 'single')
    r_pr.append(u)
    run.append(r_pr)
    run_text = OxmlElement('w:t')
    run_text.text = text
    run.append(run_text)
    fld_simple.append(run)
    paragraph._p.append(fld_simple)

# =========================
# UI — Styles communs
# =========================
class StyleHelper:
    def __init__(self, master, prefs: dict):
        self.master = master
        self.prefs = prefs
        self.style = ttk.Style()
        try: self.style.theme_use("clam")
        except Exception: pass

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

# =========================
# (Ancien) onglet Export Cartes
# =========================
class ExportCartesTab(ttk.Frame):
    def __init__(self, parent, style_helper: StyleHelper, prefs: dict):
        super().__init__(parent, padding=12)
        self.parent = parent
        self.prefs = prefs
        self.style_helper = style_helper

        self.font_title = tkfont.Font(family="Segoe UI", size=15, weight="bold")
        self.font_sub   = tkfont.Font(family="Segoe UI", size=10)
        self.font_mono  = tkfont.Font(family="Consolas", size=9)

        self.ze_shp_var   = tk.StringVar(value=self.prefs.get("ZE_SHP", ""))
        self.ae_shp_var   = tk.StringVar(value=self.prefs.get("AE_SHP", ""))
        self.cadrage_var  = tk.StringVar(value=self.prefs.get("CADRAGE_MODE", "BOTH"))
        self.overwrite_var= tk.BooleanVar(value=self.prefs.get("OVERWRITE", OVERWRITE_DEFAULT))
        self.dpi_var      = tk.IntVar(value=int(self.prefs.get("DPI", DPI_DEFAULT)))
        self.workers_var  = tk.IntVar(value=int(self.prefs.get("N_WORKERS", N_WORKERS_DEFAULT)))
        self.margin_var   = tk.DoubleVar(value=float(self.prefs.get("MARGIN_FAC", MARGIN_FAC_DEFAULT)))

        self.project_vars: dict[str, tk.IntVar] = {}
        self.all_projects: List[str] = []
        self.filtered_projects: List[str] = []
        self.total_expected = 0
        self.progress_done  = 0

        self._build_ui()
        self._populate_projects()
        self._update_counts()

    def _build_ui(self):
        header = ttk.Frame(self, style="Header.TFrame", padding=(14, 12))
        header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header, text="Export cartes — QGIS → PNG", style="Card.TLabel", font=self.font_title)\
            .grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Sélection shapefiles, choix du cadrage, export multi-projets.", style="Subtle.TLabel", font=self.font_sub)\
            .grid(row=1, column=0, sticky="w", pady=(4,0))
        header.columnconfigure(0, weight=1)

        grid = ttk.Frame(self); grid.pack(fill=tk.BOTH, expand=True)
        left = ttk.Frame(grid);  left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right = ttk.Frame(grid); right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        grid.columnconfigure(0, weight=1); grid.columnconfigure(1, weight=1); grid.rowconfigure(0, weight=1)

        # Shapefiles
        shp = ttk.Frame(left, style="Card.TFrame", padding=12); shp.pack(fill=tk.X)
        ttk.Label(shp, text="1. Couches Shapefile", style="Card.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        self._file_row(shp, 1, "📁 Zone d'étude…", self.ze_shp_var, lambda: self._select_shapefile('ZE'))
        self._file_row(shp, 2, "📁 Aire d'étude élargie…", self.ae_shp_var, lambda: self._select_shapefile('AE'))
        shp.columnconfigure(1, weight=1)

        # Options
        opt = ttk.Frame(left, style="Card.TFrame", padding=12); opt.pack(fill=tk.X, pady=(10,0))
        ttk.Label(opt, text="2. Cadrage et options", style="Card.TLabel").grid(row=0, column=0, columnspan=6, sticky="w")
        ttk.Radiobutton(opt, text="AE + ZE", variable=self.cadrage_var, value="BOTH", style="Card.TRadiobutton").grid(row=1, column=0, sticky="w", pady=(6,2))
        ttk.Radiobutton(opt, text="ZE uniquement", variable=self.cadrage_var, value="ZE", style="Card.TRadiobutton").grid(row=1, column=1, sticky="w", padx=(12,0))
        ttk.Radiobutton(opt, text="AE uniquement", variable=self.cadrage_var, value="AE", style="Card.TRadiobutton").grid(row=1, column=2, sticky="w", padx=(12,0))
        ttk.Checkbutton(opt, text="Écraser si le PNG existe", variable=self.overwrite_var, style="Card.TCheckbutton").grid(row=1, column=3, sticky="w", padx=(24, 0))

        ttk.Label(opt, text="DPI", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=(8,0))
        ttk.Spinbox(opt, from_=72, to=1200, textvariable=self.dpi_var, width=6, justify="right").grid(row=2, column=1, sticky="w", pady=(8,0))
        ttk.Label(opt, text="Workers", style="Card.TLabel").grid(row=2, column=2, sticky="w", padx=(12,0), pady=(8,0))
        ttk.Spinbox(opt, from_=1, to=max(1, (os.cpu_count() or 2)), textvariable=self.workers_var, width=6, justify="right").grid(row=2, column=3, sticky="w", pady=(8,0))
        ttk.Label(opt, text="Marge", style="Card.TLabel").grid(row=2, column=4, sticky="w", padx=(12,0), pady=(8,0))
        ttk.Spinbox(opt, from_=1.00, to=2.00, increment=0.05, textvariable=self.margin_var, width=6, justify="right").grid(row=2, column=5, sticky="w", pady=(8,0))

        # Actions
        act = ttk.Frame(left, style="Card.TFrame", padding=12); act.pack(fill=tk.X, pady=(10,0))
        self.export_button = ttk.Button(act, text="▶ Lancer l’export", style="Accent.TButton", command=self.start_export_thread)
        self.export_button.grid(row=0, column=0, sticky="w")
        obtn = ttk.Button(act, text="📂 Ouvrir le dossier de sortie", command=self._open_out_dir)
        obtn.grid(row=0, column=1, padx=(10,0)); ToolTip(obtn, OUT_IMG)
        tbtn = ttk.Button(act, text="🧪 Tester QGIS", command=self._test_qgis_threaded)
        tbtn.grid(row=0, column=2, padx=(10,0)); ToolTip(tbtn, "Vérifier l’import QGIS/Qt")

        # Projets
        proj = ttk.Frame(right, style="Card.TFrame", padding=12); proj.pack(fill=tk.BOTH, expand=True)
        ttk.Label(proj, text="3. Projets QGIS", style="Card.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(proj, text="Filtrer", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=(6,6))
        self.filter_var = tk.StringVar()
        fe = ttk.Entry(proj, textvariable=self.filter_var, width=32); fe.grid(row=1, column=1, sticky="w", pady=(6,6))
        fe.bind("<KeyRelease>", lambda _e: self._apply_filter())
        ttk.Button(proj, text="Tout", width=6, command=lambda: self._select_all(True)).grid(row=1, column=2, padx=(8,0))
        ttk.Button(proj, text="Aucun", width=6, command=lambda: self._select_all(False)).grid(row=1, column=3, padx=(6,0))

        canvas = tk.Canvas(proj, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(proj, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=2, column=0, columnspan=4, sticky="nsew", pady=(6, 6))
        scrollbar.grid(row=2, column=4, sticky="ns", padx=(6,0))
        proj.rowconfigure(2, weight=1); proj.columnconfigure(1, weight=1)

        # Bas
        bottom = ttk.Frame(self, style="Card.TFrame", padding=12); bottom.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        self.status_label = ttk.Label(bottom, text="Prêt.", style="Status.TLabel"); self.status_label.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(bottom, orient="horizontal", mode="determinate", length=220)
        self.progress.grid(row=0, column=1, sticky="e"); bottom.columnconfigure(0, weight=1)

        log_frame = ttk.Frame(bottom, style="Card.TFrame"); log_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8,0))
        bottom.rowconfigure(1, weight=1)
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state='disabled',
                                bg=self.style_helper.style.lookup("Card.TFrame", "background"),
                                fg=self.style_helper.style.lookup("TLabel", "foreground"))
        self.log_text.configure(font=self.font_mono, relief="flat")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text['yscrollcommand'] = log_scroll.set
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sys.stdout = TextRedirector(self.log_text)

    def _file_row(self, parent, row: int, label: str, var: tk.StringVar, cmd):
        btn = ttk.Button(parent, text=label, command=cmd)
        btn.grid(row=row, column=0, sticky="w", pady=(8 if row == 1 else 4, 2))
        ent = ttk.Entry(parent, textvariable=var, width=10); ent.grid(row=row, column=1, sticky="ew", padx=8)
        ent.configure(state="readonly")
        copy_btn = ttk.Button(parent, text="Copier", width=6, command=lambda: self._copy_to_clipboard(var.get()))
        copy_btn.grid(row=row, column=2, sticky="e")
        clear_btn = ttk.Button(parent, text="✖", width=3, command=lambda: var.set(""))
        clear_btn.grid(row=row, column=3, sticky="e")
        ToolTip(copy_btn, "Copier le chemin"); ToolTip(clear_btn, "Effacer")
        parent.columnconfigure(1, weight=1)

    def _copy_to_clipboard(self, text: str):
        try: self.winfo_toplevel().clipboard_clear(); self.winfo_toplevel().clipboard_append(text)
        except Exception: pass

    def _select_shapefile(self, shp_type):
        label_text = "Zone d'étude" if shp_type == 'ZE' else "Aire d'étude élargie"
        title = f"Sélectionner le shapefile pour '{label_text}'"
        base_dir = DEFAULT_SHAPE_DIR if os.path.isdir(DEFAULT_SHAPE_DIR) else os.path.expanduser("~")
        path = filedialog.askopenfilename(title=title, initialdir=base_dir, filetypes=[("Shapefile ESRI", "*.shp")])
        if path:
            if shp_type == 'ZE': self.ze_shp_var.set(path)
            else: self.ae_shp_var.set(path)

    def _populate_projects(self):
        for w in list(self.scrollable_frame.children.values()): w.destroy()
        self.project_vars.clear()
        self.all_projects = discover_projects()
        self.filtered_projects = list(self.all_projects)
        if not self.all_projects:
            ttk.Label(self.scrollable_frame, text="Aucun projet trouvé ou dossier inaccessible.", foreground="red").pack(anchor="w")
            return
        for proj_path in self.filtered_projects:
            var = tk.IntVar(value=1); self.project_vars[proj_path] = var
            ttk.Checkbutton(self.scrollable_frame, text=os.path.basename(proj_path), variable=var, style="Card.TCheckbutton")\
                .pack(anchor='w', padx=4, pady=1)

    def _apply_filter(self):
        term = normalize_name(self.filter_var.get())
        for w in list(self.scrollable_frame.children.values()): w.destroy()
        self.filtered_projects = [p for p in self.all_projects if term in normalize_name(os.path.basename(p))]
        if not self.filtered_projects:
            ttk.Label(self.scrollable_frame, text="Aucun projet ne correspond au filtre.", foreground="red").pack(anchor="w")
            self.project_vars = {}; self._update_counts(); return
        for proj_path in self.filtered_projects:
            current = self.project_vars.get(proj_path, tk.IntVar(value=1))
            self.project_vars[proj_path] = current
            ttk.Checkbutton(self.scrollable_frame, text=os.path.basename(proj_path), variable=current, style="Card.TCheckbutton")\
                .pack(anchor='w', padx=4, pady=1)
        self._update_counts()

    def _select_all(self, state: bool):
        for var in self.project_vars.values(): var.set(1 if state else 0)
        self._update_counts()

    def _selected_projects(self) -> List[str]:
        return [p for p, v in self.project_vars.items() if v.get() == 1 and p in self.filtered_projects]

    def _update_counts(self):
        selected = len(self._selected_projects()); total = len(self.filtered_projects)
        self.status_label.config(text=f"Projets sélectionnés : {selected} / {total}")

    def _open_out_dir(self):
        try:
            os.makedirs(OUT_IMG, exist_ok=True); os.startfile(OUT_IMG)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d’ouvrir le dossier de sortie : {e}")

    def _test_qgis_threaded(self):
        t = threading.Thread(target=self._test_qgis); t.daemon = True; t.start()

    def _test_qgis(self):
        try:
            log_with_time("Test QGIS : import…")
            cfg = {"QGIS_ROOT": QGIS_ROOT, "QGIS_APP": QGIS_APP, "PY_VER": PY_VER}
            worker_run(([], cfg))
            log_with_time("Test QGIS : OK")
            messagebox.showinfo("QGIS", "Import QGIS OK.")
        except Exception as e:
            log_with_time(f"Échec import QGIS : {e}")
            messagebox.showerror("QGIS", f"Échec import QGIS : {e}")

    def start_export_thread(self):
        if not self.ze_shp_var.get() or not self.ae_shp_var.get():
            messagebox.showerror("Erreur", "Sélectionnez les deux shapefiles."); return
        if not os.path.isfile(self.ze_shp_var.get()) or not os.path.isfile(self.ae_shp_var.get()):
            messagebox.showerror("Erreur", "Un shapefile est introuvable."); return
        projets = self._selected_projects()
        if not projets:
            messagebox.showerror("Erreur", "Sélectionnez au moins un projet."); return

        self._update_counts()
        self.export_button.config(state="disabled")
        self.progress_done = 0; self.progress["value"] = 0

        mode = self.cadrage_var.get(); per_project = 2 if mode == "BOTH" else 1
        self.total_expected = per_project * len(projets)
        self.progress["maximum"] = max(1, self.total_expected)

        self.prefs.update({
            "ZE_SHP": self.ze_shp_var.get(),
            "AE_SHP": self.ae_shp_var.get(),
            "CADRAGE_MODE": self.cadrage_var.get(),
            "OVERWRITE": bool(self.overwrite_var.get()),
            "DPI": int(self.dpi_var.get()),
            "N_WORKERS": int(self.workers_var.get()),
            "MARGIN_FAC": float(self.margin_var.get()),
        }); save_prefs(self.prefs)

        t = threading.Thread(target=self._run_export_logic, args=(projets,))
        t.daemon = True; t.start()

    def _run_export_logic(self, projets: List[str]):
        try:
            start = datetime.datetime.now()
            os.makedirs(OUT_IMG, exist_ok=True)

            log_with_time(f"{len(projets)} projets (attendu = calcul en cours)")
            log_with_time(f"Workers={self.workers_var.get()}, DPI={self.dpi_var.get()}, marge={self.margin_var.get():.2f}, overwrite={self.overwrite_var.get()}")

            workers = int(self.workers_var.get())
            chunks = chunk_even(projets, workers)
            cfg = {
                "QGIS_ROOT": QGIS_ROOT,
                "QGIS_APP": QGIS_APP,
                "PY_VER": PY_VER,
                "EXPORT_DIR": OUT_IMG,
                "DPI": int(self.dpi_var.get()),
                "MARGIN_FAC": float(self.margin_var.get()),
                "LAYER_AE_NAME": LAYER_AE_NAME,
                "LAYER_ZE_NAME": LAYER_ZE_NAME,
                "AE_SHP": self.ae_shp_var.get(),
                "ZE_SHP": self.ze_shp_var.get(),
                "CADRAGE_MODE": self.cadrage_var.get(),
                "OVERWRITE": bool(self.overwrite_var.get()),
                "EXPORT_TYPE": "PNG",
                "WORKERS": workers,
            }

            ok_total = 0
            ko_total = 0

            def ui_update_progress(done_inc):
                self.progress_done += done_inc
                self.progress["value"] = min(self.progress_done, self.total_expected)
                self.status_label.config(text=f"Progression : {self.progress_done}/{self.total_expected}")

            if workers <= 1:
                for chunk in chunks:
                    ok, ko = worker_run((chunk, cfg))
                    ok_total += ok
                    ko_total += ko
                    self.after(0, ui_update_progress, ok + ko)
                    log_with_time(f"Lot terminé: {ok} OK, {ko} KO")
            else:
                try:
                    import multiprocessing as mp
                    mp.set_start_method("spawn", force=True)
                except Exception:
                    pass
                with ProcessPoolExecutor(max_workers=workers) as ex:
                    futures = [ex.submit(worker_run, (chunk, cfg)) for chunk in chunks if chunk]
                    for fut in as_completed(futures):
                        try:
                            ok, ko = fut.result()
                            ok_total += ok
                            ko_total += ko
                            self.after(0, ui_update_progress, ok + ko)
                            log_with_time(f"Lot terminé: {ok} OK, {ko} KO")
                        except Exception as e:
                            log_with_time(f"Erreur worker: {e}")

            elapsed = datetime.datetime.now() - start
            log_with_time(f"FIN — OK={ok_total} | KO={ko_total} | Attendu={self.total_expected} | Durée={elapsed}")
            self.after(0, lambda: self.status_label.config(text=f"Terminé — OK={ok_total} / KO={ko_total}"))
        except Exception as e:
            log_with_time(f"Erreur critique: {e}")
            self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
        finally:
            self.after(0, lambda: self.export_button.config(state="normal"))

# =========================
# Onglet 3 — Identification Pl@ntNet (UI + logique)
# =========================
class PlantNetTab(ttk.Frame):
    def __init__(self, parent, style_helper: StyleHelper, prefs: dict):
        super().__init__(parent, padding=12)
        self.style_helper = style_helper
        self.prefs = prefs

        self.font_title = tkfont.Font(family="Segoe UI", size=15, weight="bold")
        self.font_sub   = tkfont.Font(family="Segoe UI", size=10)
        self.font_mono  = tkfont.Font(family="Consolas", size=9)

        self.folder_var = tk.StringVar(value=self.prefs.get("PLANTNET_FOLDER", ""))

        self._build_ui()

    def _build_ui(self):
        header = ttk.Frame(self, style="Header.TFrame", padding=(14, 12))
        header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header, text="Identification Pl@ntNet", style="Card.TLabel", font=self.font_title)\
            .grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Analyse un dossier d'images via l'API Pl@ntNet.", style="Subtle.TLabel", font=self.font_sub)\
            .grid(row=1, column=0, sticky="w", pady=(4,0))
        header.columnconfigure(0, weight=1)

        card = ttk.Frame(self, style="Card.TFrame", padding=12)
        card.pack(fill=tk.X)
        ttk.Label(card, text="Dossier d'images", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        row = ttk.Frame(card, style="Card.TFrame")
        row.grid(row=0, column=1, sticky="ew", padx=0)
        row.columnconfigure(0, weight=1)
        ttk.Entry(row, textvariable=self.folder_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(row, text="Parcourir…", command=self._pick_folder).grid(row=0, column=1, padx=(6,0))
        card.columnconfigure(1, weight=1)

        act = ttk.Frame(self, style="Card.TFrame", padding=12)
        act.pack(fill=tk.X, pady=(10,0))
        self.run_btn = ttk.Button(act, text="▶ Lancer l'analyse", style="Accent.TButton", command=self._start_thread)
        self.run_btn.grid(row=0, column=0, sticky="w")
        obtn = ttk.Button(act, text="📂 Ouvrir le dossier de sortie", command=self._open_out_dir)
        obtn.grid(row=0, column=1, padx=(10,0)); ToolTip(obtn, "Ouvrir le dossier cible")

        bottom = ttk.Frame(self, style="Card.TFrame", padding=12)
        bottom.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        self.log_text = tk.Text(bottom, height=12, wrap=tk.WORD, state='disabled',
                                bg=self.style_helper.style.lookup("Card.TFrame", "background"),
                                fg=self.style_helper.style.lookup("TLabel", "foreground"))
        self.log_text.configure(font=self.font_mono, relief="flat")
        log_scroll = ttk.Scrollbar(bottom, orient="vertical", command=self.log_text.yview)
        self.log_text['yscrollcommand'] = log_scroll.set
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.stdout_redirect = TextRedirector(self.log_text)

    def _pick_folder(self):
        base = self.folder_var.get() or os.path.expanduser("~")
        d = filedialog.askdirectory(title="Choisir le dossier d'images", initialdir=base if os.path.isdir(base) else os.path.expanduser("~"))
        if d:
            self.folder_var.set(d)

    def _open_out_dir(self):
        try:
            os.makedirs(OUT_IMG, exist_ok=True)
            os.startfile(OUT_IMG)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d’ouvrir le dossier : {e}")

    def _start_thread(self):
        self.run_btn.config(state="disabled")
        t = threading.Thread(target=self._run_process)
        t.daemon = True
        t.start()

    def _run_process(self):
        folder = self.folder_var.get().strip()
        if not folder:
            print("Veuillez sélectionner un dossier.", file=self.stdout_redirect)
            self.after(0, lambda: self.run_btn.config(state="normal"))
            return
        if not os.path.exists(folder):
            print(f"Le dossier à traiter n'existe pas : {folder}", file=self.stdout_redirect)
            self.after(0, lambda: self.run_btn.config(state="normal"))
            return

        self.prefs["PLANTNET_FOLDER"] = folder
        save_prefs(self.prefs)

        image_extensions = ['.jpg', '.jpeg', '.png', '.heic', '.heif']
        image_files = []
        for root, dirs, files in os.walk(folder):
            for f in files:
                if os.path.splitext(f)[1].lower() in image_extensions:
                    if '@plantnet' not in f:
                        image_files.append(os.path.join(root, f))

        if not image_files:
            print("Aucune image à traiter dans le dossier.", file=self.stdout_redirect)
            self.after(0, lambda: self.run_btn.config(state="normal"))
            return

        plant_name_counts = {}
        old_stdout = sys.stdout
        sys.stdout = self.stdout_redirect
        try:
            for image_path in image_files:
                organ = 'flower'
                plant_name = identify_plant(image_path, organ)
                if plant_name:
                    count = plant_name_counts.get(plant_name, 0) + 1
                    plant_name_counts[plant_name] = count
                    copy_and_rename_file(image_path, folder, plant_name, count)
                else:
                    print(f"Aucune identification possible pour l'image : {image_path}")
            print("Analyse terminée.")
        finally:
            sys.stdout = old_stdout
            self.after(0, lambda: self.run_btn.config(state="normal"))

# =========================
# Onglet 4 — ID contexte éco
# =========================
class IDContexteEcoTab(ttk.Frame):
    def __init__(self, parent, style_helper: StyleHelper, prefs: dict):
        super().__init__(parent, padding=12)
        self.style_helper = style_helper
        self.prefs = prefs

        self.font_title = tkfont.Font(family="Segoe UI", size=15, weight="bold")
        self.font_sub   = tkfont.Font(family="Segoe UI", size=10)
        self.font_mono  = tkfont.Font(family="Consolas", size=9)

        self.ae_var = tk.StringVar()
        self.ze_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        header = ttk.Frame(self, style="Header.TFrame", padding=(14, 12))
        header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header, text="Identification des zonages", style="Card.TLabel", font=self.font_title)\
            .grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Choisissez les shapefiles de référence puis lancez l'analyse.",
                  style="Subtle.TLabel", font=self.font_sub)\
            .grid(row=1, column=0, sticky="w", pady=(4,0))
        header.columnconfigure(0, weight=1)

        card = ttk.Frame(self, style="Card.TFrame", padding=12)
        card.pack(fill=tk.X)
        self._file_row(card, 0, "📁 Aire d'étude élargie…", self.ae_var, self._select_ae)
        self._file_row(card, 1, "📁 Zone d'étude…", self.ze_var, self._select_ze)
        card.columnconfigure(1, weight=1)

        act = ttk.Frame(self, style="Card.TFrame", padding=12)
        act.pack(fill=tk.X, pady=(10,0))
        self.run_btn = ttk.Button(act, text="▶ Lancer l'analyse", style="Accent.TButton", command=self._start_thread)
        self.run_btn.grid(row=0, column=0, sticky="w")

        bottom = ttk.Frame(self, style="Card.TFrame", padding=12)
        bottom.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        self.log_text = tk.Text(bottom, height=12, wrap=tk.WORD, state='disabled',
                                bg=self.style_helper.style.lookup("Card.TFrame", "background"),
                                fg=self.style_helper.style.lookup("TLabel", "foreground"))
        self.log_text.configure(font=self.font_mono, relief="flat")
        log_scroll = ttk.Scrollbar(bottom, orient="vertical", command=self.log_text.yview)
        self.log_text['yscrollcommand'] = log_scroll.set
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.stdout_redirect = TextRedirector(self.log_text)

    def _file_row(self, parent, row: int, label: str, var: tk.StringVar, cmd):
        btn = ttk.Button(parent, text=label, command=cmd)
        btn.grid(row=row, column=0, sticky="w", pady=(8 if row == 0 else 4, 2))
        ent = ttk.Entry(parent, textvariable=var, width=10)
        ent.grid(row=row, column=1, sticky="ew", padx=8)
        ent.configure(state="readonly")
        clear_btn = ttk.Button(parent, text="✖", width=3, command=lambda: var.set(""))
        clear_btn.grid(row=row, column=2, sticky="e")
        parent.columnconfigure(1, weight=1)

    def _select_ae(self):
        base = self.ae_var.get() or os.path.expanduser("~")
        path = filedialog.askopenfilename(title="Sélectionner l'aire d'étude élargie",
                                          initialdir=base if os.path.isdir(base) else os.path.expanduser("~"),
                                          filetypes=[("Shapefile ESRI", "*.shp")])
        if path:
            self.ae_var.set(path)

    def _select_ze(self):
        base = self.ze_var.get() or os.path.expanduser("~")
        path = filedialog.askopenfilename(title="Sélectionner la zone d'étude",
                                          initialdir=base if os.path.isdir(base) else os.path.expanduser("~"),
                                          filetypes=[("Shapefile ESRI", "*.shp")])
        if path:
            self.ze_var.set(path)

    def _start_thread(self):
        self.run_btn.config(state="disabled")
        t = threading.Thread(target=self._run_process)
        t.daemon = True
        t.start()

    def _run_process(self):
        ae = self.ae_var.get().strip()
        ze = self.ze_var.get().strip()
        if not ae or not ze:
            print("Veuillez sélectionner les deux shapefiles.", file=self.stdout_redirect)
            self.after(0, lambda: self.run_btn.config(state="normal"))
            return

        old_stdout = sys.stdout
        sys.stdout = self.stdout_redirect
        try:
            from .id_contexte_eco import run_analysis as run_id_context
            run_id_context(ae, ze)
            print("Analyse terminée.")
        except Exception as e:
            print(f"Erreur: {e}")
        finally:
            sys.stdout = old_stdout
            self.after(0, lambda: self.run_btn.config(state="normal"))

# =========================
# Nouvel onglet « Contexte éco »
# =========================
class ContexteEcoTab(ttk.Frame):
    def __init__(self, parent, style_helper: StyleHelper, prefs: dict):
        super().__init__(parent, padding=12)
        self.style_helper = style_helper
        self.prefs = prefs

        self.font_mono = tkfont.Font(family="Consolas", size=9)

        # Variables partagées
        self.ze_shp_var   = tk.StringVar(value=self.prefs.get("ZE_SHP", ""))
        self.ae_shp_var   = tk.StringVar(value=self.prefs.get("AE_SHP", ""))
        self.cadrage_var   = tk.StringVar(value=self.prefs.get("CADRAGE_MODE", "BOTH"))
        self.overwrite_var = tk.BooleanVar(value=self.prefs.get("OVERWRITE", OVERWRITE_DEFAULT))
        self.dpi_var       = tk.IntVar(value=int(self.prefs.get("DPI", DPI_DEFAULT)))
        self.workers_var   = tk.IntVar(value=int(self.prefs.get("N_WORKERS", N_WORKERS_DEFAULT)))
        self.margin_var    = tk.DoubleVar(value=float(self.prefs.get("MARGIN_FAC", MARGIN_FAC_DEFAULT)))
        self.buffer_var    = tk.DoubleVar(value=float(self.prefs.get("ID_TAMPON_KM", 5.0)))
        self.out_dir_var   = tk.StringVar(value=self.prefs.get("OUT_DIR", OUT_IMG))
        self.export_type_var = tk.StringVar(value=self.prefs.get("EXPORT_TYPE", "BOTH"))

        self.project_vars: dict[str, tk.IntVar] = {}
        self.all_projects: List[str] = []
        self.filtered_projects: List[str] = []
        self.total_expected = 0
        self.progress_done  = 0
        self.busy = False

        self._build_ui()
        self._populate_projects()
        self._update_counts()

    # ---------- Construction UI ----------
    def _build_ui(self):
        # Sélecteurs shapefiles
        shp = ttk.Frame(self, style="Card.TFrame", padding=12)
        shp.pack(fill=tk.X)
        ttk.Label(shp, text="Couches Shapefile", style="Card.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        self._file_row(shp, 1, "📁 Zone d'étude…", self.ze_shp_var, self._select_ze)
        self._file_row(shp, 2, "📁 Aire d'étude élargie…", self.ae_shp_var, self._select_ae)
        shp.columnconfigure(1, weight=1)

        # Encart Export cartes
        exp = ttk.Frame(self, style="Card.TFrame", padding=12)
        exp.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        exp.columnconfigure(0, weight=1); exp.columnconfigure(1, weight=1)

        opt = ttk.Frame(exp)
        opt.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        ttk.Label(opt, text="Cadrage", style="Card.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(opt, text="AE + ZE", variable=self.cadrage_var, value="BOTH", style="Card.TRadiobutton").grid(row=1, column=0, sticky="w")
        ttk.Radiobutton(opt, text="ZE uniquement", variable=self.cadrage_var, value="ZE", style="Card.TRadiobutton").grid(row=1, column=1, sticky="w", padx=(12,0))
        ttk.Radiobutton(opt, text="AE uniquement", variable=self.cadrage_var, value="AE", style="Card.TRadiobutton").grid(row=1, column=2, sticky="w", padx=(12,0))
        ttk.Checkbutton(opt, text="Écraser si le PNG existe", variable=self.overwrite_var, style="Card.TCheckbutton").grid(row=2, column=0, columnspan=3, sticky="w", pady=(6,0))

        ttk.Label(opt, text="DPI", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=(6,0))
        ttk.Spinbox(opt, from_=72, to=1200, textvariable=self.dpi_var, width=6, justify="right").grid(row=3, column=1, sticky="w", pady=(6,0))
        ttk.Label(opt, text="Workers", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=(6,0))
        ttk.Spinbox(opt, from_=1, to=max(1, (os.cpu_count() or 2)), textvariable=self.workers_var, width=6, justify="right").grid(row=4, column=1, sticky="w", pady=(6,0))
        ttk.Label(opt, text="Marge", style="Card.TLabel").grid(row=5, column=0, sticky="w", pady=(6,0))
        ttk.Spinbox(opt, from_=1.00, to=2.00, increment=0.05, textvariable=self.margin_var, width=6, justify="right").grid(row=5, column=1, sticky="w", pady=(6,0))

        ttk.Label(opt, text="Dossier de sortie", style="Card.TLabel").grid(row=6, column=0, sticky="w", pady=(6,0))
        out_row = ttk.Frame(opt)
        out_row.grid(row=6, column=1, columnspan=2, sticky="ew")
        out_row.columnconfigure(0, weight=1)
        ttk.Entry(out_row, textvariable=self.out_dir_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(out_row, text="Parcourir…", command=self._select_out_dir).grid(row=0, column=1, padx=(6,0))

        ttk.Label(opt, text="Exporter", style="Card.TLabel").grid(row=7, column=0, sticky="w", pady=(6,0))
        exp_row = ttk.Frame(opt)
        exp_row.grid(row=7, column=1, columnspan=2, sticky="w")
        ttk.Radiobutton(exp_row, text="PNG + QGIS", variable=self.export_type_var, value="BOTH", style="Card.TRadiobutton").pack(side=tk.LEFT)
        ttk.Radiobutton(exp_row, text="PNG uniquement", variable=self.export_type_var, value="PNG", style="Card.TRadiobutton").pack(side=tk.LEFT, padx=(8,0))
        ttk.Radiobutton(exp_row, text="QGIS uniquement", variable=self.export_type_var, value="QGS", style="Card.TRadiobutton").pack(side=tk.LEFT, padx=(8,0))

        self.export_button = ttk.Button(opt, text="Lancer l’export cartes", style="Accent.TButton", command=self.start_export_thread)
        self.export_button.grid(row=8, column=0, columnspan=3, sticky="w", pady=(10,0))

        proj = ttk.Frame(exp)
        proj.grid(row=0, column=1, sticky="nsew", padx=(8,0))
        ttk.Label(proj, text="Projets QGIS", style="Card.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(proj, text="Filtrer", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=(6,6))
        self.filter_var = tk.StringVar()
        fe = ttk.Entry(proj, textvariable=self.filter_var, width=32)
        fe.grid(row=1, column=1, sticky="w", pady=(6,6))
        fe.bind("<KeyRelease>", lambda _e: self._apply_filter())
        ttk.Button(proj, text="Tout", width=6, command=lambda: self._select_all(True)).grid(row=1, column=2, padx=(8,0))
        ttk.Button(proj, text="Aucun", width=6, command=lambda: self._select_all(False)).grid(row=1, column=3, padx=(6,0))

        canvas = tk.Canvas(proj, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(proj, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=2, column=0, columnspan=4, sticky="nsew", pady=(6, 6))
        scrollbar.grid(row=2, column=4, sticky="ns", padx=(6,0))
        proj.rowconfigure(2, weight=1); proj.columnconfigure(1, weight=1)

        # Encart ID contexte éco
        idf = ttk.Frame(self, style="Card.TFrame", padding=12)
        idf.pack(fill=tk.X, pady=(10,0))
        ttk.Label(idf, text="Tampon ZE (km)", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(idf, from_=0.0, to=50.0, increment=0.5, textvariable=self.buffer_var, width=6, justify="right").grid(row=0, column=1, sticky="w", padx=(8,0))
        self.id_button = ttk.Button(idf, text="Lancer l’ID Contexte éco", style="Accent.TButton", command=self.start_id_thread)
        self.id_button.grid(row=0, column=2, sticky="w", padx=(12,0))
        self.wiki_button = ttk.Button(idf, text="Wikipedia", style="Accent.TButton", command=self.start_wiki_thread)
        self.wiki_button.grid(row=0, column=3, sticky="w", padx=(12,0))
        self.rlt_button = ttk.Button(idf, text="Remonter le temps", style="Accent.TButton", command=self.start_rlt_thread)
        self.rlt_button.grid(row=0, column=4, sticky="w", padx=(12,0))
        self.maps_button = ttk.Button(idf, text="Ouvrir Google Maps", style="Accent.TButton", command=self.open_gmaps)
        self.maps_button.grid(row=0, column=5, sticky="w", padx=(12,0))
        self.bassin_button = ttk.Button(idf, text="Bassin versant", style="Accent.TButton", command=self.start_bassin_thread)
        self.bassin_button.grid(row=0, column=6, sticky="w", padx=(12,0))
        # Section Wikipedia
        wiki = ttk.Frame(self, style="Card.TFrame", padding=12)
        wiki.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        header = ttk.Frame(wiki)
        header.pack(fill=tk.X)
        self.wiki_visible = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            header,
            text="Wikipedia",
            variable=self.wiki_visible,
            command=self._toggle_wiki_section,
            style="Card.TCheckbutton",
        ).pack(anchor="w")
        self.wiki_content = ttk.Frame(wiki)
        self.wiki_content.pack(fill=tk.BOTH, expand=True, pady=(8,0))
        self._init_wiki_table(self.wiki_content)

        # Console + progression
        bottom = ttk.Frame(self, style="Card.TFrame", padding=12)
        bottom.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        self.status_label = ttk.Label(bottom, text="Prêt.", style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(bottom, orient="horizontal", mode="determinate", length=220)
        self.progress.grid(row=0, column=1, sticky="e")
        bottom.columnconfigure(0, weight=1)

        log_frame = ttk.Frame(bottom, style="Card.TFrame")
        log_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8,0))
        bottom.rowconfigure(1, weight=1)
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state='disabled',
                                bg=self.style_helper.style.lookup("Card.TFrame", "background"),
                                fg=self.style_helper.style.lookup("TLabel", "foreground"))
        self.log_text.configure(font=self.font_mono, relief="flat")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text['yscrollcommand'] = log_scroll.set
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.stdout_redirect = TextRedirector(self.log_text)

        obtn = ttk.Button(bottom, text="Ouvrire output", command=self._open_out_dir)
        obtn.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8,0))

    # ---------- Helpers UI ----------
    def _file_row(self, parent, row: int, label: str, var: tk.StringVar, cmd):
        btn = ttk.Button(parent, text=label, command=cmd)
        btn.grid(row=row, column=0, sticky="w", pady=(8 if row == 1 else 4, 2))
        ent = ttk.Entry(parent, textvariable=var, width=10)
        ent.grid(row=row, column=1, sticky="ew", padx=8)
        ent.configure(state="readonly")
        clear_btn = ttk.Button(parent, text="✖", width=3, command=lambda: var.set(""))
        clear_btn.grid(row=row, column=2, sticky="e")
        parent.columnconfigure(1, weight=1)

    def _select_ze(self):
        base = self.ze_shp_var.get() or os.path.expanduser("~")
        path = filedialog.askopenfilename(title="Sélectionner la zone d'étude",
                                          initialdir=base if os.path.isdir(base) else os.path.expanduser("~"),
                                          filetypes=[("Shapefile ESRI", "*.shp")])
        if path:
            # Normaliser pour gérer les chemins réseau ou trop longs
            self.ze_shp_var.set(to_long_unc(os.path.normpath(path)))

    def _select_ae(self):
        base = self.ae_shp_var.get() or os.path.expanduser("~")
        path = filedialog.askopenfilename(title="Sélectionner l'aire d'étude élargie",
                                          initialdir=base if os.path.isdir(base) else os.path.expanduser("~"),
                                          filetypes=[("Shapefile ESRI", "*.shp")])
        if path:
            # Normaliser pour gérer les chemins réseau ou trop longs
            self.ae_shp_var.set(to_long_unc(os.path.normpath(path)))

    def _select_out_dir(self):
        base = self.out_dir_var.get() or OUT_IMG
        d = filedialog.askdirectory(title="Choisir le dossier de sortie",
                                    initialdir=base if os.path.isdir(base) else os.path.expanduser("~"))
        if d:
            self.out_dir_var.set(d)

    def _open_out_dir(self):
        try:
            out_dir = self.out_dir_var.get() or OUT_IMG
            os.makedirs(out_dir, exist_ok=True)
            os.startfile(out_dir)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d’ouvrir le dossier : {e}")

    # ---------- Section Wikipedia ----------
    def _init_wiki_table(self, parent: ttk.Frame) -> None:
        tbl = ttk.Frame(parent)
        tbl.pack(fill=tk.BOTH, expand=True)

        ttk.Label(tbl, text="Climat", style="Card.TLabel").grid(row=0, column=0, sticky="nw")
        self.wiki_climat_text = tk.Text(tbl, height=5, wrap=tk.WORD, state="disabled")
        clim_scroll = ttk.Scrollbar(tbl, orient="vertical", command=self.wiki_climat_text.yview)
        self.wiki_climat_text.configure(yscrollcommand=clim_scroll.set)
        self.wiki_climat_text.grid(row=0, column=1, sticky="nsew")
        clim_scroll.grid(row=0, column=2, sticky="ns")

        ttk.Label(tbl, text="Corine Land Cover", style="Card.TLabel").grid(row=1, column=0, sticky="nw", pady=(6,0))
        self.wiki_corine_text = tk.Text(tbl, height=5, wrap=tk.WORD, state="disabled")
        cor_scroll = ttk.Scrollbar(tbl, orient="vertical", command=self.wiki_corine_text.yview)
        self.wiki_corine_text.configure(yscrollcommand=cor_scroll.set)
        self.wiki_corine_text.grid(row=1, column=1, sticky="nsew", pady=(6,0))
        cor_scroll.grid(row=1, column=2, sticky="ns", pady=(6,0))

        tbl.columnconfigure(1, weight=1)

    def _toggle_wiki_section(self):
        if self.wiki_visible.get():
            self.wiki_content.pack(fill=tk.BOTH, expand=True, pady=(8,0))
        else:
            self.wiki_content.forget()

    def render_wikipedia_section(self, data: dict) -> None:
        self.wiki_climat_text.config(state="normal")
        self.wiki_climat_text.delete("1.0", tk.END)
        self.wiki_climat_text.insert(tk.END, data.get("climat", ""))
        self.wiki_climat_text.config(state="disabled")

        self.wiki_corine_text.config(state="normal")
        self.wiki_corine_text.delete("1.0", tk.END)
        self.wiki_corine_text.insert(tk.END, data.get("corine", ""))
        self.wiki_corine_text.config(state="disabled")

    def start_wiki_thread(self):
        if not self.ze_shp_var.get().strip():
            messagebox.showerror("Erreur", "Sélectionner la Zone d'étude.")
            return
        self.wiki_button.config(state="disabled")
        t = threading.Thread(target=self._run_wiki)
        t.daemon = True
        t.start()

    def _run_wiki(self):
        try:
            ze_path = self.ze_shp_var.get()
            gdf = gpd.read_file(ze_path)
            if gdf.crs is None:
                raise ValueError("CRS non défini")
            gdf = gdf.to_crs("EPSG:4326")
            centroid = gdf.geometry.unary_union.centroid
            lat, lon = centroid.y, centroid.x
            commune, dep = self._detect_commune(lat, lon)
            commune_label = f"{commune} ({dep})"
            print(f"[Wiki] Requête : {commune_label}", file=self.stdout_redirect)
            data = run_wikipedia_scrape(commune_label)
            print(f"[Wiki] Page Wikipédia : {data['url']}", file=self.stdout_redirect)
            print("[Wiki] CLIMAT :", file=self.stdout_redirect)
            print(data.get("climat", "Donnée non disponible"), file=self.stdout_redirect)
            print("[Wiki] OCCUPATION DES SOLS :", file=self.stdout_redirect)
            print(data.get("corine", "Donnée non disponible"), file=self.stdout_redirect)
            self.after(0, lambda d=data: self.render_wikipedia_section(d))
        except Exception as e:
            print(f"[Wiki] Erreur : {e}", file=self.stdout_redirect)
        finally:
            self.after(0, lambda: self.wiki_button.config(state="normal"))

    # --- Boutons ajoutés ---
    def start_rlt_thread(self):
        if not self.ze_shp_var.get().strip():
            messagebox.showerror("Erreur", "Sélectionner la Zone d'étude.")
            return
        self.rlt_button.config(state="disabled")
        t = threading.Thread(target=self._run_rlt)
        t.daemon = True
        t.start()

    def start_bassin_thread(self):
        if not self.ze_shp_var.get().strip():
            messagebox.showerror("Erreur", "Sélectionner la Zone d'étude.")
            return
        self.bassin_button.config(state="disabled")
        t = threading.Thread(target=self._run_bassin)
        t.daemon = True
        t.start()

    def open_gmaps(self):
        if not self.ze_shp_var.get().strip():
            messagebox.showerror("Erreur", "Sélectionner la Zone d'étude.")
            return
        try:
            gdf = gpd.read_file(self.ze_shp_var.get())
            if gdf.crs is None:
                raise ValueError("CRS non défini")
            gdf = gdf.to_crs("EPSG:4326")
            centroid = gdf.geometry.unary_union.centroid
            lat, lon = centroid.y, centroid.x
            url = f"https://www.google.com/maps/@{lat},{lon},17z"
            print(f"[Maps] {url}", file=self.stdout_redirect)
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d’ouvrir Google Maps : {e}")

    def _run_rlt(self):
        try:
            ze_path = self.ze_shp_var.get()
            gdf = gpd.read_file(ze_path)
            if gdf.crs is None:
                raise ValueError("CRS non défini")
            gdf = gdf.to_crs("EPSG:4326")
            centroid = gdf.geometry.unary_union.centroid
            lat_dd, lon_dd = centroid.y, centroid.x
            commune, _dep = self._detect_commune(lat_dd, lon_dd)
            wait_s = WAIT_TILES_DEFAULT
            out_dir = OUTPUT_DIR_RLT
            os.makedirs(out_dir, exist_ok=True)
            comment_txt = COMMENT_TEMPLATE.format(commune=commune)

            drv_opts = webdriver.ChromeOptions()
            drv_opts.add_argument("--log-level=3")
            drv_opts.add_experimental_option('excludeSwitches', ['enable-logging'])
            drv_opts.add_argument("--disable-extensions")

            print(f"[IGN] Lancement Chrome…", file=self.stdout_redirect)
            driver = webdriver.Chrome(options=drv_opts)
            try:
                driver.maximize_window()
            except Exception:
                pass

            images = []
            viewport = (By.CSS_SELECTOR, "div.ol-viewport")
            for title, layer_val in LAYERS:
                url = URL.format(lon=f"{lon_dd:.6f}", lat=f"{lat_dd:.6f}", layer=layer_val)
                print(f"[IGN] {title} → {url}", file=self.stdout_redirect)
                driver.get(url)
                WebDriverWait(driver, 20).until(EC.visibility_of_element_located(viewport))
                time.sleep(wait_s)
                tgt = driver.find_element(*viewport)
                img_path = os.path.join(out_dir, f"{title}.png")
                if tgt.screenshot(img_path):
                    img = Image.open(img_path)
                    w, h = img.size
                    left, right = int(w * 0.05), int(w * 0.95)
                    img.crop((left, 0, right, h)).save(img_path)
                    images.append((title, img_path))
                    print(f"[IGN] Capture OK : {img_path}", file=self.stdout_redirect)
                else:
                    print(f"[IGN] Capture échouée : {title}", file=self.stdout_redirect)

            driver.quit()

            if not images:
                print("[IGN] Aucune image → pas de doc.", file=self.stdout_redirect)
                messagebox.showwarning("IGN", "Aucune image capturée.")
                return

            print("[IGN] Génération du Word…", file=self.stdout_redirect)
            doc = Document()
            style_normal = doc.styles['Normal']
            style_normal.font.name = 'Calibri'
            style_normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
            style_tbl = doc.styles['Table Grid']
            style_tbl.font.name = 'Calibri'
            style_tbl._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')

            sec = doc.sections[0]
            sec.orientation = WD_ORIENT.LANDSCAPE
            sec.page_width, sec.page_height = sec.page_height, sec.page_width
            for m in (sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin):
                m = Cm(1.5)

            cap_par = doc.add_paragraph()
            cap_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_hyperlink(cap_par, "https://remonterletemps.ign.fr/",
                          f"Comparaison temporelle — {commune} (source : IGN – RemonterLeTemps)")

            table = doc.add_table(rows=2, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = "Table Grid"
            table.autofit = False

            for idx, (title, path) in enumerate(images):
                r, c = divmod(idx, 2)
                cell = table.cell(r, c)
                p_t = cell.paragraphs[0]
                run_t = p_t.add_run(title); run_t.bold = True
                p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.add_paragraph()
                p_img = cell.add_paragraph()
                if os.path.exists(path):
                    p_img.add_run().add_picture(path, width=IMG_WIDTH)
                else:
                    p_img.add_run(f"[image manquante : {title}]")
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph()
            p_comm = doc.add_paragraph(comment_txt)
            p_comm.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            doc_path = os.path.join(out_dir, WORD_FILENAME)
            doc.save(doc_path)
            print(f"[IGN] Document généré : {doc_path}", file=self.stdout_redirect)
            self._set_status(f"Terminé — {doc_path}")
        except Exception as e:
            print(f"[IGN] Erreur : {e}", file=self.stdout_redirect)
            messagebox.showerror("IGN", str(e))
        finally:
            self.after(0, lambda: self.rlt_button.config(state="normal"))

    def _run_bassin(self):
        try:
            ze_path = self.ze_shp_var.get()
            gdf = gpd.read_file(ze_path)
            if gdf.crs is None:
                raise ValueError("CRS non défini")
            gdf = gdf.to_crs("EPSG:4326")
            centroid = gdf.geometry.unary_union.centroid
            lat_dd, lon_dd = centroid.y, centroid.x
            user_address = dd_to_dms(lat_dd, lon_dd)
            download_dir = OUT_IMG
            target_folder_name = "Bassin versant"
            target_path = os.path.join(download_dir, target_folder_name)
            os.makedirs(download_dir, exist_ok=True)
            print(f"[BV] Coordonnées : {lat_dd:.6f}, {lon_dd:.6f}", file=self.stdout_redirect)

            options = webdriver.ChromeOptions()
            options.add_argument("--log-level=3")
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_argument("--disable-extensions")
            prefs = {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "profile.default_content_settings.popups": 0,
                "download.directory_upgrade": True
            }
            options.add_experimental_option("prefs", prefs)

            print("[BV] Initialisation du navigateur...", file=self.stdout_redirect)
            driver = webdriver.Chrome(options=options)
            try:
                driver.maximize_window()
            except Exception:
                pass

            url = "https://mghydro.com/watersheds/"
            print(f"[BV] Navigation vers {url}...", file=self.stdout_redirect)
            driver.get(url)
            wait = WebDriverWait(driver, 2)

            opts_button = wait.until(EC.element_to_be_clickable((By.ID, "opts_click")))
            opts_button.click(); time.sleep(1)
            downloadable_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "downloadable")))
            if not downloadable_checkbox.is_selected():
                downloadable_checkbox.click()

            search_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".leaflet-control-search .search-button")))
            search_icon.click()
            search_input = wait.until(EC.visibility_of_element_located((By.ID, "searchtext84")))
            search_input.clear(); time.sleep(0.3)
            search_input.send_keys(user_address); time.sleep(1.5)
            search_input.send_keys(Keys.ARROW_DOWN); time.sleep(0.3)
            search_input.send_keys(Keys.ENTER); time.sleep(1.5)

            map_element = wait.until(EC.presence_of_element_located((By.ID, "map")))
            ActionChains(driver).move_to_element(map_element).click().perform(); time.sleep(0.8)
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.leaflet-popup")))
            delineate_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".leaflet-popup .gobutton")))
            delineate_button.click(); time.sleep(1.5)

            downloads_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'ui-selectmenu-text') and contains(text(), 'Watershed Boundary')]")))
            downloads_button.click(); time.sleep(0.8)
            ActionChains(driver).send_keys(Keys.ARROW_DOWN).pause(0.3).send_keys(Keys.ARROW_DOWN).pause(0.3).send_keys(Keys.ENTER).perform()
            time.sleep(1.5)

        except Exception as e:
            print(f"[BV] Erreur Selenium : {e}", file=self.stdout_redirect)
        finally:
            try:
                driver.quit()
            except Exception:
                pass

        try:
            print("[BV] Attente du téléchargement du ZIP...", file=self.stdout_redirect)
            zip_file_path = None
            wait_time = 30
            start_time = time.time()
            while time.time() - start_time < wait_time:
                zip_candidates = [f for f in os.listdir(download_dir) if f.lower().endswith('.zip') and not f.lower().endswith('.crdownload')]
                if zip_candidates:
                    zip_candidates_full = [os.path.join(download_dir, z) for z in zip_candidates]
                    zip_file_path = max(zip_candidates_full, key=os.path.getmtime)
                    size1 = os.path.getsize(zip_file_path); time.sleep(1); size2 = os.path.getsize(zip_file_path)
                    if size1 == size2:
                        print(f"[BV] ZIP détecté : {os.path.basename(zip_file_path)}", file=self.stdout_redirect)
                        break
                time.sleep(1)

            if not zip_file_path:
                print("[BV] Aucun fichier ZIP trouvé.", file=self.stdout_redirect)
            else:
                if os.path.exists(target_path):
                    print(f"[BV] Remplacement du dossier '{target_folder_name}'", file=self.stdout_redirect)
                    shutil.rmtree(target_path, ignore_errors=True)
                os.makedirs(target_path, exist_ok=True)
                with zipfile.ZipFile(zip_file_path, 'r') as zf:
                    zf.extractall(path=target_path)
                os.remove(zip_file_path)
                print(f"[BV] Décompression terminée dans '{target_folder_name}'", file=self.stdout_redirect)
                self._set_status(f"Bassin versant : {target_path}")
        except Exception as e:
            print(f"[BV] Erreur décompression : {e}", file=self.stdout_redirect)
        finally:
            self.after(0, lambda: self.bassin_button.config(state="normal"))

    def _set_status(self, txt: str):
        self.after(0, lambda: self.status_label.config(text=txt))

    def _detect_commune(self, lat: float, lon: float) -> Tuple[str, str]:
        try:
            url = ("https://nominatim.openstreetmap.org/reverse?format=json"
                   f"&lat={lat}&lon={lon}&zoom=10&addressdetails=1")
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.load(resp)
            addr = data.get("address", {})
            commune = "Inconnue"
            for key in ("city", "town", "village", "municipality"):
                if key in addr:
                    commune = addr[key]
                    break
            dept_name = addr.get("county", "")
            dep_rev = {v: k for k, v in DEP.items()}
            dep = dep_rev.get(dept_name, "")
            if not dep:
                postcode = addr.get("postcode", "")
                if len(postcode) >= 2:
                    dep = postcode[:2]
            return commune, dep
        except Exception as e:
            print(f"[Wiki] Détection commune échouée : {e}", file=self.stdout_redirect)
            return "Inconnue", ""

    # ---------- Gestion projets QGIS ----------
    def _populate_projects(self):
        for w in list(self.scrollable_frame.children.values()): w.destroy()
        self.project_vars = {}
        self.all_projects = discover_projects()
        self.filtered_projects = list(self.all_projects)
        if not self.all_projects:
            ttk.Label(self.scrollable_frame, text="Aucun projet trouvé ou dossier inaccessible.", foreground="red").pack(anchor="w")
            return
        for proj_path in self.filtered_projects:
            var = tk.IntVar(value=1); self.project_vars[proj_path] = var
            ttk.Checkbutton(self.scrollable_frame, text=os.path.basename(proj_path), variable=var, style="Card.TCheckbutton").pack(anchor='w', padx=4, pady=1)

    def _apply_filter(self):
        term = normalize_name(self.filter_var.get())
        for w in list(self.scrollable_frame.children.values()): w.destroy()
        self.filtered_projects = [p for p in self.all_projects if term in normalize_name(os.path.basename(p))]
        if not self.filtered_projects:
            ttk.Label(self.scrollable_frame, text="Aucun projet ne correspond au filtre.", foreground="red").pack(anchor="w")
            self.project_vars = {}; self._update_counts(); return
        for proj_path in self.filtered_projects:
            current = self.project_vars.get(proj_path, tk.IntVar(value=1))
            self.project_vars[proj_path] = current
            ttk.Checkbutton(self.scrollable_frame, text=os.path.basename(proj_path), variable=current, style="Card.TCheckbutton").pack(anchor='w', padx=4, pady=1)
        self._update_counts()

    def _select_all(self, state: bool):
        for var in self.project_vars.values(): var.set(1 if state else 0)
        self._update_counts()

    def _selected_projects(self) -> List[str]:
        return [p for p, v in self.project_vars.items() if v.get() == 1 and p in self.filtered_projects]

    def _update_counts(self):
        selected = len(self._selected_projects()); total = len(self.filtered_projects)
        self.status_label.config(text=f"Projets sélectionnés : {selected} / {total}")

    # ---------- Lancement export ----------
    def start_export_thread(self):
        if self.busy:
            print("Une action est déjà en cours.", file=self.stdout_redirect)
            return
        if not self.ze_shp_var.get() or not self.ae_shp_var.get():
            messagebox.showerror("Erreur", "Sélectionnez les deux shapefiles."); return
        if not os.path.isfile(self.ze_shp_var.get()) or not os.path.isfile(self.ae_shp_var.get()):
            messagebox.showerror("Erreur", "Un shapefile est introuvable."); return
        projets = self._selected_projects()
        if not projets:
            messagebox.showerror("Erreur", "Sélectionnez au moins un projet."); return

        self.busy = True
        self.export_button.config(state="disabled")
        self.id_button.config(state="disabled")
        mode = self.cadrage_var.get()
        exp_type = self.export_type_var.get()
        png_exports = 2 if (exp_type in ("PNG", "BOTH") and mode == "BOTH") else (1 if exp_type in ("PNG", "BOTH") else 0)
        qgs_exports = 1 if exp_type in ("QGS", "BOTH") else 0
        per_project = png_exports + qgs_exports
        self.total_expected = per_project * len(projets)
        self.progress_done = 0
        self.progress.config(mode="determinate", maximum=max(1, self.total_expected), value=0)
        self.status_label.config(text=f"Progression : 0/{self.total_expected}")

        self.prefs.update({
            "ZE_SHP": self.ze_shp_var.get(),
            "AE_SHP": self.ae_shp_var.get(),
            "CADRAGE_MODE": self.cadrage_var.get(),
            "OVERWRITE": bool(self.overwrite_var.get()),
            "DPI": int(self.dpi_var.get()),
            "N_WORKERS": int(self.workers_var.get()),
            "MARGIN_FAC": float(self.margin_var.get()),
            "OUT_DIR": self.out_dir_var.get(),
            "EXPORT_TYPE": exp_type,
        }); save_prefs(self.prefs)

        t = threading.Thread(target=self._run_export_logic, args=(projets,), daemon=True)
        t.start()

    def _run_export_logic(self, projets: List[str]):
        old_stdout = sys.stdout
        sys.stdout = self.stdout_redirect
        try:
            start = datetime.datetime.now()
            out_dir = self.out_dir_var.get() or OUT_IMG
            os.makedirs(out_dir, exist_ok=True)
            log_with_time(f"{len(projets)} projets (attendu = calcul en cours)")
            log_with_time(f"Workers={self.workers_var.get()}, DPI={self.dpi_var.get()}, marge={self.margin_var.get():.2f}, overwrite={self.overwrite_var.get()}")
            workers = int(self.workers_var.get())
            chunks = chunk_even(projets, workers)
            cfg = {
                "QGIS_ROOT": QGIS_ROOT,
                "QGIS_APP": QGIS_APP,
                "PY_VER": PY_VER,
                "EXPORT_DIR": out_dir,
                "DPI": int(self.dpi_var.get()),
                "MARGIN_FAC": float(self.margin_var.get()),
                "LAYER_AE_NAME": LAYER_AE_NAME,
                "LAYER_ZE_NAME": LAYER_ZE_NAME,
                "AE_SHP": self.ae_shp_var.get(),
                "ZE_SHP": self.ze_shp_var.get(),
                "CADRAGE_MODE": self.cadrage_var.get(),
                "OVERWRITE": bool(self.overwrite_var.get()),
                "EXPORT_TYPE": self.export_type_var.get(),
                "WORKERS": workers,
            }
            ok_total = 0
            ko_total = 0
            def ui_update_progress(done_inc):
                self.progress_done += done_inc
                self.progress["value"] = min(self.progress_done, self.total_expected)
                self.status_label.config(text=f"Progression : {self.progress_done}/{self.total_expected}")
            if workers <= 1:
                for chunk in chunks:
                    ok, ko = worker_run((chunk, cfg))
                    ok_total += ok
                    ko_total += ko
                    self.after(0, ui_update_progress, ok + ko)
                    log_with_time(f"Lot terminé: {ok} OK, {ko} KO")
            else:
                try:
                    import multiprocessing as mp
                    mp.set_start_method("spawn", force=True)
                except Exception:
                    pass
                with ProcessPoolExecutor(max_workers=workers) as ex:
                    futures = [ex.submit(worker_run, (chunk, cfg)) for chunk in chunks if chunk]
                    for fut in as_completed(futures):
                        try:
                            ok, ko = fut.result()
                            ok_total += ok
                            ko_total += ko
                            self.after(0, ui_update_progress, ok + ko)
                            log_with_time(f"Lot terminé: {ok} OK, {ko} KO")
                        except Exception as e:
                            log_with_time(f"Erreur worker: {e}")
            elapsed = datetime.datetime.now() - start
            log_with_time(f"FIN — OK={ok_total} | KO={ko_total} | Attendu={self.total_expected} | Durée={elapsed}")
            self.after(0, lambda: self.status_label.config(text=f"Terminé — OK={ok_total} / KO={ko_total}"))
        except Exception as e:
            log_with_time(f"Erreur critique: {e}")
            self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
        finally:
            sys.stdout = old_stdout
            self.after(0, self._run_finished)

    # ---------- Lancement ID contexte ----------
    def start_id_thread(self):
        if self.busy:
            print("Une action est déjà en cours.", file=self.stdout_redirect)
            return
        ae = to_long_unc(os.path.normpath(self.ae_shp_var.get().strip()))
        ze = to_long_unc(os.path.normpath(self.ze_shp_var.get().strip()))
        if not ae or not ze:
            messagebox.showerror("Erreur", "Sélectionnez les deux shapefiles."); return
        if not os.path.isfile(ae) or not os.path.isfile(ze):
            messagebox.showerror("Erreur", "Un shapefile est introuvable."); return

        self.busy = True
        self.export_button.config(state="disabled")
        self.id_button.config(state="disabled")
        self.progress.config(mode="indeterminate")
        self.progress.start()
        self.status_label.config(text="Analyse en cours…")

        self.prefs.update({
            "ZE_SHP": ze,
            "AE_SHP": ae,
            "ID_TAMPON_KM": float(self.buffer_var.get()),
        }); save_prefs(self.prefs)

        t = threading.Thread(target=self._run_id_logic, args=(ae, ze, float(self.buffer_var.get())), daemon=True)
        t.start()

    def _run_id_logic(self, ae: str, ze: str, buffer_km: float):
        old_stdout = sys.stdout
        sys.stdout = self.stdout_redirect
        try:
            from .id_contexte_eco import run_analysis as run_id_context
            run_id_context(ae, ze, buffer_km)
            log_with_time("Analyse terminée.")
            self.after(0, lambda: self.status_label.config(text="Terminé"))
        except Exception as e:
            log_with_time(f"Erreur: {e}")
            self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
        finally:
            sys.stdout = old_stdout
            self.after(0, lambda: (self.progress.stop(), self.progress.config(mode="determinate", value=0)))
            self.after(0, self._run_finished)

    def _run_finished(self):
        self.export_button.config(state="normal")
        self.id_button.config(state="normal")
        self.busy = False

# =========================
# App principale avec Notebook
# =========================
class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Contexte éco — Outils")
        self.root.geometry("1060x760"); self.root.minsize(900, 640)

        self.prefs = load_prefs()
        self.style_helper = StyleHelper(root, self.prefs)
        self.theme_var = tk.StringVar(value=self.prefs.get("theme", "light"))
        self.style_helper.apply(self.theme_var.get())

        self.wiki_driver = None

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
        self.tab_plant = PlantNetTab(nb, self.style_helper, self.prefs)

        nb.add(self.tab_ctx, text="Contexte éco")
        nb.add(self.tab_plant, text="Pl@ntNet")

        # Raccourcis utiles
        root.bind("<Control-1>", lambda _e: nb.select(0))
        root.bind("<Control-2>", lambda _e: nb.select(1))

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

# =========================
# Main
# =========================
def launch():
    """Lance l'interface principale."""
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()


if __name__ == "__main__":
    launch()
