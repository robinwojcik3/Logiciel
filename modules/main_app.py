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

chromedriver dans PATH.

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

import webbrowser
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx.shared import Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

from PIL import Image

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from tkinter import font as tkfont

from typing import List, Optional, Tuple

from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed

import requests

from io import BytesIO

import pillow_heif

import zipfile

import traceback

import subprocess

import geopandas as gpd
import pandas as pd
from bs4 import BeautifulSoup

# --- Helper functions for Word document generation ---
def dms_to_dd(text: str) -> float:
    pat = r'(\d{1,3})[°d]\s*(\d{1,2})[\'m]\s*([\d\.]+)[\"s]?\s*([NSEW])'
    alt = r"(\d{1,3})\s+(\d{1,2})\s+([\d\.]+)\s*([NSEW])"
    m = re.search(pat, text, re.I) or re.search(alt, text, re.I)
    if not m:
        raise ValueError(f"Format DMS invalide : {text}")
    deg, mn, sc, hemi = m.groups()
    dd = float(deg) + float(mn)/60 + float(sc)/3600
    return -dd if hemi.upper() in ("S", "W") else dd

def add_hyperlink(paragraph, url: str, text: str, italic: bool = True):
    """
    Ajoute un lien hypertexte cliquable (python-docx ne fournit pas
    d’API dédiée ; on passe par l’XML bas niveau).
    """
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




# ==== Imports supplémentaires pour l'onglet Contexte éco ====

# Note: geopandas n'est pas utilisé directement dans ce module.

# Les traitements géospatiaux sont effectués dans des modules dédiés

# (ex: id_contexte_eco) afin d'éviter de charger des dépendances lourdes

# au lancement de l'UI principale.



# Assurer que le dossier racine du projet est dans sys.path quand ce fichier est exécuté directement
try:
    _THIS_DIR = os.path.dirname(__file__)
    _PROJECT_ROOT = os.path.abspath(os.path.join(_THIS_DIR, '..'))  # .. = dossier Logiciel/
    if _PROJECT_ROOT not in sys.path:
        sys.path.insert(0, _PROJECT_ROOT)
except Exception:
    pass


# Import du scraper Wikipédia

try:
    # When run as package (via Start.py)
    from .wikipedia_scraper import DEP, get_wikipedia_extracts
except Exception:
    # Fallback when running this file directly
    from modules.wikipedia_scraper import DEP, get_wikipedia_extracts
  
  
  # Import du worker QGIS externalisé
 
try:
    from .export_worker import worker_run
except Exception:
    from modules.export_worker import worker_run
  
  
  
  
  

# ==== Imports spécifiques onglet 2 (gardés en tête de fichier comme le script source) ====

from selenium import webdriver

from selenium.webdriver.chrome.service import Service

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
from docx.oxml import OxmlElement



from PIL import Image

from bs4 import BeautifulSoup



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



OUT_IMG    = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'output'))



# Dossier par défaut pour la sélection des shapefiles (onglet 1)

DEFAULT_SHAPE_DIR = r"C:\Users\utilisateur\Mon Drive\1 - Bota & Travail\+++++++++  BOTA  +++++++++\---------------------- 2) CARTO terrain"



# QGIS

QGIS_ROOT = r"C:\Program Files\QGIS 3.40.3"

QGIS_APP  = os.path.join(QGIS_ROOT, "apps", "qgis")

PY_VER    = "Python312"

# --- Global Paths & Constants ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, '..')) 
PREFS_PATH = os.path.join(BASE_DIR, "prefs.json")
WAIT_TILES_DEFAULT = 0.5



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


def from_long_unc(path: str) -> str:

    r"""Convertit un chemin Windows étendu (\\?\ ou \\?\UNC) en chemin standard.

    - \\?\UNC\server\share\... -> \\server\share\...
    - \\?\C:\path -> C:\path
    """

    p = path or ""

    if p.startswith("\\\\?\\UNC"):
        return "\\\\" + p[8:]

    if p.startswith("\\\\?\\"):
        return p[4:]

    return p



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



def qgis_multiprocessing_ok() -> bool:

    """Quick check: can QGIS's Python import _multiprocessing? Avoids noisy spawn failures."""

    try:

        qgis_py = os.path.join(QGIS_ROOT, "apps", PY_VER, "python.exe")

        if not os.path.isfile(qgis_py):

            return False

        env = os.environ.copy()

        # Prefer clean env and explicit paths

        qgis_py_root = os.path.join(QGIS_ROOT, "apps", PY_VER)

        qgis_lib = os.path.join(qgis_py_root, "Lib")

        qgis_dlls = os.path.join(qgis_py_root, "DLLs")

        qgis_site = os.path.join(qgis_lib, "site-packages")

        qgis_app_py = os.path.join(QGIS_APP, "python")

        env.pop("PYTHONHOME", None); env.pop("PYTHONPATH", None); env["PYTHONNOUSERSITE"] = "1"

        env["PYTHONHOME"] = qgis_py_root

        env["PYTHONPATH"] = os.pathsep.join([qgis_py_root, qgis_lib, qgis_dlls, qgis_site, qgis_app_py])

        code = "import multiprocessing, multiprocessing.connection, _multiprocessing; print('OK')"

        r = subprocess.run([qgis_py, "-c", code], stdout=subprocess.PIPE, stderr=subprocess.STDOUT,

                           text=True, timeout=6, env=env, cwd=os.path.dirname(__file__))

        return r.returncode == 0 and (r.stdout or "").strip().endswith("OK")

    except Exception:

        return False



def run_worker_subprocess(projects: List[str], cfg: dict) -> tuple[int, int]:

    """Exécute un lot via QGIS Python dans un sous-processus.



    Cette approche évite totalement multiprocessing dans l'interpréteur QGIS,

    supprimant l'erreur `_multiprocessing`.

    """

    import json as _json

    tmp = None

    try:

        qgis_py = os.path.join(QGIS_ROOT, "apps", PY_VER, "python.exe")

        if not os.path.isfile(qgis_py):

            raise RuntimeError(f"Python QGIS introuvable: {qgis_py}")



        repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))

        qgis_py_root = os.path.join(QGIS_ROOT, "apps", PY_VER)

        qgis_lib  = os.path.join(qgis_py_root, "Lib")

        qgis_dlls = os.path.join(qgis_py_root, "DLLs")

        qgis_site = os.path.join(qgis_lib, "site-packages")

        qgis_app_py = os.path.join(QGIS_APP, "python")



        data = {"projects": list(projects or []), "cfg": dict(cfg or {})}

        fd, tmp = tempfile.mkstemp(prefix="qgis_worker_", suffix=".json")

        os.close(fd)

        with open(tmp, "w", encoding="utf-8") as f:

            _json.dump(data, f)



        env = os.environ.copy()

        for k in ("PYTHONPATH", "PYTHONHOME", "PYTHONSTARTUP"):

            env.pop(k, None)

        env["PYTHONNOUSERSITE"] = "1"

        env["PYTHONHOME"] = qgis_py_root

        env["PYTHONPATH"] = os.pathsep.join([repo_root, qgis_py_root, qgis_lib, qgis_dlls, qgis_site, qgis_app_py])



        cmd = [qgis_py, "-m", "modules.export_worker_cli", tmp]

        r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, env=env)

        out = (r.stdout or "").strip()

        if r.returncode == 0 and "," in out:

            try:

                last = out.splitlines()[-1]

                s_ok, s_ko = last.split(",", 1)

                return int(s_ok), int(s_ko)

            except Exception:

                pass

        log_with_time(f"Worker subproc KO (code={r.returncode}): {out[:400]}")

        return 0, len(projects or [])

    finally:

        if tmp and os.path.isfile(tmp):

            try:

                os.remove(tmp)

            except Exception:

                pass



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
        # Ensure destination directory exists
        os.makedirs(dest_folder, exist_ok=True)

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
            return [from_long_unc(os.path.join(d, f)) for f in sorted(qgz)]

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

    fld_simple = OxmlElement('w:hyperlink')

    fld_simple.set(qn('r:id'), r_id)



    run = OxmlElement('w:r')

    r_pr = OxmlElement('w:rPr')

    if italic:

        i = OxmlElement('w:i')

        r_pr.append(i)

    u = OxmlElement('w:u')

    u.set(qn('w:val'), 'single')

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

        # Card-like containers

        s.configure("Card.TFrame", background=card_bg, bordercolor=border, relief="solid", borderwidth=1, padding=8)

        s.configure("Header.TFrame", background=card_bg, bordercolor=border, relief="flat", padding=6)

        s.configure("TLabel", background=bg, foreground=fg)

        s.configure("Card.TLabel", background=card_bg, foreground=fg)

        s.configure("Subtle.TLabel", background=bg, foreground=subfg)

        s.configure("Tooltip.TLabel", background="#111827", foreground="#F9FAFB")

        # Buttons: compact, consistent

        s.configure("Accent.TButton", padding=(12, 6), background=accent, foreground="#FFFFFF")

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

        self.shared_driver = None

        self.style_helper = style_helper



        self.font_title = tkfont.Font(family="Segoe UI", size=15, weight="bold")

        self.font_sub   = tkfont.Font(family="Segoe UI", size=10)

        self.font_mono  = tkfont.Font(family="Consolas", size=9)



        self.ze_shp_var   = tk.StringVar(value=from_long_unc(self.prefs.get("ZE_SHP", "")))

        self.ae_shp_var   = tk.StringVar(value=from_long_unc(self.prefs.get("AE_SHP", "")))

        self.cadrage_var  = tk.StringVar(value=self.prefs.get("CADRAGE_MODE", "BOTH"))

        self.overwrite_var= tk.BooleanVar(value=self.prefs.get("OVERWRITE", OVERWRITE_DEFAULT))

        self.dpi_var      = tk.IntVar(value=int(self.prefs.get("DPI", DPI_DEFAULT)))

        self.workers_var  = tk.IntVar(value=int(self.prefs.get("N_WORKERS", N_WORKERS_DEFAULT)))

        self.wait_tiles_var = tk.DoubleVar(value=float(self.prefs.get("WAIT_TILES", WAIT_TILES_DEFAULT)))

        self.commune_var = tk.StringVar(value="")

        self.margin_var   = tk.DoubleVar(value=float(self.prefs.get("MARGIN_FAC", MARGIN_FAC_DEFAULT)))
        self.id_buffer_km_var = tk.DoubleVar(value=float(self.prefs.get("ID_BUFFER_KM", 5.0)))



        self.project_vars: dict[str, tk.IntVar] = {}

        self.all_projects: List[str] = []
        self.filtered_projects: List[str] = []

        self.total_expected = 0

        self.progress_done  = 0
        self.shared_driver = None

        self._build_ui()

        self._populate_projects()

        self._update_counts()

    def _build_ui(self):
        # State
        self.busy = False
        self.wiki_last_url = ""
        self._after_run_callback = None
        self._report_active = False
        self._report_iter = None

        # Vars used elsewhere
        self.out_dir_var = tk.StringVar(value=self.prefs.get("OUT_DIR", OUT_IMG))
        # Alias for buffer var used by start_id_thread
        self.buffer_var = tk.DoubleVar(value=float(self.prefs.get("ID_TAMPON_KM", self.id_buffer_km_var.get())))
        # Export type
        self.export_type_var = tk.StringVar(value=self.prefs.get("EXPORT_TYPE", "BOTH"))  # PNG, QGS, BOTH
        # Wiki query manual override
        self.wiki_query_var = tk.StringVar(value=self.prefs.get("WIKI_QUERY", ""))

        # Root layout using resizable panes
        root = ttk.Panedwindow(self, orient=tk.VERTICAL)
        root.pack(fill="both", expand=True)

        top_panes = ttk.Panedwindow(root, orient=tk.HORIZONTAL)
        left = ttk.Frame(top_panes, style="Card.TFrame", padding=6)
        right = ttk.Frame(top_panes, style="Card.TFrame", padding=6)
        top_panes.add(left, weight=1)
        top_panes.add(right, weight=2)

        bottom = ttk.Frame(root, style="Card.TFrame", padding=6)

        root.add(top_panes, weight=3)
        root.add(bottom, weight=1)

        # --- Left: parameters and actions ---
        ttk.Label(left, text="Contexte éco — Paramètres", font=self.font_title).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0,8))

        # Shapefile selectors
        ttk.Label(left, text="Shapefile Zone d'étude (ZE)").grid(row=1, column=0, sticky="w")
        ze_entry = ttk.Entry(left, textvariable=self.ze_shp_var, width=48, style="Card.TEntry")
        ze_entry.grid(row=1, column=1, sticky="ew", padx=(6,6))
        ttk.Button(left, text="Parcourir…", command=lambda: self._browse_file(self.ze_shp_var)).grid(row=1, column=2)

        ttk.Label(left, text="Shapefile Aire d'étude (AE)").grid(row=2, column=0, sticky="w")
        ae_entry = ttk.Entry(left, textvariable=self.ae_shp_var, width=48, style="Card.TEntry")
        ae_entry.grid(row=2, column=1, sticky="ew", padx=(6,6))
        ttk.Button(left, text="Parcourir…", command=lambda: self._browse_file(self.ae_shp_var)).grid(row=2, column=2)

        ttk.Button(left, text="Identifier commune", command=self._identify_commune).grid(row=3, column=0, sticky="w")
        ttk.Label(left, textvariable=self.commune_var).grid(row=3, column=1, columnspan=2, sticky="w")

        # Export options
        opt_frm = ttk.Frame(left)
        opt_frm.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8,4))
        for i in range(6):
            opt_frm.columnconfigure(i, weight=1)

        ttk.Label(opt_frm, text="DPI").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(opt_frm, from_=72, to=1200, increment=10, textvariable=self.dpi_var, width=6).grid(row=0, column=1, sticky="w")

        ttk.Label(opt_frm, text="Workers").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(opt_frm, from_=1, to=max(1, (os.cpu_count() or 2)), textvariable=self.workers_var, width=4).grid(row=0, column=3, sticky="w")

        ttk.Label(opt_frm, text="Marge (fac.)").grid(row=0, column=4, sticky="w")
        ttk.Spinbox(opt_frm, from_=1.0, to=2.0, increment=0.05, textvariable=self.margin_var, width=6).grid(row=0, column=5, sticky="w")

        ttk.Checkbutton(opt_frm, text="Écraser existants", variable=self.overwrite_var, style="Card.TCheckbutton").grid(row=1, column=0, columnspan=2, sticky="w", pady=(6,0))

        ttk.Label(opt_frm, text="Type d'export").grid(row=1, column=2, sticky="w", pady=(6,0))
        types = ttk.Frame(opt_frm)
        types.grid(row=1, column=3, columnspan=3, sticky="w", pady=(6,0))
        ttk.Radiobutton(types, text="PNG", value="PNG", variable=self.export_type_var, style="Card.TRadiobutton").pack(side="left")
        ttk.Radiobutton(types, text="QGS", value="QGS", variable=self.export_type_var, style="Card.TRadiobutton").pack(side="left", padx=(8,0))
        ttk.Radiobutton(types, text="Les deux", value="BOTH", variable=self.export_type_var, style="Card.TRadiobutton").pack(side="left", padx=(8,0))

        # Output directory
        ttk.Label(left, text="Dossier de sortie").grid(row=5, column=0, sticky="w", pady=(4,0))
        out_frm = ttk.Frame(left)
        out_frm.grid(row=5, column=1, columnspan=2, sticky="ew", padx=(6,0), pady=(4,0))
        out_frm.columnconfigure(0, weight=1)
        out_entry = ttk.Entry(out_frm, textvariable=self.out_dir_var, width=48, style="Card.TEntry")
        out_entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(out_frm, text="Choisir…", command=lambda: self._browse_dir(self.out_dir_var)).grid(row=0, column=1, padx=(6,0))
        btn_open_out = ttk.Button(out_frm, text="Ouvrir", command=self._open_output_dir)
        btn_open_out.grid(row=0, column=2, padx=(6,0))
        try:
            ToolTip(btn_open_out, "Ouvrir le dossier d'export dans l'explorateur")
        except Exception:
            pass

        # Actions: Export, Identification
        act_frm = ttk.Frame(left)
        act_frm.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10,4))
        act_frm.columnconfigure(0, weight=1)
        act_frm.columnconfigure(1, weight=1)
        act_frm.columnconfigure(2, weight=1)
        self.export_button = ttk.Button(act_frm, text="Exporter cartes", style="Accent.TButton", command=self.start_export_thread)
        self.export_button.grid(row=0, column=0, sticky="ew", padx=(0,6))
        self.id_button = ttk.Button(act_frm, text="ID Contexte éco", command=self.start_id_thread)
        self.id_button.grid(row=0, column=1, sticky="ew", padx=(0,6))
        self.report_button = ttk.Button(
            act_frm,
            text="Générer rapport Word",
            command=self.start_report_sequence,
        )
        self.report_button.grid(row=0, column=2, sticky="ew")

        # ID buffer
        id_frm = ttk.Frame(left)
        id_frm.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(4,8))
        ttk.Label(id_frm, text="Tampon (km)").pack(side="left")
        ttk.Spinbox(id_frm, from_=0.0, to=20.0, increment=0.5, textvariable=self.buffer_var, width=6).pack(side="left", padx=(6,0))

        # Quick tools: RLT, Maps, BV
        tools_frm = ttk.Frame(left)
        tools_frm.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(2,10))
        self.rlt_button = ttk.Button(tools_frm, text="Remonter le temps (IGN)", command=self.start_rlt_thread)
        self.rlt_button.pack(side="left")
        ttk.Button(tools_frm, text="Google Maps", command=self.open_gmaps).pack(side="left", padx=6)
        self.bassin_button = ttk.Button(tools_frm, text="Bassin versant", command=self.start_bassin_thread)
        self.bassin_button.pack(side="left")

        # Wikipedia + Veg/Sol scraping controls
        wiki_frm = ttk.LabelFrame(left, text="Scraping Wikipedia / Veg&Sol")
        wiki_frm.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(8,4))
        wiki_frm.columnconfigure(0, weight=1)

        wiki_sub_frm = ttk.Frame(wiki_frm)
        wiki_sub_frm.grid(row=0, column=0, columnspan=2, sticky="ew", padx=4, pady=4)
        wiki_sub_frm.columnconfigure(0, weight=1)

        self.wiki_button = ttk.Button(wiki_sub_frm, text="Lancer scraping (Climat, Occ. sol, Altitude, Végét., Sol)", command=self.start_full_scrape_thread)
        self.wiki_button.grid(row=0, column=0, sticky="ew", padx=(0,4))

        self.wiki_status_var = tk.StringVar(value="Prêt.")
        status_label = ttk.Label(wiki_sub_frm, textvariable=self.wiki_status_var, style="Status.TLabel", anchor="e")
        status_label.grid(row=0, column=1, sticky="e")

        # Manual override for Wikipedia query
        ttk.Label(wiki_frm, text="Commune (si différent de la ZE)").grid(row=1, column=0, sticky="w", padx=4, pady=(4,0))
        ttk.Entry(wiki_frm, textvariable=self.wiki_query_var, style="Card.TEntry").grid(row=1, column=1, sticky="ew", padx=4, pady=(4,0))

        # Results tree for Wikipedia/Végétation/Sol
        res_frm = ttk.Frame(wiki_frm)
        res_frm.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4, pady=(4,6))
        res_frm.columnconfigure(0, weight=1)
        self.results_tree = ttk.Treeview(res_frm, columns=("label", "value"), show="headings", height=5)
        self.results_tree.heading("label", text="Rubrique")
        self.results_tree.heading("value", text="Résultat")
        self.results_tree.column("label", width=150, anchor="w")
        self.results_tree.column("value", anchor="w")
        self.results_tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(res_frm, orient="vertical", command=self.results_tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.results_tree.configure(yscrollcommand=yscroll.set)
        # Initialize rows with placeholders and fixed iids expected by _update_results_tree
        try:
            self.results_tree.insert("", "end", values=("Climat", "—"), iid="climat")
            self.results_tree.insert("", "end", values=("Occupation sols", "—"), iid="occupation_sols")
            self.results_tree.insert("", "end", values=("Altitude", "—"), iid="altitude")
            self.results_tree.insert("", "end", values=("Végétation", "—"), iid="vegetation")
            self.results_tree.insert("", "end", values=("Sols", "—"), iid="sols")
        except Exception:
            pass

        # --- Right: projects list with filter and scroll ---
        ttk.Label(right, text="Projets QGIS — Export", font=self.font_title).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0,8), padx=6)
        # Filter controls
        self.filter_var = tk.StringVar(value="")
        filt_frm = ttk.Frame(right)
        filt_frm.grid(row=1, column=0, columnspan=3, sticky="ew", padx=6)
        filt_frm.columnconfigure(1, weight=1)
        ttk.Label(filt_frm, text="Filtrer").grid(row=0, column=0, sticky="w")
        ent = ttk.Entry(filt_frm, textvariable=self.filter_var, style="Card.TEntry")
        ent.grid(row=0, column=1, sticky="ew", padx=(6,6))
        ent.bind("<Return>", lambda e: self._apply_filter())
        ttk.Button(filt_frm, text="Appliquer", command=self._apply_filter).grid(row=0, column=2)
        ttk.Button(filt_frm, text="Tout sélectionner", command=lambda: self._select_all(True)).grid(row=0, column=3, padx=(6,0))
        ttk.Button(filt_frm, text="Tout désélectionner", command=lambda: self._select_all(False)).grid(row=0, column=4, padx=(6,0))

        # Scrollable list of projects
        proj_frm = ttk.Frame(right)
        proj_frm.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=6, pady=(6,6))
        right.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)
        # Canvas + scrollbar
        canvas = tk.Canvas(proj_frm, highlightthickness=0)
        vbar = ttk.Scrollbar(proj_frm, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        proj_frm.rowconfigure(0, weight=1)
        proj_frm.columnconfigure(0, weight=1)
        # Inner frame
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        self._projects_window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        # Make the inner frame width track the canvas width
        def _on_canvas_configure(event):
            try:
                canvas.itemconfigure(self._projects_window, width=event.width)
            except Exception:
                pass
        canvas.bind("<Configure>", _on_canvas_configure)

        # Status and progress
        status_frm = ttk.Frame(right)
        status_frm.grid(row=3, column=0, columnspan=3, sticky="ew", padx=6, pady=(0,6))
        status_frm.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(status_frm, text="Projets sélectionnés : 0 / 0", style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(status_frm, mode="determinate")
        self.progress.grid(row=0, column=1, sticky="ew", padx=(8,0))

        # --- Bottom: console output ---
        bottom.columnconfigure(0, weight=1)
        bottom.rowconfigure(0, weight=1)
        ttk.Label(bottom, text="Console", style="Subtle.TLabel").grid(row=0, column=0, sticky="w", padx=6, pady=(4,0))
        ttk.Button(bottom, text="Ouput", command=self._open_output_dir).grid(row=0, column=1, sticky="e", padx=6, pady=(4,0))
        console_frm = ttk.Frame(bottom)
        console_frm.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0,6))
        console_frm.columnconfigure(0, weight=1)
        console_frm.rowconfigure(0, weight=1)
        self.console = tk.Text(console_frm, height=10, state="disabled", wrap="word")
        self.console.grid(row=0, column=0, sticky="nsew")
        cons_scroll = ttk.Scrollbar(console_frm, orient="vertical", command=self.console.yview)
        cons_scroll.grid(row=0, column=1, sticky="ns")
        self.console.configure(yscrollcommand=cons_scroll.set, font=self.font_mono)
        # Redirect stdout to console
        try:
            self.stdout_redirect = TextRedirector(self.console)
        except Exception:
            self.stdout_redirect = sys.stdout

    def open_gmaps(self):
        """Ouvre Google Maps centré sur le centroïde de la ZE."""
        centroid = self._get_centroid_wgs84()
        if not centroid:
            return
        lat, lon = centroid
        url = f"https://www.google.com/maps/@{lat},{lon},16z"
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir Google Maps: {e}")

    def start_rlt_thread(self):
        """Ouvre IGN Remonter le temps dans le navigateur (thread pour ne pas bloquer l'UI)."""
        t = threading.Thread(target=self._open_rlt_links, daemon=True)
        t.start()

    def _open_rlt_links(self):
        try:
            centroid = self._get_centroid_wgs84()
            if not centroid:
                return
            lat_dd, lon_dd = centroid

            # Déterminer la commune à partir du centroïde
            commune = ""
            try:
                url = (
                    "https://geo.api.gouv.fr/communes"
                    f"?lat={lat_dd}&lon={lon_dd}&fields=nom,codeDepartement&format=json"
                )
                resp = requests.get(url, timeout=10)
                resp.raise_for_status()
                data = resp.json()
                if data:
                    nom = data[0].get("nom")
                    dep = data[0].get("codeDepartement")
                    if nom and dep:
                        commune = f"{nom} ({dep})"
            except Exception:
                pass

            # Fallback: tenter de récupérer le nom depuis le shapefile
            if not commune:
                try:
                    shp_path = self.ze_shp_var.get()
                    gdf = gpd.read_file(shp_path)
                    for col in ["commune", "nom", "NOM_COM", "NOM"]:
                        if col in gdf.columns:
                            commune = str(gdf.iloc[0][col])
                            break
                except Exception:
                    pass

            output_dir = os.path.join(self.out_dir_var.get() or OUT_IMG, "Remonter le temps")
            word_filename = "Comparaison_temporelle_Paysage.docx"
            os.makedirs(output_dir, exist_ok=True)

            print(f"Coordonnées utilisées : {lat_dd:.6f}, {lon_dd:.6f}")

            # --- 2. Capture des images avec Selenium ---
            print("Lancement de la capture d'images...")
            driver = self._get_or_create_driver()
            if not driver:
                return
            
            try:
                driver.minimize_window()
            except Exception:
                pass
            images = []
            viewport = (By.CSS_SELECTOR, "div.ol-viewport")

            # Utiliser les couches de l'ancien script
            layers_ign = [
                ("Aujourd’hui", "10"),
                ("2000-2005", "18"),
                ("1965-1980", "20"),
                ("1950-1965", "19"),
            ]
            url_template = "https://remonterletemps.ign.fr/comparer/?lon={lon}&lat={lat}&z=17&layer1={layer}&layer2=19&mode=dub1"

            for title, layer_val in layers_ign:
                url = url_template.format(lon=f"{lon_dd:.6f}", lat=f"{lat_dd:.6f}", layer=layer_val)
                driver.get(url)
                WebDriverWait(driver, 20).until(EC.visibility_of_element_located(viewport))
                time.sleep(self.wait_tiles_var.get() or 1.5)
                
                tgt = driver.find_element(*viewport)
                img_path = os.path.join(output_dir, f"{title}.png")
                if tgt.screenshot(img_path):
                    img = Image.open(img_path)
                    w, h = img.size
                    left, right = int(w * 0.05), int(w * 0.95)
                    img.crop((left, 0, right, h)).save(img_path)
                    images.append((title, img_path))
                else:
                    print(f"Capture échouée : {title}")

            # --- 3. Création du document Word ---
            if not images:
                messagebox.showwarning("Aucune image", "Aucune image n'a été capturée, le document Word ne sera pas généré.")
                return

            print("Génération du document Word...")
            doc = Document()
            style_normal = doc.styles['Normal']
            style_normal.font.name = 'Calibri'
            style_normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')

            sec = doc.sections[0]
            sec.orientation = WD_ORIENT.LANDSCAPE
            sec.page_width, sec.page_height = sec.page_height, sec.page_width
            sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = [Cm(1.5)] * 4

            caption_text = f"Tableau 3 : Évolution de l’occupation du sol du territoire de la zone d’étude (source : IGN – RemonterLeTemps)"
            cap_par = doc.add_paragraph()
            cap_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_hyperlink(cap_par, "https://remonterletemps.ign.fr/", caption_text)

            table = doc.add_table(rows=2, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = "Table Grid"
            table.autofit = False

            for idx, (title, path) in enumerate(images):
                r, c = divmod(idx, 2)
                cell = table.cell(r, c)
                p_t = cell.paragraphs[0]
                run_t = p_t.add_run(title)
                run_t.bold = True
                p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img = cell.add_paragraph()
                p_img.add_run().add_picture(path, width=Cm(12.5 * 0.8))
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph()
            comment_template = (
                "Rédige un commentaire synthétique de l'évolution de l'occupation du sol observée "
                "sur les images aériennes de la commune de {commune}, aux différentes dates indiquées "
                "(1950–1965, 1965–1980, 2000–2005, aujourd’hui). Concentre-toi sur les grandes "
                "dynamiques d'aménagement (urbanisation, artificialisation, évolution des milieux "
                "ouverts ou boisés), en identifiant les principales transformations visibles. "
                "Fais ta réponse en un seul court paragraphe. Intègre les éléments de contexte "
                "historique et territorial propres à la commune de {commune} pour interpréter ces évolutions."
            )
            p_comm = doc.add_paragraph(comment_template.format(commune=commune))
            p_comm.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            doc_path = os.path.join(output_dir, word_filename)
            doc.save(doc_path)
            messagebox.showinfo("Succès", f"Document Word généré :\n{doc_path}")

        except Exception as e:
            error_message = f"Une erreur est survenue lors du processus Remonter le temps: {e}"
            print(error_message)
            traceback.print_exc()
            messagebox.showerror("Erreur", error_message)

    def start_bassin_thread(self):
        """Ouvre une recherche utile pour le bassin versant autour du centroïde."""
        centroid = self._get_centroid_wgs84()
        if not centroid:
            return
        lat, lon = centroid
        try:
            # Heuristic: open a Google Maps search around the centroid
            url = f"https://www.google.com/maps/search/bassin+versant/@{lat},{lon},12z"
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir la recherche bassin versant: {e}")

    def _open_output_dir(self):
        """Ouvre le dossier de sortie dans l'explorateur. Le crée s'il n'existe pas."""
        try:
            path = (self.out_dir_var.get() or "").strip() or OUT_IMG
            if not path:
                messagebox.showerror("Erreur", "Aucun dossier de sortie défini.")
                return
            os.makedirs(path, exist_ok=True)
            # Windows: ouvrir dans l'explorateur
            if sys.platform.startswith("win"):
                os.startfile(path)
            else:
                # Fallback pour autres OS
                webbrowser.open(f"file://{path}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le dossier: {e}")

    def _browse_file(self, var: tk.StringVar) -> None:
        """Sélectionne un fichier shapefile et met à jour la variable liée.
        Persiste aussi la préférence ZE_SHP/AE_SHP selon le champ ciblé."""
        try:
            initial_dir = DEFAULT_SHAPE_DIR if os.path.isdir(DEFAULT_SHAPE_DIR) else os.path.dirname(var.get() or "")
            path = filedialog.askopenfilename(
                title="Sélectionner un shapefile",
                initialdir=initial_dir or os.getcwd(),
                filetypes=(("Shapefile", "*.shp"), ("Tous les fichiers", "*.*")),
            )
            if not path:
                return
            if not path.lower().endswith(".shp"):
                messagebox.showerror("Fichier invalide", "Veuillez sélectionner un fichier .shp valide.")
                return
            var.set(path)
            # Mémoriser dans les préférences
            if var is self.ze_shp_var:
                self.prefs["ZE_SHP"] = path
            elif var is self.ae_shp_var:
                self.prefs["AE_SHP"] = path
            save_prefs(self.prefs)
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de la sélection du shapefile: {e}")

    def _browse_dir(self, var: tk.StringVar) -> None:
        """Sélectionne un dossier de sortie et met à jour la variable + préférences."""
        try:
            initial_dir = var.get() or OUT_IMG
            path = filedialog.askdirectory(
                title="Choisir un dossier",
                initialdir=initial_dir if os.path.isdir(initial_dir) else os.getcwd(),
            )
            if not path:
                return
            var.set(path)
            # Persiste OUT_DIR
            if var is self.out_dir_var:
                self.prefs["OUT_DIR"] = path
                save_prefs(self.prefs)
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de la sélection du dossier: {e}")

    def _get_or_create_driver(self):
        """Initialise et retourne un driver Selenium, en le réutilisant s'il existe déjà."""
        if self.shared_driver:
            try:
                # Vérifier si le driver est toujours actif
                _ = self.shared_driver.window_handles
                print("Réutilisation du WebDriver existant.")
                return self.shared_driver
            except Exception:
                print("Le WebDriver partagé n'est plus valide, création d'un nouveau.")
                self._cleanup_driver()

        try:
            print("Création d'un nouveau WebDriver...")
            options = webdriver.ChromeOptions()
            # Ajoutez ici des options si nécessaire (ex: headless)
            # options.add_argument('--headless')
            service = Service() # Assurez-vous que chromedriver est dans le PATH
            driver = webdriver.Chrome(service=service, options=options)
            try:
                driver.minimize_window()
            except Exception:
                pass
            self.shared_driver = driver
            return driver
        except Exception as e:
            print(f"Erreur lors de la création du WebDriver: {e}")
            traceback.print_exc()
            messagebox.showerror("Erreur WebDriver", f"Impossible de démarrer Selenium. Assurez-vous que ChromeDriver est installé et dans votre PATH. Erreur: {e}")
            return None

    def _cleanup_driver(self):
        if self.shared_driver:
            try:
                self.shared_driver.quit()
            except Exception:
                pass
            finally:
                self.shared_driver = None

    def _get_centroid_wgs84(self) -> Optional[Tuple[float, float]]:
        shp_path = self.ze_shp_var.get()
        if not shp_path or not os.path.isfile(shp_path):
            messagebox.showerror("Shapefile manquant", "Veuillez sélectionner un shapefile de Zone d'étude valide.")
            return None
        try:
            gdf = gpd.read_file(shp_path)
            if gdf.crs is None:
                messagebox.showwarning("CRS manquant", "Le shapefile n'a pas de système de coordonnées défini. L'application va supposer Lambert-93.")
                gdf.set_crs("EPSG:2154", inplace=True)
            
            gdf_wgs84 = gdf.to_crs(epsg=4326)
            
            # Utiliser union_all() qui remplace unary_union déprécié
            centroid = gdf_wgs84.geometry.union_all().centroid
            return (centroid.y, centroid.x) # lat, lon
        except Exception as e:
            messagebox.showerror("Erreur Géospatiale", f"Impossible de calculer le centroïde du shapefile. Erreur: {e}")
            return None

    def _identify_commune(self) -> None:
        centroid = self._get_centroid_wgs84()
        if not centroid:
            return
        lat, lon = centroid
        try:
            url = (
                "https://geo.api.gouv.fr/communes"
                f"?lat={lat}&lon={lon}&fields=nom,codeDepartement&format=json"
            )
            resp = requests.get(url, timeout=10)
            resp.raise_for_status()
            data = resp.json()
            if data:
                nom = data[0].get("nom")
                dep = data[0].get("codeDepartement")
                if nom and dep:
                    self.commune_var.set(f"{nom} ({dep})")
                    return
            messagebox.showwarning(
                "Commune introuvable", "Aucune commune trouvée pour ces coordonnées."
            )
        except Exception as e:
            messagebox.showerror(
                "Erreur", f"Impossible d'identifier la commune: {e}"
            )

    def _run_wiki_scrape(self, driver, query: str) -> dict:
        print(f"Lancement du scraping Wikipedia pour: '{query}'")
        try:
            # La fonction get_wikipedia_extracts gère déjà l'ouverture de l'URL
            # On doit juste s'assurer qu'elle utilise le driver partagé
            extracts = get_wikipedia_extracts(query, driver)
            self.wiki_last_url = driver.current_url
            print("Scraping Wikipedia terminé.")
            return extracts
        except Exception as e:
            print(f"Erreur lors du scraping Wikipedia: {e}")
            traceback.print_exc()
            return {}

    def _run_altitude(self, driver, lon, lat) -> str:
        url = f"https://www.geoportail.gouv.fr/carte?lon={lon}&lat={lat}&z=15"
        altitude = "Non trouvée"
        try:
            print(f"Ouverture de la carte d'altitude dans un nouvel onglet: {url}")
            # Ouvrir dans un nouvel onglet
            driver.execute_script(f"window.open('{url}', '_blank');")
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(3) # Attendre le chargement initial
            
            wait = WebDriverWait(driver, 10)
            # Attendre que le conteneur des coordonnées soit visible
            wait.until(EC.visibility_of_element_located((By.ID, "gp-coordinates-container")))
            
            # Extraire l'altitude
            alt_element = driver.find_element(By.CSS_SELECTOR, "div.gp-coords-altitude span.gp-coords-value")
            altitude = alt_element.text.strip()
            print(f"Altitude trouvée: {altitude}")

        except Exception as e:
            print(f"Erreur lors de la récupération de l'altitude: {e}")
            return "Erreur scraping"

        return altitude

    def _run_vegsol(self, driver, lon, lat) -> tuple[str, str]:
        base_url = "https://www.geoportail.gouv.fr/carte"
        params = {
            "lon": lon,
            "lat": lat,
            "z": 15,
            "layers": "GEOGRAPHICALGRIDSYSTEMS.MAPS.3D,sol,vegetation",
            "ch": "GEOGRAPHICALGRIDSYSTEMS.MAPS.3D,sol,vegetation",
            "gp-access-lib": "x6j23m5wpr15nucen18p8hfn"
        }
        url = f"{base_url}?{urllib.parse.urlencode(params)}"

        veg = "Non trouvée"
        soil = "Non trouvé"

        try:
            # Ouvrir la carte dans un nouvel onglet
            print(f"Ouverture de la carte Végétation/Sol dans un nouvel onglet: {url}")
            driver.execute_script(f"window.open('{url}', '_blank');")
            
            # Attendre et basculer vers le nouvel onglet
            time.sleep(1) # Laisser le temps à l'onglet de s'ouvrir
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(4) # Attendre que la page et les scripts se chargent

            # Attendre que les informations soient potentiellement chargées
            wait = WebDriverWait(driver, 15) # Attente max de 15 secondes
            
            # Utiliser une attente explicite pour la végétation
            try:
                veg_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#vegetation-info")))
                veg = veg_element.text.strip()
                print(f"Végétation trouvée: {veg}")
            except Exception:
                print("L'élément d'information sur la végétation n'a pas été trouvé dans le temps imparti.")

            # Utiliser une attente explicite pour le sol
            try:
                soil_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#soil-info")))
                soil = soil_element.text.strip()
                print(f"Sol trouvé: {soil}")
            except Exception:
                print("L'élément d'information sur le sol n'a pas été trouvé dans le temps imparti.")

        except Exception as e:
            print(f"Erreur majeure lors du scraping de la carte Végétation/Sol: {e}")
            return "Erreur scraping", "Erreur scraping"
        
        return veg, soil

    def _run_full_scrape(self):
        """Orchestrates the entire scraping process in sequence."""
        self.after(0, lambda: self.wiki_status_var.set("Démarrage..."))
        
        centroid = self._get_centroid_wgs84()
        if not centroid:
            self.after(0, lambda: self.wiki_status_var.set("Erreur: Centroïde"))
            return

        lat, lon = centroid # Note: _get_centroid_wgs84 returns lat, lon
        
        # Détecter la commune pour le scraping
        commune_detected, _ = self._detect_commune(lat, lon)
        query = self.wiki_query_var.get().strip() or commune_detected
        if not query:
            self.after(0, lambda: self.wiki_status_var.set("Erreur: Commune"))
            return

        driver = self._get_or_create_driver()
        if not driver:
            self.after(0, lambda: self.wiki_status_var.set("Erreur: Driver"))
            return

        try:
            try:
                driver.minimize_window()
            except Exception:
                pass
            # 1. Wikipedia
            self.after(0, lambda: self.wiki_status_var.set("Scraping Wikipedia..."))
            extracts = self._run_wiki_scrape(driver, query)

            # 2. Altitude
            self.after(0, lambda: self.wiki_status_var.set("Scraping Altitude..."))
            altitude = self._run_altitude(driver, lon, lat)

            # 3. Végétation et Sol
            self.after(0, lambda: self.wiki_status_var.set("Scraping Végétation/Sol..."))
            veg, soil = self._run_vegsol(driver, lon, lat)

            # Switch back to the first tab to leave it clean
            driver.switch_to.window(driver.window_handles[0])

            # Prepare results
            payload = {
                'climat': extracts.get('climat', 'Non trouvé'),
                'occupation_sols': extracts.get('occupation_sols', 'Non trouvé'),
                'altitude': altitude,
                'vegetation': veg,
                'sols': soil
            }
            self.after(0, self._update_results_tree, payload)
            self.after(0, lambda: self.wiki_status_var.set("Terminé !"))

        except Exception as e:
            error_message = f"Erreur scraping: {e}"
            print(error_message)
            traceback.print_exc()
            self.after(0, lambda: self.wiki_status_var.set("Erreur"))

    def start_full_scrape_thread(self):
        """Starts the full scraping process in a separate thread to avoid freezing the UI."""
        t = threading.Thread(target=self._run_full_scrape, daemon=True)
        t.start()

    def _update_results_tree(self, data: dict) -> None:
        """Update rows in self.results_tree. Keys are expected to be iids:
        'climat', 'occupation_sols', 'altitude', 'vegetation', 'sols'.
        """
        try:
            if not hasattr(self, 'results_tree'):
                return
            labels = {
                'climat': 'Climat',
                'occupation_sols': 'Occupation sols',
                'altitude': 'Altitude',
                'vegetation': 'Végétation',
                'sols': 'Sols',
            }
            for iid, text in (data or {}).items():
                if iid not in labels:
                    continue
                try:
                    # Preserve first column label
                    cur = self.results_tree.item(iid, 'values')
                    label = cur[0] if (isinstance(cur, (list, tuple)) and len(cur) >= 1) else labels[iid]
                    self.results_tree.item(iid, values=(label, text))
                except Exception:
                    # If row doesn't exist yet, insert it
                    self.results_tree.insert("", "end", values=(labels[iid], text), iid=iid)
        except Exception:
            pass

    # ... (rest of the code remains the same)
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

        # Display checkboxes in two columns for compactness

        for i, proj_path in enumerate(self.filtered_projects):

            var = tk.IntVar(value=1); self.project_vars[proj_path] = var

            r, c = divmod(i, 2)

            cb = ttk.Checkbutton(self.scrollable_frame, text=os.path.basename(proj_path), variable=var, style="Card.TCheckbutton")

            cb.grid(row=r, column=c, sticky='w', padx=4, pady=2)

        try:

            self.scrollable_frame.columnconfigure(0, weight=1)

            self.scrollable_frame.columnconfigure(1, weight=1)

        except Exception:

            pass



    def _apply_filter(self):

        term = normalize_name(self.filter_var.get())

        for w in list(self.scrollable_frame.children.values()): w.destroy()

        self.filtered_projects = [p for p in self.all_projects if term in normalize_name(os.path.basename(p))]

        if not self.filtered_projects:

            ttk.Label(self.scrollable_frame, text="Aucun projet ne correspond au filtre.", foreground="red").pack(anchor="w")

            self.project_vars = {}; self._update_counts(); return

        for i, proj_path in enumerate(self.filtered_projects):

            current = self.project_vars.get(proj_path, tk.IntVar(value=1))

            self.project_vars[proj_path] = current

            r, c = divmod(i, 2)

            cb = ttk.Checkbutton(self.scrollable_frame, text=os.path.basename(proj_path), variable=current, style="Card.TCheckbutton")

            cb.grid(row=r, column=c, sticky='w', padx=4, pady=2)

        try:

            self.scrollable_frame.columnconfigure(0, weight=1)

            self.scrollable_frame.columnconfigure(1, weight=1)

        except Exception:

            pass

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
        try:
            if hasattr(self, 'id_button') and self.id_button:
                self.id_button.config(state="disabled")
        except Exception:
            pass

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

            # Désactive provisoirement le multiprocessing pour éviter les erreurs _multiprocessing

            workers = workers

            chunks = chunk_even(projets, workers)

            # Forcer au moins 2 workers pour utiliser ProcessPoolExecutor

            # (et donc le Python de QGIS configuré ci-dessous)

            workers = max(1, workers)

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

                    ok, ko = run_worker_subprocess(chunk, cfg)

                    ok_total += ok

                    ko_total += ko

                    self.after(0, ui_update_progress, ok + ko)

                    log_with_time(f"Lot terminé: {ok} OK, {ko} KO")

            else:

                try:

                    import multiprocessing as mp

                    # Nettoyage de l'environnement hérité pour éviter collisions Python 3.12/3.13

                    for _k in ("PYTHONHOME", "PYTHONPATH", "PYTHONSTARTUP"):

                        try:

                            os.environ.pop(_k, None)

                        except Exception:

                            pass

                    os.environ["PYTHONNOUSERSITE"] = "1"

                    # Fixer PYTHONHOME et PYTHONPATH sur le Python de QGIS

                    try:

                        qgis_py_root = os.path.join(QGIS_ROOT, "apps", PY_VER)

                        if os.path.isdir(qgis_py_root):

                            os.environ["PYTHONHOME"] = qgis_py_root

                            qgis_lib   = os.path.join(qgis_py_root, "Lib")

                            qgis_dlls  = os.path.join(qgis_py_root, "DLLs")

                            qgis_site  = os.path.join(qgis_lib, "site-packages")

                            qgis_app_py = os.path.join(QGIS_APP, "python")

                            py_paths = [qgis_py_root, qgis_lib, qgis_dlls, qgis_site, qgis_app_py]

                            os.environ["PYTHONPATH"] = os.pathsep.join(py_paths)

                            # Préfixer le PATH avec les dossiers Python QGIS pour la résolution des DLLs

                            os.environ["PATH"] = os.pathsep.join([qgis_py_root, qgis_dlls, os.environ.get("PATH", "")])

                            log_with_time(f"PYTHONHOME={qgis_py_root}")

                    except Exception:

                        pass

                    # Fixer PYTHONHOME sur le Python de QGIS pour que l'interprète trouve sa stdlib

                    try:

                        qgis_py_root = os.path.join(QGIS_ROOT, "apps", PY_VER)

                        if os.path.isdir(qgis_py_root):

                            os.environ["PYTHONHOME"] = qgis_py_root

                            log_with_time(f"PYTHONHOME={qgis_py_root}")

                    except Exception:

                        pass

                    ctx = mp.get_context("spawn")

                    try:

                        qgis_py = os.path.join(QGIS_ROOT, "apps", PY_VER, "python.exe")

                        if os.path.isfile(qgis_py):

                            ctx.set_executable(qgis_py)

                            log_with_time(f"MP exe: {qgis_py}")

                        else:

                            log_with_time(f"Python QGIS introuvable: {qgis_py}")

                    except Exception as e:

                        log_with_time(f"set_executable échec: {e}")

                except Exception as e:

                    log_with_time(f"init multiprocessing: {e}")

                # Ajuster temporairement sys.path pour privilégier les libs QGIS

                old_syspath = list(sys.path)

                try:

                    qgis_py_root = os.path.join(QGIS_ROOT, "apps", PY_VER)

                    qgis_py_lib = os.path.join(qgis_py_root, "Lib")

                    qgis_site = os.path.join(qgis_py_lib, "site-packages")

                    qgis_app_py = os.path.join(QGIS_APP, "python")

                    def _keep_path(p: str) -> bool:

                        if not isinstance(p, str):

                            return False

                        l = p.lower()

                        if "python313" in l or "python311" in l or "python310" in l or "python39" in l:

                            return False

                        if ".venv" in l:

                            return False

                        return True

                    sys.path = [qgis_py_root, qgis_py_lib, qgis_site, qgis_app_py] + [p for p in old_syspath if _keep_path(p)]

                except Exception as e:

                    log_with_time(f"sys.path cleanup skip: {e}")

                with ThreadPoolExecutor(max_workers=workers) as ex:

                    futures = [ex.submit(run_worker_subprocess, chunk, cfg) for chunk in chunks if chunk]

                    for fut in as_completed(futures):

                        try:

                            ok, ko = fut.result()

                            ok_total += ok

                            ko_total += ko

                            self.after(0, ui_update_progress, ok + ko)

                            log_with_time(f"Lot terminé: {ok} OK, {ko} KO")

                        except Exception as e:

                            log_with_time(f"Erreur worker: {e}")

                # Restaure le sys.path initial

                sys.path = old_syspath

            elapsed = datetime.datetime.now() - start

            log_with_time(f"FIN — OK={ok_total} | KO={ko_total} | Attendu={self.total_expected} | Durée={elapsed}")

            self.after(0, lambda: self.status_label.config(text=f"Terminé — OK={ok_total} / KO={ko_total}"))

        except Exception as e:

            log_with_time(f"Erreur critique: {e}")

            _err = str(e)

            self.after(0, lambda msg=_err: messagebox.showerror("Erreur", msg))

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
        # Certaines configurations n'affichent pas le bouton d'identification
        try:
            if hasattr(self, 'id_button') and self.id_button:
                self.id_button.config(state="disabled")
        except Exception:
            pass

        self.progress.config(mode="indeterminate")

        self.progress.start()

        self.status_label.config(text="Analyse en cours…")



        # Enregistrer des chemins standard (sans préfixe long UNC) dans les préférences
        self.prefs.update({
            "ZE_SHP": from_long_unc(ze),
            "AE_SHP": from_long_unc(ae),
            "ID_TAMPON_KM": float(self.buffer_var.get()),
        }); save_prefs(self.prefs)



        t = threading.Thread(target=self._run_id_logic, args=(ae, ze, float(self.buffer_var.get())), daemon=True)

        t.start()



    def _run_id_logic(self, ae: str, ze: str, buffer_km: float):

        old_stdout = sys.stdout

        sys.stdout = self.stdout_redirect

        try:
            # Import robuste: relatif (via package) puis absolu (script direct)
            try:
                from .id_contexte_eco import run_analysis as run_id_context
            except Exception:
                from id_contexte_eco import run_analysis as run_id_context

            run_id_context(ae, ze, buffer_km)

            log_with_time("Analyse terminée.")

            self.after(0, lambda: self.status_label.config(text="Terminé"))

        except Exception as e:

            log_with_time(f"Erreur: {e}")

            _err = str(e)

            self.after(0, lambda msg=_err: messagebox.showerror("Erreur", msg))

        finally:

            sys.stdout = old_stdout

            self.after(0, lambda: (self.progress.stop(), self.progress.config(mode="determinate", value=0)))

            self.after(0, self._run_finished)



    def _run_finished(self):
        if not getattr(self, '_report_active', False):
            self.export_button.config(state="normal")
            try:
                if hasattr(self, 'id_button') and self.id_button:
                    self.id_button.config(state="normal")
            except Exception:
                pass
            try:
                if hasattr(self, 'report_button') and self.report_button:
                    self.report_button.config(state="normal")
            except Exception:
                pass
            self.busy = False
        cb = getattr(self, '_after_run_callback', None)
        if cb:
            self._after_run_callback = None
            self.after(0, cb)

    # ---------- Rapport automatique ----------

    def start_report_sequence(self):
        if self.busy:
            print("Une action est déjà en cours.", file=self.stdout_redirect)
            return
        self._report_active = True
        self._report_iter = iter(['id', 'export', 'word'])
        self.export_button.config(state="disabled")
        try:
            if hasattr(self, 'id_button') and self.id_button:
                self.id_button.config(state="disabled")
        except Exception:
            pass
        self.report_button.config(state="disabled")
        self.busy = True
        self._run_next_report_step()

    def _run_next_report_step(self):
        try:
            step = next(self._report_iter)
        except StopIteration:
            self._report_active = False
            self._run_finished()
            self.status_label.config(text="Terminé")
            return
        if step == 'id':
            self._after_run_callback = self._run_next_report_step
            self.start_id_thread()
        elif step == 'export':
            self._after_run_callback = self._run_next_report_step
            self.start_export_thread()
        elif step == 'word':
            try:
                self.generate_report()
            except Exception as e:
                print(f"Erreur génération rapport: {e}", file=self.stdout_redirect)
                self.after(0, lambda msg=str(e): messagebox.showerror("Erreur", msg))
            self._run_next_report_step()

    def generate_report(self):
        xlsx = os.path.join(OUT_IMG, 'ID zonages.xlsx')
        if not os.path.isfile(xlsx):
            raise FileNotFoundError("Fichier d'analyse introuvable")
        tpl_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Template word  Contexte éco'))
        tpl_path = os.path.join(tpl_dir, '1 Template Contexte éco.docx')
        if not os.path.isfile(tpl_path):
            raise FileNotFoundError("Template Word introuvable")
        out_doc = os.path.join(OUT_IMG, f"Rapport Contexte eco {datetime.datetime.now():%Y%m%d_%H%M%S}.docx")
        shutil.copy(tpl_path, out_doc)
        doc = Document(out_doc)
        xls = pd.ExcelFile(xlsx)

        def find_sheets(patterns):
            def norm(s):
                return s.lower().replace(" ", "")

            found = []
            for pat in patterns:
                npat = norm(pat)
                for name in xls.sheet_names:
                    if npat in norm(name):
                        found.append(name)
            return found

        table_map = {
            'TABLEAU NATURA2000': ['Natura 2000', 'N2000'],
            'TABLEAU ZNIEFF': ['ZNIEFF de Type I', 'ZNIEFF de Type II'],
            'TABLEAU APPB': ['APPB'],
            'TABLEAU ENS': ['ENS'],
            'TABLEAU PNN': ['PNN', 'Parc National', 'PN'],
            'TABLEAU PRN': ['PRN', 'Parc Naturel Régional', 'PR'],
            'TABLEAU ZH': ['ZH'],
            'TABLEAU PELOUSES': ['Pelouses']
        }
        image_map = {
            'CARTE NATURA2000': 'Contexte éco - N2000__AE.png',
            'CARTE ZNIEFF': 'Contexte éco - ZNIEFF__AE.png',
            'CARTE APPB': 'Contexte éco - APPB__AE.png',
            'CARTE ENS': 'Contexte éco - ENS__AE.png',
            'CARTE PNN': 'Contexte éco - Parc National__AE.png',
            'CARTE PRN': 'Contexte éco - Parc Naturel Régional__AE.png',
            'CARTE ZH': 'Contexte éco - ZH avérées__AE.png',
            'CARTE PELOUSES': 'Contexte éco - Pelouses sèches__AE.png'
        }

        for para in list(doc.paragraphs):
            text = para.text.strip()
            if text.startswith('TABLEAU'):
                patterns = table_map.get(text)
                if not patterns:
                    patterns = [text.split(' ', 1)[1]] if ' ' in text else []
                sheets = find_sheets(patterns)
                dfs = [pd.read_excel(xls, sheet) for sheet in sheets]
                if dfs:
                    df = pd.concat(dfs, ignore_index=True)
                    self._insert_table_from_df(doc, para, df)
            elif text.startswith('CARTE'):
                img_name = image_map.get(text)
                if not img_name and ' ' in text:
                    theme = text.split(' ', 1)[1]
                    img_name = f"Contexte éco - {theme}__AE.png"
                if img_name:
                    img_path = os.path.join(OUT_IMG, img_name)
                    if os.path.isfile(img_path):
                        self._insert_image(para, img_path)
        doc.save(out_doc)

    def _insert_table_from_df(self, doc, paragraph, df):
        table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])
        table.style = 'Table Grid'
        for j, col in enumerate(df.columns):
            table.cell(0, j).text = str(col)
        for i, row in df.iterrows():
            for j, val in enumerate(row):
                table.cell(i + 1, j).text = str(val)
        parent = paragraph._p.getparent()
        idx = parent.index(paragraph._p)
        parent.insert(idx + 1, table._tbl)
        parent.remove(paragraph._p)

    def _insert_image(self, paragraph, img_path):
        img_par = paragraph.insert_paragraph_before()
        run = img_par.add_run()
        run.add_picture(img_path, width=Cm(16))
        img_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        parent = paragraph._p.getparent()
        parent.remove(paragraph._p)



# =========================
# Import de l'onglet Carto
# =========================
try:
    from .carto_tab import CartoTab
    CARTO_AVAILABLE = True
except ImportError as e:
    print(f"Onglet Carto non disponible: {e}")
    CARTO_AVAILABLE = False

# =========================
# Onglets PlantNet et Biodiv (stubs)
# =========================
class PlantNetTab(ttk.Frame):
    def __init__(self, parent, style_helper, prefs):
        super().__init__(parent)
        self.style_helper = style_helper
        self.prefs = prefs
        
        # Simple placeholder UI
        ttk.Label(self, text="Pl@ntNet - Identification des plantes", 
                 font=("Segoe UI", 15, "bold")).pack(pady=20)
        ttk.Label(self, text="Fonctionnalité à venir...", 
                 style="Card.TLabel").pack(pady=10)

class BiodivTab(ttk.Frame):
    def __init__(self, parent, style_helper, prefs):
        super().__init__(parent)
        self.style_helper = style_helper
        self.prefs = prefs
        
        # Simple placeholder UI
        ttk.Label(self, text="Biodiv'AURA - Base de données", 
                 font=("Segoe UI", 15, "bold")).pack(pady=20)
        ttk.Label(self, text="Fonctionnalité à venir...", 
                 style="Card.TLabel").pack(pady=10)

# =========================
# App principale avec Notebook
# =========================
class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.prefs = load_prefs()

        self.title("Bota-Logiciel | Assistant Contexte Éco & Identification")
        self.geometry(self.prefs.get("GEOMETRY", "1200x900"))

        self.style_helper = StyleHelper(self, self.prefs)
        self.style_helper.apply(self.prefs.get("THEME", "light"))

        self.notebook = ttk.Notebook(self)
        self.export_tab = ExportCartesTab(self.notebook, self.style_helper, self.prefs)
        self.plantnet_tab = PlantNetTab(self.notebook, self.style_helper, self.prefs)
        self.biodiv_tab = BiodivTab(self.notebook, self.style_helper, self.prefs)
        
        # Ajouter l'onglet Carto si disponible
        if CARTO_AVAILABLE:
            self.carto_tab = CartoTab(self.notebook, self.style_helper, self.prefs)
        else:
            self.carto_tab = None

        self.notebook.add(self.export_tab, text="Contexte Écologique & Cartes")
        if CARTO_AVAILABLE:
            self.notebook.add(self.carto_tab, text="  Carto  ")
        self.notebook.add(self.plantnet_tab, text="  Pl@ntNet  ")
        self.notebook.add(self.biodiv_tab, text="Biodiv'AURA")

        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)

        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _on_closing(self):
        try:
            g = self.geometry()
            if "x" in g:
                self.prefs["GEOMETRY"] = g
            save_prefs(self.prefs)
        except Exception:
            pass

        # Call the cleanup method for the shared browser
        if hasattr(self, 'export_tab'):
            self.export_tab._cleanup_driver()
        
        # Cleanup Carto tab resources
        if hasattr(self, 'carto_tab') and self.carto_tab:
            self.carto_tab.cleanup()

        self.destroy()


def launch():
    """Point d'entrée programmatique pour lancer l'application GUI.
    Utilisé par `Start.py` via `modules.main_app.launch()`.
    """
    app = Application()
    app.mainloop()


if __name__ == '__main__':
    launch()
















