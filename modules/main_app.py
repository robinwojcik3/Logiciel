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

from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed

import requests

from io import BytesIO

import pillow_heif

import zipfile

import traceback

import subprocess

import geopandas as gpd



# ==== Imports supplémentaires pour l'onglet Contexte éco ====

# Note: geopandas n'est pas utilisé directement dans ce module.

# Les traitements géospatiaux sont effectués dans des modules dédiés

# (ex: id_contexte_eco) afin d'éviter de charger des dépendances lourdes

# au lancement de l'UI principale.



# Import du scraper Wikipédia

from .wikipedia_scraper import DEP, get_wikipedia_extracts



# Import du worker QGIS externalisé

from .export_worker import worker_run





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
        # Zone haute défilante + console basse fixe (PlantNet)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        top_container = ttk.Frame(self)
        top_container.grid(row=0, column=0, sticky="nsew")
        canvas = tk.Canvas(top_container, highlightthickness=0, borderwidth=0)
        vscroll = ttk.Scrollbar(top_container, orient="vertical", command=canvas.yview)
        hscroll = ttk.Scrollbar(top_container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        top = ttk.Frame(canvas)
        _win = canvas.create_window((0, 0), window=top, anchor="nw")
        def _on_frame_config(_=None):
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                pass
        def _on_canvas_config(e):
            try:
                pass
            except Exception:
                pass
        top.bind("<Configure>", _on_frame_config)
        canvas.bind("<Configure>", _on_canvas_config)
        def _mw(e):
            try:
                delta = -1 * (e.delta // 120)
            except Exception:
                delta = -1 if getattr(e, 'num', 0) == 4 else (1 if getattr(e, 'num', 0) == 5 else 0)
            if delta:
                canvas.yview_scroll(delta, "units")
        canvas.bind_all("<MouseWheel>", _mw)
        canvas.bind_all("<Button-4>", _mw)
        canvas.bind_all("<Button-5>", _mw)
        # Layout root of the tab: top scrollable content + fixed bottom console
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        top_container = ttk.Frame(self)
        top_container.grid(row=0, column=0, sticky="nsew")
        canvas = tk.Canvas(top_container, highlightthickness=0, borderwidth=0)
        vscroll = ttk.Scrollbar(top_container, orient="vertical", command=canvas.yview)
        hscroll = ttk.Scrollbar(top_container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        top = ttk.Frame(canvas)
        _win = canvas.create_window((0, 0), window=top, anchor="nw")
        def _on_frame_config(_e=None):
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                pass
        def _on_canvas_config(e):
            try:
                pass
            except Exception:
                pass
        top.bind("<Configure>", _on_frame_config)
        canvas.bind("<Configure>", _on_canvas_config)
        def _mw(e):
            try:
                delta = -1 * (e.delta // 120)
            except Exception:
                delta = -1 if getattr(e, 'num', 0) == 4 else (1 if getattr(e, 'num', 0) == 5 else 0)
            if delta:
                canvas.yview_scroll(delta, "units")
        canvas.bind_all("<MouseWheel>", _mw)
        canvas.bind_all("<Button-4>", _mw)
        canvas.bind_all("<Button-5>", _mw)
        header = ttk.Frame(top, style="Header.TFrame", padding=(14, 12))

        header.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(header, text="Export cartes — QGIS ? PNG", style="Card.TLabel", font=self.font_title).grid(row=0, column=0, sticky="w")

        ttk.Label(header, text="Sélection shapefiles, choix du cadrage, export multi-projets.", style="Subtle.TLabel", font=self.font_sub).grid(row=1, column=0, sticky="w", pady=(4,0))

        header.columnconfigure(0, weight=1)



        grid = ttk.Frame(self); grid.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(grid);  left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        right = ttk.Frame(grid); right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        grid.columnconfigure(0, weight=1); grid.columnconfigure(1, weight=1); grid.rowconfigure(0, weight=1)



        # Shapefiles

        shp = ttk.Frame(left, style="Card.TFrame", padding=12); shp.pack(fill=tk.X)

        ttk.Label(shp, text="1. Couches Shapefile", style="Card.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")

        self._file_row(shp, 1, "?? Zone d'étude…", self.ze_shp_var, lambda: self._select_shapefile('ZE'))

        self._file_row(shp, 2, "?? Aire d'étude élargie…", self.ae_shp_var, lambda: self._select_shapefile('AE'))

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

        self.export_button = ttk.Button(act, text="â–¶ Lancer l’export", style="Accent.TButton", command=self.start_export_thread)

        self.export_button.grid(row=0, column=0, sticky="w")

        obtn = ttk.Button(act, text="?? Ouvrir le dossier de sortie", command=self._open_out_dir)

        obtn.grid(row=0, column=1, padx=(10,0)); ToolTip(obtn, OUT_IMG)

        tbtn = ttk.Button(act, text="?? Tester QGIS", command=self._test_qgis_threaded)

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

        canvas.configure(yscrollcommand=scrollbar.set, height=220)

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

        clear_btn = ttk.Button(parent, text="?", width=3, command=lambda: var.set(""))

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

            run_worker_subprocess([], cfg)

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

            # Désactive provisoirement le multiprocessing pour éviter les erreurs _multiprocessing

            # keep configured workers

            chunks = chunk_even(projets, workers)

            # Forcer au moins 2 workers pour utiliser ProcessPoolExecutor

            # (et donc le Python de QGIS configuré ci-dessous)

            workers = max(1, workers)

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



            # Pas de preflight multiprocessing: on utilise des sous-processus QGIS Python



            if workers <= 1:

                for chunk in chunks:

                    ok, ko = run_worker_subprocess(chunk, cfg)

                    ok_total += ok

                    ko_total += ko

                    self.after(0, ui_update_progress, ok + ko)

                    log_with_time(f"Lot terminé: {ok} OK, {ko} KO")

            else:

                ctx = None

                try:

                    import multiprocessing as mp

                    # Nettoyer l'environnement Python hérité pour éviter le mélange de versions (3.12/3.13)

                    for _k in ("PYTHONHOME", "PYTHONPATH", "PYTHONSTARTUP"):

                        if os.environ.get(_k):

                            try:

                                os.environ.pop(_k, None)

                                log_with_time(f"Env nettoye: unset {_k} pour les workers")

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



                    ctx = mp.get_context("spawn")

                    # Forcer l'utilisation du Python de QGIS pour les sous-processus (compat CPython/Qt)

                    try:

                        qgis_py = os.path.join(QGIS_ROOT, "apps", PY_VER, "python.exe")

                        if os.path.isfile(qgis_py):

                            ctx.set_executable(qgis_py)

                            log_with_time(f"MP exe: {qgis_py}")

                        else:

                            log_with_time(f"Python QGIS introuvable: {qgis_py}")

                    except Exception as e:

                        log_with_time(f"set_executable ï¿½chec: {e}")

                except Exception:

                    pass

                # Ajuster temporairement sys.path pour privilégier les libs QGIS

                old_syspath = list(sys.path)

                try:

                    qgis_py_root = os.path.join(QGIS_ROOT, "apps", PY_VER)

                    qgis_py_lib = os.path.join(qgis_py_root, "Lib")

                    qgis_dlls = os.path.join(qgis_py_root, "DLLs")

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

                    # Important: include QGIS DLLs so extension modules like _multiprocessing are found

                    sys.path = [qgis_py_root, qgis_py_lib, qgis_dlls, qgis_site, qgis_app_py] + [p for p in old_syspath if _keep_path(p)]

                except Exception as e:

                    log_with_time(f"sys.path cleanup skip: {e}")

                # Pool de threads qui pilotent des sous-processus QGIS Python (pas de multiprocessing)

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

                # Fallback séquentiel si aucun résultat (pool KO)

                if (ok_total + ko_total) == 0 and chunks:

                    log_with_time("Tous les workers ont échoué -> bascule en mode séquentiel")

                    for chunk in chunks:

                        try:

                            ok, ko = run_worker_subprocess(chunk, cfg)

                            ok_total += ok

                            ko_total += ko

                            self.after(0, ui_update_progress, ok + ko)

                            log_with_time(f"Lot terminé (fallback): {ok} OK, {ko} KO")

                        except Exception as e:

                            log_with_time(f"Erreur fallback: {e}")



            # Si aucun résultat n'a été produit (workers plantés), on retente en séquentiel

            if (ok_total + ko_total) == 0 and chunks:

                log_with_time("Tous les workers ont échoué — bascule en mode séquentiel…")

                for chunk in chunks:

                    try:

                        ok, ko = run_worker_subprocess(chunk, cfg)

                        ok_total += ok

                        ko_total += ko

                        self.after(0, ui_update_progress, ok + ko)

                        log_with_time(f"Lot terminé (fallback): {ok} OK, {ko} KO")

                    except Exception as e:

                        log_with_time(f"Erreur fallback: {e}")

            # Si aucun résultat n'a été produit (workers plantés), on retente en séquentiel

            if (ok_total + ko_total) == 0 and chunks:

                log_with_time("Tous les workers ont échoué — bascule en mode séquentiel…")

                for chunk in chunks:

                    try:

                        ok, ko = run_worker_subprocess(chunk, cfg)

                        ok_total += ok

                        ko_total += ko

                        self.after(0, ui_update_progress, ok + ko)

                        log_with_time(f"Lot terminé (fallback): {ok} OK, {ko} KO")

                    except Exception as e:

                        log_with_time(f"Erreur fallback: {e}")

            elapsed = datetime.datetime.now() - start

            log_with_time(f"FIN — OK={ok_total} | KO={ko_total} | Attendu={self.total_expected} | Durée={elapsed}")

            self.after(0, lambda: self.status_label.config(text=f"Terminé — OK={ok_total} / KO={ko_total}"))

        except Exception as e:

            log_with_time(f"Erreur critique: {e}")

            _err = str(e)

            self.after(0, lambda msg=_err: messagebox.showerror("Erreur", msg))

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
        # Zone haute défilante + console basse fixe
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        top_container = ttk.Frame(self)
        top_container.grid(row=0, column=0, sticky="nsew")
        canvas = tk.Canvas(top_container, highlightthickness=0, borderwidth=0)
        vscroll = ttk.Scrollbar(top_container, orient="vertical", command=canvas.yview)
        hscroll = ttk.Scrollbar(top_container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        top = ttk.Frame(canvas)
        _win = canvas.create_window((0, 0), window=top, anchor="nw")
        def _on_frame_config(_=None):
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                pass
        def _on_canvas_config(e):
            try:
                pass
            except Exception:
                pass
        top.bind("<Configure>", _on_frame_config)
        canvas.bind("<Configure>", _on_canvas_config)
        def _mw(e):
            try:
                delta = -1 * (e.delta // 120)
            except Exception:
                delta = -1 if getattr(e, 'num', 0) == 4 else (1 if getattr(e, 'num', 0) == 5 else 0)
            if delta:
                canvas.yview_scroll(delta, "units")
        canvas.bind_all("<MouseWheel>", _mw)
        canvas.bind_all("<Button-4>", _mw)
        canvas.bind_all("<Button-5>", _mw)

        header = ttk.Frame(top, style="Header.TFrame", padding=(14, 12))

        header.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(header, text="Identification Pl@ntNet", style="Card.TLabel", font=self.font_title).grid(row=0, column=0, sticky="w")

        ttk.Label(header, text="Analyse un dossier d'images via l'API Pl@ntNet.", style="Subtle.TLabel", font=self.font_sub).grid(row=1, column=0, sticky="w", pady=(4,0))

        header.columnconfigure(0, weight=1)



        card = ttk.Frame(top, style="Card.TFrame", padding=12)

        card.pack(fill=tk.X)

        ttk.Label(card, text="Dossier d'images", style="Card.TLabel").grid(row=0, column=0, sticky="w")

        row = ttk.Frame(card, style="Card.TFrame")

        row.grid(row=0, column=1, sticky="ew", padx=0)

        row.columnconfigure(0, weight=1)

        ttk.Entry(row, textvariable=self.folder_var).grid(row=0, column=0, sticky="ew")

        ttk.Button(row, text="Parcourir…", command=self._pick_folder).grid(row=0, column=1, padx=(6,0))

        card.columnconfigure(1, weight=1)



        act = ttk.Frame(top, style="Card.TFrame", padding=12)

        act.pack(fill=tk.X, pady=(10,0))

        self.run_btn = ttk.Button(act, text="â–¶ Lancer l'analyse", style="Accent.TButton", command=self._start_thread)

        self.run_btn.grid(row=0, column=0, sticky="w")

        obtn = ttk.Button(act, text="?? Ouvrir le dossier de sortie", command=self._open_out_dir)

        obtn.grid(row=0, column=1, padx=(10,0)); ToolTip(obtn, "Ouvrir le dossier cible")



        bottom = ttk.Frame(self, style="Card.TFrame", padding=12)
        bottom.grid(row=1, column=0, sticky="nsew", pady=(10,0))
        bottom.columnconfigure(0, weight=1)

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

        ttk.Label(header, text="Identification des zonages", style="Card.TLabel", font=self.font_title).grid(row=0, column=0, sticky="w")

        ttk.Label(header, text="Choisissez les shapefiles de référence puis lancez l'analyse.", style="Subtle.TLabel", font=self.font_sub).grid(row=1, column=0, sticky="w", pady=(4,0))

        header.columnconfigure(0, weight=1)



        card = ttk.Frame(self, style="Card.TFrame", padding=12)

        card.pack(fill=tk.X)

        self._file_row(card, 0, "?? Aire d'étude élargie…", self.ae_var, self._select_ae)

        self._file_row(card, 1, "?? Zone d'étude…", self.ze_var, self._select_ze)

        card.columnconfigure(1, weight=1)



        act = ttk.Frame(self, style="Card.TFrame", padding=12)

        act.pack(fill=tk.X, pady=(10,0))

        self.run_btn = ttk.Button(act, text="â–¶ Lancer l'analyse", style="Accent.TButton", command=self._start_thread)

        self.run_btn.grid(row=0, column=0, sticky="w")



        bottom = ttk.Frame(self, style="Card.TFrame", padding=12)
        bottom.grid(row=1, column=0, sticky="nsew", pady=(10,0))
        bottom.columnconfigure(0, weight=1)

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

        clear_btn = ttk.Button(parent, text="?", width=3, command=lambda: var.set(""))

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

        self.ze_shp_var   = tk.StringVar(value=from_long_unc(self.prefs.get("ZE_SHP", "")))

        self.ae_shp_var   = tk.StringVar(value=from_long_unc(self.prefs.get("AE_SHP", "")))

        self.cadrage_var   = tk.StringVar(value=self.prefs.get("CADRAGE_MODE", "BOTH"))

        self.overwrite_var = tk.BooleanVar(value=self.prefs.get("OVERWRITE", OVERWRITE_DEFAULT))

        self.dpi_var       = tk.IntVar(value=int(self.prefs.get("DPI", DPI_DEFAULT)))

        self.workers_var   = tk.IntVar(value=int(self.prefs.get("N_WORKERS", N_WORKERS_DEFAULT)))

        self.margin_var    = tk.DoubleVar(value=float(self.prefs.get("MARGIN_FAC", MARGIN_FAC_DEFAULT)))

        self.buffer_var    = tk.DoubleVar(value=float(self.prefs.get("ID_TAMPON_KM", 5.0)))

        self.out_dir_var   = tk.StringVar(value=self.prefs.get("OUT_DIR", OUT_IMG))

        self.export_type_var = tk.StringVar(value=self.prefs.get("EXPORT_TYPE", "BOTH"))

        # Resultats Wikipedia (affichage sous forme de tableau)

        self.wiki_climat_var = tk.StringVar(value="")

        self.wiki_occ_var = tk.StringVar(value="")

        self.wiki_last_url = ""

        self.wiki_query_var = tk.StringVar(value=self.prefs.get("WIKI_QUERY", ""))



        # Résultats Cartes végétation/sols

        self.veg_alt_var = tk.StringVar(value="")

        self.veg_veg_var = tk.StringVar(value="")

        self.veg_soil_var = tk.StringVar(value="")



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
        # Layout racine: contenu haut défilant (vertical uniquement) + console fixe en bas
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        top_container = ttk.Frame(self)
        top_container.grid(row=0, column=0, sticky="nsew")

        # Zone défilante verticale sans barre horizontale (tout le contenu
        # s'adapte en largeur au canevas pour éviter le défilement latéral)
        canvas = tk.Canvas(top_container, highlightthickness=0, borderwidth=0)
        vscroll = ttk.Scrollbar(top_container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        top = ttk.Frame(canvas)
        _win = canvas.create_window((0, 0), window=top, anchor="nw")

        def _on_frame_config(_e=None):
            # Met à jour la zone défilante lorsque le contenu change
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                pass

        def _on_canvas_config(e):
            # Force l'adaptation en largeur du contenu au canevas
            try:
                canvas.itemconfigure(_win, width=e.width)
            except Exception:
                pass

        top.bind("<Configure>", _on_frame_config)
        canvas.bind("<Configure>", _on_canvas_config)

        # Défilement à la molette uniquement quand la souris est sur le canevas
        def _mw(e):
            try:
                delta = -1 * (e.delta // 120)
            except Exception:
                delta = -1 if getattr(e, 'num', 0) == 4 else (1 if getattr(e, 'num', 0) == 5 else 0)
            if delta:
                canvas.yview_scroll(delta, "units")

        canvas.bind("<MouseWheel>", _mw)
        canvas.bind("<Button-4>", _mw)
        canvas.bind("<Button-5>", _mw)

        # Sélecteurs shapefiles

        shp = ttk.Frame(top, style="Card.TFrame", padding=12)

        shp.pack(fill=tk.X)

        ttk.Label(shp, text="Couches Shapefile", style="Card.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")

        self._file_row(shp, 1, "?? Zone d'étude…", self.ze_shp_var, self._select_ze)

        self._file_row(shp, 2, "?? Aire d'étude élargie…", self.ae_shp_var, self._select_ae)

        shp.columnconfigure(1, weight=1)



        # Encart Export cartes

        exp = ttk.Frame(top, style="Card.TFrame", padding=12)

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

        idf = ttk.Frame(top, style="Card.TFrame", padding=12)

        idf.pack(fill=tk.X, pady=(10,0))

        ttk.Label(idf, text="Tampon ZE (km)", style="Card.TLabel").grid(row=0, column=0, sticky="w")

        ttk.Spinbox(idf, from_=0.0, to=50.0, increment=0.5, textvariable=self.buffer_var, width=6, justify="right").grid(row=0, column=1, sticky="w", padx=(8,0))

        self.id_button = ttk.Button(idf, text="Lancer l’ID Contexte éco", style="Accent.TButton", command=self.start_id_thread)

        self.id_button.grid(row=0, column=2, sticky="w", padx=(12,0))



        self.wiki_button = ttk.Button(

            idf,

            text="Scraper Wikipedia",

            style="Accent.TButton",

            command=self.start_wiki_thread,

        )

        self.wiki_button.grid(row=0, column=3, sticky="w", padx=(12,0))


        # Nouveau bouton Photo Biodiv'AURA (a cote de Wikipedia)
        self.biodiv_button = ttk.Button(
            idf,
            text="Photo Biodiv'AURA",
            style="Accent.TButton",
            command=self.open_biodiv_dialog,
        )
        self.biodiv_button.grid(row=0, column=4, sticky="w", padx=(12,0))


        self.vegsol_button = ttk.Button(

            idf,

            text="Cartes végétation/sols",

            style="Accent.TButton",

            command=self.start_vegsol_thread,

        )

        self.vegsol_button.grid(row=0, column=5, sticky="w", padx=(12,0))



        self.rlt_button = ttk.Button(idf, text="Remonter le temps", style="Accent.TButton", command=self.start_rlt_thread)

        self.rlt_button.grid(row=0, column=6, sticky="w", padx=(12,0))

        self.maps_button = ttk.Button(idf, text="Ouvrir Google Maps", style="Accent.TButton", command=self.open_gmaps)

        self.maps_button.grid(row=0, column=7, sticky="w", padx=(12,0))

        self.bassin_button = ttk.Button(idf, text="Bassin versant", style="Accent.TButton", command=self.start_bassin_thread)

        self.bassin_button.grid(row=0, column=8, sticky="w", padx=(12,0))



        # Champ de requête Wikipedia optionnel (permet d'écrire la commune à la main)

        ttk.Label(idf, text="Commune (optionnel)", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=(8,0))

        wiki_q_row = ttk.Frame(idf)

        wiki_q_row.grid(row=1, column=1, columnspan=3, sticky="ew", pady=(8,0))

        wiki_q_row.columnconfigure(0, weight=1)

        ttk.Entry(wiki_q_row, textvariable=self.wiki_query_var).grid(row=0, column=0, sticky="ew")



        # Tableau Wikipedia (2 lignes, 2 colonnes)

        wiki_res = ttk.Frame(top, style="Card.TFrame", padding=12)

        wiki_res.pack(fill=tk.X, pady=(8,0))

        ttk.Label(wiki_res, text="Wikipedia", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=(0,6))

        # Bouton pour ouvrir l'article dans le navigateur

        self.wiki_open_button = ttk.Button(wiki_res, text="Ouvrir Wikipédia", command=self.open_wiki_url, state="disabled")

        self.wiki_open_button.grid(row=0, column=1, sticky="e", pady=(0,6))

        ttk.Label(wiki_res, text="Climat", style="Card.TLabel").grid(row=1, column=0, sticky="nw")

        ttk.Label(wiki_res, text="Corine Land Cover", style="Card.TLabel").grid(row=2, column=0, sticky="nw")

        # Cellules scrollables pour les textes

        clim_cell = ttk.Frame(wiki_res)

        clim_cell.grid(row=1, column=1, sticky="nsew")

        self.wiki_climat_txt = tk.Text(clim_cell, height=3, wrap=tk.WORD, state='disabled', relief='flat')

        clim_scroll = ttk.Scrollbar(clim_cell, orient="vertical", command=self.wiki_climat_txt.yview)

        self.wiki_climat_txt.configure(yscrollcommand=clim_scroll.set)

        self.wiki_climat_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        clim_scroll.pack(side=tk.RIGHT, fill=tk.Y)



        occ_cell = ttk.Frame(wiki_res)

        occ_cell.grid(row=2, column=1, sticky="nsew")

        self.wiki_occ_txt = tk.Text(occ_cell, height=3, wrap=tk.WORD, state='disabled', relief='flat')

        occ_scroll = ttk.Scrollbar(occ_cell, orient="vertical", command=self.wiki_occ_txt.yview)

        self.wiki_occ_txt.configure(yscrollcommand=occ_scroll.set)

        self.wiki_occ_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        occ_scroll.pack(side=tk.RIGHT, fill=tk.Y)



        wiki_res.columnconfigure(1, weight=1)



        # Tableau Cartes végétation/sols (3 lignes)

        vegsol_res = ttk.Frame(wiki_res, style="Card.TFrame", padding=4)

        vegsol_res.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(6,0))

        ttk.Label(vegsol_res, text="Cartes végétation/sols", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=(0,6))

        ttk.Label(vegsol_res, text="Altitude", style="Card.TLabel").grid(row=1, column=0, sticky="nw")

        ttk.Label(vegsol_res, text="Végétation", style="Card.TLabel").grid(row=2, column=0, sticky="nw")

        ttk.Label(vegsol_res, text="Sols", style="Card.TLabel").grid(row=3, column=0, sticky="nw")



        alt_cell = ttk.Frame(vegsol_res); alt_cell.grid(row=1, column=1, sticky="nsew")

        self.veg_alt_txt = tk.Text(alt_cell, height=2, wrap=tk.WORD, state='disabled', relief='flat')

        alt_scroll = ttk.Scrollbar(alt_cell, orient="vertical", command=self.veg_alt_txt.yview)

        self.veg_alt_txt.configure(yscrollcommand=alt_scroll.set)

        self.veg_alt_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        alt_scroll.pack(side=tk.RIGHT, fill=tk.Y)



        veg_cell = ttk.Frame(vegsol_res); veg_cell.grid(row=2, column=1, sticky="nsew")

        self.veg_veg_txt = tk.Text(veg_cell, height=3, wrap=tk.WORD, state='disabled', relief='flat')

        veg_scroll = ttk.Scrollbar(veg_cell, orient="vertical", command=self.veg_veg_txt.yview)

        self.veg_veg_txt.configure(yscrollcommand=veg_scroll.set)

        self.veg_veg_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        veg_scroll.pack(side=tk.RIGHT, fill=tk.Y)



        soil_cell = ttk.Frame(vegsol_res); soil_cell.grid(row=3, column=1, sticky="nsew")

        self.veg_soil_txt = tk.Text(soil_cell, height=3, wrap=tk.WORD, state='disabled', relief='flat')

        soil_scroll = ttk.Scrollbar(soil_cell, orient="vertical", command=self.veg_soil_txt.yview)

        self.veg_soil_txt.configure(yscrollcommand=soil_scroll.set)

        self.veg_soil_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        soil_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        vegsol_res.columnconfigure(1, weight=1)



        # Console + progression

        bottom = ttk.Frame(self, style="Card.TFrame", padding=12)
        bottom.grid(row=1, column=0, sticky="nsew", pady=(10,0))
        bottom.columnconfigure(0, weight=1)

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



        obtn = ttk.Button(bottom, text="?? Ouvrir le dossier de sortie", command=self._open_out_dir)
        obtn.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8,0))
        try:
            ToolTip(obtn, "Ouvrir le dossier cible")
        except Exception:
            pass



    # ---------- Helpers UI ----------

    def _file_row(self, parent, row: int, label: str, var: tk.StringVar, cmd):

        btn = ttk.Button(parent, text=label, command=cmd)

        btn.grid(row=row, column=0, sticky="w", pady=(8 if row == 1 else 4, 2))

        ent = ttk.Entry(parent, textvariable=var, width=10)

        ent.grid(row=row, column=1, sticky="ew", padx=8)

        ent.configure(state="readonly")

        clear_btn = ttk.Button(parent, text="?", width=3, command=lambda: var.set(""))

        clear_btn.grid(row=row, column=2, sticky="e")

        parent.columnconfigure(1, weight=1)



    def _select_ze(self):

        base = self.ze_shp_var.get() or os.path.expanduser("~")

        path = filedialog.askopenfilename(title="Sélectionner la zone d'étude",

                                          initialdir=base if os.path.isdir(base) else os.path.expanduser("~"),

                                          filetypes=[("Shapefile ESRI", "*.shp")])

        if path:
            # Normaliser pour l'affichage (sans préfixe long UNC)
            self.ze_shp_var.set(os.path.normpath(path))



    def _select_ae(self):

        base = self.ae_shp_var.get() or os.path.expanduser("~")

        path = filedialog.askopenfilename(title="Sélectionner l'aire d'étude élargie",

                                          initialdir=base if os.path.isdir(base) else os.path.expanduser("~"),

                                          filetypes=[("Shapefile ESRI", "*.shp")])

        if path:
            # Normaliser pour l'affichage (sans préfixe long UNC)
            self.ae_shp_var.set(os.path.normpath(path))



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



    def start_wiki_thread(self):

        if (not self.wiki_query_var.get().strip()) and (not self.ze_shp_var.get().strip()):

            messagebox.showerror("Erreur", "Sélectionner la Zone d'étude ou saisir une commune.")

            return

        print("[Wiki] Bouton Wikipédia cliqué", file=self.stdout_redirect)

        self.wiki_button.config(state="disabled")

        # Réinitialiser le tableau et le bouton avant un nouveau scraping

        try:

            self.wiki_climat_var.set("")

            self.wiki_occ_var.set("")

            self.wiki_last_url = ""

            if hasattr(self, 'wiki_open_button'):

                self.wiki_open_button.config(state="disabled")

        except Exception:

            pass

        t = threading.Thread(target=self._run_wiki)

        t.daemon = True

        t.start()



    def open_wiki_url(self) -> None:

        try:

            url = (self.wiki_last_url or "").strip()

            if not url:

                messagebox.showinfo("Wikipedia", "Aucune URL disponible. Lancez d'abord le scraping.")

                return

            print(f"[Wiki] Ouverture dans le navigateur : {url}", file=self.stdout_redirect)

            webbrowser.open(url)

        except Exception as e:

            messagebox.showerror("Wikipedia", f"Impossible d'ouvrir l'URL : {e}")



    def _run_wiki(self):

        try:

            print("[Wiki] Lancement du scraping Wikipedia", file=self.stdout_redirect)

            ze_path = self.ze_shp_var.get()

            gdf = gpd.read_file(ze_path)

            if gdf.crs is None:

                raise ValueError("CRS non défini")

            gdf = gdf.to_crs("EPSG:4326")

            centroid = gdf.geometry.unary_union.centroid

            lat, lon = centroid.y, centroid.x

            manual = (self.wiki_query_var.get() or "").strip()

            if manual:

                query = manual

                try:

                    self.prefs["WIKI_QUERY"] = manual; save_prefs(self.prefs)

                except Exception:

                    pass

            else:

                commune, dep = self._detect_commune(lat, lon)

                query = f"{commune} {dep}".strip()

            print(f"[Wiki] Requête : {query}", file=self.stdout_redirect)

            data = get_wikipedia_extracts(query)

            # Mettre à jour le tableau Wikipedia (dès que les données sont disponibles)

            self._update_wiki_table(data)

            if "error" in data:

                print(f"[Wiki] {data['error']}", file=self.stdout_redirect)

            else:

                print(f"[Wiki] Page Wikipédia : {data['url']}", file=self.stdout_redirect)

                print("[Wiki] CLIMAT :", file=self.stdout_redirect)

                if data['climat_p1'] != 'Non trouvé':

                    print(data['climat_p1'], file=self.stdout_redirect)

                if data['climat_p2'] != 'Non trouvé':

                    print(data['climat_p2'], file=self.stdout_redirect)

                print("[Wiki] OCCUPATION DES SOLS :", file=self.stdout_redirect)

                if data['occupation_p1'] != 'Non trouvé':

                    print(data['occupation_p1'], file=self.stdout_redirect)



        except Exception as e:

            print(f"[Wiki] Erreur : {e}", file=self.stdout_redirect)

        finally:

            self.after(0, lambda: self.wiki_button.config(state="normal"))



    def _update_wiki_table(self, data: dict) -> None:

        try:

            clim_txt = data.get('climat_p1', '')

            occ_txt = data.get('occupation_p1', '')

            url_txt = data.get('url', '')

            # Compat: utiliser les nouvelles clés si présentes

            clim_txt2 = data.get('climat') or clim_txt

            occ_txt2 = data.get('occup_sols') or occ_txt

            def _norm(s):

                try:

                    return s if (isinstance(s, str) and not s.lower().startswith('non trouv')) else ''

                except Exception:

                    return ''

            # Mettre à jour aussi les zones scrollables

            def _fill(widget, s):

                try:

                    widget.delete('1.0', tk.END)

                    s2 = s if (isinstance(s, str) and not s.lower().startswith('non trouv')) else 'Non trouvé'

                    widget.insert(tk.END, s2)

                    widget.config(state='disabled')

                    #

                except Exception:

                    pass

            self.after(0, lambda: _fill(self.wiki_climat_txt, clim_txt2 or ''))

            self.after(0, lambda: _fill(self.wiki_occ_txt, occ_txt2 or ''))

            self.after(0, lambda: self.wiki_climat_var.set(_norm(clim_txt2)))

            self.after(0, lambda: self.wiki_occ_var.set(_norm(occ_txt2)))

            # Mettre à jour l'URL et l'état du bouton d'ouverture

            def _upd_url():

                try:

                    self.wiki_last_url = url_txt or ""

                    if hasattr(self, 'wiki_open_button'):

                        self.wiki_open_button.config(state=("normal" if self.wiki_last_url else "disabled"))

                except Exception:

                    pass

            self.after(0, _upd_url)

        except Exception:

            pass



    def start_vegsol_thread(self):

        if not self.ze_shp_var.get().strip():

            messagebox.showerror("Erreur", "Sélectionner la Zone d'étude.")

            return

        print("[Cartes] Bouton cartes cliqué", file=self.stdout_redirect)

        self.vegsol_button.config(state="disabled")

        t = threading.Thread(target=self._run_vegsol)

        t.daemon = True

        t.start()



    def _run_vegsol(self):

        try:

            print("[Cartes] Lancement du scraping des cartes", file=self.stdout_redirect)

            ze_path = self.ze_shp_var.get()

            gdf = gpd.read_file(ze_path)

            if gdf.crs is None:

                raise ValueError("CRS non défini")

            gdf = gdf.to_crs("EPSG:4326")

            centroid = gdf.geometry.unary_union.centroid

            lat, lon = centroid.y, centroid.x

            coords_dms = dd_to_dms(lat, lon)

            options = webdriver.ChromeOptions()

            options.add_experimental_option("excludeSwitches", ["enable-logging"])

            options.add_argument("--log-level=3")

            options.add_argument("--disable-extensions")

            options.add_argument("--disable-gpu")

            options.add_argument("--no-sandbox")

            options.add_argument("--disable-dev-shm-usage")

            # Respect APP_HEADLESS env var (default: visible)

            try:

                if os.environ.get("APP_HEADLESS", "0").lower() in ("1", "true", "yes"):

                    options.add_argument("--headless=new")

            except Exception:

                try:

                    if os.environ.get("APP_HEADLESS", "0").lower() in ("1", "true", "yes"):

                        options.add_argument("--headless")

                except Exception:

                    pass

            # Driver local si présent

            local_driver = os.path.join(REPO_ROOT if 'REPO_ROOT' in globals() else os.path.abspath(os.path.join(os.path.dirname(__file__), '..')), 'tools', 'chromedriver.exe')

            if os.path.isfile(local_driver):

                self.vegsol_driver = webdriver.Chrome(service=Service(local_driver), options=options)

            else:

                self.vegsol_driver = webdriver.Chrome(options=options)

            self.vegsol_driver.maximize_window()



            # Nouveau flux automatisé (import shapefile + scraping pop-up)

            try:

                wait = WebDriverWait(self.vegsol_driver, 10)

                # 1) Ouvrir l'URL

                self.vegsol_driver.get("https://floreapp.netlify.app/biblio-patri.html")

                time.sleep(0.75)



                # 3) Cliquer sur Importer shapefile

                try:

                    btn_upload = wait.until(EC.element_to_be_clickable((By.ID, "upload-shapefile-btn")))

                    btn_upload.click()

                except Exception:

                    pass

                time.sleep(0.75)



                # 5) Cliquer sur Zone d’étude

                try:

                    btn_zone = wait.until(EC.element_to_be_clickable((By.ID, "import-zone-btn")))

                    btn_zone.click()

                except Exception:

                    pass



                # Préparer la liste des fichiers du shapefile

                def _from_long_unc(p: str) -> str:

                    p = p or ""

                    if p.startswith("\\\\?\\UNC"):

                        return "\\\\" + p[8:]

                    if p.startswith("\\\\?\\"):

                        return p[4:]

                    return p



                ze_shp = (self.ze_shp_var.get() or "").strip()

                base_no_ext, _ = os.path.splitext(_from_long_unc(ze_shp))

                exts = [".cpg", ".dbf", ".prj", ".qmd", ".shp", ".shx"]

                files = [base_no_ext + e for e in exts if os.path.isfile(base_no_ext + e)]

                if not files:

                    raise ValueError("Fichiers du shapefile introuvables pour l'import")



                # Envoyer les fichiers à l'input[type=file]

                inputs = self.vegsol_driver.find_elements(By.CSS_SELECTOR, "input[type='file']")

                target_input = inputs[-1] if inputs else None

                if not target_input:

                    raise RuntimeError("Champ d'import fichier introuvable")

                try:

                    target_input.send_keys("\n".join(files))

                except Exception as e:

                    print(f"[Cartes] Envoi fichiers échoué: {e}", file=self.stdout_redirect)



                time.sleep(0.75)



                # 8) Clic droit au centre de la carte

                map_el = wait.until(EC.visibility_of_element_located((By.ID, "map")))

                ActionChains(self.vegsol_driver).move_to_element(map_el).context_click(map_el).perform()

                time.sleep(0.5)



                # 10) Cliquer sur 'Ressources'

                try:

                    btn_res = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Ressources')]")))

                    btn_res.click()

                except Exception:

                    pass



                time.sleep(0.75)



                # 12) Scraper les éléments de la pop-up

                # Attendre la présence des éléments de la pop-up

                try:

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.altitude-info")))

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.tooltip-pill.vegetation-pill")))

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.tooltip-pill.soil-pill")))

                except Exception:

                    pass



                html = self.vegsol_driver.page_source

                soup = BeautifulSoup(html, "lxml")

                def _txt(sel):

                    el = soup.select_one(sel)

                    return el.get_text(" ", strip=True) if el else "Non trouvé"

                alt = _txt("div.altitude-info")

                veg = _txt("div.tooltip-pill.vegetation-pill")

                soil = _txt("div.tooltip-pill.soil-pill")

                self._update_vegsol_table({"altitude": alt, "vegetation": veg, "sol": soil})

                return

            except Exception as flow_err:

                print(f"[Cartes] Flux import/scraping échoué: {flow_err}", file=self.stdout_redirect)



            def _open_layer(layer_label: str) -> None:

                try:

                    wait = WebDriverWait(self.vegsol_driver, 0.5)

                    self.vegsol_driver.execute_script(

                        "window.open('https://floreapp.netlify.app/biblio-patri.html','_blank');"

                    )

                    self.vegsol_driver.switch_to.window(self.vegsol_driver.window_handles[-1])

                    addr = wait.until(EC.element_to_be_clickable((By.ID, "address-input")))

                    addr.click()

                    addr.clear()

                    addr.send_keys(coords_dms)

                    wait.until(

                        EC.element_to_be_clickable((By.ID, "search-address-btn"))

                    ).click()

                    wait.until(

                        EC.element_to_be_clickable(

                            (By.CSS_SELECTOR, "a.leaflet-control-layers-toggle")

                        )

                    ).click()

                    checkbox = wait.until(

                        EC.element_to_be_clickable(

                            (By.XPATH, f"//label[contains(.,'{layer_label}')]/input")

                        )

                    )

                    if not checkbox.is_selected():

                        checkbox.click()

                except Exception as fe:

                    print(

                        f"[Cartes] Étapes {layer_label} échouées : {fe}",

                        file=self.stdout_redirect,

                    )



            _open_layer("Carte de la végétation")

            _open_layer("Carte des sols")

        except Exception as e:

            print(f"[Cartes] Erreur : {e}", file=self.stdout_redirect)

        finally:

            self.after(0, lambda: self.vegsol_button.config(state="normal"))



    def _update_vegsol_table(self, data: dict) -> None:

        try:

            alt = data.get('altitude', '')

            veg = data.get('vegetation', '')

            soil = data.get('sol', '')

            def _fill(widget: tk.Text, s: str):

                try:

                    widget.config(state='normal')

                    widget.delete('1.0', tk.END)

                    s2 = s if (isinstance(s, str) and s.strip()) else 'Non trouvé'

                    widget.insert(tk.END, s2)

                    widget.config(state='disabled')

                except Exception:

                    pass

            self.after(0, lambda: _fill(self.veg_alt_txt, alt))

            self.after(0, lambda: _fill(self.veg_veg_txt, veg))

            self.after(0, lambda: _fill(self.veg_soil_txt, soil))

        except Exception:

            pass



    # --- Boutons ajoutés ---


    def open_biodiv_dialog(self) -> None:
        try:
            if hasattr(self, 'biodiv_win') and self.biodiv_win and tk.Toplevel.winfo_exists(self.biodiv_win):
                self.biodiv_win.lift()
                return
        except Exception:
            pass

        win = tk.Toplevel(self)
        win.title("Photo Biodiv'AURA")
        win.geometry("640x520")
        self.biodiv_win = win

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Collez 1 espèce par ligne (20 max)", style="Card.TLabel").pack(anchor='w')

        txt = tk.Text(frm, height=20, wrap=tk.NONE)
        txt.pack(fill=tk.BOTH, expand=True, pady=(6,6))
        self.biodiv_text = txt

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X)

        self.biodiv_launch_btn = ttk.Button(btns, text="Lancer le scraping", style="Accent.TButton",
                                            command=self.start_biodiv_thread)
        self.biodiv_launch_btn.pack(side=tk.LEFT)

        ttk.Button(btns, text="Fermer", command=win.destroy).pack(side=tk.RIGHT)

    def start_biodiv_thread(self) -> None:
        try:
            raw = self.biodiv_text.get('1.0', tk.END)
        except Exception:
            raw = ''

        species = [s.strip() for s in (raw.splitlines() if raw else [])]
        species = [s for s in species if s]
        if not species:
            messagebox.showerror("Biodiv'AURA", "Veuillez saisir au moins une espèce.")
            return

        if len(species) > 20:
            species = species[:20]

        print(f"[Biodiv] Lancement pour {len(species)} espèce(s)", file=self.stdout_redirect)

        try:
            if hasattr(self, 'biodiv_launch_btn'):
                self.biodiv_launch_btn.config(state='disabled')
        except Exception:
            pass

        t = threading.Thread(target=self._run_biodiv, args=(species,))
        t.daemon = True
        t.start()

    def _run_biodiv(self, species_list: list[str]) -> None:
        try:
            os.makedirs(OUT_IMG, exist_ok=True)
            out_dir = os.path.join(OUT_IMG, "Photo BiodivAURA")
            os.makedirs(out_dir, exist_ok=True)

            # Document unique pour toutes les especes
            doc = Document()
            try:
                style_normal = doc.styles['Normal']
                style_normal.font.name = 'Calibri'
                style_normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
            except Exception:
                pass

            options = webdriver.ChromeOptions()
            options.add_experimental_option("excludeSwitches", ["enable-logging"])
            options.add_argument("--log-level=3")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            try:
                if os.environ.get("APP_HEADLESS", "0").lower() in ("1", "true", "yes"):
                    options.add_argument("--headless=new")
            except Exception:
                pass

            local_driver = os.path.join(REPO_ROOT if 'REPO_ROOT' in globals() else os.path.abspath(os.path.join(os.path.dirname(__file__), '..')), 'tools', 'chromedriver.exe')
            if os.path.isfile(local_driver):
                driver = webdriver.Chrome(service=Service(local_driver), options=options)
            else:
                driver = webdriver.Chrome(options=options)

            wait = WebDriverWait(driver, 10)
            print(f"[Biodiv] Scraping de {len(species_list)} espece(s)...", file=self.stdout_redirect)

            # Collecter toutes les données d'espèces avec leurs images
            species_data = []
            WAIT_SHORT = 1.0  # secondes
            
            for idx, sp in enumerate(species_list, start=1):
                try:
                    print(f"[Biodiv] ({idx}/{len(species_list)}) {sp}", file=self.stdout_redirect)
                    driver.get("https://atlas.biodiversite-auvergne-rhone-alpes.fr/")

                    inp = wait.until(EC.element_to_be_clickable((By.ID, "searchTaxons")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", inp)
                    time.sleep(WAIT_SHORT)

                    inp.clear()
                    inp.send_keys(sp)
                    time.sleep(WAIT_SHORT)

                    # Cliquer sur le premier resultat
                    try:
                        first_result = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".search-results .result-item:first-child")))
                        driver.execute_script("arguments[0].click();", first_result)
                        time.sleep(WAIT_SHORT * 2)
                    except Exception:
                        print(f"[Biodiv] Pas de resultat pour {sp}", file=self.stdout_redirect)
                        species_data.append({'name': sp, 'image_path': None, 'url': None})
                        continue

                    # Chercher une image
                    img_bytes = None
                    tmp_path = None
                    try:
                        img_elem = driver.find_element(By.CSS_SELECTOR, ".species-photo img, .photo-gallery img, img[src*='photo'], img[src*='image']")
                        img_url = img_elem.get_attribute("src")
                        if img_url and img_url.startswith("http"):
                            print(f"[Biodiv] Image trouvee: {img_url[:80]}...", file=self.stdout_redirect)
                            r = requests.get(img_url, timeout=10)
                            if r.ok:
                                img_bytes = r.content
                                tmp_path = os.path.join(tempfile.gettempdir(), f"biodiv_{int(time.time()*1000)}_{idx}.jpg")
                                with open(tmp_path, 'wb') as f:
                                    f.write(img_bytes)
                    except Exception as de:
                        print(f"[Biodiv] Download echoue pour {sp}: {de}", file=self.stdout_redirect)

                    # Récupérer l'URL de la page espèce
                    try:
                        url_sp = driver.current_url
                    except Exception:
                        url_sp = None

                    species_data.append({
                        'name': sp,
                        'image_path': tmp_path,
                        'url': url_sp
                    })

                except Exception as sp_err:
                    print(f"[Biodiv] Erreur espece {sp}: {sp_err}", file=self.stdout_redirect)
                    species_data.append({'name': sp, 'image_path': None, 'url': None})

            try:
                driver.quit()
            except Exception:
                pass

            # Créer le tableau avec les images (2x3 format)
            self._create_species_table(doc, species_data)

            ts = datetime.datetime.now().strftime('%Y%m%d-%H%M')
            doc_path = os.path.join(out_dir, f"Photos_BiodivAURA_{ts}.docx")
            try:
                doc.save(doc_path)
                print(f"[Biodiv] Document genere: {doc_path}", file=self.stdout_redirect)
            except Exception as se:
                print(f"[Biodiv] Sauvegarde echouee: {se}", file=self.stdout_redirect)

            self._set_status(f"Termine - {doc_path}")

        except Exception as e:
            print(f"[Biodiv] Erreur: {e}", file=self.stdout_redirect)
        finally:
            try:
                if hasattr(self, 'biodiv_launch_btn'):
                    self.after(0, lambda: self.biodiv_launch_btn.config(state='normal'))
            except Exception:
                pass

    def _create_species_table(self, doc, species_data):
        """Crée un tableau 2x3 avec les images d'espèces comme dans le screenshot"""
        if not species_data:
            return

        # Traiter les espèces par groupes de 6 (2 lignes x 3 colonnes)
        for page_start in range(0, len(species_data), 6):
            page_species = species_data[page_start:page_start + 6]
            
            # Calculer le nombre de lignes nécessaires
            rows_needed = (len(page_species) + 2) // 3  # +2 pour arrondir vers le haut
            
            # Créer le tableau
            table = doc.add_table(rows=rows_needed, cols=3)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            # Configurer le style du tableau
            table.style = 'Table Grid'
            
            # Remplir le tableau
            for i, species in enumerate(page_species):
                row_idx = i // 3
                col_idx = i % 3
                cell = table.cell(row_idx, col_idx)
                
                # Vider la cellule
                cell.text = ''
                
                # Ajouter l'image
                if species['image_path'] and os.path.exists(species['image_path']):
                    try:
                        # Créer un paragraphe pour l'image
                        img_paragraph = cell.paragraphs[0]
                        img_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = img_paragraph.runs[0] if img_paragraph.runs else img_paragraph.add_run()
                        
                        # Ajuster la taille de l'image pour le tableau (plus petite)
                        img_width = Cm(4.5)  # Largeur réduite pour s'adapter au tableau
                        run.add_picture(species['image_path'], width=img_width)
                        
                    except Exception as e:
                        print(f"[Biodiv] Erreur ajout image {species['name']}: {e}", file=self.stdout_redirect)
                        cell.paragraphs[0].add_run("[Image non disponible]")
                else:
                    cell.paragraphs[0].add_run("[Image non disponible]")
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Ajouter le nom de l'espèce en dessous
                name_paragraph = cell.add_paragraph()
                name_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                name_run = name_paragraph.add_run(species['name'])
                name_run.bold = True
                
                # Ajouter le lien si disponible
                if species['url']:
                    try:
                        link_paragraph = cell.add_paragraph()
                        link_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        add_hyperlink(link_paragraph, species['url'], "Voir sur Biodiv'AURA", italic=True)
                    except Exception:
                        pass
            
            # Ajuster la largeur des cellules
            for row in table.rows:
                for cell in row.cells:
                    # Définir la largeur des cellules
                    cell.width = Cm(5.5)
            
            # Ajouter un saut de page si ce n'est pas la dernière page
            if page_start + 6 < len(species_data):
                doc.add_page_break()

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



    def start_rlt_thread(self):
        if not self.ze_shp_var.get().strip():
            messagebox.showerror("Erreur", "Sélectionner la Zone d'étude.")
            return
        self.rlt_button.config(state="disabled")
        t = threading.Thread(target=self._run_rlt)
        t.daemon = True
        t.start()


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

                print(f"[IGN] {title} ? {url}", file=self.stdout_redirect)

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

                print("[IGN] Aucune image ? pas de doc.", file=self.stdout_redirect)

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

        self.id_button.config(state="disabled")

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

            from .id_contexte_eco import run_analysis as run_id_context

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

        self.vegsol_driver = None



        # Header global + bouton thème

        top = ttk.Frame(root, style="Header.TFrame", padding=(12, 8))

        top.pack(fill=tk.X)

        ttk.Label(top, text='Contexte éco — Suite d’outils', style='Card.TLabel',

                  font=tkfont.Font(family="Segoe UI", size=16, weight="bold")).pack(side=tk.LEFT)

        btn_theme = ttk.Button(top, text='Changer de thème', command=self._toggle_theme)

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






















