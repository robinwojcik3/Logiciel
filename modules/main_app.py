#!/usr/bin/env "C:/Program Files/QGIS 3.40.3/apps/Python312/python.exe"
# -*- coding: utf-8 -*-
r"""
Application à onglets (ttk.Notebook) :

Onglet « Contexte éco » :
    - Export des mises en page QGIS en PNG.
    - Identification des zonages (ID Contexte éco) avec tampon configurable.
    - Sélection commune des shapefiles ZE/AE et console partagée.

Onglet « Remonter le temps & Bassin versant » :
    - Reprend le workflow IGN (capture + rapport Word) mais
      DEMANDE à l’utilisateur :
        • Coordonnées en DMS (LAT puis LON, ex : 45°09'30" N 5°43'12" E)
      La commune est désormais détectée automatiquement au lieu d'être saisie.
    - Options : dossier de sortie, headless, tempo de chargement.
    - Produit un Word 2×2 avec les vues temporelles, + commentaire.

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
from .utils import (log_with_time, normalize_name, to_long_unc, chunk_even, load_prefs, save_prefs, TextRedirector, ToolTip, StyleHelper, OUT_IMG)
from .tabs.remonter import RemonterLeTempsTab
from .tabs.plantnet import PlantNetTab
from concurrent.futures import ProcessPoolExecutor, as_completed
import zipfile
import traceback


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


# Dossier par défaut pour la sélection des shapefiles (onglet 1)
DEFAULT_SHAPE_DIR = r"C:\Users\utilisateur\Mon Drive\1 - Bota & Travail\+++++++++  BOTA  +++++++++\---------------------- 2) CARTO terrain"

# QGIS
QGIS_ROOT = r"C:\Program Files\QGIS 3.40.3"
QGIS_APP  = os.path.join(QGIS_ROOT, "apps", "qgis")
PY_VER    = "Python312"


# =========================
# Fonctions QGIS (onglet 1) — inchangées
# =========================
def worker_run(args: Tuple[List[str], dict]) -> Tuple[int, int]:
    projects, cfg = args

    os.environ["OSGEO4W_ROOT"] = cfg["QGIS_ROOT"]
    os.environ["QGIS_PREFIX_PATH"] = cfg["QGIS_APP"]
    os.environ.setdefault("GDAL_DATA", os.path.join(cfg["QGIS_ROOT"], "share", "gdal"))
    os.environ.setdefault("PROJ_LIB", os.path.join(cfg["QGIS_ROOT"], "share", "proj"))
    os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")

    qt_base = None
    for name in ("Qt6", "Qt5"):
        base = os.path.join(cfg["QGIS_ROOT"], "apps", name)
        if os.path.isdir(base):
            qt_base = base
            break
    if qt_base is None:
        raise RuntimeError("Qt introuvable")

    platform_dir = os.path.join(qt_base, "plugins", "platforms")
    os.environ["QT_PLUGIN_PATH"] = os.path.join(qt_base, "plugins")
    os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = platform_dir
    qpa = "windows" if os.path.isfile(os.path.join(platform_dir, "qwindows.dll")) \
        else ("minimal" if os.path.isfile(os.path.join(platform_dir, "qminimal.dll")) else "offscreen")
    os.environ["QT_QPA_PLATFORM"] = qpa

    os.environ["PATH"] = os.pathsep.join([
        os.path.join(qt_base, "bin"),
        os.path.join(cfg["QGIS_APP"], "bin"),
        os.path.join(cfg["QGIS_ROOT"], "bin"),
        os.environ.get("PATH", ""),
    ])

    sys.path.insert(0, os.path.join(cfg["QGIS_APP"], "python"))
    sys.path.insert(0, os.path.join(cfg["QGIS_ROOT"], "apps", cfg["PY_VER"], "Lib", "site-packages"))

    from qgis.core import (
        QgsApplication, QgsProject, QgsLayoutExporter, QgsLayoutItemMap, QgsRectangle,
        QgsCoordinateTransform
    )

    qgs = QgsApplication([], False)
    qgs.setPrefixPath(cfg["QGIS_APP"], True)
    qgs.initQgis()

    ok = 0
    ko = 0

    def adjust_extent_to_item_ratio(ext: 'QgsRectangle', target_ratio: float, margin: float) -> 'QgsRectangle':
        if ext.width() <= 0 or ext.height() <= 0:
            return ext
        cx, cy = ext.center().x(), ext.center().y()
        w, h = ext.width(), ext.height()
        if (w / h) > target_ratio:
            new_h = w / target_ratio
            dh = (new_h - h) / 2.0
            xmin, xmax = ext.xMinimum(), ext.xMaximum()
            ymin, ymax = cy - h/2.0 - dh, cy + h/2.0 + dh
        else:
            new_w = h * target_ratio
            dw = (new_w - w) / 2.0
            ymin, ymax = ext.yMinimum(), ext.yMaximum()
            xmin, xmax = cx - w/2.0 - dw, cx + w/2.0 + dw
        cw, ch = (xmax - xmin), (ymax - ymin)
        mx, my = (margin - 1.0) * cw / 2.0, (margin - 1.0) * ch / 2.0
        return QgsRectangle(xmin - mx, ymin - my, xmax + mx, ymax + my)

    def extent_in_project_crs(prj: 'QgsProject', lyr) -> Optional['QgsRectangle']:
        ext = lyr.extent()
        try:
            if lyr.crs() != prj.crs():
                ct = QgsCoordinateTransform(lyr.crs(), prj.crs(), prj)
                ext = ct.transformBoundingBox(ext)
        except Exception:
            pass
        return ext

    def apply_extent_and_export(layout, lyr_extent: 'QgsRectangle', out_png: str) -> bool:
        maps = [it for it in layout.items() if isinstance(it, QgsLayoutItemMap)]
        if not maps: return False
        for m in maps:
            size = m.sizeWithUnits()
            target_ratio = max(1e-9, float(size.width()) / float(size.height()))
            adj_extent = adjust_extent_to_item_ratio(lyr_extent, target_ratio, cfg["MARGIN_FAC"])
            m.setExtent(adj_extent); m.refresh()
        img = QgsLayoutExporter.ImageExportSettings()
        img.dpi = cfg["DPI"]
        for attr in ("antialiasing", "antiAliasing"):
            if hasattr(img, attr):
                setattr(img, attr, True)
        try:
            flag_val = 0
            for name in dir(img.__class__):
                if "UseAdvancedEffects" in name:
                    flag_val |= int(getattr(img.__class__, name))
            if flag_val and hasattr(img, "flags"): img.flags = flag_val
        except Exception:
            pass
        if hasattr(img, "generateWorldFile"): img.generateWorldFile = False
        exp = QgsLayoutExporter(layout)
        res = exp.exportToImage(out_png, img)
        return res == QgsLayoutExporter.Success

    def relink_layer(prj: 'QgsProject', layer_name: str, shp_path: str) -> Optional[object]:
        layers = prj.mapLayersByName(layer_name)
        if not layers: return None
        lyr = layers[0]
        try:
            lyr.setDataSource(shp_path, lyr.name(), "ogr")
            return lyr
        except Exception:
            return None

    def export_views(projet_path: str) -> Tuple[int, int]:
        okc = 0; koc = 0
        nom = os.path.splitext(os.path.basename(projet_path))[0]
        out_ae = os.path.join(cfg["OUT_IMG"], f"{nom}__AE.png")
        out_ze = os.path.join(cfg["OUT_IMG"], f"{nom}__ZE.png")

        mode = cfg.get("CADRAGE_MODE", "BOTH")
        expected_exports = (1 if mode in ("AE", "BOTH") else 0) + (1 if mode in ("ZE", "BOTH") else 0)

        prj = QgsProject.instance(); prj.clear()

        opened = False
        for pth in (projet_path, to_long_unc(projet_path)):
            try:
                if prj.read(pth): opened = True; break
            except Exception:
                pass
        if not opened:
            return 0, expected_exports

        lm = prj.layoutManager()
        layouts = lm.layouts()
        if not layouts:
            prj.clear(); return 0, expected_exports
        layout = layouts[0]

        lyr_ae = relink_layer(prj, cfg["LAYER_AE_NAME"], cfg["AE_SHP"])
        lyr_ze = relink_layer(prj, cfg["LAYER_ZE_NAME"], cfg["ZE_SHP"])

        if mode in ("AE", "BOTH"):
            if (not cfg["OVERWRITE"]) and os.path.exists(out_ae):
                okc += 1
            else:
                if lyr_ae:
                    ext_ae = extent_in_project_crs(prj, lyr_ae)
                    if ext_ae and apply_extent_and_export(layout, ext_ae, out_ae): okc += 1
                    else: koc += 1
                else:
                    koc += 1

        if mode in ("ZE", "BOTH"):
            if (not cfg["OVERWRITE"]) and os.path.exists(out_ze):
                okc += 1
            else:
                if lyr_ze:
                    ext_ze = extent_in_project_crs(prj, lyr_ze)
                    if ext_ze and apply_extent_and_export(layout, ext_ze, out_ze): okc += 1
                    else: koc += 1
                else:
                    koc += 1

        prj.clear()
        return okc, koc

    for p in projects:
        ok_c, ko_c = export_views(p)
        ok += ok_c; ko += ko_c

    qgs.exitQgis()
    return ok, ko

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
            qt_base = None
            for name in ("Qt6", "Qt5"):
                base = os.path.join(QGIS_ROOT, "apps", name)
                if os.path.isdir(base): qt_base = base; break
            if not qt_base: raise RuntimeError("Répertoire Qt introuvable")
            sys.path.insert(0, os.path.join(QGIS_APP, "python"))
            sys.path.insert(0, os.path.join(QGIS_ROOT, "apps", PY_VER, "Lib", "site-packages"))
            from qgis.core import QgsApplication  # noqa: F401
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

            chunks = chunk_even(projets, self.workers_var.get())
            cfg = {
                "QGIS_ROOT": QGIS_ROOT, "QGIS_APP": QGIS_APP, "PY_VER": PY_VER,
                "OUT_IMG": OUT_IMG, "DPI": int(self.dpi_var.get()),
                "MARGIN_FAC": float(self.margin_var.get()),
                "LAYER_AE_NAME": LAYER_AE_NAME, "LAYER_ZE_NAME": LAYER_ZE_NAME,
                "AE_SHP": self.ae_shp_var.get(), "ZE_SHP": self.ze_shp_var.get(),
                "CADRAGE_MODE": self.cadrage_var.get(), "OVERWRITE": bool(self.overwrite_var.get()),
            }

            ok_total = 0; ko_total = 0
            def ui_update_progress(done_inc):
                self.progress_done += done_inc
                self.progress["value"] = min(self.progress_done, self.total_expected)
                self.status_label.config(text=f"Progression : {self.progress_done}/{self.total_expected}")

            with ProcessPoolExecutor(max_workers=int(self.workers_var.get())) as ex:
                futures = [ex.submit(worker_run, (chunk, cfg)) for chunk in chunks if chunk]
                for fut in as_completed(futures):
                    try:
                        ok, ko = fut.result()
                        ok_total += ok; ko_total += ko
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
            from id_contexte_eco import run_analysis as run_id_context
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
        self.cadrage_var  = tk.StringVar(value=self.prefs.get("CADRAGE_MODE", "BOTH"))
        self.overwrite_var= tk.BooleanVar(value=self.prefs.get("OVERWRITE", OVERWRITE_DEFAULT))
        self.dpi_var      = tk.IntVar(value=int(self.prefs.get("DPI", DPI_DEFAULT)))
        self.workers_var  = tk.IntVar(value=int(self.prefs.get("N_WORKERS", N_WORKERS_DEFAULT)))
        self.margin_var   = tk.DoubleVar(value=float(self.prefs.get("MARGIN_FAC", MARGIN_FAC_DEFAULT)))
        self.buffer_var   = tk.DoubleVar(value=float(self.prefs.get("ID_TAMPON_KM", 5.0)))

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

        self.export_button = ttk.Button(opt, text="Lancer l’export cartes", style="Accent.TButton", command=self.start_export_thread)
        self.export_button.grid(row=6, column=0, columnspan=3, sticky="w", pady=(10,0))

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
            self.ze_shp_var.set(path)

    def _select_ae(self):
        base = self.ae_shp_var.get() or os.path.expanduser("~")
        path = filedialog.askopenfilename(title="Sélectionner l'aire d'étude élargie",
                                          initialdir=base if os.path.isdir(base) else os.path.expanduser("~"),
                                          filetypes=[("Shapefile ESRI", "*.shp")])
        if path:
            self.ae_shp_var.set(path)

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
        mode = self.cadrage_var.get(); per_project = 2 if mode == "BOTH" else 1
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
        }); save_prefs(self.prefs)

        t = threading.Thread(target=self._run_export_logic, args=(projets,), daemon=True)
        t.start()

    def _run_export_logic(self, projets: List[str]):
        old_stdout = sys.stdout
        sys.stdout = self.stdout_redirect
        try:
            start = datetime.datetime.now()
            os.makedirs(OUT_IMG, exist_ok=True)
            log_with_time(f"{len(projets)} projets (attendu = calcul en cours)")
            log_with_time(f"Workers={self.workers_var.get()}, DPI={self.dpi_var.get()}, marge={self.margin_var.get():.2f}, overwrite={self.overwrite_var.get()}")
            chunks = chunk_even(projets, self.workers_var.get())
            cfg = {
                "QGIS_ROOT": QGIS_ROOT, "QGIS_APP": QGIS_APP, "PY_VER": PY_VER,
                "OUT_IMG": OUT_IMG, "DPI": int(self.dpi_var.get()),
                "MARGIN_FAC": float(self.margin_var.get()),
                "LAYER_AE_NAME": LAYER_AE_NAME, "LAYER_ZE_NAME": LAYER_ZE_NAME,
                "AE_SHP": self.ae_shp_var.get(), "ZE_SHP": self.ze_shp_var.get(),
                "CADRAGE_MODE": self.cadrage_var.get(), "OVERWRITE": bool(self.overwrite_var.get()),
            }
            ok_total = 0; ko_total = 0
            def ui_update_progress(done_inc):
                self.progress_done += done_inc
                self.progress["value"] = min(self.progress_done, self.total_expected)
                self.status_label.config(text=f"Progression : {self.progress_done}/{self.total_expected}")
            with ProcessPoolExecutor(max_workers=int(self.workers_var.get())) as ex:
                futures = [ex.submit(worker_run, (chunk, cfg)) for chunk in chunks if chunk]
                for fut in as_completed(futures):
                    try:
                        ok, ko = fut.result()
                        ok_total += ok; ko_total += ko
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
        ae = self.ae_shp_var.get().strip()
        ze = self.ze_shp_var.get().strip()
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
            from id_contexte_eco import run_analysis as run_id_context
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


