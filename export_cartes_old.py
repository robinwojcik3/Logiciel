#!/usr/bin/env "C:/Program Files/QGIS 3.40.3/apps/Python312/python.exe"
# -*- coding: utf-8 -*-
r"""
Export PNG des mises en page QGIS, 2 zooms par carte, avec sélection interactive des shapefiles.
- Une interface graphique Tkinter centralise les sélections et le lancement.
- L'utilisateur choisit les .shp pour "Zone d'étude" et "Aire d'étude élargie".
- L'utilisateur choisit le type de cadrage (AE, ZE, ou les deux).
- L'utilisateur sélectionne les projets QGIS à traiter depuis une liste.
- Relie chaque projet QGIS aux sources de données choisies (setDataSource) avant export.
- Lancement des exports en parallèle via ProcessPoolExecutor.
"""

import os
import sys
import datetime
import threading
from typing import List, Optional, Tuple
from concurrent.futures import ProcessPoolExecutor, as_completed

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    print("Tkinter n'est pas installé. L'interface graphique ne peut pas être lancée.")
    sys.exit(1)

DPI = 300
N_WORKERS = max(1, min((os.cpu_count() or 2) - 1, 6))
MARGIN_FAC = 1.15

LAYER_AE_NAME = "Aire d'étude élargie"
LAYER_ZE_NAME = "Zone d'étude"

BASE_SHARE = r"\\192.168.1.240\commun\PARTAGE"
SUBPATH = r"Espace_RWO\CARTO ROBIN"

OUT_IMG = r"C:\Users\utilisateur\Mon Drive\1 - Bota & Travail\+++++++++  BOTA  +++++++++\---------------------- 3) BDD\PYTHON\2) Contexte éco\OUTPUT"

QGIS_ROOT = r"C:\Program Files\QGIS 3.40.3"
QGIS_APP = os.path.join(QGIS_ROOT, "apps", "qgis")
PY_VER = "Python312"

def log_with_time(msg: str) -> None:
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")

def normalize_name(s: str) -> str:
    s2 = s.replace("\u00A0", " ").replace("\u202F", " ")
    while "  " in s2:
        s2 = s2.replace("  ", " ")
    return s2.strip().lower()

def to_long_unc(path: str) -> str:
    if path.startswith("\\\\?\\"): return path
    if path.startswith("\\\\"): return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path

def chunk_even(lst: List[str], k: int) -> List[List[str]]:
    k = max(1, min(k, len(lst)))
    base = len(lst) // k
    extra = len(lst) % k
    out = []
    start = 0
    for i in range(k):
        size = base + (1 if i < extra else 0)
        out.append(lst[start:start+size])
        start += size
    return out

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

    def adjust_extent_to_item_ratio(ext: QgsRectangle, target_ratio: float, margin: float) -> QgsRectangle:
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

    def extent_in_project_crs(prj: QgsProject, lyr) -> Optional[QgsRectangle]:
        ext = lyr.extent()
        try:
            if lyr.crs() != prj.crs():
                ct = QgsCoordinateTransform(lyr.crs(), prj.crs(), prj)
                ext = ct.transformBoundingBox(ext)
        except Exception:
            pass
        return ext

    def apply_extent_and_export(layout, lyr_extent: QgsRectangle, out_png: str) -> bool:
        maps = [it for it in layout.items() if isinstance(it, QgsLayoutItemMap)]
        if not maps:
            return False
        for m in maps:
            size = m.sizeWithUnits()
            target_ratio = max(1e-9, float(size.width()) / float(size.height()))
            adj_extent = adjust_extent_to_item_ratio(lyr_extent, target_ratio, cfg["MARGIN_FAC"])
            m.setExtent(adj_extent)
            m.refresh()
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
            if flag_val and hasattr(img, "flags"):
                img.flags = flag_val
        except Exception:
            pass
        if hasattr(img, "generateWorldFile"):
            img.generateWorldFile = False
        exp = QgsLayoutExporter(layout)
        res = exp.exportToImage(out_png, img)
        return res == QgsLayoutExporter.Success

    def relink_layer(prj: QgsProject, layer_name: str, shp_path: str):
        layers = prj.mapLayersByName(layer_name)
        if not layers:
            return None
        lyr = layers[0]
        try:
            lyr.setDataSource(shp_path, lyr.name(), "ogr")
            return lyr
        except Exception:
            return None

    def export_views(projet_path: str) -> Tuple[int, int]:
        okc = 0
        koc = 0
        nom = os.path.splitext(os.path.basename(projet_path))[0]
        out_ae = os.path.join(cfg["OUT_IMG"], f"{nom}__AE.png")
        out_ze = os.path.join(cfg["OUT_IMG"], f"{nom}__ZE.png")

        mode = cfg.get("CADRAGE_MODE", "BOTH")
        expected_exports = 0
        if mode in ("AE", "BOTH"): expected_exports += 1
        if mode in ("ZE", "BOTH"): expected_exports += 1

        prj = QgsProject.instance()
        prj.clear()

        opened = False
        for pth in (projet_path, to_long_unc(projet_path)):
            try:
                if prj.read(pth):
                    opened = True
                    break
            except Exception:
                pass
        if not opened:
            return 0, expected_exports

        lm = prj.layoutManager()
        layouts = lm.layouts()
        if not layouts:
            prj.clear()
            return 0, expected_exports
        layout = layouts[0]

        lyr_ae = relink_layer(prj, cfg["LAYER_AE_NAME"], cfg["AE_SHP"])
        lyr_ze = relink_layer(prj, cfg["LAYER_ZE_NAME"], cfg["ZE_SHP"])

        if mode in ("AE", "BOTH"):
            if os.path.exists(out_ae):
                okc += 1
            else:
                if lyr_ae:
                    ext_ae = extent_in_project_crs(prj, lyr_ae)
                    if ext_ae and apply_extent_and_export(layout, ext_ae, out_ae):
                        okc += 1
                    else:
                        koc += 1
                else:
                    koc += 1

        if mode in ("ZE", "BOTH"):
            if os.path.exists(out_ze):
                okc += 1
            else:
                if lyr_ze:
                    ext_ze = extent_in_project_crs(prj, lyr_ze)
                    if ext_ze and apply_extent_and_export(layout, ext_ze, out_ze):
                        okc += 1
                    else:
                        koc += 1
                else:
                    koc += 1

        prj.clear()
        return okc, koc

    for p in projects:
        ok_c, ko_c = export_views(p)
        ok += ok_c
        ko += ko_c

    qgs.exitQgis()
    return ok, ko

def discover_projects() -> List[str]:
    partage = BASE_SHARE
    base_dir: Optional[str] = None

    try:
        if os.path.isdir(partage):
            entries = os.listdir(partage)
            for e in entries:
                if normalize_name(e) == normalize_name("Espace_RWO"):
                    base_espace = os.path.join(partage, e)
                    subs = os.listdir(base_espace)
                    for s in subs:
                        if normalize_name(s) == normalize_name("CARTO ROBIN"):
                            base_dir = os.path.join(base_espace, s)
                            break
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
            projets = [os.path.join(d, f) for f in sorted(qgz)]
            return projets
        except Exception:
            continue
    return []

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

class ExportApp:
    def __init__(self, master):
        self.master = master
        master.title("Export cartes contexte éco")
        master.geometry("800x650")
        master.minsize(600, 500)

        style = ttk.Style()
        style.configure("TLabel", padding=5)
        style.configure("TButton", padding=5)
        style.configure("TRadiobutton", padding=5)

        self.ze_shp_var = tk.StringVar()
        self.ae_shp_var = tk.StringVar()
        self.cadrage_var = tk.StringVar(value="BOTH")
        self.project_vars = {}

        self._create_widgets()
        self._populate_projects()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        shp_frame = ttk.LabelFrame(main_frame, text="1. Sélection des couches Shapefile", padding="10")
        ttk.Button(shp_frame, text="Zone d'étude (.shp)", command=lambda: self._select_shapefile('ZE')).grid(row=0, column=0, sticky="ew")
        ttk.Label(shp_frame, textvariable=self.ze_shp_var, wraplength=400).grid(row=0, column=1, padx=10, sticky="w")
        ttk.Button(shp_frame, text="Aire d'étude élargie (.shp)", command=lambda: self._select_shapefile('AE')).grid(row=1, column=0, sticky="ew", pady=5)
        ttk.Label(shp_frame, textvariable=self.ae_shp_var, wraplength=400).grid(row=1, column=1, padx=10, sticky="w")
        shp_frame.columnconfigure(1, weight=1)

        cadrage_frame = ttk.LabelFrame(main_frame, text="2. Choix du cadrage", padding="10")
        ttk.Radiobutton(cadrage_frame, text="Cadrage sur les deux", variable=self.cadrage_var, value="BOTH").pack(anchor='w')
        ttk.Radiobutton(cadrage_frame, text="Cadrage uniquement sur la Zone d'étude", variable=self.cadrage_var, value="ZE").pack(anchor='w')
        ttk.Radiobutton(cadrage_frame, text="Cadrage uniquement sur l'Aire d'étude élargie", variable=self.cadrage_var, value="AE").pack(anchor='w')

        proj_frame = ttk.LabelFrame(main_frame, text="3. Sélection des projets QGIS", padding="10")
        canvas = tk.Canvas(proj_frame)
        scrollbar = ttk.Scrollbar(proj_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.export_button = ttk.Button(main_frame, text="Lancer l'export", command=self.start_export_thread)

        log_frame = ttk.LabelFrame(main_frame, text="Console", padding="10")
        log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, state='disabled')
        log_text_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
        log_text['yscrollcommand'] = log_text_scroll.set
        log_text_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sys.stdout = TextRedirector(log_text)

        shp_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        cadrage_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        log_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, pady=(10, 0))
        self.export_button.pack(side=tk.BOTTOM, pady=10)
        proj_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=5)

    def _select_shapefile(self, shp_type):
        label_text = "Zone d'étude" if shp_type == 'ZE' else "Aire d'étude élargie"
        title = f"Sélectionner le shapefile pour '{label_text}'"
        path = filedialog.askopenfilename(title=title, filetypes=[("Shapefile ESRI", "*.shp")])
        if path:
            if shp_type == 'ZE':
                self.ze_shp_var.set(path)
            else:
                self.ae_shp_var.set(path)

    def _populate_projects(self):
        projects = discover_projects()
        if not projects:
            ttk.Label(self.scrollable_frame, text="Aucun projet trouvé ou dossier inaccessible.", foreground="red").pack()
            return
        for proj_path in projects:
            var = tk.IntVar(value=1)
            proj_name = os.path.basename(proj_path)
            self.project_vars[proj_path] = var
            cb = ttk.Checkbutton(self.scrollable_frame, text=proj_name, variable=var)
            cb.pack(anchor='w', padx=5)

    def start_export_thread(self):
        if not self.ze_shp_var.get() or not self.ae_shp_var.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner les deux shapefiles avant de lancer l'export.")
            return
        selected_projects = [p for p, var in self.project_vars.items() if var.get() == 1]
        if not selected_projects:
            messagebox.showerror("Erreur", "Veuillez sélectionner au moins un projet QGIS.")
            return
        self.export_button.config(state="disabled")
        thread = threading.Thread(target=self._run_export_logic, args=(selected_projects,))
        thread.daemon = True
        thread.start()

    def _run_export_logic(self, projets):
        try:
            start = datetime.datetime.now()
            os.makedirs(OUT_IMG, exist_ok=True)

            cadrage_mode = self.cadrage_var.get()
            exports_per_project = 2 if cadrage_mode == "BOTH" else 1
            total_expected = exports_per_project * len(projets)

            log_with_time(f"{len(projets)} projets sélectionnés (exports attendus = {total_expected})")
            log_with_time(f"Parallélisation: {N_WORKERS} worker(s), DPI={DPI}, marge={MARGIN_FAC}")

            chunks = chunk_even(projets, N_WORKERS)
            cfg = {
                "QGIS_ROOT": QGIS_ROOT,
                "QGIS_APP": QGIS_APP,
                "PY_VER": PY_VER,
                "OUT_IMG": OUT_IMG,
                "DPI": DPI,
                "MARGIN_FAC": MARGIN_FAC,
                "LAYER_AE_NAME": LAYER_AE_NAME,
                "LAYER_ZE_NAME": LAYER_ZE_NAME,
                "AE_SHP": self.ae_shp_var.get(),
                "ZE_SHP": self.ze_shp_var.get(),
                "CADRAGE_MODE": cadrage_mode,
            }

            ok_total = 0
            ko_total = 0

            with ProcessPoolExecutor(max_workers=N_WORKERS) as ex:
                futures = [ex.submit(worker_run, (chunk, cfg)) for chunk in chunks if chunk]
                for fut in as_completed(futures):
                    try:
                        ok, ko = fut.result()
                        ok_total += ok
                        ko_total += ko
                        log_with_time(f"Chunk terminé: {ok} OK, {ko} KO")
                    except Exception as e:
                        log_with_time(f"Erreur d'un worker: {e}")

            elapsed = datetime.datetime.now() - start
            log_with_time(f"FIN — Exports OK={ok_total} | KO={ko_total} | Attendu={total_expected} | Durée={elapsed}")

        except Exception as e:
            log_with_time(f"Une erreur critique est survenue: {e}")
        finally:
            self.master.after(0, lambda: self.export_button.config(state="normal"))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExportApp(root)
    root.mainloop()
