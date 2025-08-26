#!/usr/bin/env "C:/Program Files/QGIS 3.40.3/apps/Python312/python.exe"
# -*- coding: utf-8 -*-
"""Module dédié à l'export des cartes (processus séparé).
Contient la fonction worker_run utilisée par l'onglet « Export Cartes ».
Issu de la version précédente qui fonctionnait correctement."""

import os
import sys
from typing import List, Optional, Tuple


def to_long_unc(path: str) -> str:
    """Convertit un chemin en version UNC longue pour Windows."""
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path


def worker_run(args: Tuple[List[str], dict]) -> Tuple[int, int]:
    """Effectue l'export des projets QGIS vers des PNG.

    Args:
        args: Tuple contenant (liste de projets, configuration).
    Returns:
        Tuple (ok_exports, ko_exports).
    """
    projects, cfg = args

    # --- Environnement QGIS/Qt ---
    os.environ["OSGEO4W_ROOT"] = cfg["QGIS_ROOT"]
    os.environ["QGIS_PREFIX_PATH"] = cfg["QGIS_APP"]
    os.environ.setdefault("GDAL_DATA", os.path.join(cfg["QGIS_ROOT"], "share", "gdal"))
    os.environ.setdefault("PROJ_LIB", os.path.join(cfg["QGIS_ROOT"], "share", "proj"))
    os.environ.setdefault("QT_QPA_FONTDIR", r"C:\\Windows\\Fonts")

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

    def relink_layer(prj: QgsProject, layer_name: str, shp_path: str) -> Optional[object]:
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
        if mode in ("AE", "BOTH"):
            expected_exports += 1
        if mode in ("ZE", "BOTH"):
            expected_exports += 1

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
            if (not cfg["OVERWRITE"]) and os.path.exists(out_ae):
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
            if (not cfg["OVERWRITE"]) and os.path.exists(out_ze):
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
