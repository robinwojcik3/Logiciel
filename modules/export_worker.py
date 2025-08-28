# -*- coding: utf-8 -*-
"""Worker QGIS isolé pour l'export des cartes."""
import os
import sys
import datetime
from typing import List, Dict, Tuple


def log_with_time(msg: str) -> None:
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}")


def to_long_unc(path: str) -> str:
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC" + path[1:]
    return "\\\\?\\" + path


def _prepare_qgis_env(cfg: Dict) -> str:
    """Prépare l'environnement QGIS/Qt pour un sous-processus."""
    # Assainir les variables d'environnement Python pour éviter les mélanges 3.12/3.13
    for _k in ("PYTHONHOME", "PYTHONPATH", "PYTHONSTARTUP"):
        try:
            os.environ.pop(_k, None)
        except Exception:
            pass
    os.environ["PYTHONNOUSERSITE"] = "1"

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
        raise RuntimeError("Qt introuvable sous QGIS_ROOT/apps")

    platform_dir = os.path.join(qt_base, "plugins", "platforms")
    os.environ["QT_PLUGIN_PATH"] = os.path.join(qt_base, "plugins")
    os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = platform_dir

    if os.path.isfile(os.path.join(platform_dir, "qwindows.dll")):
        qpa = "windows"
    elif os.path.isfile(os.path.join(platform_dir, "qminimal.dll")):
        qpa = "minimal"
    else:
        qpa = "offscreen"
    os.environ["QT_QPA_PLATFORM"] = qpa

    sys.path.insert(0, os.path.join(cfg["QGIS_APP"], "python"))
    sys.path.insert(0, os.path.join(cfg["QGIS_ROOT"], "apps", cfg["PY_VER"], "Lib", "site-packages"))

    if hasattr(os, "add_dll_directory"):
        for d in (
            os.path.join(qt_base, "bin"),
            os.path.join(cfg["QGIS_APP"], "bin"),
            os.path.join(cfg["QGIS_ROOT"], "bin"),
        ):
            if os.path.isdir(d):
                os.add_dll_directory(d)

    os.environ["PATH"] = os.pathsep.join([
        os.path.join(qt_base, "bin"),
        os.path.join(cfg["QGIS_APP"], "bin"),
        os.path.join(cfg["QGIS_ROOT"], "bin"),
        os.environ.get("PATH", ""),
    ])

    return qt_base


# ---------- Helpers QGIS ----------

def adjust_extent_to_item_ratio(ext, target_ratio: float, margin: float):
    from qgis.core import QgsRectangle
    if ext.width() <= 0 or ext.height() <= 0:
        return ext
    cx, cy = ext.center().x(), ext.center().y()
    w, h = ext.width(), ext.height()
    if (w / h) > target_ratio:
        new_h = w / target_ratio
        dh = (new_h - h) / 2.0
        xmin, xmax = ext.xMinimum(), ext.xMaximum()
        ymin, ymax = cy - h / 2.0 - dh, cy + h / 2.0 + dh
    else:
        new_w = h * target_ratio
        dw = (new_w - w) / 2.0
        ymin, ymax = ext.yMinimum(), ext.yMaximum()
        xmin, xmax = cx - w / 2.0 - dw, cx + w / 2.0 + dw
    cw, ch = (xmax - xmin), (ymax - ymin)
    mx, my = (margin - 1.0) * cw / 2.0, (margin - 1.0) * ch / 2.0
    return QgsRectangle(xmin - mx, ymin - my, xmax + mx, ymax + my)


def extent_in_project_crs(prj, lyr):
    ext = lyr.extent()
    try:
        if lyr.crs() != prj.crs():
            from qgis.core import QgsCoordinateTransform
            ct = QgsCoordinateTransform(lyr.crs(), prj.crs(), prj)
            ext = ct.transformBoundingBox(ext)
    except Exception:
        pass
    return ext


def apply_extent_and_export(layout, lyr_extent, out_png: str, cfg: Dict) -> bool:
    from qgis.core import QgsLayoutItemMap, QgsLayoutExporter
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


def relink_layer(prj, layer_name: str, shp_path: str):
    layers = prj.mapLayersByName(layer_name)
    if not layers:
        return None
    lyr = layers[0]
    try:
        lyr.setDataSource(shp_path, lyr.name(), "ogr")
        return lyr
    except Exception:
        return None


def export_views(projet_path: str, cfg: Dict) -> Tuple[int, int]:
    from qgis.core import QgsProject
    okc = 0
    koc = 0
    nom = os.path.splitext(os.path.basename(projet_path))[0]
    out_ae = os.path.join(cfg["EXPORT_DIR"], f"{nom}__AE.png")
    out_ze = os.path.join(cfg["EXPORT_DIR"], f"{nom}__ZE.png")
    out_proj = os.path.join(cfg["EXPORT_DIR"], f"{nom}{os.path.splitext(projet_path)[1]}")

    mode = cfg.get("CADRAGE_MODE", "BOTH")
    expected_exports = 0
    if cfg.get("EXPORT_TYPE", "PNG") in ("PNG", "BOTH"):
        expected_exports += (2 if mode == "BOTH" else 1)
    if cfg.get("EXPORT_TYPE", "PNG") in ("QGS", "BOTH"):
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

    if cfg.get("EXPORT_TYPE", "PNG") in ("PNG", "BOTH"):
        if mode in ("AE", "BOTH"):
            if (not cfg["OVERWRITE"]) and os.path.exists(out_ae):
                okc += 1
            else:
                if lyr_ae:
                    ext_ae = extent_in_project_crs(prj, lyr_ae)
                    if ext_ae and apply_extent_and_export(layout, ext_ae, out_ae, cfg):
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
                    if ext_ze and apply_extent_and_export(layout, ext_ze, out_ze, cfg):
                        okc += 1
                    else:
                        koc += 1
                else:
                    koc += 1

    if cfg.get("EXPORT_TYPE", "PNG") in ("QGS", "BOTH"):
        if (not cfg["OVERWRITE"]) and os.path.exists(out_proj):
            okc += 1
        else:
            try:
                if prj.write(out_proj):
                    okc += 1
                else:
                    koc += 1
            except Exception:
                koc += 1

    prj.clear()
    return okc, koc


def worker_run(args: Tuple[List[str], Dict]) -> Tuple[int, int]:
    projects, cfg = args
    log_with_time(f"Worker start: {len(projects)} projets")
    qt_base = _prepare_qgis_env(cfg)

    global QgsApplication, QgsProject, QgsLayoutExporter, QgsLayoutItemMap, QgsRectangle, QgsCoordinateTransform
    from qgis.core import (
        QgsApplication, QgsProject, QgsLayoutExporter, QgsLayoutItemMap,
        QgsRectangle, QgsCoordinateTransform
    )
    log_with_time(f"QGIS import OK; QPA={os.environ.get('QT_QPA_PLATFORM')}; Qt={'Qt6' if 'Qt6' in qt_base else 'Qt5'}")

    qgs = QgsApplication([], False)
    qgs.initQgis()
    try:
        ok = 0
        ko = 0
        for path in projects:
            try:
                ok_c, ko_c = export_views(path, cfg)
                ok += ok_c
                ko += ko_c
            except Exception:
                ko += 1
        log_with_time(f"Export OK/KO: {ok}/{ko}")
        return ok, ko
    finally:
        qgs.exitQgis()
