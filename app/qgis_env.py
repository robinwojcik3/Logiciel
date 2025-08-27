from __future__ import annotations
import os
import platform
import sys


def _dedup_dirs(dirs):
    seen = set()
    out = []
    for d in dirs:
        if not d:
            continue
        d = os.path.abspath(d)
        if os.path.isdir(d) and d not in seen:
            seen.add(d)
            out.append(d)
    return out


def prepare_qgis_env(cfg: dict | None = None) -> None:
    """Idempotent. À appeler dans CHAQUE process AVANT 'from qgis.core import ...'."""
    if platform.system() != "Windows":
        return
    cfg = cfg or {}

    candidates = []
    for key in ("QGIS_ROOT", "QGIS_APP", "QT_DIR"):
        base = cfg.get(key) or os.environ.get(key, "")
        if base:
            candidates += [base, os.path.join(base, "bin")]

    qt_plugins = cfg.get("QT_QPA_PLATFORM_PLUGIN_PATH") or os.environ.get("QT_QPA_PLATFORM_PLUGIN_PATH", "")
    if not qt_plugins and (cfg.get("QT_DIR") or os.environ.get("QT_DIR")):
        qt_plugins = os.path.join(cfg.get("QT_DIR") or os.environ.get("QT_DIR"), "plugins", "platforms")

    dll_dirs = _dedup_dirs(candidates)

    for d in dll_dirs:
        try:
            os.add_dll_directory(d)
        except Exception:
            pass

    os.environ["PATH"] = os.pathsep.join(dll_dirs + [os.environ.get("PATH", "")])

    if cfg.get("QGIS_PREFIX_PATH") or os.environ.get("QGIS_PREFIX_PATH"):
        os.environ["QGIS_PREFIX_PATH"] = cfg.get("QGIS_PREFIX_PATH") or os.environ["QGIS_PREFIX_PATH"]

    if qt_plugins and os.path.isdir(qt_plugins):
        os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = qt_plugins


def qgis_smoke_test(cfg) -> tuple[bool, str]:
    try:
        prepare_qgis_env(cfg)
        from qgis.core import QgsApplication  # noqa: F401
        return True, "Import QGIS OK"
    except Exception as e:
        import traceback
        return False, f"Import QGIS KO: {e}\n{traceback.format_exc()}"
