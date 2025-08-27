from __future__ import annotations
import os, sys, platform, traceback


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

    # Collecte des candidats DLL
    candidates = []
    for key in ("QGIS_ROOT", "QGIS_APP", "QT_DIR"):
        base = cfg.get(key) or os.environ.get(key, "")
        if base:
            candidates += [base, os.path.join(base, "bin")]

    # Ajout des répertoires connus de plugins Qt/QGIS si fournis
    qt_plugins = cfg.get("QT_QPA_PLATFORM_PLUGIN_PATH") or os.environ.get("QT_QPA_PLATFORM_PLUGIN_PATH", "")
    if not qt_plugins:
        qt_dir = cfg.get("QT_DIR") or os.environ.get("QT_DIR")
        if qt_dir:
            qt_plugins = os.path.join(qt_dir, "plugins", "platforms")

    dll_dirs = _dedup_dirs(candidates)

    # Déclare les dossiers DLL au loader (Python >=3.8)
    for d in dll_dirs:
        try:
            os.add_dll_directory(d)  # crucial sous Windows
        except Exception:
            pass  # compatible avec Py<3.8 ou environnements restreints

    # Sécurise PATH en plus (utile pour certaines dépendances secondaires)
    os.environ["PATH"] = os.pathsep.join(dll_dirs + [os.environ.get("PATH", "")])

    # QGIS prefix
    qgis_prefix = cfg.get("QGIS_PREFIX_PATH") or os.environ.get("QGIS_PREFIX_PATH")
    if qgis_prefix:
        os.environ["QGIS_PREFIX_PATH"] = qgis_prefix

    # Qt platform plugins
    if qt_plugins and os.path.isdir(qt_plugins):
        os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = qt_plugins


def qgis_smoke_test(cfg: dict | None):
    try:
        prepare_qgis_env(cfg)
        from qgis.core import QgsApplication  # noqa: F401
        return True, "Import QGIS OK"
    except Exception as e:
        return False, f"Import QGIS KO: {e}\n{traceback.format_exc()}"
