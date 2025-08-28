import os
from pathlib import Path

# Repo root (modules/..)
REPO_ROOT = Path(__file__).resolve().parent.parent

# Ne pas importer main_app ici pour éviter de charger des dépendances lourdes
# (les sous-modules sont importés à la demande: `from modules import main_app` fonctionne)

# Repo-first overrides (non-destructive, keep defaults as fallback)
try:
    # Sorties
    main_app.OUT_IMG = str(REPO_ROOT / "output")
    os.makedirs(main_app.OUT_IMG, exist_ok=True)
except Exception:
    pass

try:
    # Dossier de sélection de shapefiles
    main_app.DEFAULT_SHAPE_DIR = str(REPO_ROOT / "data")
except Exception:
    pass

try:
    # QGIS local prioritaire (env var QGIS_ROOT, sinon third_party/QGIS)
    env_qgis = os.environ.get("QGIS_ROOT")
    if env_qgis and os.path.isdir(env_qgis):
        main_app.QGIS_ROOT = env_qgis
        main_app.QGIS_APP = os.path.join(main_app.QGIS_ROOT, "apps", "qgis")
    else:
        local_qgis = REPO_ROOT / "third_party" / "QGIS"
        if local_qgis.is_dir():
            main_app.QGIS_ROOT = str(local_qgis)
            main_app.QGIS_APP = os.path.join(main_app.QGIS_ROOT, "apps", "qgis")
except Exception:
    pass

# Découverte des projets: préfère le dossier local si disponible
def _discover_projects_repo_first():
    candidates = [
        REPO_ROOT / "Cartes contexte éco export",
    ]
    # Ajoute les emplacements réseau connus (inchangés)
    try:
        base_dir = os.path.join(main_app.BASE_SHARE, main_app.SUBPATH)
        candidates[:0] = [Path(base_dir), Path(main_app.to_long_unc(base_dir))]
    except Exception:
        pass
    for d in candidates:
        try:
            if not d.is_dir():
                continue
            files = [f for f in os.listdir(d) if f.lower().endswith(".qgz")]
            if files:
                return [str(d / f) for f in sorted(files)]
        except Exception:
            continue
    return []

# Remplace la fonction si elle existe
try:
    main_app.discover_projects = _discover_projects_repo_first
except Exception:
    pass
