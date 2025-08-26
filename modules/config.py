import os

# =========================
# Paramètres globaux
# =========================
DPI_DEFAULT        = 300
N_WORKERS_DEFAULT  = max(1, min((os.cpu_count() or 2) - 1, 6))
MARGIN_FAC_DEFAULT = 1.15
OVERWRITE_DEFAULT  = False

LAYER_AE_NAME = "Aire d'étude élargie"
LAYER_ZE_NAME = "Zone d'étude"

BASE_SHARE = r"\\\\192.168.1.240\\commun\\PARTAGE"
SUBPATH    = r"Espace_RWO\\CARTO ROBIN"

OUT_IMG    = r"C:\\Users\\utilisateur\\Mon Drive\\1 - Bota & Travail\\+++++++++  BOTA  +++++++++\\---------------------- 3) BDD\\PYTHON\\2) Contexte éco\\OUTPUT"

# Dossier par défaut pour la sélection des shapefiles (onglet 1)
DEFAULT_SHAPE_DIR = r"C:\\Users\\utilisateur\\Mon Drive\\1 - Bota & Travail\\+++++++++  BOTA  +++++++++\\---------------------- 2) CARTO terrain"

# QGIS
QGIS_ROOT = r"C:\\Program Files\\QGIS 3.40.3"
QGIS_APP  = os.path.join(QGIS_ROOT, "apps", "qgis")
PY_VER    = "Python312"

# Préférences
PREFS_PATH = os.path.join(os.path.expanduser("~"), "ExportCartesContexteEco.config.json")

# Onglet 2 — Remonter le temps & Bassin versant
LAYERS = [
    ("Aujourd’hui",   "10"),
    ("2000-2005",     "18"),
    ("1965-1980",     "20"),
    ("1950-1965",     "19"),
]
URL = ("https://remonterletemps.ign.fr/comparer/?lon={lon}&lat={lat}"
       "&z=17&layer1={layer}&layer2=19&mode=dub1")
WAIT_TILES_DEFAULT = 1.5
try:
    from docx.shared import Cm
    IMG_WIDTH = Cm(12.5 * 0.8)
except Exception:  # pragma: no cover - dépendance optionnelle
    IMG_WIDTH = None
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
