Installation et exécution isolées
---------------------------------

Objectif: préparer un environnement virtuel local (.venv) dans le dépôt et y installer toutes les dépendances Python nécessaires.

Prérequis système
- Python 3.11+ (3.13 supporté)
- QGIS 3.x installé (requis pour l'export des cartes via le worker QGIS)
- Google Chrome (requis pour l'automatisation Selenium)

Étapes
- Ouvrir PowerShell dans le dossier racine du dépôt.
- Exécuter:

  scripts\setup.ps1

- Lancer ensuite l'application:

  scripts\run.ps1

Notes
- Les dépendances Python sont listées dans requirements.txt et installées dans .venv.
- Selenium télécharge automatiquement le driver Chrome si nécessaire.
- Le chemin QGIS utilisé par le worker est défini dans modules/main_app.py (variables QGIS_ROOT, QGIS_APP, PY_VER).

