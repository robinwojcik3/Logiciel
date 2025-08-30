# Logiciel (autonome)

Ce dépôt contient une application Python avec interface graphique (plusieurs onglets, dont “Contexte éco” et “Identification Pl@ntNet”). L’application est conçue pour être autonome: toutes les dépendances Python sont installées localement dans le dépôt afin d’éviter toute dépendance à l’environnement de la machine.

Principe clé: le logiciel doit être autonome et isolé. Toutes les bibliothèques utilisées par l’application doivent être installées à l’intérieur du dépôt et référencées depuis celui‑ci (pas d’installations globales, pas de chemins système en dur pour les dépendances Python).

## Installation (environnement local dans le dépôt)

Sous PowerShell à la racine du dépôt:

1) `scripts\\setup.ps1` — crée `.venv` dans le dépôt et installe toutes les dépendances Python listées dans `requirements.txt`.
2) `scripts\\run.ps1` — lance l’application en utilisant l’interpréteur de `.venv`.

Voir aussi `docs/INSTALL.md` pour les détails et prérequis non‑Python (QGIS/Chrome).

## Lancer l’application

- Méthode recommandée: `scripts\\run.ps1`
- Ou manuellement: `.\\.venv\\Scripts\\python.exe Start.py`

## Structure du projet

- `Start.py`: point d’entrée.
- `modules/`: logique métier (UI, export QGIS, scraping Wikipédia, etc.).
- `requirements.txt`: dépendances Python installées dans `.venv` (dans le dépôt).
- `scripts/`: scripts d’installation/ lancement (`setup.ps1`, `run.ps1`).
- `docs/INSTALL.md`: procédure d’installation isolée.
- Ressources: `Shapefile_Flore_Patri/`, `Template word  Contexte éco/`, etc.

## Pré-requis non‑Python

- QGIS 3.x installé (requis pour l’export via PyQGIS). Les chemins par défaut pointent vers `C:\\Program Files\\QGIS ...`. Pour un confinement total, il est possible d’utiliser une distribution portable de QGIS placée dans le dépôt et d’ajuster `QGIS_ROOT`/`QGIS_APP`/`PY_VER` dans `modules/main_app.py`.
- Google Chrome (Selenium). Par défaut Selenium Manager gère le driver. Si vous exigez 100% local, placez un `chromedriver` dans le repo et initialisez le WebDriver avec ce binaire. Les fenêtres Chrome ouvertes par l’automatisation sont réduites automatiquement pour ne pas s’afficher en plein écran.

## Verrouillage des versions (optionnel)

- Geler l’état exact: `.\\.venv\\Scripts\\pip.exe freeze > requirements.lock.txt`
- Réinstaller à l’identique: `.\\.venv\\Scripts\\pip.exe install -r requirements.lock.txt`

## Bonnes pratiques d’isolement

- N’utiliser que `.venv` pour installer/mettre à jour des paquets.
- Référencer des chemins relatifs au dépôt dans le code.
- Documenter toute nouvelle dépendance dans `requirements.txt` et relancer `scripts\\setup.ps1`.

