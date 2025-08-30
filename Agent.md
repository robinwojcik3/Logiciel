# Agent.md

Objectif: rappeler et garantir que cette application doit fonctionner en tant que logiciel autonome, sans dépendances globales externes, et que tout ce qui est nécessaire à son exécution est contenu dans le dépôt.

Principes d’isolement
- Autonome: toutes les dépendances Python sont installées dans un environnement virtuel local situé dans le dépôt (`.venv`). Aucune installation globale n’est requise ni utilisée.
- Local-first: les fonctions de l’app doivent référencer des ressources, binaires et configurations présents dans le dépôt (chemins relatifs ou ancrés à la racine du repo), pas des chemins systèmes.
- Reproductible: les versions des dépendances sont contrôlées via `requirements.txt`. Le dépôt peut également contenir un fichier de lock (`requirements.lock.txt`) si vous le souhaitez pour figer les versions exactes.

Ce qui est inclus dans le dépôt
- `.venv/`: environnement Python local contenant tous les paquets nécessaires (créé par `scripts/setup.ps1`).
- `requirements.txt`: liste des dépendances Python (geopandas, requests, selenium, pillow, pillow-heif, python-docx, openpyxl, bs4, lxml, etc.).
- `scripts/setup.ps1`: prépare `.venv` et installe toutes les dépendances dans le dépôt.
- `scripts/run.ps1`: lance l’application avec l’interpréteur de `.venv`.

Pré-requis non-Python
- QGIS 3.x: requis pour l’onglet “Contexte éco” (export carto via PyQGIS). Par défaut, l’app pointe vers une installation système (variables `QGIS_ROOT`, `QGIS_APP`, `PY_VER` dans `modules/main_app.py`). Pour un isolement maximal, vous pouvez placer une version portable de QGIS dans un sous-dossier du dépôt (ex: `third_party/QGIS`) et ajuster ces variables pour pointer vers ce chemin local.
- Google Chrome: requis par Selenium. Selenium Manager gère le driver automatiquement. Pour un confinement 100% local, vous pouvez déposer un `chromedriver` dans le dépôt et adapter l’initialisation du `webdriver` afin d’utiliser ce binaire local.
  Les scripts Selenium doivent lancer Chrome en fenêtre minimisée pour ne pas perturber l’utilisateur.

Règles de contribution (isolation garantie)
- Ne jamais installer un paquet globalement pour les besoins de l’app. Utiliser exclusivement `.venv` (via `scripts/setup.ps1` ou `.\.venv\Scripts\pip.exe`).
- Ne pas coder de chemins absolus vers des dossiers système. Utiliser des chemins relatifs à la racine du repo (via `os.path.join` et `__file__`) et des variables centralisées.
- Documenter toute nouvelle dépendance dans `requirements.txt` et ré-exécuter `scripts/setup.ps1` pour la rendre disponible dans `.venv`.
- Si vous introduisez des binaires tiers (drivers, outils CLI), placez-les dans un sous-dossier du repo (ex: `tools/` ou `third_party/`) et référencez-les depuis le code via des chemins relatifs.

Procédure d’installation locale
1) Ouvrir PowerShell à la racine du dépôt.
2) Exécuter: `scripts\\setup.ps1` (crée `.venv` et installe les dépendances dans le dépôt).
3) Lancer l’app: `scripts\\run.ps1`.

Option de verrouillage des versions
- Pour capturer l’état exact des dépendances utilisées: `.\.venv\Scripts\pip.exe freeze > requirements.lock.txt`
- Pour réinstaller à l’identique: `.\.venv\Scripts\pip.exe install -r requirements.lock.txt`

Notes
- Certaines fonctionnalités nécessitent QGIS (composants C++/Qt). L’inclusion d’une distribution portable de QGIS dans le dépôt est possible mais volumineuse; par défaut on s’appuie sur une installation locale documentée.
- Selenium peut télécharger un driver en cache. Si vous exigez 0 dépendance hors dépôt, fournissez le binaire `chromedriver` dans le repo et utilisez-le explicitement dans le code.

