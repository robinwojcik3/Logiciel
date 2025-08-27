# Logiciel

Ce dépôt contient une application Python à interface graphique permettant de manipuler différents outils « Contexte éco ».

## Lancer l'application

Assurez-vous d'utiliser Python 3. Pour démarrer l'interface, exécutez :

```bash
python Start.py
```

L'application s'ouvrira avec plusieurs onglets proposant les fonctionnalités principales.

## Structure du projet

- `Start.py` : point d'entrée qui ouvre l'interface graphique.
- `modules/` : contient la logique métier et les composants de l'interface.
- `Shapefile_Flore_Patri/` et `Template word Contexte éco/` : ressources utilisées par l'application.

## Dépendances

Les modules standard de Python suffisent pour lancer l'application de base. Certaines fonctionnalités peuvent requérir des bibliothèques supplémentaires, installables via `pip`.

## Export cartes — Windows + Python 3.12

Sur Windows avec Python 3.12, l'import de QGIS peut échouer avec « DLL load failed ». Le module `modules/qgis_env.py` prépare explicitement les répertoires de DLL grâce à `os.add_dll_directory` avant chaque `from qgis.core import ...`.

Les chemins sont récupérés à partir de variables déjà présentes dans l'environnement ou la configuration :

- `QGIS_ROOT`
- `QGIS_APP`
- `QGIS_PREFIX_PATH`
- `QT_DIR`
- `QT_QPA_PLATFORM_PLUGIN_PATH`

Assurez-vous que ces variables pointent vers votre installation locale de QGIS/Qt lorsque vous utilisez l'export cartographique.

