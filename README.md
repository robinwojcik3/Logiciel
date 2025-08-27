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


## Export cartographique sous Windows

Sous Windows 10/11 avec Python 3.12, l'export des cartes QGIS nécessite que les dossiers contenant les bibliothèques Qt et QGIS soient déclarés explicitement. L'application utilise `os.add_dll_directory` pour enregistrer ces dossiers dans chaque processus avant d'importer `qgis.core`.

Les chemins sont récupérés via les variables d'environnement déjà présentes (`QGIS_ROOT`, `QGIS_PREFIX_PATH`, `QGIS_APP`, `QT_DIR`, `QT_QPA_PLATFORM_PLUGIN_PATH`). Aucune de ces variables n'est obligatoire : le code ignore celles qui manquent.

En cas de problème d'import (message « DLL load failed »), l'application tente automatiquement de relancer l'export avec un seul processus pour maximiser les chances de succès.
