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

## Export cartes sous Windows

Avec Python 3.12 sur Windows, l'export cartographique utilise le module QGIS. Depuis Python 3.8, le chargeur ne cherche plus automatiquement les DLL de QGIS et Qt dans les processus enfants. Pour assurer un import correct, l'application ajoute explicitement les dossiers de DLL via `os.add_dll_directory` et lit les variables d'environnement suivantes si elles sont définies :

- `QGIS_ROOT`
- `QGIS_APP`
- `QGIS_PREFIX_PATH`
- `QT_DIR`
- `QT_QPA_PLATFORM_PLUGIN_PATH`

Si l'import QGIS échoue, un test de fumée (accessible depuis l'onglet Export cartes) affiche un message d'erreur lisible. Dans certains cas, l'application se replie automatiquement sur un seul worker pour contourner des problèmes de chargement de DLL.

