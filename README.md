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

Lors des exports de cartes, l'application utilise QGIS en tâche de fond.
Depuis Python 3.8 sous Windows, le système ne cherche plus automatiquement les DLL de Qt/QGIS.
Pour éviter l'erreur « DLL load failed », chaque processus ajoute maintenant explicitement
les dossiers indiqués par les variables suivantes :

- `QGIS_ROOT`
- `QGIS_APP`
- `QGIS_PREFIX_PATH`
- `QT_DIR`
- `QT_QPA_PLATFORM_PLUGIN_PATH`

La fonction utilitaire `prepare_qgis_env` se charge de déclarer ces dossiers via
`os.add_dll_directory`. Un test rapide de l'environnement est disponible avec
`qgis_smoke_test` avant de lancer les exports.

