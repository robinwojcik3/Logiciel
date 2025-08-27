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

Sous Windows avec Python 3.12, l'export cartographique peut nécessiter une configuration explicite des chemins de DLL de QGIS et Qt.
L'application s'appuie sur `os.add_dll_directory` via `prepare_qgis_env` pour charger correctement les bibliothèques avant tout import `qgis.core`.
Les variables suivantes peuvent être définies dans la configuration ou l'environnement :

- `QGIS_ROOT`
- `QGIS_APP`
- `QGIS_PREFIX_PATH`
- `QT_DIR`
- `QT_QPA_PLATFORM_PLUGIN_PATH`

Un test de fumée est disponible pour vérifier l'import de QGIS. En cas d'échec, l'export n'est pas lancé et le message détaillé s'affiche.

