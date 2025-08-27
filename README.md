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


## Démarrage rapide

La fenêtre principale s'affiche immédiatement. Les onglets sont construits uniquement lors du premier clic, évitant ainsi les imports lourds au démarrage.

## Variables d'environnement

- `LAZY_TABS` : vaut `1` par défaut. Mettre `0` pour recréer le comportement ancien où tous les onglets sont chargés dès l'ouverture.
- `APP_SPLASH` : vaut `0` par défaut. Mettre `1` pour afficher un petit écran de démarrage « Initialisation… ».
