# Logiciel

Ce dépôt contient une application Python à interface graphique permettant de manipuler différents outils « Contexte éco ».

## Démarrage rapide

Assurez-vous d'utiliser Python 3. Pour démarrer l'interface, exécutez :

```bash
python Start.py
```

L'application s'ouvrira avec plusieurs onglets proposant les fonctionnalités principales.

## Variables d'environnement

- `LAZY_TABS` : mettre `0` pour créer tous les onglets dès le démarrage (par défaut `1`).
- `APP_SPLASH` : mettre `1` pour afficher une petite fenêtre de chargement.

## Structure du projet

- `Start.py` : point d'entrée qui ouvre l'interface graphique.
- `modules/` : contient la logique métier et les composants de l'interface.
- `Shapefile_Flore_Patri/` et `Template word Contexte éco/` : ressources utilisées par l'application.

## Dépendances

Les modules standard de Python suffisent pour lancer l'application de base. Certaines fonctionnalités peuvent requérir des bibliothèques supplémentaires, installables via `pip`.

