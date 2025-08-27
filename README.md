# Logiciel

Ce dépôt contient une application Python à interface graphique permettant de manipuler différents outils « Contexte éco ».

## Démarrage rapide

Assurez-vous d'utiliser Python 3. Pour démarrer l'interface, exécutez :

```bash
python Start.py
```

L'interface s'ouvre immédiatement puis charge le premier onglet en tâche de fond.

## Variables d'environnement

- `LAZY_TABS` : mettre à `0` pour créer tous les onglets dès le démarrage (par défaut `1`).
- `APP_SPLASH` : mettre à `1` pour afficher un petit écran de chargement au lancement.

## Structure du projet

- `Start.py` : point d'entrée qui ouvre l'interface graphique.
- `modules/` : contient la logique métier et les composants de l'interface.
- `utils/` : helpers de cache et d'accès fichier.
- `tests/` : tests unitaires.

## Dépendances

Les modules standard de Python suffisent pour lancer l'application de base. Certaines fonctionnalités peuvent requérir des bibliothèques supplémentaires, installables via `pip`.
