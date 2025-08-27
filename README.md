# Logiciel

Ce dépôt contient une application Python à interface graphique permettant de manipuler différents outils « Contexte éco ».

## Démarrage rapide

```bash
python Start.py
```

L'interface principale s'ouvre immédiatement puis charge les onglets au besoin.

## Variables d'environnement

- `LAZY_TABS` : activer le chargement différé des onglets (1 par défaut).
- `APP_SPLASH` : afficher un écran de démarrage minimal (0 par défaut).

## Structure du projet

- `Start.py` : point d'entrée qui ouvre l'interface graphique.
- `modules/` : composants de l'interface et logique métier.
- `utils/` : fonctions utilitaires (cache, accès réseau).
- `tests/` : tests unitaires.

## Dépendances

Les modules standard de Python suffisent pour lancer l'application de base. Certaines fonctionnalités peuvent nécessiter des bibliothèques supplémentaires, installables via `pip`.
