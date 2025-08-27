# Logiciel

Ce dépôt contient une application Python à interface graphique permettant de manipuler différents outils « Contexte éco ».

## Démarrage rapide

Assurez-vous d'utiliser Python 3. Pour démarrer l'interface, exécutez :

```bash
python Start.py
```

Par défaut, seuls les onglets nécessaires sont créés au premier clic afin
d'accélérer l'ouverture. Les variables d'environnement suivantes permettent
d'ajuster le comportement :

| Variable      | Valeur par défaut | Effet |
|---------------|------------------|-------|
| `LAZY_TABS`   | `1`              | Si `0`, tous les onglets sont chargés immédiatement. |
| `APP_SPLASH`  | `0`              | Si `1`, affiche une petite fenêtre de démarrage. |

L'application s'ouvrira avec plusieurs onglets proposant les fonctionnalités principales.

## Structure du projet

- `Start.py` : point d'entrée qui ouvre l'interface graphique.
- `modules/` : contient la logique métier et les composants de l'interface.
- `Shapefile_Flore_Patri/` et `Template word Contexte éco/` : ressources utilisées par l'application.

## Dépendances

Les modules standard de Python suffisent pour lancer l'application de base. Certaines fonctionnalités peuvent requérir des bibliothèques supplémentaires, installables via `pip`.

