# Logiciel

Ce dépôt contient une application Python à interface graphique permettant de manipuler différents outils « Contexte éco ».

## Démarrage rapide

Assurez-vous d'utiliser Python 3. Pour démarrer l'interface :

```bash
python Start.py
```

Par défaut, les onglets sont chargés à la demande afin d'accélérer l'ouverture.

### Variables d'environnement

- `LAZY_TABS=0` pour désactiver le chargement différé des onglets.
- `APP_SPLASH=1` pour afficher une petite fenêtre de chargement.

## Structure du projet

- `Start.py` : point d'entrée qui ouvre l'interface graphique.
- `modules/` : contient la logique métier et les composants de l'interface.
- `utils/` : helpers pour le cache et les accès fichiers.
- `Shapefile_Flore_Patri/` et `Template word Contexte éco/` : ressources utilisées par l'application.

## Dépendances

Les modules standard de Python suffisent pour lancer l'application de base. Certaines fonctionnalités peuvent requérir des bibliothèques supplémentaires, installables via `pip`.

