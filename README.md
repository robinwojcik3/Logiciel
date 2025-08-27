# Logiciel

Ce dépôt contient une application Python à interface graphique permettant de manipuler différents outils « Contexte éco ».

## Lancer l'application

Assurez-vous d'utiliser Python 3. Pour démarrer l'interface, exécutez :

```bash
python Start.py
```

L'application s'ouvrira avec plusieurs onglets proposant les fonctionnalités principales.

## Démarrage rapide

L'ouverture de la fenêtre est immédiate : les onglets sont créés seulement lorsqu'on clique dessus et le scan réseau se fait en arrière-plan.

## Variables d'environnement

- `LAZY_TABS` (défaut `1`) : mettre à `0` pour charger tous les onglets dès le démarrage.
- `APP_SPLASH` (défaut `0`) : mettre à `1` pour afficher une petite fenêtre d'attente.

## Structure du projet

- `Start.py` : point d'entrée qui ouvre l'interface graphique.
- `modules/` : contient la logique métier et les composants de l'interface.
- `Shapefile_Flore_Patri/` et `Template word Contexte éco/` : ressources utilisées par l'application.

## Dépendances

Les modules standard de Python suffisent pour lancer l'application de base. Certaines fonctionnalités peuvent requérir des bibliothèques supplémentaires, installables via `pip`.

