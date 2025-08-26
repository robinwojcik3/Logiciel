# Règles de développement

Cette application contient plusieurs boutons qui lancent des actions Python (export, analyses, etc.).

- **Boutons marche / arrêt** : chaque bouton d'action doit pouvoir être pressé de nouveau pour annuler l'exécution en cours. L'appui répété arrête le traitement, réinitialise le bouton et nettoie le terminal ainsi que la zone de log.
- **Futur développement** : lors de l'ajout de nouveaux boutons déclenchant une tâche longue, appliquer la même logique de bascule (démarrer/annuler) afin de garder l'interface cohérente.

Ces règles devront être respectées pour toute évolution de ce dépôt.
