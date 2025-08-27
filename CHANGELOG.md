# Changelog

## Unreleased
- Amélioration des performances au démarrage :
  - Chargement différé des onglets.
  - Scan des projets en arrière-plan avec cache de 5 min.
  - Imports lourds déplacés dans les fonctions qui en ont besoin.
  - Option d'écran de démarrage via `APP_SPLASH`.
- Flags d'environnement : `LAZY_TABS`, `APP_SPLASH`.
