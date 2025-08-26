"""Fonctions liées au calcul du contexte écologique."""

from .utils import log_with_time


def analyse_zonages(ae_shp: str, ze_shp: str, buffer_km: float = 5.0) -> None:
    """Appelle le module d'analyse des zonages."""
    log_with_time("Démarrage de l'analyse des zonages")
    from id_contexte_eco import run_analysis
    run_analysis(ae_shp, ze_shp, buffer_km)
