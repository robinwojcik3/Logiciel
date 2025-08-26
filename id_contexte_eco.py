#!/usr/bin/env "C:\\USERS\\UTILISATEUR\\Mon Drive\\1 - Bota & Travail\\+++++++++  BOTA  +++++++++\\---------------------- 3) BDD\\PYTHON\\2) Contexte éco\\INPUT\\Configuration Python\\venv_geopandas\\Scripts\\python.exe"
# -*- coding: utf-8 -*-
"""Module d'identification des zonages.

Ce module reprend le script fourni mais expose une fonction
`run_id_contexte_eco` qui attend les chemins des shapefiles
"Aire d'étude élargie" et "Zone d'étude". Les chemins ne sont
plus codés en dur.
"""

import datetime

def log_with_time(message: str) -> None:
    now = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] {message}")

def run_id_contexte_eco(couche_reference1: str, couche_reference2: str) -> None:
    """Exécute le traitement principal.

    Cette version est un extract simplifié ; le reste du script original
    doit être intégré ici. Les chemins des deux couches de référence sont
    passés en paramètres et non plus écrits en dur.
    """
    log_with_time(f"Utilisation de '{couche_reference1}' pour l'aire d'étude élargie")
    log_with_time(f"Utilisation de '{couche_reference2}' pour la zone d'étude")
    # Ici viendrait le code original utilisant les deux chemins fournis.
    log_with_time("Traitement terminé (script raccourci).")
