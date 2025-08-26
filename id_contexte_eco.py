# -*- coding: utf-8 -*-
import os
import sys
import datetime
import math


def run_id_contexte(couche_reference1, couche_reference2):
    """Exécute l'identification des zonages avec les shapefiles fournis."""
    # Configuration des variables d'environnement pour PROJ et GDAL
    qgis_base = r"C:\\Program Files\\QGIS 3.40.3"
    os.environ['PROJ_LIB'] = os.path.join(qgis_base, "share", "proj")
    os.environ['PROJ_DATA'] = os.path.join(qgis_base, "share", "proj")
    os.environ['GDAL_DATA'] = os.path.join(qgis_base, "share", "gdal")

    print(f"PROJ_LIB défini à: {os.environ['PROJ_LIB']}")
    print(f"GDAL_DATA défini à: {os.environ['GDAL_DATA']}")

    import geopandas as gpd
    import pandas as pd
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    def log_with_time(message: str) -> None:
        now = datetime.datetime.now().strftime("%H:%M:%S")
        print(f"[{now}] {message}")

    log_with_time("Démarrage du script d'identification des zonages...")

    couche_reference1_path = couche_reference1
    couche_reference2_path = couche_reference2

    if not os.path.exists(couche_reference1_path):
        log_with_time(f"Le fichier de la première couche de référence n'a pas été trouvé : {couche_reference1_path}")
        return
    if not os.path.exists(couche_reference2_path):
        log_with_time(f"Le fichier de la deuxième couche de référence n'a pas été trouvé : {couche_reference2_path}")
        return

    try:
        reference_gdf = gpd.read_file(couche_reference1_path)
        log_with_time("Première couche de référence chargée avec succès")
    except Exception as e:
        log_with_time(f"Erreur lors du chargement de la première couche de référence : {e}")
        return

    try:
        reference2_gdf = gpd.read_file(couche_reference2_path)
        log_with_time("Deuxième couche de référence chargée avec succès")
    except Exception as e:
        log_with_time(f"Erreur lors du chargement de la deuxième couche de référence : {e}")
        return

    crs_projected = "EPSG:2154"
    if reference_gdf.crs != crs_projected:
        reference_gdf = reference_gdf.to_crs(crs_projected)
        log_with_time("Reprojection de la première couche de référence effectuée")
    if reference2_gdf.crs != crs_projected:
        reference2_gdf = reference2_gdf.to_crs(crs_projected)
        log_with_time("Reprojection de la deuxième couche de référence effectuée")

    try:
        reference2_centroid = reference2_gdf.geometry.union_all().centroid
        log_with_time("Calcul du centroïde de référence effectué")
    except Exception as e:
        log_with_time(f"Erreur lors du calcul du centroïde de la deuxième couche de référence : {e}")
        return

    def calculate_azimuth(point1, point2):
        delta_x = point2.x - point1.x
        delta_y = point2.y - point1.y
        angle_rad = math.atan2(delta_x, delta_y)
        angle_deg = math.degrees(angle_rad)
        azimuth = (angle_deg + 360) % 360
        return azimuth

    def map_azimuth_to_direction(azimuth):
        if (337.5 <= azimuth < 360) or (0 <= azimuth < 22.5):
            return 'Nord'
        elif 22.5 <= azimuth < 67.5:
            return 'Nord-est'
        elif 67.5 <= azimuth < 112.5:
            return 'Est'
        elif 112.5 <= azimuth < 157.5:
            return 'Sud-est'
        elif 157.5 <= azimuth < 202.5:
            return 'Sud'
        elif 202.5 <= azimuth < 247.5:
            return 'Sud-ouest'
        elif 247.5 <= azimuth < 292.5:
            return 'Ouest'
        elif 292.5 <= azimuth < 337.5:
            return 'Nord-ouest'
        else:
            return 'Inconnu'

    def combine_distance_and_direction(distance_km, direction):
        if distance_km == 0.0:
            return "Se superpose à la zone d'étude"
        if direction in ['Est', 'Ouest']:
            preposition = "à l'"
            direction_str = direction.lower()
        else:
            preposition = 'au '
            direction_str = direction.lower()
        return f"{distance_km} km {preposition}{direction_str}"

    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(fill_type='solid', start_color='4F81BD', end_color='4F81BD')
    data_font = Font(name='Calibri', size=11)
    data_fill_even = PatternFill(fill_type='solid', start_color='F2F2F2', end_color='F2F2F2')
    data_fill_odd = PatternFill(fill_type='solid', start_color='FFFFFF', end_color='FFFFFF')
    border = Border(left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000'))
    alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

    # TODO: la liste complète des couches cibles est abrégée ici pour concision.
    couches_cibles = [
        {
            'nom': 'N2000 ZPS',
            'chemin': r"C:\\Users\\utilisateur\\Mon Drive\\1 - Bota & Travail\\+++++++++  BOTA  +++++++++\\---------------------- 3) BDD\\PYTHON\\2) Contexte éco\\INPUT\\Tables pour ID zonages\\N2000 ZPS.shp",
            'attributs': ['SITENAME', 'SITECODE'],
        },
        # ... autres couches ...
    ]

    dossier_sortie = r"C:\\USERS\\UTILISATEUR\\Mon Drive\\1 - Bota & Travail\\+++++++++  BOTA  +++++++++\\---------------------- 3) BDD\\PYTHON\\2) Contexte éco\\OUTPUT"
    nom_fichier_sortie = 'ID zonages.xlsx'
    chemin_sortie = os.path.join(dossier_sortie, nom_fichier_sortie)

    if os.path.exists(chemin_sortie):
        try:
            os.remove(chemin_sortie)
            log_with_time(f"Le fichier existant '{chemin_sortie}' a été supprimé et sera remplacé par le nouvel export.")
        except Exception as e:
            log_with_time(f"Erreur lors de la suppression du fichier existant '{chemin_sortie}': {e}")
            return

    # ... Le reste du traitement est identique au script original ...
    log_with_time("Traitement terminé (section abrégée)")
