# -*- coding: utf-8 -*-
"""
Utilitaires pour l'onglet Carto - Gestion des couches et services externes

Reproduit les fonctionnalités des modules JavaScript:
- arcgis.js: Requêtes vers les services ArcGIS REST
- vegetationLayer.js: Gestion de la couche végétation
- labelUtils.js: Calcul des centroïdes et gestion des labels
- ignWmtsLayers.js: Configuration des couches WMTS IGN
"""

import os
import json
import requests
import tempfile
from typing import Dict, List, Optional, Tuple, Any
from urllib.parse import urlencode
import geopandas as gpd
import shapely.geometry as geom
from shapely.ops import transform
import pyproj
from functools import partial


class ArcGISService:
    """Utilitaires pour interroger des services ArcGIS REST"""
    
    def __init__(self, base_url: str):
        self.base_url = base_url.rstrip('/')
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Bota-Logiciel/1.0'
        })
    
    def build_query_url(self, bbox: Tuple[float, float, float, float], 
                       where_clause: str = "1=1", 
                       out_fields: str = "*",
                       out_sr: int = 4326) -> str:
        """Construit une URL de requête avec bounding box et critères"""
        xmin, ymin, xmax, ymax = bbox
        
        params = {
            'where': where_clause,
            'geometry': f"{xmin},{ymin},{xmax},{ymax}",
            'geometryType': 'esriGeometryEnvelope',
            'spatialRel': 'esriSpatialRelIntersects',
            'outFields': out_fields,
            'returnGeometry': 'true',
            'outSR': out_sr,
            'f': 'geojson'
        }
        
        return f"{self.base_url}/query?" + urlencode(params)
    
    def fetch_page(self, url: str, result_offset: int = 0, 
                   result_record_count: int = 1000) -> Dict[str, Any]:
        """Récupère une page de résultats"""
        params = {
            'resultOffset': result_offset,
            'resultRecordCount': result_record_count
        }
        
        full_url = url + '&' + urlencode(params)
        
        try:
            response = self.session.get(full_url, timeout=30)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            print(f"Erreur lors de la requête ArcGIS: {e}")
            return {'features': [], 'exceededTransferLimit': False}
    
    def fetch_all_pages(self, bbox: Tuple[float, float, float, float],
                       where_clause: str = "1=1") -> List[Dict[str, Any]]:
        """Pagine et récupère tous les résultats pour une emprise donnée"""
        url = self.build_query_url(bbox, where_clause)
        all_features = []
        offset = 0
        page_size = 1000
        
        while True:
            page_data = self.fetch_page(url, offset, page_size)
            features = page_data.get('features', [])
            
            if not features:
                break
                
            all_features.extend(features)
            
            # Vérifier si il y a plus de données
            if not page_data.get('exceededTransferLimit', False):
                break
                
            offset += page_size
            
            # Limite de sécurité
            if len(all_features) > 10000:
                print(f"Limite de 10000 features atteinte pour la requête ArcGIS")
                break
        
        return all_features


class VegetationLayer:
    """Gestion de la couche végétation potentielle"""
    
    def __init__(self, service_url: str):
        self.arcgis_service = ArcGISService(service_url)
        self.cache = {}  # Cache par emprise
        self.color_palette = self._load_color_palette()
    
    def _load_color_palette(self) -> Dict[str, str]:
        """Charge la palette de couleurs pour les unités de végétation"""
        # Palette simplifiée - à adapter selon les données réelles
        return {
            'UCV1': '#228B22',  # Forêt de feuillus
            'UCV2': '#006400',  # Forêt de conifères
            'UCV3': '#32CD32',  # Forêt mixte
            'UCV4': '#9ACD32',  # Landes
            'UCV5': '#ADFF2F',  # Prairies
            'UCV6': '#7CFC00',  # Pelouses
            'UCV7': '#00FF7F',  # Zones humides
            'default': '#90EE90'
        }
    
    def get_style_for_ucv(self, ucv_code: str) -> Dict[str, Any]:
        """Retourne le style pour un code UCV donné"""
        color = self.color_palette.get(ucv_code, self.color_palette['default'])
        return {
            'color': color,
            'weight': 2,
            'opacity': 0.8,
            'fillColor': color,
            'fillOpacity': 0.3
        }
    
    def fetch_vegetation_data(self, bbox: Tuple[float, float, float, float]) -> List[Dict[str, Any]]:
        """Récupère les données de végétation pour une emprise"""
        cache_key = f"{bbox[0]:.4f},{bbox[1]:.4f},{bbox[2]:.4f},{bbox[3]:.4f}"
        
        if cache_key in self.cache:
            return self.cache[cache_key]
        
        features = self.arcgis_service.fetch_all_pages(bbox)
        
        # Traitement des features pour ajouter le style
        for feature in features:
            properties = feature.get('properties', {})
            ucv_code = properties.get('UCV', 'default')
            properties['style'] = self.get_style_for_ucv(ucv_code)
        
        self.cache[cache_key] = features
        return features


class WFSService:
    """Service pour les couches WFS (contexte écologique)"""
    
    def __init__(self, base_url: str):
        self.base_url = base_url.rstrip('/')
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Bota-Logiciel/1.0'
        })
    
    def fetch_wfs_data(self, layer_name: str, bbox: Tuple[float, float, float, float],
                      srs_name: str = "EPSG:4326") -> Dict[str, Any]:
        """Récupère les données WFS pour une couche et une emprise"""
        xmin, ymin, xmax, ymax = bbox
        bbox_str = f"{xmin},{ymin},{xmax},{ymax}"
        
        params = {
            'service': 'WFS',
            'request': 'GetFeature',
            'version': '2.0.0',
            'typeNames': layer_name,
            'srsName': srs_name,
            'bbox': bbox_str,
            'outputFormat': 'application/json'
        }
        
        try:
            response = self.session.get(self.base_url, params=params, timeout=30)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            print(f"Erreur lors de la requête WFS pour {layer_name}: {e}")
            return {'features': []}


class LabelUtils:
    """Utilitaires pour le calcul des centroïdes et gestion des labels"""
    
    @staticmethod
    def calculate_centroid(geometry: Dict[str, Any]) -> Optional[Tuple[float, float]]:
        """Calcule le centroïde d'une géométrie GeoJSON"""
        try:
            geom_obj = geom.shape(geometry)
            centroid = geom_obj.centroid
            return (centroid.x, centroid.y)
        except Exception as e:
            print(f"Erreur lors du calcul du centroïde: {e}")
            return None
    
    @staticmethod
    def avoid_label_collision(labels: List[Dict[str, Any]], 
                            min_distance: float = 0.001) -> List[Dict[str, Any]]:
        """Évite les collisions entre les labels"""
        if len(labels) <= 1:
            return labels
        
        # Algorithme simple de déplacement des labels qui se chevauchent
        adjusted_labels = []
        
        for i, label in enumerate(labels):
            pos = label.get('position', (0, 0))
            adjusted_pos = list(pos)
            
            # Vérifier les collisions avec les labels précédents
            for prev_label in adjusted_labels:
                prev_pos = prev_label.get('position', (0, 0))
                distance = ((adjusted_pos[0] - prev_pos[0])**2 + 
                           (adjusted_pos[1] - prev_pos[1])**2)**0.5
                
                if distance < min_distance:
                    # Déplacer le label
                    adjusted_pos[0] += min_distance
                    adjusted_pos[1] += min_distance * 0.5
            
            adjusted_label = label.copy()
            adjusted_label['position'] = tuple(adjusted_pos)
            adjusted_labels.append(adjusted_label)
        
        return adjusted_labels
    
    @staticmethod
    def scale_labels_for_zoom(labels: List[Dict[str, Any]], 
                            zoom_level: int) -> List[Dict[str, Any]]:
        """Ajuste la taille des labels selon le niveau de zoom"""
        base_size = 12
        scale_factor = max(0.5, min(2.0, zoom_level / 10.0))
        
        scaled_labels = []
        for label in labels:
            scaled_label = label.copy()
            scaled_label['fontSize'] = int(base_size * scale_factor)
            scaled_labels.append(scaled_label)
        
        return scaled_labels


class IGNWMTSLayers:
    """Configuration des couches WMTS IGN historiques"""
    
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.getenv('IGN_API_KEY')
        self.base_url = "https://wxs.ign.fr/{key}/geoportail/wmts"
    
    def build_ign_wmts_historical_layers(self) -> Dict[str, Dict[str, Any]]:
        """Retourne un dictionnaire de couches WMTS historiques"""
        if not self.api_key:
            print("Clé API IGN manquante - couches WMTS désactivées")
            return {}
        
        layers = {
            'ortho_1950_1965': {
                'name': 'Orthophotos 1950-1965',
                'layer': 'ORTHOIMAGERY.ORTHOPHOTOS.1950-1965',
                'style': 'normal',
                'format': 'image/jpeg',
                'url': self.base_url.format(key=self.api_key)
            },
            'ortho_1965_1980': {
                'name': 'Orthophotos 1965-1980',
                'layer': 'ORTHOIMAGERY.ORTHOPHOTOS.1965-1980',
                'style': 'normal',
                'format': 'image/jpeg',
                'url': self.base_url.format(key=self.api_key)
            },
            'ortho_2000_2005': {
                'name': 'Orthophotos 2000-2005',
                'layer': 'ORTHOIMAGERY.ORTHOPHOTOS.2000-2005',
                'style': 'normal',
                'format': 'image/jpeg',
                'url': self.base_url.format(key=self.api_key)
            }
        }
        
        return layers


class ElevationProfile:
    """Calcul du profil d'altitude le long d'une ligne"""
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Bota-Logiciel/1.0'
        })
    
    def get_elevation_data(self, coordinates: List[Tuple[float, float]]) -> List[Dict[str, Any]]:
        """Récupère les données d'altitude pour une liste de coordonnées"""
        # Utiliser l'API Open-Meteo pour les altitudes
        elevations = []
        
        for i, (lon, lat) in enumerate(coordinates):
            try:
                url = f"https://api.open-meteo.com/v1/elevation?latitude={lat}&longitude={lon}"
                response = self.session.get(url, timeout=10)
                response.raise_for_status()
                data = response.json()
                
                elevation = data.get('elevation', [0])[0]
                elevations.append({
                    'distance': i * 100,  # Distance approximative en mètres
                    'elevation': elevation,
                    'coordinates': [lon, lat]
                })
                
            except Exception as e:
                print(f"Erreur lors de la récupération de l'altitude pour {lat}, {lon}: {e}")
                elevations.append({
                    'distance': i * 100,
                    'elevation': 0,
                    'coordinates': [lon, lat]
                })
        
        return elevations
    
    def calculate_profile_from_line(self, line_geojson: Dict[str, Any], 
                                  sample_distance: float = 100.0) -> List[Dict[str, Any]]:
        """Calcule le profil d'altitude le long d'une ligne"""
        try:
            # Convertir en objet Shapely
            line_geom = geom.shape(line_geojson['geometry'])
            
            # Échantillonner la ligne à intervalles réguliers
            total_length = line_geom.length
            num_samples = max(10, int(total_length / sample_distance))
            
            coordinates = []
            for i in range(num_samples + 1):
                distance = (i / num_samples) * total_length
                point = line_geom.interpolate(distance)
                coordinates.append((point.x, point.y))
            
            # Récupérer les altitudes
            return self.get_elevation_data(coordinates)
            
        except Exception as e:
            print(f"Erreur lors du calcul du profil d'altitude: {e}")
            return []


class ShapefileHandler:
    """Gestion de l'import/export de shapefiles"""
    
    @staticmethod
    def import_shapefile(file_path: str) -> Optional[Dict[str, Any]]:
        """Importe un shapefile et le convertit en GeoJSON"""
        try:
            if file_path.endswith('.zip'):
                # Extraire le ZIP et trouver le .shp
                import zipfile
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    temp_dir = tempfile.mkdtemp()
                    zip_ref.extractall(temp_dir)
                    
                    # Trouver le fichier .shp
                    shp_files = [f for f in os.listdir(temp_dir) if f.endswith('.shp')]
                    if not shp_files:
                        raise ValueError("Aucun fichier .shp trouvé dans l'archive")
                    
                    shp_path = os.path.join(temp_dir, shp_files[0])
            else:
                shp_path = file_path
            
            # Lire avec GeoPandas
            gdf = gpd.read_file(shp_path)
            
            # Convertir en WGS84 si nécessaire
            if gdf.crs and gdf.crs != 'EPSG:4326':
                gdf = gdf.to_crs('EPSG:4326')
            
            # Convertir en GeoJSON
            return json.loads(gdf.to_json())
            
        except Exception as e:
            print(f"Erreur lors de l'import du shapefile: {e}")
            return None
    
    @staticmethod
    def export_geojson_to_shapefile(geojson_data: Dict[str, Any], 
                                   output_path: str) -> bool:
        """Exporte des données GeoJSON vers un shapefile"""
        try:
            # Créer un GeoDataFrame depuis le GeoJSON
            gdf = gpd.GeoDataFrame.from_features(geojson_data['features'])
            
            # Définir le CRS
            gdf.crs = 'EPSG:4326'
            
            # Exporter
            gdf.to_file(output_path, driver='ESRI Shapefile')
            return True
            
        except Exception as e:
            print(f"Erreur lors de l'export du shapefile: {e}")
            return False


# Configuration des couches de contexte écologique
CONTEXT_LAYERS_CONFIG = {
    'znieff1': {
        'name': 'ZNIEFF de type I',
        'wfs_layer': 'PROTECTEDAREAS.ZNIEFF1',
        'color': '#ff0000',
        'group': 'Contexte écologique'
    },
    'znieff2': {
        'name': 'ZNIEFF de type II', 
        'wfs_layer': 'PROTECTEDAREAS.ZNIEFF2',
        'color': '#ff8800',
        'group': 'Contexte écologique'
    },
    'natura2000': {
        'name': 'Natura 2000',
        'wfs_layer': 'PROTECTEDAREAS.SIC',
        'color': '#00ff00',
        'group': 'Contexte écologique'
    },
    'sols': {
        'name': 'Carte des sols',
        'wfs_layer': 'GEOLOGIE.SOLS',
        'color': '#8B4513',
        'group': 'Contexte écologique'
    }
}

# URL de base pour les services WFS IGN
WFS_BASE_URL = "https://wxs.ign.fr/environnement/geoportail/wfs"

# URL pour le service de végétation ArcGIS (exemple)
VEGETATION_SERVICE_URL = "https://services.arcgis.com/example/vegetation/FeatureServer/0"
