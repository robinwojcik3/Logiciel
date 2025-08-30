# Onglet Carto - Documentation

## Vue d'ensemble

L'onglet **Carto** reproduit toutes les fonctionnalités de l'interface cartographique de FloreApp dans votre application Python. Il offre une carte interactive avec des couches de contexte écologique, des outils de dessin, et des capacités d'import/export de données géospatiales.

## Fonctionnalités implémentées

### 🗺️ Carte interactive
- **Carte Leaflet** intégrée via QWebEngine
- **Couches de base** : OpenTopoMap, ESRI Imagery, IGN Orthophotos historiques
- **Projection Lambert 93** avec conversion automatique vers WGS84
- **Contrôles de navigation** et zoom

### 🌿 Couches de contexte écologique
- **ZNIEFF** de type I et II
- **Natura 2000**
- **Carte des sols**
- **Végétation potentielle** (service ArcGIS)
- **Chargement dynamique** selon l'emprise visible
- **Mise en cache** pour optimiser les performances

### ✏️ Outils de dessin
- **Polygones** et **lignes** interactifs
- **Sélection de zones** d'étude
- **Buffer automatique** pour créer des bandes d'étude
- **Effacement** et modification des sélections

### 📁 Import/Export
- **Import de shapefiles** (format .shp ou .zip)
- **Export vers shapefile** des géométries sélectionnées
- **Support GeoJSON** pour l'échange de données
- **Conversion automatique** des projections

### 📊 Profil d'altitude
- **Calcul automatique** le long des lignes tracées
- **API Open-Meteo** pour les données d'altitude
- **Graphique interactif** avec étages de végétation
- **Échantillonnage régulier** des points

## Architecture technique

### Modules créés

1. **`carto_tab.py`** - Interface utilisateur principale
   - Classe `CartoTab` héritant de `ttk.Frame`
   - Intégration QWebEngine pour la carte Leaflet
   - Contrôles de couches et outils de dessin
   - Communication Python ↔ JavaScript via QWebChannel

2. **`carto_utils.py`** - Utilitaires et services
   - `ArcGISService` : Requêtes vers services ArcGIS REST
   - `VegetationLayer` : Gestion couche végétation potentielle
   - `WFSService` : Accès aux couches WFS IGN
   - `ElevationProfile` : Calcul profils d'altitude
   - `ShapefileHandler` : Import/export shapefiles

### Dépendances ajoutées

```
PyQt5              # Interface graphique Qt
PyQtWebEngine      # Moteur web pour Leaflet
geopandas[all]     # Manipulation données géospatiales (déjà présent)
requests           # Requêtes HTTP (déjà présent)
```

## Configuration

### Variables d'environnement (optionnelles)

```bash
# Clé API IGN pour les orthophotos historiques
IGN_API_KEY=votre_cle_ign
```

### URLs des services

Les URLs des services sont configurables dans `carto_utils.py` :

```python
# Services WFS IGN
WFS_BASE_URL = "https://wxs.ign.fr/environnement/geoportail/wfs"

# Service végétation ArcGIS (à adapter selon vos données)
VEGETATION_SERVICE_URL = "https://services.arcgis.com/example/vegetation/FeatureServer/0"
```

## Utilisation

### Lancement
L'onglet Carto apparaît automatiquement dans l'application si PyQt5 est installé. Sinon, une interface de fallback permet d'ouvrir la carte dans le navigateur.

### Contrôles de couches
- **Couches de base** : Sélection exclusive (une seule active)
- **Couches de contexte** : Sélection multiple possible
- **Activation/désactivation** en temps réel

### Outils de dessin
1. Cliquer sur "Dessiner polygone" ou "Dessiner ligne"
2. Tracer sur la carte
3. La géométrie est automatiquement transmise à Python
4. Utiliser "Effacer sélection" pour recommencer

### Import/Export
- **Import** : Sélectionner un fichier .shp ou .zip
- **Export** : Dessiner une géométrie puis exporter
- **Formats supportés** : Shapefile, GeoJSON

## Fallback sans PyQt5

Si PyQt5 n'est pas disponible, l'onglet propose :
- Message informatif sur les dépendances manquantes
- Bouton pour ouvrir la carte dans le navigateur web
- Fonctionnalités limitées mais utilisables

## Points d'extension

### Ajout de nouvelles couches
Modifier `CONTEXT_LAYERS_CONFIG` dans `carto_utils.py` :

```python
'nouvelle_couche': {
    'name': 'Nom affiché',
    'wfs_layer': 'IDENTIFIANT.WFS',
    'color': '#couleur',
    'group': 'Groupe'
}
```

### Services personnalisés
Créer de nouvelles classes héritant des services de base :

```python
class MonServicePersonnalise(WFSService):
    def __init__(self):
        super().__init__("https://mon-service.fr/wfs")
```

### Styles et couleurs
Personnaliser les palettes dans `VegetationLayer._load_color_palette()`.

## Dépannage

### PyQt5 non installé
```bash
pip install PyQt5 PyQtWebEngine
```

### Erreurs de projection
Vérifier que `pyproj` est à jour :
```bash
pip install --upgrade pyproj
```

### Services WFS inaccessibles
- Vérifier la connectivité réseau
- Adapter les URLs dans la configuration
- Utiliser les données locales du dossier `Cartes contexte éco export/`

## Performances

- **Cache local** pour les requêtes WFS et ArcGIS
- **Pagination automatique** pour les gros datasets
- **Limite de sécurité** à 10 000 features par couche
- **Chargement à la demande** selon l'emprise visible

## Sécurité

- **Validation des entrées** utilisateur
- **Gestion des timeouts** pour les requêtes réseau
- **Nettoyage automatique** des fichiers temporaires
- **Protection contre les injections** dans les requêtes WFS
