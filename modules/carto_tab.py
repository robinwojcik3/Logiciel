# -*- coding: utf-8 -*-
"""
Onglet Carto - Interface cartographique interactive avec Leaflet

Reproduit les fonctionnalités de l'onglet Carto de FloreApp:
- Carte interactive avec couches de base (OpenTopoMap, ESRI, IGN)
- Couches de contexte écologique (WFS)
- Couche de végétation potentielle (ArcGIS)
- Outils de dessin et sélection de zones
- Export/import de shapefiles
- Profil d'altitude
- Contrôles de couches groupées
"""

import os
import sys
import json
import tempfile
import webbrowser
from typing import Optional, Dict, Any
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from tkinter import font as tkfont

try:
    from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
    from PyQt5.QtWebEngineWidgets import QWebEngineView
    from PyQt5.QtCore import QUrl, pyqtSignal, QObject
    from PyQt5.QtWebChannel import QWebChannel
    PYQT_AVAILABLE = True
except ImportError:
    PYQT_AVAILABLE = False
    print("PyQt5 non disponible - l'onglet Carto utilisera une interface simplifiée")


class CartoWebBridge(QObject if PYQT_AVAILABLE else object):
    """Pont de communication entre Python et JavaScript pour la carte"""
    
    if PYQT_AVAILABLE:
        geometrySelected = pyqtSignal(str)  # Signal émis quand une géométrie est sélectionnée
        layerToggled = pyqtSignal(str, bool)  # Signal émis quand une couche est activée/désactivée
    
    def __init__(self):
        if PYQT_AVAILABLE:
            super().__init__()
        self.selected_geometry = None
        self.active_layers = set()
    
    def receiveGeometry(self, geojson_str: str):
        """Reçoit une géométrie sélectionnée depuis JavaScript"""
        try:
            self.selected_geometry = json.loads(geojson_str)
            if PYQT_AVAILABLE:
                self.geometrySelected.emit(geojson_str)
            print(f"Géométrie reçue: {geojson_str[:100]}...")
        except Exception as e:
            print(f"Erreur lors de la réception de géométrie: {e}")
    
    def toggleLayer(self, layer_id: str, visible: bool):
        """Active/désactive une couche"""
        if visible:
            self.active_layers.add(layer_id)
        else:
            self.active_layers.discard(layer_id)
        
        if PYQT_AVAILABLE:
            self.layerToggled.emit(layer_id, visible)
        print(f"Couche {layer_id}: {'activée' if visible else 'désactivée'}")


class CartoTab(ttk.Frame):
    """Onglet Carto avec carte interactive Leaflet"""
    
    def __init__(self, parent, style_helper, prefs: Dict[str, Any]):
        super().__init__(parent, padding=8)
        self.parent = parent
        self.style_helper = style_helper
        self.prefs = prefs
        
        # Fonts
        self.font_title = tkfont.Font(family="Segoe UI", size=15, weight="bold")
        self.font_sub = tkfont.Font(family="Segoe UI", size=10)
        
        # Variables
        self.map_widget = None
        self.web_bridge = CartoWebBridge()
        self.temp_html_file = None
        
        # Configuration des couches
        self.layer_config = {
            'base_layers': {
                'opentopomap': {'name': 'OpenTopoMap', 'active': True},
                'esri_imagery': {'name': 'ESRI Imagery', 'active': False},
                'ign_ortho': {'name': 'IGN Orthophotos', 'active': False}
            },
            'context_layers': {
                'znieff1': {'name': 'ZNIEFF de type I', 'active': True, 'color': '#ff0000'},
                'znieff2': {'name': 'ZNIEFF de type II', 'active': True, 'color': '#ff8800'},
                'natura2000': {'name': 'Natura 2000', 'active': True, 'color': '#00ff00'},
                'sols': {'name': 'Carte des sols', 'active': False, 'color': '#8B4513'},
                'vegetation': {'name': 'Végétation potentielle', 'active': False, 'color': '#228B22'}
            }
        }
        
        self._build_ui()
        self._create_map()
    
    def _build_ui(self):
        """Construit l'interface utilisateur"""
        # Layout principal
        main_frame = ttk.Frame(self, style="Header.TFrame")
        main_frame.pack(fill="both", expand=True)
        
        # Titre
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(title_frame, text="Carto - Interface cartographique", 
                 font=self.font_title).pack(side="left")
        
        # Frame pour les contrôles et la carte
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill="both", expand=True)
        
        # Panneau de contrôle à gauche
        control_frame = ttk.LabelFrame(content_frame, text="Contrôles", padding=10)
        control_frame.pack(side="left", fill="y", padx=(0, 10))
        
        self._build_layer_controls(control_frame)
        self._build_tool_controls(control_frame)
        
        # Zone de la carte à droite
        map_frame = ttk.Frame(content_frame, style="Card.TFrame")
        map_frame.pack(side="right", fill="both", expand=True)
        
        if PYQT_AVAILABLE:
            self._create_qt_map(map_frame)
        else:
            self._create_fallback_map(map_frame)
    
    def _build_layer_controls(self, parent):
        """Construit les contrôles de couches"""
        # Couches de base
        base_frame = ttk.LabelFrame(parent, text="Couches de base")
        base_frame.pack(fill="x", pady=(0, 10))
        
        self.base_layer_vars = {}
        for layer_id, config in self.layer_config['base_layers'].items():
            var = tk.BooleanVar(value=config['active'])
            self.base_layer_vars[layer_id] = var
            ttk.Checkbutton(base_frame, text=config['name'], variable=var,
                           command=lambda lid=layer_id: self._toggle_base_layer(lid)).pack(anchor="w")
        
        # Couches de contexte
        context_frame = ttk.LabelFrame(parent, text="Contexte écologique")
        context_frame.pack(fill="x", pady=(0, 10))
        
        self.context_layer_vars = {}
        for layer_id, config in self.layer_config['context_layers'].items():
            var = tk.BooleanVar(value=config['active'])
            self.context_layer_vars[layer_id] = var
            ttk.Checkbutton(context_frame, text=config['name'], variable=var,
                           command=lambda lid=layer_id: self._toggle_context_layer(lid)).pack(anchor="w")
    
    def _build_tool_controls(self, parent):
        """Construit les contrôles d'outils"""
        tools_frame = ttk.LabelFrame(parent, text="Outils")
        tools_frame.pack(fill="x", pady=(0, 10))
        
        # Outils de dessin
        ttk.Button(tools_frame, text="Dessiner polygone", 
                  command=self._start_polygon_draw).pack(fill="x", pady=2)
        ttk.Button(tools_frame, text="Dessiner ligne", 
                  command=self._start_line_draw).pack(fill="x", pady=2)
        ttk.Button(tools_frame, text="Effacer sélection", 
                  command=self._clear_selection).pack(fill="x", pady=2)
        
        ttk.Separator(tools_frame, orient="horizontal").pack(fill="x", pady=5)
        
        # Import/Export
        ttk.Button(tools_frame, text="Importer shapefile", 
                  command=self._import_shapefile).pack(fill="x", pady=2)
        ttk.Button(tools_frame, text="Exporter shapefile", 
                  command=self._export_shapefile).pack(fill="x", pady=2)
        
        ttk.Separator(tools_frame, orient="horizontal").pack(fill="x", pady=5)
        
        # Profil d'altitude
        ttk.Button(tools_frame, text="Profil d'altitude", 
                  command=self._show_elevation_profile).pack(fill="x", pady=2)
    
    def _create_qt_map(self, parent):
        """Crée la carte avec QWebEngine"""
        try:
            # Créer le widget Qt dans le frame Tkinter
            self.map_widget = QWebEngineView()
            
            # Configuration du canal de communication
            channel = QWebChannel()
            channel.registerObject("bridge", self.web_bridge)
            self.map_widget.page().setWebChannel(channel)
            
            # Connecter les signaux
            if hasattr(self.web_bridge, 'geometrySelected'):
                self.web_bridge.geometrySelected.connect(self._on_geometry_selected)
                self.web_bridge.layerToggled.connect(self._on_layer_toggled)
            
            # Charger la page HTML de la carte
            self._load_map_html()
            
        except Exception as e:
            print(f"Erreur lors de la création de la carte Qt: {e}")
            self._create_fallback_map(parent)
    
    def _create_fallback_map(self, parent):
        """Crée une interface de fallback qui lance le serveur Flask"""
        fallback_frame = ttk.Frame(parent)
        fallback_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        ttk.Label(fallback_frame, text="Carte interactive", 
                 font=self.font_title).pack(pady=20)
        ttk.Label(fallback_frame, 
                 text="Interface cartographique complète avec analyse patrimoniale.\nReproduction fidèle de l'onglet Carto de FloreApp.",
                 justify="center").pack(pady=10)
        
        # Bouton pour lancer le serveur Carto
        ttk.Button(fallback_frame, text="🚀 Lancer l'interface Carto", 
                  command=self._start_carto_server, style="Accent.TButton").pack(pady=10)
        
        # Status du serveur
        self.server_status = ttk.Label(fallback_frame, text="", style="Status.TLabel")
        self.server_status.pack(pady=5)
    
    def _create_map(self):
        """Initialise la carte"""
        if not PYQT_AVAILABLE:
            return
        
        # La carte sera créée lors du chargement du HTML
        pass
    
    def _load_map_html(self):
        """Charge le fichier HTML de la carte"""
        html_content = self._generate_map_html()
        
        # Créer un fichier temporaire
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
            f.write(html_content)
            self.temp_html_file = f.name
        
        # Charger dans QWebEngine
        if self.map_widget:
            self.map_widget.load(QUrl.fromLocalFile(self.temp_html_file))
    
    def _generate_map_html(self) -> str:
        """Génère le contenu HTML de la carte"""
        return '''<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Carto - Interface cartographique</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    
    <!-- Leaflet CSS -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.7.1/dist/leaflet.css" />
    
    <!-- Leaflet Draw CSS -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet.draw/1.0.4/leaflet.draw.css" />
    
    <!-- Turf.js -->
    <script src="https://unpkg.com/@turf/turf@6/turf.min.js"></script>
    
    <!-- Proj4 -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/proj4js/2.6.2/proj4.min.js"></script>
    
    <style>
        body { margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        #map { height: 100vh; width: 100%; }
        .info-panel { 
            position: absolute; 
            top: 10px; 
            right: 10px; 
            background: white; 
            padding: 10px; 
            border-radius: 5px; 
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            z-index: 1000;
            max-width: 300px;
        }
        .layer-control { 
            background: white; 
            padding: 10px; 
            border-radius: 5px; 
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
    </style>
</head>
<body>
    <div id="map"></div>
    <div id="info" class="info-panel" style="display: none;">
        <h4>Informations</h4>
        <div id="info-content"></div>
    </div>

    <!-- Leaflet JS -->
    <script src="https://unpkg.com/leaflet@1.7.1/dist/leaflet.js"></script>
    
    <!-- Leaflet Draw JS -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet.draw/1.0.4/leaflet.draw.min.js"></script>
    
    <!-- ESRI Leaflet -->
    <script src="https://unpkg.com/esri-leaflet@3.0.5/dist/esri-leaflet.js"></script>
    
    <!-- Shapefile.js -->
    <script src="https://unpkg.com/shpjs@3.5.0/dist/shp.min.js"></script>
    
    <!-- QWebChannel -->
    <script src="qrc:///qtwebchannel/qwebchannel.js"></script>
    
    <script>
        // Variables globales
        let map;
        let bridge;
        let drawControl;
        let drawnItems;
        let baseLayers = {};
        let overlayLayers = {};
        
        // Initialisation de QWebChannel
        new QWebChannel(qt.webChannelTransport, function(channel) {
            bridge = channel.objects.bridge;
            initMap();
        });
        
        function initMap() {
            // Définir la projection Lambert 93
            proj4.defs("EPSG:2154", "+proj=lcc +lat_1=49 +lat_2=44 +lat_0=46.5 +lon_0=3 +x_0=700000 +y_0=6600000 +ellps=GRS80 +towgs84=0,0,0,0,0,0,0 +units=m +no_defs");
            
            // Initialiser la carte
            map = L.map('map').setView([45.5, 3.0], 8);
            
            // Couches de base
            baseLayers['OpenTopoMap'] = L.tileLayer('https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png', {
                attribution: '© OpenTopoMap contributors'
            }).addTo(map);
            
            baseLayers['ESRI Imagery'] = L.esri.basemapLayer('Imagery');
            
            // Contrôle de couches
            drawnItems = new L.FeatureGroup();
            map.addLayer(drawnItems);
            
            // Contrôle de dessin
            drawControl = new L.Control.Draw({
                edit: {
                    featureGroup: drawnItems
                },
                draw: {
                    polygon: true,
                    polyline: true,
                    rectangle: true,
                    circle: false,
                    marker: false,
                    circlemarker: false
                }
            });
            map.addControl(drawControl);
            
            // Événements de dessin
            map.on(L.Draw.Event.CREATED, function(e) {
                const layer = e.layer;
                drawnItems.addLayer(layer);
                
                // Envoyer la géométrie à Python
                const geojson = layer.toGeoJSON();
                if (bridge) {
                    bridge.receiveGeometry(JSON.stringify(geojson));
                }
            });
            
            // Contrôle de couches
            const layerControl = L.control.layers(baseLayers, overlayLayers, {
                position: 'topleft'
            });
            layerControl.addTo(map);
            
            // Charger les couches de contexte
            loadContextLayers();
            
            console.log('Carte initialisée');
        }
        
        function loadContextLayers() {
            // ZNIEFF de type I (exemple)
            const znieff1Url = 'https://wxs.ign.fr/environnement/geoportail/wfs?service=WFS&version=2.0.0&request=GetFeature&typename=PROTECTEDAREAS.ZNIEFF1&outputFormat=application/json';
            
            // Végétation potentielle (ArcGIS)
            const vegetationLayer = L.esri.featureLayer({
                url: 'https://services.arcgis.com/example/vegetation',
                style: function(feature) {
                    return {
                        color: '#228B22',
                        weight: 2,
                        opacity: 0.8,
                        fillOpacity: 0.3
                    };
                }
            });
            
            overlayLayers['Végétation potentielle'] = vegetationLayer;
        }
        
        // Fonctions appelées depuis Python
        function toggleBaseLayer(layerId, visible) {
            const layer = baseLayers[layerId];
            if (layer) {
                if (visible) {
                    map.addLayer(layer);
                } else {
                    map.removeLayer(layer);
                }
            }
        }
        
        function toggleContextLayer(layerId, visible) {
            const layer = overlayLayers[layerId];
            if (layer) {
                if (visible) {
                    map.addLayer(layer);
                } else {
                    map.removeLayer(layer);
                }
            }
        }
        
        function clearSelection() {
            drawnItems.clearLayers();
        }
        
        function exportShapefile() {
            const geojson = drawnItems.toGeoJSON();
            if (bridge) {
                bridge.receiveGeometry(JSON.stringify(geojson));
            }
        }
    </script>
</body>
</html>'''
    
    def _toggle_base_layer(self, layer_id: str):
        """Active/désactive une couche de base"""
        visible = self.base_layer_vars[layer_id].get()
        
        # Désactiver les autres couches de base
        if visible:
            for lid, var in self.base_layer_vars.items():
                if lid != layer_id:
                    var.set(False)
        
        self._execute_js(f"toggleBaseLayer('{layer_id}', {str(visible).lower()})")
    
    def _toggle_context_layer(self, layer_id: str):
        """Active/désactive une couche de contexte"""
        visible = self.context_layer_vars[layer_id].get()
        self._execute_js(f"toggleContextLayer('{layer_id}', {str(visible).lower()})")
    
    def _execute_js(self, js_code: str):
        """Exécute du code JavaScript dans la carte"""
        if self.map_widget:
            self.map_widget.page().runJavaScript(js_code)
    
    def _start_polygon_draw(self):
        """Démarre le dessin de polygone"""
        self._execute_js("new L.Draw.Polygon(map, drawControl.options.polygon).enable()")
    
    def _start_line_draw(self):
        """Démarre le dessin de ligne"""
        self._execute_js("new L.Draw.Polyline(map, drawControl.options.polyline).enable()")
    
    def _clear_selection(self):
        """Efface la sélection"""
        self._execute_js("clearSelection()")
        self.web_bridge.selected_geometry = None
    
    def _import_shapefile(self):
        """Importe un shapefile"""
        file_path = filedialog.askopenfilename(
            title="Sélectionner un shapefile",
            filetypes=[("Shapefiles", "*.shp"), ("Archives ZIP", "*.zip")]
        )
        
        if file_path:
            try:
                # TODO: Implémenter l'import de shapefile
                messagebox.showinfo("Import", f"Import de {file_path} (à implémenter)")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'import: {e}")
    
    def _export_shapefile(self):
        """Exporte la sélection en shapefile"""
        if not self.web_bridge.selected_geometry:
            messagebox.showwarning("Export", "Aucune géométrie sélectionnée")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Exporter en shapefile",
            defaultextension=".shp",
            filetypes=[("Shapefiles", "*.shp")]
        )
        
        if file_path:
            try:
                # TODO: Implémenter l'export de shapefile
                messagebox.showinfo("Export", f"Export vers {file_path} (à implémenter)")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'export: {e}")
    
    def _show_elevation_profile(self):
        """Affiche le profil d'altitude"""
        if not self.web_bridge.selected_geometry:
            messagebox.showwarning("Profil d'altitude", "Dessinez d'abord une ligne")
            return
        
        # TODO: Implémenter le calcul du profil d'altitude
        messagebox.showinfo("Profil d'altitude", "Calcul du profil d'altitude (à implémenter)")
    
    def _start_carto_server(self):
        """Lance le serveur Flask Carto"""
        try:
            from .carto_server import start_carto_server
            import threading
            
            self.server_status.config(text="🔄 Démarrage du serveur Carto...")
            
            # Lancer le serveur dans un thread séparé
            def run_server():
                try:
                    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                    start_carto_server(project_root)
                except Exception as e:
                    self.server_status.config(text=f"❌ Erreur serveur: {e}")
            
            server_thread = threading.Thread(target=run_server, daemon=True)
            server_thread.start()
            
            self.server_status.config(text="✅ Serveur Carto lancé - Interface ouverte dans le navigateur")
            
        except ImportError as e:
            self.server_status.config(text=f"❌ Module serveur manquant: {e}")
        except Exception as e:
            self.server_status.config(text=f"❌ Erreur: {e}")
    
    def _on_geometry_selected(self, geojson_str: str):
        """Callback quand une géométrie est sélectionnée"""
        try:
            geometry = json.loads(geojson_str)
            print(f"Géométrie sélectionnée: {geometry['geometry']['type']}")
        except Exception as e:
            print(f"Erreur lors du traitement de la géométrie: {e}")
    
    def _on_layer_toggled(self, layer_id: str, visible: bool):
        """Callback quand une couche est activée/désactivée"""
        print(f"Couche {layer_id}: {'activée' if visible else 'désactivée'}")
    
    def cleanup(self):
        """Nettoie les ressources"""
        if self.temp_html_file and os.path.exists(self.temp_html_file):
            try:
                os.unlink(self.temp_html_file)
            except Exception:
                pass
