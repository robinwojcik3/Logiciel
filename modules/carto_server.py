# -*- coding: utf-8 -*-
"""
Serveur Flask local pour l'onglet Carto
Reproduit l'architecture de FloreApp en local avec proxy GBIF et service des fichiers statiques
"""

import os
import sys
import json
import requests
import tempfile
import webbrowser
from typing import Dict, Any, Optional
from flask import Flask, render_template, request, Response, jsonify, send_from_directory
from flask_cors import CORS
import threading
import time

# Configuration
FLASK_PORT = 5000
FLASK_HOST = '127.0.0.1'

class CartoServer:
    """Serveur Flask pour l'onglet Carto"""
    
    def __init__(self, project_root: str):
        self.project_root = project_root
        self.app = Flask(__name__, 
                        template_folder=os.path.join(project_root, 'carto_web', 'templates'),
                        static_folder=os.path.join(project_root, 'carto_web', 'static'))
        
        # Activer CORS pour les requêtes locales
        CORS(self.app)
        
        # Configuration
        self.app.config['SECRET_KEY'] = 'carto-local-dev'
        
        # Variables d'environnement
        self.ign_api_key = os.getenv('IGN_API_KEY', 'essentiels')
        
        self._setup_routes()
    
    def _setup_routes(self):
        """Configure les routes Flask"""
        
        @self.app.route('/')
        def index():
            return render_template('index.html')
        
        @self.app.route('/carto')
        def carto():
            """Page principale de l'onglet Carto"""
            return render_template('biblio-patri.html', ign_api_key=self.ign_api_key)
        
        @self.app.route('/api/gbif')
        def gbif_proxy():
            """Proxy pour l'API GBIF - reproduit gbif-proxy.js de Netlify"""
            endpoint = request.args.get('endpoint')
            
            if not endpoint or endpoint not in ['match', 'search', 'synonyms']:
                return jsonify({'error': 'Invalid or missing endpoint'}), 400
            
            try:
                if endpoint == 'synonyms':
                    usage_key = request.args.get('usageKey')
                    if not usage_key:
                        return jsonify({'error': 'Missing usageKey'}), 400
                    
                    url = f'https://api.gbif.org/v1/species/{usage_key}/synonyms'
                    params = {k: v for k, v in request.args.items() 
                             if k not in ['endpoint', 'usageKey']}
                
                elif endpoint == 'match':
                    url = 'https://api.gbif.org/v1/species/match'
                    params = {k: v for k, v in request.args.items() if k != 'endpoint'}
                
                else:  # search
                    url = 'https://api.gbif.org/v1/occurrence/search'
                    params = {k: v for k, v in request.args.items() if k != 'endpoint'}
                
                # Ajouter User-Agent pour GBIF
                headers = {'User-Agent': 'Bota-Logiciel/1.0 (Local Application)'}
                
                response = requests.get(url, params=params, headers=headers, timeout=30)
                
                return Response(
                    response.content,
                    status=response.status_code,
                    content_type=response.headers.get('Content-Type', 'application/json')
                )
                
            except requests.exceptions.Timeout:
                return jsonify({'error': 'GBIF API timeout'}), 504
            except requests.exceptions.RequestException as e:
                return jsonify({'error': f'GBIF API error: {str(e)}'}), 502
        
        @self.app.route('/api/config')
        def get_config():
            """Retourne la configuration pour le client"""
            return jsonify({
                'ign_api_key': self.ign_api_key,
                'server_url': f'http://{FLASK_HOST}:{FLASK_PORT}'
            })
        
        @self.app.route('/data/<path:filename>')
        def serve_data(filename):
            """Sert les fichiers de données (shapefiles, JSON, CSV)"""
            data_dir = os.path.join(self.project_root, 'Bases de données')
            if os.path.exists(data_dir):
                return send_from_directory(data_dir, filename)
            return "File not found", 404
        
        @self.app.route('/shapefiles/<path:filename>')
        def serve_shapefiles(filename):
            """Sert les shapefiles patrimoniaux"""
            shp_dir = os.path.join(self.project_root, 'Shapefile_Flore_Patri')
            if os.path.exists(shp_dir):
                return send_from_directory(shp_dir, filename)
            return "Shapefile not found", 404
    
    def run(self, debug=False, open_browser=True):
        """Lance le serveur Flask"""
        if open_browser:
            # Ouvrir le navigateur après un court délai
            def open_browser_delayed():
                time.sleep(1.5)
                webbrowser.open(f'http://{FLASK_HOST}:{FLASK_PORT}/carto')
            
            threading.Thread(target=open_browser_delayed, daemon=True).start()
        
        print(f"🗺️  Serveur Carto démarré sur http://{FLASK_HOST}:{FLASK_PORT}")
        print(f"📍 Interface Carto: http://{FLASK_HOST}:{FLASK_PORT}/carto")
        
        self.app.run(host=FLASK_HOST, port=FLASK_PORT, debug=debug, use_reloader=False)


def start_carto_server(project_root: str = None):
    """Point d'entrée pour démarrer le serveur Carto"""
    if not project_root:
        # Détecter automatiquement le répertoire du projet
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(current_dir)  # Remonte d'un niveau depuis modules/
    
    server = CartoServer(project_root)
    server.run()


if __name__ == '__main__':
    start_carto_server()
