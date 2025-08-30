#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de démarrage rapide pour l'interface Carto
Lance directement le serveur Flask sans passer par l'application principale
"""

import os
import sys

# Ajouter le répertoire parent au path pour les imports
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

def main():
    """Point d'entrée principal"""
    print("🗺️  Démarrage de l'interface Carto...")
    print("=" * 50)
    
    try:
        from modules.carto_server import start_carto_server
        
        # Démarrer le serveur avec le répertoire du projet
        start_carto_server(project_root)
        
    except ImportError as e:
        print(f"❌ Erreur d'import: {e}")
        print("Vérifiez que Flask est installé: pip install Flask Flask-CORS")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n🛑 Arrêt du serveur Carto")
        sys.exit(0)
    except Exception as e:
        print(f"❌ Erreur: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
