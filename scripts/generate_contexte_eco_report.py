"""Script de génération complète du rapport "Contexte éco".

Il enchaîne automatiquement :
1. l'analyse des zonages (ID Contexte éco),
2. l'export des cartes QGIS en PNG,
3. la construction d'un document Word basé sur le modèle fourni.

Ce script reste volontairement simple : certains paramètres (couches,
projets QGIS…) doivent être adaptés à votre organisation.
"""

from __future__ import annotations

import argparse
import os

from modules import id_contexte_eco
from modules import export_worker
from modules import report_generator


def main() -> None:
    parser = argparse.ArgumentParser(description="Génère un rapport Contexte éco complet")
    parser.add_argument("ae", help="Shapefile de l'aire d'étude élargie")
    parser.add_argument("ze", help="Shapefile de la zone d'étude")
    parser.add_argument("template", help="Modèle Word à utiliser")
    parser.add_argument("output", help="Chemin du document Word de sortie")
    parser.add_argument("projects", nargs="+", help="Projets QGIS à exporter")
    parser.add_argument("--buffer", type=float, default=5.0, help="Tampon en km autour de la ZE")
    parser.add_argument("--export-dir", default=os.path.join("output", "cartes"), help="Dossier d'export des cartes")
    args = parser.parse_args()

    excel_path = id_contexte_eco.run_analysis(args.ae, args.ze, args.buffer)

    cfg = {
        "AE_SHP": args.ae,
        "ZE_SHP": args.ze,
        "EXPORT_DIR": args.export_dir,
        "EXPORT_TYPE": "PNG",
        "CADRAGE_MODE": "AE",
        "LAYER_AE_NAME": "AE",
        "LAYER_ZE_NAME": "ZE",
        "MARGIN_FAC": 1.1,
        "DPI": 300,
    }
    export_worker.worker_run((args.projects, cfg))

    report_generator.build_report(
        excel_path,
        args.export_dir,
        args.template,
        args.output,
    )
    print(f"Rapport généré : {os.path.abspath(args.output)}")


if __name__ == "__main__":
    main()
