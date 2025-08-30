"""Script CLI pour générer un rapport Word complet.

Usage:
    python scripts/generate_report.py chemin/resultats.xlsx dossier_images modele.docx rapport.docx
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from modules.report_generator import generate_report


def main() -> int:
    p = argparse.ArgumentParser(description="Génère le rapport Contexte éco")
    p.add_argument("excel", help="Fichier Excel de l'analyse ID Contexte éco")
    p.add_argument("images", help="Dossier contenant les cartes exportées")
    p.add_argument("template", help="Modèle Word de départ")
    p.add_argument("output", help="Fichier Word de sortie")
    p.add_argument(
        "--mapping",
        help="JSON décrivant les correspondances feuilles/marqueurs", 
    )
    args = p.parse_args()

    if args.mapping:
        mapping = json.loads(Path(args.mapping).read_text(encoding="utf-8"))
    else:
        mapping = {
            "Natura 2000": {
                "table": "TABLEAU NATURA2000",
                "image": "CARTE NATURA2000",
                "png": "Contexte éco - N2000__AE.png",
            },
            "ENS": {
                "table": "TABLEAU ENS",
                "image": "CARTE ENS",
                "png": "Contexte éco - ENS__AE.png",
            },
            "APPB": {
                "table": "TABLEAU APPB",
                "image": "CARTE APPB",
                "png": "Contexte éco - APPB__AE.png",
            },
        }

    generate_report(args.excel, args.images, args.template, args.output, mapping)
    print("Rapport généré dans", args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
