"""Génération automatique du rapport Word pour le contexte écologique.

Ce module combine les résultats de l'analyse "ID Contexte éco" (fichier Excel)
avec les cartes exportées en PNG afin de produire un document Word basé sur un
modèle. Les marqueurs présents dans le modèle sont remplacés par les tableaux
et images correspondants.

Les marqueurs supportés suivent la structure :
- ``TABLEAU <NOM>`` pour les tableaux issus de l'Excel
- ``CARTE <NOM>``   pour les images PNG

Exemple de noms : ``NATURA2000``, ``ENS``, ``APPB``.
"""

from __future__ import annotations

import os
from typing import Dict, Iterable, List

import pandas as pd
from docx import Document
from docx.shared import Cm


def _load_dataframe(excel_path: str, patterns: Iterable[str]) -> pd.DataFrame:
    """Charge et fusionne les feuilles dont le nom contient un des motifs."""
    xls = pd.ExcelFile(excel_path)
    frames: List[pd.DataFrame] = []
    for sheet in xls.sheet_names:
        if any(pat.lower() in sheet.lower() for pat in patterns):
            frames.append(xls.parse(sheet))
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def _replace_paragraph_with_table(paragraph, df: pd.DataFrame, doc: Document) -> None:
    """Remplace le paragraphe par un tableau construit depuis ``df``."""
    table = doc.add_table(rows=len(df.index) + 1, cols=len(df.columns))
    for j, col in enumerate(df.columns):
        table.cell(0, j).text = str(col)
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        for j, val in enumerate(row):
            table.cell(i, j).text = str(val)
    paragraph._p.addnext(table._tbl)
    paragraph._element.getparent().remove(paragraph._element)


def _replace_paragraph_with_image(paragraph, image_path: str) -> None:
    """Remplace le paragraphe par une image."""
    for run in paragraph.runs:
        run.clear()
    run = paragraph.add_run()
    if os.path.isfile(image_path):
        run.add_picture(image_path, width=Cm(12))
    paragraph.alignment = None


def build_report(
    excel_path: str,
    maps_dir: str,
    template_path: str,
    output_path: str,
    mapping: Dict[str, Dict[str, List[str]]] | None = None,
) -> str:
    """Construit le rapport Word à partir du modèle et retourne le chemin créé."""
    if mapping is None:
        mapping = {
            "NATURA2000": {
                "sheets": ["N2000"],
                "image": "Contexte éco - N2000__AE.png",
            },
            "ENS": {
                "sheets": ["ENS"],
                "image": "Contexte éco - ENS__AE.png",
            },
            "APPB": {
                "sheets": ["APPB"],
                "image": "Contexte éco - APPB__AE.png",
            },
        }

    doc = Document(template_path)

    for para in list(doc.paragraphs):
        text = para.text.strip()
        for name, cfg in mapping.items():
            if text == f"TABLEAU {name}":
                df = _load_dataframe(excel_path, cfg.get("sheets", []))
                _replace_paragraph_with_table(para, df, doc)
                break
            if text == f"CARTE {name}":
                img = os.path.join(maps_dir, cfg.get("image", ""))
                if os.path.isfile(img):
                    _replace_paragraph_with_image(para, img)
                else:
                    _replace_paragraph_with_image(para, img)  # laisse un emplacement vide
                break

    out = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)
    return out


__all__ = ["build_report"]
