"""Assemblage automatique du rapport Word pour le contexte écologique.

Ce module combine les résultats de l'analyse "ID Contexte éco" et les
cartes exportées afin de générer un document Word prêt à l'emploi.  Il
cherche des marqueurs textuels dans un modèle `.docx` et y insère
les tableaux et cartes correspondants.
"""
from __future__ import annotations

import os
from typing import Dict

import pandas as pd
from docx import Document
from docx.shared import Inches


def _insert_table(paragraph, df: pd.DataFrame) -> None:
    """Insert a DataFrame as a table just after the given paragraph."""
    doc = paragraph._parent
    rows, cols = df.shape
    table = doc.add_table(rows=rows + 1, cols=cols)
    table.style = "Light List"

    for j, col in enumerate(df.columns):
        table.cell(0, j).text = str(col)
    for i in range(rows):
        for j in range(cols):
            table.cell(i + 1, j).text = str(df.iat[i, j])

    paragraph._p.addnext(table._tbl)
    paragraph._element.getparent().remove(paragraph._element)


def _insert_image(paragraph, img_path: str) -> None:
    """Insert an image in place of the paragraph marker."""
    paragraph.text = ""
    run = paragraph.add_run()
    run.add_picture(img_path, width=Inches(6))


def generate_report(excel_path: str, images_dir: str, template_path: str,
                    output_path: str, mapping: Dict[str, Dict[str, str]]) -> str:
    """Génère un rapport Word à partir du modèle.

    Parameters
    ----------
    excel_path : str
        Chemin vers le fichier Excel produit par l'analyse des zonages.
    images_dir : str
        Dossier contenant les cartes exportées au format PNG.
    template_path : str
        Modèle Word à utiliser comme base. Le fichier n'est pas modifié.
    output_path : str
        Chemin du document Word généré.
    mapping : dict
        Dictionnaire décrivant les correspondances entre les feuilles
        Excel, les marqueurs dans le document et les noms de fichiers PNG.
        Exemple::

            {
                "Natura 2000": {
                    "table": "TABLEAU NATURA2000",
                    "image": "CARTE NATURA2000",
                    "png": "Contexte éco - N2000__AE.png",
                },
            }

    Returns
    -------
    str
        Chemin du fichier Word généré.
    """
    doc = Document(template_path)
    xls = pd.ExcelFile(excel_path)

    for sheet, cfg in mapping.items():
        if sheet in xls.sheet_names:
            df = xls.parse(sheet)
            marker = cfg.get("table")
            for p in doc.paragraphs:
                if p.text.strip() == marker:
                    _insert_table(p, df)
                    break
        img_name = cfg.get("png")
        img_marker = cfg.get("image")
        img_path = os.path.join(images_dir, img_name)
        if os.path.isfile(img_path):
            for p in doc.paragraphs:
                if p.text.strip() == img_marker:
                    _insert_image(p, img_path)
                    break
    doc.save(output_path)
    return output_path

