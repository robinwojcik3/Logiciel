import re
from docx.oxml import OxmlElement, ns

def dms_to_dd(text: str) -> float:
    """Convertir des coordonnées DMS en degrés décimaux."""
    pat = r"(\d{1,3})[°d]\s*(\d{1,2})['m]\s*([\d\.]+)[\"s]?\s*([NSEW])"
    alt = r"(\d{1,3})\s+(\d{1,2})\s+([\d\.]+)\s*([NSEW])"
    m = re.search(pat, text, re.I) or re.search(alt, text, re.I)
    if not m:
        raise ValueError(f"Format DMS invalide : {text}")
    deg, mn, sc, hemi = m.groups()
    dd = float(deg) + float(mn)/60 + float(sc)/3600
    return -dd if hemi.upper() in ("S", "W") else dd


def add_hyperlink(paragraph, url: str, text: str, italic: bool = True):
    """Ajouter un lien hypertexte dans un paragraphe docx."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    fld_simple = OxmlElement('w:hyperlink')
    fld_simple.set(ns.qn('r:id'), r_id)
    run = OxmlElement('w:r')
    r_pr = OxmlElement('w:rPr')
    if italic:
        i = OxmlElement('w:i')
        r_pr.append(i)
    u = OxmlElement('w:u')
    u.set(ns.qn('w:val'), 'single')
    r_pr.append(u)
    run.append(r_pr)
    run_text = OxmlElement('w:t')
    run_text.text = text
    run.append(run_text)
    fld_simple.append(run)
    paragraph._p.append(fld_simple)
