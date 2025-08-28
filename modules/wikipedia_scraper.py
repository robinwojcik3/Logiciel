# -*- coding: utf-8 -*-
"""Outils de scraping Wikipédia utilisés par l'application.

Ce module ouvre la page d'une commune française sur ``fr.wikipedia.org`` puis
extrait deux paragraphes normalisés : un sur le climat et un sur l'occupation
des sols (données Corine Land Cover). Le scraping réutilise un navigateur
Selenium partagé qui reste ouvert après l'exécution.
"""

from __future__ import annotations

import logging
import re
import time
import unicodedata
from typing import Dict, Iterable, List, Optional

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

LOGGER = logging.getLogger(__name__)

# Départements courants (ajoutez si besoin). Gardé pour compatibilité avec
# ``main_app`` qui l'importe.
DEP: Dict[str, str] = {
    "01": "Ain",
    "03": "Allier",
    "04": "Alpes-de-Haute-Provence",
    "05": "Hautes-Alpes",
    "06": "Alpes-Maritimes",
    "07": "Ardèche",
    "09": "Ariège",
    "11": "Aude",
    "13": "Bouches-du-Rhône",
    "15": "Cantal",
    "21": "Côte-d'Or",
    "26": "Drôme",
    "30": "Gard",
    "31": "Haute-Garonne",
    "34": "Hérault",
    "38": "Isère",
    "39": "Jura",
    "42": "Loire",
    "43": "Haute-Loire",
    "63": "Puy-de-Dôme",
    "69": "Rhône",
    "73": "Savoie",
    "74": "Haute-Savoie",
    "75": "Paris",
    "83": "Var",
    "84": "Vaucluse",
    "90": "Territoire de Belfort",
}


# Paramètres généraux
TIMEOUT = 15
RETRIES = 2
TYPE_DELAY = 0.5


_shared_driver: Optional[webdriver.Chrome] = None


def get_shared_driver() -> webdriver.Chrome:
    """Retourne une instance unique de ``webdriver.Chrome``.

    Le navigateur reste ouvert entre les appels : il ne faut donc **jamais**
    appeler ``quit`` ou ``close`` sur l'objet renvoyé.
    """

    global _shared_driver
    if _shared_driver is None:
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        _shared_driver = webdriver.Chrome(options=options)
        _shared_driver.maximize_window()
    return _shared_driver


def wait_css(drv: webdriver.Chrome, selector: str, timeout: int = TIMEOUT):
    return WebDriverWait(drv, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
    )


def wait_any(drv: webdriver.Chrome, selectors: Iterable[str], timeout: int = TIMEOUT):
    end = time.time() + timeout
    while time.time() < end:
        for sel in selectors:
            els = drv.find_elements(By.CSS_SELECTOR, sel)
            if els:
                return els[0]
        time.sleep(0.1)
    raise TimeoutException(f"Aucun sélecteur trouvé parmi {selectors}")


def ensure_fr_wikipedia_home(drv: webdriver.Chrome) -> None:
    """S'assure que le navigateur est sur ``fr.wikipedia.org``."""

    if not drv.current_url.startswith("https://fr.wikipedia.org"):
        drv.get("https://fr.wikipedia.org")
    try:
        wait_css(drv, "#searchInput")
    except TimeoutException:
        drv.get("https://fr.wikipedia.org")
        wait_css(drv, "#searchInput")


def resolve_search_results_if_needed(drv: webdriver.Chrome, commune_label: str) -> None:
    """Résout une éventuelle page de résultats/homonymie.

    On cherche un lien dont le titre contient le nom de la commune et le numéro
    de département. À défaut, on ouvre les résultats un par un jusqu'à trouver
    une infobox mentionnant "Commune".
    """

    wait = WebDriverWait(drv, TIMEOUT)
    name_match = re.match(r"^(.*)\s*\((\d{2})\)$", commune_label)
    commune = name_match.group(1) if name_match else commune_label
    dep = name_match.group(2) if name_match else ""

    if drv.current_url.startswith("https://fr.wikipedia.org/wiki/"):
        return  # déjà sur un article

    try:
        results = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.mw-search-result-heading a"))
        )
    except TimeoutException:
        return

    target_index = 0
    for i, link in enumerate(results):
        title = (link.get_attribute("title") or "").lower()
        if commune.lower() in title and dep in title:
            target_index = i
            break

    for i, link in enumerate(results[target_index:]):
        link.click()
        try:
            wait_any(drv, ["h1#firstHeading", "h1 .mw-page-title-main"])  # page chargée
            infobox = drv.find_elements(By.CSS_SELECTOR, "table.infobox")
            if infobox and "commune" in infobox[0].text.lower():
                return
        except TimeoutException:
            pass
        drv.back()
        wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.mw-search-result-heading a"))
        )


def extract_paragraph_by_prefixes(
    soup: BeautifulSoup,
    section_hint,
    prefixes: List[str],
    contains_all: List[str],
) -> str:
    """Extrait un paragraphe selon un jeu de préfixes ou de mots-clés."""

    def _find_section() -> Optional[BeautifulSoup]:
        if isinstance(section_hint, str):
            span = soup.find(
                "span", class_="mw-headline", string=re.compile(section_hint, re.I)
            )
            return span.find_parent("h2") if span else None
        if isinstance(section_hint, tuple) and len(section_hint) == 2:
            span = soup.find(
                "span", class_="mw-headline", string=re.compile(section_hint[0], re.I)
            )
            h2 = span.find_parent("h2") if span else None
            if not h2:
                return None
            for sib in h2.find_next_siblings():
                if sib.name == "h2":
                    break
                if sib.name == "h3" and re.search(section_hint[1], sib.get_text(), re.I):
                    return sib
            return None
        return None

    sect = _find_section()
    if not sect:
        return ""

    for sib in sect.find_next_siblings():
        if sib.name in {sect.name, "h2"}:
            break
        if sib.name == "p":
            txt = sib.get_text(" ", strip=True)
            for pref in prefixes:
                if re.search(pref, txt, re.I):
                    return txt

    for sib in sect.find_next_siblings():
        if sib.name in {sect.name, "h2"}:
            break
        if sib.name == "p":
            txt = sib.get_text(" ", strip=True)
            if all(re.search(c, txt, re.I) for c in contains_all):
                return txt
    return ""


def clean_text(txt: str) -> str:
    if not txt:
        return ""
    txt = re.sub(r"\[\d+\]", "", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return unicodedata.normalize("NFC", txt)


CLIMAT_PREFIXES = [
    r"^Pour la période 1971-2000, la température annuelle",
    r"^Pour la période 1981-2010, la température annuelle",
]
CLIMAT_FALLBACK_CONTAINS = [r"température annuelle", r"précipitations"]

OCCUP_PREFIXES = [r"^L'occupation des sols de la commune, telle qu'elle"]
OCCUP_FALLBACK_CONTAINS = [r"Corine Land Cover", r"occupation des sols"]


def run_wikipedia_scrape(commune_label: str) -> Dict[str, str]:
    """Scrape Wikipédia et renvoie les paragraphes demandés.

    ``commune_label`` doit être de la forme ``"Nom (num_dep)"``.
    """

    drv = get_shared_driver()
    for attempt in range(1, RETRIES + 1):
        try:
            LOGGER.info("Navigation Wikipédia pour %s", commune_label)
            ensure_fr_wikipedia_home(drv)
            box = wait_css(drv, "#searchInput")
            box.clear()
            time.sleep(TYPE_DELAY)
            box.send_keys(commune_label)
            box.send_keys(Keys.ENTER)

            resolve_search_results_if_needed(drv, commune_label)

            wait_any(drv, ["h1#firstHeading", "h1 .mw-page-title-main"])
            url = drv.current_url

            soup = BeautifulSoup(drv.page_source, "lxml")

            climat = extract_paragraph_by_prefixes(
                soup,
                section_hint="Climat",
                prefixes=CLIMAT_PREFIXES,
                contains_all=CLIMAT_FALLBACK_CONTAINS,
            )
            corine = extract_paragraph_by_prefixes(
                soup,
                section_hint=("Urbanisme", "Occupation des sols"),
                prefixes=OCCUP_PREFIXES,
                contains_all=OCCUP_FALLBACK_CONTAINS,
            )

            climat = clean_text(climat) or "Donnée non disponible"
            corine = clean_text(corine) or "Donnée non disponible"

            return {
                "climat": climat,
                "corine": corine,
                "url": url,
                "commune": commune_label,
            }
        except Exception as exc:  # pragma: no cover - logging
            LOGGER.error("Erreur scraping (%s/%s) : %s", attempt, RETRIES, exc)
            if attempt >= RETRIES:
                break
    return {
        "climat": "Donnée non disponible",
        "corine": "Donnée non disponible",
        "url": "",
        "commune": commune_label,
    }


# Ancienne API conservée pour compatibilité éventuelle ---------------------
def fetch_wikipedia_info(commune_query: str):  # pragma: no cover - compat
    """Alias rétrocompatible de ``run_wikipedia_scrape``."""

    return run_wikipedia_scrape(commune_query), get_shared_driver()


__all__ = [
    "DEP",
    "run_wikipedia_scrape",
    "fetch_wikipedia_info",
    "get_shared_driver",
]

