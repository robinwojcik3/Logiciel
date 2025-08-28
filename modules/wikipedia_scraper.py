# -*- coding: utf-8 -*-
"""Outils de scraping pour récupérer des informations sur Wikipédia.

Ce module ouvre les pages des communes françaises sur ``fr.wikipedia.org`` et
extrait deux paragraphes normalisés concernant le climat et l'occupation des
sols (Corine Land Cover). Il réutilise une instance unique de ``webdriver`` afin
de ne jamais fermer la fenêtre du navigateur durant la session de
l'application.
"""

from __future__ import annotations

import logging
import re
import time
import unicodedata
from typing import Dict, List, Optional, Tuple

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)

# Départements courants (ajoutez si besoin)
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


# Paramètres généraux du scraping
TIMEOUT = 15
RETRIES = 2
TYPE_DELAY = 0.5

CLIMAT_PREFIXES = [
    r"^Pour la période 1971-2000, la température annuelle",
    r"^Pour la période 1981-2010, la température annuelle",
]
CLIMAT_FALLBACK_CONTAINS = [r"température annuelle", r"précipitations"]

OCCUP_PREFIXES = [r"^L'occupation des sols de la commune, telle qu'elle"]
OCCUP_FALLBACK_CONTAINS = [r"Corine Land Cover", r"occupation des sols"]

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

_SHARED_DRIVER: Optional[webdriver.Chrome] = None


# ---------------------------------------------------------------------------
# Utilitaires Selenium
# ---------------------------------------------------------------------------
def get_shared_driver() -> webdriver.Chrome:
    """Retourne l'instance globale de ``webdriver`` (créée au besoin)."""

    global _SHARED_DRIVER
    if _SHARED_DRIVER is None:
        options = webdriver.ChromeOptions()
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        options.add_argument("--log-level=3")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        _SHARED_DRIVER = webdriver.Chrome(options=options)
        try:
            _SHARED_DRIVER.maximize_window()
        except Exception:
            pass
    return _SHARED_DRIVER


def wait_css(driver: webdriver.Chrome, selector: str, timeout: int) -> webdriver.Remote:
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
    )


def wait_any(driver: webdriver.Chrome, selectors: List[str], timeout: int):
    wait = WebDriverWait(driver, timeout)

    def _inner(drv):
        for sel in selectors:
            els = drv.find_elements(By.CSS_SELECTOR, sel)
            if els:
                return els[0]
        return False

    return wait.until(_inner)


def ensure_fr_wikipedia_home(driver: webdriver.Chrome) -> None:
    """Charge la page d'accueil de fr.wikipedia si nécessaire."""

    if "fr.wikipedia.org" not in (driver.current_url or ""):
        driver.get("https://fr.wikipedia.org")
    try:
        wait_css(driver, "#searchInput", TIMEOUT)
    except TimeoutException:
        driver.get("https://fr.wikipedia.org")
        wait_css(driver, "#searchInput", TIMEOUT)


def resolve_search_results_if_needed(drv: webdriver.Chrome, commune_label: str) -> None:
    """Résout la page de résultats si la recherche ne mène pas directement à l'article."""

    wait = WebDriverWait(drv, 10)
    try:
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div.mw-search-result-heading a")
            )
        )
    except TimeoutException:
        return

    name, dep = commune_label.split("(", 1)
    name = name.strip()
    dep = dep.strip(" )")
    links = drv.find_elements(By.CSS_SELECTOR, "div.mw-search-result-heading a")

    # Priorité au lien contenant le nom et le département
    for link in links:
        title = link.get_attribute("title") or link.text
        if name.lower() in title.lower() and dep in title:
            link.click()
            return

    # Fallback : premier lien menant à une infobox « Commune »
    for link in links:
        link.click()
        try:
            wait_any(drv, ["table.infobox"], 5)
            if re.search(r"Commune", drv.page_source, re.IGNORECASE):
                return
        finally:
            drv.back()
            wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.mw-search-result-heading a")
                )
            )

    if links:
        links[0].click()


# ---------------------------------------------------------------------------
# Extraction HTML
# ---------------------------------------------------------------------------
def extract_paragraph_by_prefixes(
    soup: BeautifulSoup,
    section_hint: str | Tuple[str, str],
    prefixes: List[str],
    contains_all: List[str],
) -> str:
    """Retourne le texte d'un paragraphe correspondant aux critères donnés."""

    section = None
    if isinstance(section_hint, str):
        section = soup.find(
            "h2",
            string=lambda t: t and section_hint.lower() in t.lower(),
        )
    else:
        h2 = soup.find(
            "h2", string=lambda t: t and section_hint[0].lower() in t.lower()
        )
        if h2:
            for h3 in h2.find_all_next("h3"):
                if section_hint[1].lower() in h3.get_text(strip=True).lower():
                    section = h3
                    break

    if not section:
        return ""

    def _iterate_paragraphs(start_tag):
        for sib in start_tag.find_next_siblings():
            if sib.name and sib.name.startswith("h"):
                break
            if sib.name == "p":
                yield sib

    # Priorité : préfixes
    for p in _iterate_paragraphs(section):
        text = p.get_text(" ", strip=True)
        if any(re.search(pref, text, re.IGNORECASE) for pref in prefixes):
            return text

    # Fallback : tokens obligatoires
    for p in _iterate_paragraphs(section):
        text = p.get_text(" ", strip=True)
        if all(re.search(tok, text, re.IGNORECASE) for tok in contains_all):
            return text

    return ""


def clean_text(txt: str) -> str:
    if not txt:
        return ""
    txt = unicodedata.normalize("NFC", txt)
    txt = re.sub(r"\[\d+\]", "", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


# ---------------------------------------------------------------------------
# Fonction principale
# ---------------------------------------------------------------------------
def run_wikipedia_scrape(commune_label: str) -> Dict[str, str]:
    """Lance le scraping Wikipédia pour ``commune_label``.

    ``commune_label`` doit être de la forme ``"Nom (38)"``.
    La fonction renvoie un dictionnaire : ``{"climat": str, "corine": str,
    "url": str, "commune": str}``.
    """

    drv = get_shared_driver()

    for attempt in range(RETRIES):
        try:
            logger.info("Recherche Wikipédia pour %s (tentative %s)", commune_label, attempt + 1)
            ensure_fr_wikipedia_home(drv)
            srch = wait_css(drv, "#searchInput", TIMEOUT)
            srch.clear()
            time.sleep(TYPE_DELAY)
            srch.send_keys(commune_label + Keys.ENTER)

            resolve_search_results_if_needed(drv, commune_label)

            wait_any(drv, ["h1#firstHeading", "h1 .mw-page-title-main"], TIMEOUT)
            url = drv.current_url

            soup = BeautifulSoup(drv.page_source, "lxml")

            climat = extract_paragraph_by_prefixes(
                soup,
                "Climat",
                CLIMAT_PREFIXES,
                CLIMAT_FALLBACK_CONTAINS,
            )
            if not climat:
                logger.warning("Section Climat non trouvée pour %s", commune_label)

            corine = extract_paragraph_by_prefixes(
                soup,
                ("Urbanisme", "Occupation des sols"),
                OCCUP_PREFIXES,
                OCCUP_FALLBACK_CONTAINS,
            )
            if not corine:
                logger.warning("Section Occupation des sols non trouvée pour %s", commune_label)

            climat = clean_text(climat) or "Donnée non disponible"
            corine = clean_text(corine) or "Donnée non disponible"

            return {
                "climat": climat,
                "corine": corine,
                "url": url,
                "commune": commune_label,
            }
        except (
            TimeoutException,
            NoSuchElementException,
            StaleElementReferenceException,
        ) as e:
            logger.error("Erreur lors de la navigation Wikipédia : %s", e)

    # En cas d'échec total
    return {
        "climat": "Donnée non disponible",
        "corine": "Donnée non disponible",
        "url": "",
        "commune": commune_label,
    }


# ---------------------------------------------------------------------------
# Compatibilité ascendante
# ---------------------------------------------------------------------------
def fetch_wikipedia_info(commune_query: str):
    """Alias pour maintenir l'ancienne API.

    Renvoie ``(data, driver)`` afin de ne pas casser les imports existants.
    """

    data = run_wikipedia_scrape(commune_query)
    return data, get_shared_driver()


__all__ = [
    "DEP",
    "run_wikipedia_scrape",
    "fetch_wikipedia_info",
]

