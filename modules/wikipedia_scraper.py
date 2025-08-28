# -*- coding: utf-8 -*-
"""Outils de scraping Wikipédia pour l'application.

Ce module expose une fonction principale : :func:`run_wikipedia_scrape` qui
récupère deux paragraphes normalisés (climat et occupation des sols) depuis
la page Wikipédia d'une commune française. Le même ``webdriver`` est partagé
entre les appels afin de laisser la fenêtre ouverte après usage.
"""

from __future__ import annotations

import logging
import re
import time
import unicodedata
from typing import Dict, Iterable, Tuple, Union

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)

# ---------------------- Constantes & paramètres ----------------------
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

_driver: webdriver.Chrome | None = None

# ---------------------- Helpers Selenium ----------------------

def get_shared_driver() -> webdriver.Chrome:
    """Retourne une unique instance de ``webdriver.Chrome``."""
    global _driver
    if _driver is None:
        options = webdriver.ChromeOptions()
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        _driver = webdriver.Chrome(options=options)
        _driver.maximize_window()
    return _driver


def ensure_fr_wikipedia_home(driver: webdriver.Chrome):
    """Ouvre la page d'accueil fr.wikipedia.org et renvoie l'input de recherche."""
    if "fr.wikipedia.org" not in driver.current_url:
        driver.get("https://fr.wikipedia.org")
    try:
        return WebDriverWait(driver, TIMEOUT).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#searchInput"))
        )
    except TimeoutException:
        driver.get("https://fr.wikipedia.org")
        return WebDriverWait(driver, TIMEOUT).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#searchInput"))
        )


def resolve_search_results_if_needed(driver: webdriver.Chrome, commune_label: str) -> None:
    """Gère la page de résultats ou d'homonymie si nécessaire."""
    wait = WebDriverWait(driver, 10)
    try:
        results = driver.find_elements(By.CSS_SELECTOR, "div.mw-search-result-heading a")
        if not results:
            return
        commune_lower = commune_label.lower()
        target = None
        for a in results:
            title = (a.text or "").lower()
            if commune_lower in title:
                target = a
                break
        if target:
            target.click()
            return
        # Fallback : ouvrir chaque résultat et vérifier la présence d'une infobox de commune
        for a in results:
            href = a.get_attribute("href")
            if not href:
                continue
            driver.get(href)
            try:
                infobox = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table.infobox"))
                )
                if "Commune" in infobox.text:
                    return
            except TimeoutException:
                driver.back()
                continue
        # Dernier recours : premier résultat
        results[0].click()
    except Exception as e:  # pragma: no cover - protection large
        logger.warning("Résolution page de résultats échouée: %s", e)


# ---------------------- Extraction HTML ----------------------

def extract_paragraph_by_prefixes(
    soup: BeautifulSoup,
    section_hint: Union[str, Tuple[str, str]],
    prefixes: Iterable[str],
    contains_all: Iterable[str],
) -> str:
    """Retourne un paragraphe selon différents critères."""
    node = None
    if isinstance(section_hint, str):
        span = soup.find(
            "span", class_="mw-headline", string=re.compile(section_hint, re.I)
        )
        node = span.find_parent("h2") if span else None
    else:
        span = soup.find(
            "span", class_="mw-headline", string=re.compile(section_hint[0], re.I)
        )
        if span:
            h2 = span.find_parent("h2")
            for sib in h2.find_all_next(["h2", "h3"]):
                if sib.name == "h2":
                    break
                if sib.name == "h3" and re.search(section_hint[1], sib.get_text(), re.I):
                    node = sib
                    break
    paragraphs: list[str] = []
    if node:
        for sib in node.find_all_next():
            if sib.name in ["h2", "h3"] and sib.name == node.name:
                break
            if sib.name == "p":
                paragraphs.append(sib.get_text(" ", strip=True))
    else:
        paragraphs = [p.get_text(" ", strip=True) for p in soup.find_all("p")]
    for p in paragraphs:
        for pref in prefixes:
            if re.match(pref, p, flags=re.I):
                return p
    for p in paragraphs:
        if all(re.search(tok, p, flags=re.I) for tok in contains_all):
            logger.warning("Paragraphe fallback utilisé pour %s", section_hint)
            return p
    return ""


def clean_text(txt: str) -> str:
    if not txt:
        return ""
    txt = unicodedata.normalize("NFC", txt)
    txt = re.sub(r"\[\d+\]", "", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


# ---------------------- API principale ----------------------

def run_wikipedia_scrape(commune_label: str) -> Dict[str, str]:
    """Scrape la page Wikipédia de ``commune_label``.

    Retourne un dictionnaire : ``{'climat': str, 'corine': str, 'url': str, 'commune': str}``.
    En cas d'échec, les valeurs textuelles vaudront « Donnée non disponible ».
    """
    driver = get_shared_driver()
    for attempt in range(1, RETRIES + 1):
        try:
            logger.info("Scraping %s (tentative %s)", commune_label, attempt)
            box = ensure_fr_wikipedia_home(driver)
            box.clear()
            time.sleep(TYPE_DELAY)
            box.send_keys(commune_label)
            box.send_keys(Keys.ENTER)

            resolve_search_results_if_needed(driver, commune_label)

            WebDriverWait(driver, TIMEOUT).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "h1#firstHeading, h1 .mw-page-title-main")
                )
            )
            url = driver.current_url
            soup = BeautifulSoup(driver.page_source, "lxml")
            climat = extract_paragraph_by_prefixes(
                soup,
                "Climat",
                CLIMAT_PREFIXES,
                CLIMAT_FALLBACK_CONTAINS,
            )
            corine = extract_paragraph_by_prefixes(
                soup,
                ("Urbanisme", "Occupation des sols"),
                OCCUP_PREFIXES,
                OCCUP_FALLBACK_CONTAINS,
            )
            climat = clean_text(climat) or "Donnée non disponible"
            corine = clean_text(corine) or "Donnée non disponible"
            return {"climat": climat, "corine": corine, "url": url, "commune": commune_label}
        except (
            TimeoutException,
            NoSuchElementException,
            StaleElementReferenceException,
        ) as e:
            logger.error("Erreur scraping: %s", e)
    logger.error("Échec du scraping pour %s", commune_label)
    return {
        "climat": "Donnée non disponible",
        "corine": "Donnée non disponible",
        "url": "",
        "commune": commune_label,
    }

