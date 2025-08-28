# -*- coding: utf-8 -*-
"""Scraping Wikipédia pour les communes françaises.

Ce module ouvre une page Wikipédia dans une instance Chrome partagée,
extrait les paragraphes relatifs au climat et à l'occupation des sols
("Corine Land Cover") puis renvoie les données normalisées.
"""
from __future__ import annotations

import logging
import re
import time
import unicodedata
from typing import Optional

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

# Paramètres généraux
TIMEOUT = 15
RETRIES = 2
TYPE_DELAY = 0.5

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# --------- Gestion driver partagé ---------
_DRIVER: Optional[webdriver.Chrome] = None

def _get_driver() -> webdriver.Chrome:
    global _DRIVER
    if _DRIVER is None:
        opts = webdriver.ChromeOptions()
        opts.add_experimental_option("excludeSwitches", ["enable-logging"])
        _DRIVER = webdriver.Chrome(options=opts)
        try:
            _DRIVER.maximize_window()
        except Exception:
            pass
    return _DRIVER

# --------- Utilitaires Selenium ---------

def _wait_css(driver: webdriver.Chrome, selector: str, timeout: int = TIMEOUT):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
    )

def _ensure_fr_home(driver: webdriver.Chrome) -> None:
    if not driver.current_url.startswith("https://fr.wikipedia.org"):
        logger.info("Ouverture de la page d'accueil Wikipédia FR")
        driver.get("https://fr.wikipedia.org")
        try:
            _wait_css(driver, "#searchInput")
        except TimeoutException:
            logger.warning("Champ de recherche introuvable sur la page d'accueil")

# --------- Résolution des résultats/homonymies ---------

def resolve_search_results_if_needed(driver: webdriver.Chrome, commune_label: str) -> None:
    """Clique un résultat pertinent si la recherche affiche une page de résultats."""
    try:
        results = driver.find_elements(By.CSS_SELECTOR, "div.mw-search-result-heading a")
        if not results:
            return
        logger.info("Page de résultats détectée, tentative de sélection du bon lien")
        target = None
        for link in results:
            title = link.get_attribute("title") or ""
            if all(part.lower() in title.lower() for part in commune_label.split()):
                target = link
                break
        if target is None:
            target = results[0]
        target.click()
        WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h1#firstHeading, h1 .mw-page-title-main"))
        )
    except Exception as e:
        logger.warning("Échec sélection résultat : %s", e)

# --------- Extraction DOM ---------

CLIMAT_PREFIXES = [
    r"^Pour la période 1971-2000, la température annuelle",
    r"^Pour la période 1981-2010, la température annuelle",
]
CLIMAT_FALLBACK_CONTAINS = [r"température annuelle", r"précipitations"]

OCCUP_PREFIXES = [r"^L'occupation des sols de la commune, telle qu'elle"]
OCCUP_FALLBACK_CONTAINS = [r"Corine Land Cover", r"occupation des sols"]

def _find_section(soup: BeautifulSoup, section_hint):
    if isinstance(section_hint, tuple):
        h2 = _find_section(soup, section_hint[0])
        if not h2:
            return None
        for tag in h2.find_all_next(["h2", "h3"]):
            if tag.name == "h2":
                break
            if tag.name == "h3" and section_hint[1].lower() in tag.get_text(" ").lower():
                return tag
        return None
    for h in soup.select("h2 span.mw-headline"):
        if section_hint.lower() in h.get_text(" ").lower():
            return h.parent
    return None

def extract_paragraph_by_prefixes(soup: BeautifulSoup, section_hint, prefixes, contains_all):
    sec = _find_section(soup, section_hint)
    if sec:
        for p in sec.find_all_next("p"):
            if p.find_previous(sec.name) != sec:
                break
            txt = p.get_text(" ", strip=True)
            for pref in prefixes:
                if re.search(pref, txt, re.IGNORECASE):
                    return txt
            if all(t in txt.lower() for t in [c.lower() for c in contains_all]):
                return txt
    return ""

def clean_text(txt: str) -> str:
    if not txt:
        return ""
    txt = re.sub(r"\[\d+\]", "", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return unicodedata.normalize("NFC", txt)

# --------- Fonction principale ---------

def run_wikipedia_scrape(commune_label: str) -> dict:
    """Scrape Wikipédia pour ``commune_label`` et renvoie les paragraphes utiles."""
    data = {"climat": "Donnée non disponible", "corine": "Donnée non disponible", "url": "", "commune": commune_label}
    driver = _get_driver()
    for attempt in range(1, RETRIES + 1):
        try:
            _ensure_fr_home(driver)
            search = _wait_css(driver, "#searchInput")
            search.clear()
            time.sleep(TYPE_DELAY)
            search.send_keys(commune_label + Keys.ENTER)
            resolve_search_results_if_needed(driver, commune_label)
            _wait_css(driver, "h1#firstHeading, h1 .mw-page-title-main")
            data["url"] = driver.current_url
            soup = BeautifulSoup(driver.page_source, "lxml")
            climat = extract_paragraph_by_prefixes(
                soup, "Climat", CLIMAT_PREFIXES, CLIMAT_FALLBACK_CONTAINS
            )
            corine = extract_paragraph_by_prefixes(
                soup, ("Urbanisme", "Occupation des sols"), OCCUP_PREFIXES, OCCUP_FALLBACK_CONTAINS
            )
            climat = clean_text(climat)
            corine = clean_text(corine)
            data["climat"] = climat or "Donnée non disponible"
            data["corine"] = corine or "Donnée non disponible"
            logger.info("Scraping terminé : %s", commune_label)
            return data
        except (TimeoutException, NoSuchElementException, StaleElementReferenceException) as e:
            logger.error("Tentative %s échouée: %s", attempt, e)
    return data
