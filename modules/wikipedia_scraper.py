# -*- coding: utf-8 -*-
"""Utilitaire pour extraire quelques sections des pages Wikipédia de communes françaises.

Ce module reprend le script fourni et l'adapte sous forme de fonction
facilement réutilisable par l'application.
"""

from __future__ import annotations

import re
from typing import Dict, Tuple

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

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


def _find_section_heading(soup: BeautifulSoup, heading_text: str):
    span = soup.find(
        "span",
        class_="mw-headline",
        string=lambda t: t and heading_text.lower() in t.lower(),
    )
    return span.find_parent(["h2", "h3"]) if span else None


def _scrape_sections(driver: webdriver.Chrome) -> Dict[str, str]:
    out = {
        "climat_p1": "Non trouvé",
        "climat_p2": "Non trouvé",
        "occupation_p1": "Non trouvé",
    }
    soup = BeautifulSoup(driver.page_source, "html.parser")

    h = _find_section_heading(soup, "Climat")
    if h:
        start = None
        for p in h.find_all_next("p"):
            t = p.get_text(strip=True)
            if t.startswith("En 2010, le climat de la commune est de type") or "climat de la commune est de type" in t:
                start = p
                break
        if start:
            fol = start.find_next_siblings("p", limit=2)
            if len(fol) >= 1:
                out["climat_p1"] = fol[0].get_text(strip=True)
            if len(fol) >= 2:
                out["climat_p2"] = fol[1].get_text(strip=True)

    h = _find_section_heading(soup, "Occupation des sols")
    if h:
        for p in h.find_all_next("p"):
            t = p.get_text(strip=True)
            if t.startswith("L'occupation des sols de la commune, telle qu'elle") or "L'occupation des sols de la commune, telle qu'elle ressort" in t:
                out["occupation_p1"] = t
                break
    return out


def _normalize_query(s: str) -> str:
    s = s.strip()
    m = re.match(r"^(.*?)[\s,;_-]*\(?(\d{2})\)?$", s)
    if m:
        base = m.group(1).strip()
        return f"{base} ({m.group(2)})"
    return s


def _open_article(driver: webdriver.Chrome, query: str, wait: WebDriverWait) -> bool:
    driver.get("https://fr.wikipedia.org/")
    try:
        driver.find_element(
            By.XPATH,
            "//button[contains(.,'Accepter') or contains(.,'Tout accepter') or contains(.,\"J'ai compris\") or contains(.,'J’ai compris')]",
        ).click()
    except Exception:
        pass

    box = WebDriverWait(driver, 0.5).until(
        EC.element_to_be_clickable((By.ID, "searchInput"))
    )
    box.clear()
    box.send_keys(query)
    try:
        WebDriverWait(driver, 2).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".suggestions-results"))
        )
        ActionChains(driver).move_to_element_with_offset(
            box, 0, box.size["height"] + 10
        ).click().perform()
    except TimeoutException:
        box.send_keys(Keys.ENTER)

    try:
        wait.until(EC.presence_of_element_located((By.ID, "firstHeading")))
        if "Spécial:Recherche" in driver.current_url or "Special:Search" in driver.current_url:
            link = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div.mw-search-result-heading a"))
            )
            link.click()
            wait.until(EC.presence_of_element_located((By.ID, "firstHeading")))
        return True
    except TimeoutException:
        return False


def fetch_wikipedia_info(commune_query: str) -> Tuple[Dict[str, str], webdriver.Chrome]:
    """Ouvre la page Wikipédia correspondant à ``commune_query`` et en extrait
    quelques sections utiles. La fonction renvoie également l'objet ``driver``
    afin que l'utilisateur décide quand fermer la fenêtre du navigateur.

    ``commune_query`` peut être de la forme ``"Vizille 38"`` ou
    ``"Vizille (38)``.
    """

    query = _normalize_query(commune_query)
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--log-level=3")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    wait = WebDriverWait(driver, 10)

    ok = _open_article(driver, query, wait)
    if not ok:
        alt = f"{query} (commune)"
        ok = _open_article(driver, alt, wait)
    if not ok:
        return {"error": "Article introuvable"}, driver

    data = _scrape_sections(driver)
    data["url"] = driver.current_url
    return data, driver

