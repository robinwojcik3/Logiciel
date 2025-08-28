# -*- coding: utf-8 -*-
"""Utilitaire pour extraire quelques sections des pages Wikipédia de communes françaises.

Ce module reprend le script fourni et l'adapte sous forme de fonction
facilement réutilisable par l'application.
"""

from __future__ import annotations

import re
from typing import Dict, Tuple
import time
import urllib.parse
import requests

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
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
        # Recherche prioritaire du paragraphe commençant par
        # "Pour la période 1971-2000, la température annuelle ..."
        target = None
        for p in h.find_all_next("p"):
            t = p.get_text(strip=True)
            if t.startswith("Pour la période 1971-2000"):
                target = t
                break
        if target:
            out["climat_p1"] = target
        else:
            # Fallback: même logique qu'avant (paragraphe(s) suivant(s) l'explication du type de climat)
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


def _scrape_sections_from_html(html: str) -> Dict[str, str]:
    """Version hors-Selenium: extrait les mêmes sections depuis du HTML.

    Utilisé en repli si la navigation Selenium échoue ou ne trouve pas
    les paragraphes souhaités.
    """
    out = {
        "climat_p1": "Non trouvé",
        "climat_p2": "Non trouvé",
        "occupation_p1": "Non trouvé",
    }
    soup = BeautifulSoup(html, "html.parser")

    h = _find_section_heading(soup, "Climat")
    if h:
        target = None
        for p in h.find_all_next("p"):
            t = p.get_text(strip=True)
            if t.startswith("Pour la période 1971-2000"):
                target = t
                break
        if target:
            out["climat_p1"] = target
        else:
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
    """Ouvre la page de recherche avancée puis l'article correspondant."""

    search_url = (
        "https://fr.wikipedia.org/w/index.php?search=&title=Sp%C3%A9cial%3ARecherche"
        "&profile=advanced&fulltext=1&ns0=1"
    )
    driver.get(search_url)

    # Gestion de la bannière cookies éventuelle
    try:
        btn = WebDriverWait(driver, 0.5).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//button[contains(.,'Accepter') or contains(.,'Tout accepter') or contains(.,\"J'ai compris\") or contains(.,'J’ai compris')]",
                )
            )
        )
        btn.click()
    except TimeoutException:
        pass

    box = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "input.oo-ui-inputWidget-input[name='search']")
        )
    )
    box.clear()
    # Réduction du délai d'attente avant saisie: 0.5 s
    time.sleep(0.5)
    box.send_keys(query)

    # Cliquer sur le bouton "Rechercher" (fall back sur Entrée au besoin)
    try:
        btn_search = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//span[@class='oo-ui-labelElement-label' and text()='Rechercher']/ancestor::button",
                )
            )
        )
        btn_search.click()
    except TimeoutException:
        box.send_keys(Keys.ENTER)

    # Ouvrir le premier résultat de la recherche
    try:
        link = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "div.mw-search-result-heading a")
            )
        )
        link.click()
        wait.until(EC.presence_of_element_located((By.ID, "firstHeading")))
        return True
    except TimeoutException:
        return False


def _http_fetch_article_and_parse(query: str) -> Tuple[Dict[str, str], str]:
    """Recherche via l'API MediaWiki et extrait les sections depuis l'HTML.

    Retourne (data, url) où data contient les clés climat_p1, climat_p2,
    occupation_p1. Lève aucune exception: en cas d'échec, renvoie des valeurs
    "Non trouvé" et une URL vide.
    """
    try:
        api = "https://fr.wikipedia.org/w/api.php"
        # 1) Recherche du titre
        r = requests.get(
            api,
            params={
                "action": "query",
                "list": "search",
                "srsearch": query,
                "utf8": 1,
                "format": "json",
                "srlimit": 1,
                "srprop": "",
            },
            timeout=10,
            headers={"User-Agent": "ContexteEco/1.0 (scraper)"},
        )
        r.raise_for_status()
        js = r.json()
        hits = js.get("query", {}).get("search", [])
        if not hits:
            return {
                "climat_p1": "Non trouvé",
                "climat_p2": "Non trouvé",
                "occupation_p1": "Non trouvé",
            }, ""
        title = hits[0]["title"]
        # 2) Récupération du HTML de la page via action=parse (formatversion=2)
        r2 = requests.get(
            api,
            params={
                "action": "parse",
                "page": title,
                "prop": "text",
                "format": "json",
                "formatversion": 2,
                "utf8": 1,
            },
            timeout=10,
            headers={"User-Agent": "ContexteEco/1.0 (scraper)"},
        )
        r2.raise_for_status()
        html = r2.json().get("parse", {}).get("text", "")
        if not html:
            return {
                "climat_p1": "Non trouvé",
                "climat_p2": "Non trouvé",
                "occupation_p1": "Non trouvé",
            }, ""
        data = _scrape_sections_from_html(html)
        url = f"https://fr.wikipedia.org/wiki/{urllib.parse.quote(title.replace(' ', '_'))}"
        return data, url
    except Exception:
        return {
            "climat_p1": "Non trouvé",
            "climat_p2": "Non trouvé",
            "occupation_p1": "Non trouvé",
        }, ""


def fetch_wikipedia_info(commune_query: str) -> Tuple[Dict[str, str], webdriver.Chrome]:
    """Ouvre la page Wikipédia correspondant à ``commune_query`` et en extrait
    quelques sections utiles. La fonction renvoie également l'objet ``driver``
    afin que l'utilisateur décide quand fermer la fenêtre du navigateur.

    ``commune_query`` peut être de la forme ``"Vizille 38"`` ou
    ``"Vizille (38)``.
    """

    query = _normalize_query(commune_query)
    import os
    from pathlib import Path
    # Repo root (modules/..)
    REPO_ROOT = Path(__file__).resolve().parent.parent
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--log-level=3")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    # Par défaut on affiche le navigateur (APP_HEADLESS=0)
    # Mettre APP_HEADLESS=1 pour exécuter en mode headless si souhaité.
    if os.environ.get("APP_HEADLESS", "0").lower() in ("1", "true", "yes"):  # opt-in via APP_HEADLESS=1
        try:
            options.add_argument("--headless=new")
        except Exception:
            options.add_argument("--headless")
    # Driver local (repo) prioritaire si présent
    # Ensure browser is visible if APP_HEADLESS=0 (ou non défini)
    import os as _os
    try:
        if _os.environ.get("APP_HEADLESS", "0").lower() in ("0", "false", "no", ""):
            if hasattr(options, "arguments"):
                options.arguments = [a for a in options.arguments if not str(a).startswith("--headless")]
    except Exception:
        pass
    local_driver = REPO_ROOT / "tools" / "chromedriver.exe"
    if local_driver.is_file():
        driver = webdriver.Chrome(service=Service(str(local_driver)), options=options)
    else:
        driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    wait = WebDriverWait(driver, 10)

    ok = _open_article(driver, query, wait)
    if not ok:
        alt = f"{query} (commune)"
        ok = _open_article(driver, alt, wait)

    data: Dict[str, str]
    if ok:
        data = _scrape_sections(driver)
        data["url"] = driver.current_url
        # Si le scraping via Selenium ne trouve rien, on tente le repli HTTP
        if all(data.get(k, "").startswith("Non trouv") for k in ("climat_p1", "occupation_p1")):
            http_data, http_url = _http_fetch_article_and_parse(query)
            # Remplacer uniquement les champs manquants
            for k in ("climat_p1", "climat_p2", "occupation_p1"):
                if data.get(k, "").startswith("Non trouv") and not http_data.get(k, "").startswith("Non trouv"):
                    data[k] = http_data[k]
            if http_url:
                data["url"] = http_url
        return data, driver

    # Selenium a échoué: repli 100% HTTP
    data, http_url = _http_fetch_article_and_parse(query)
    if http_url:
        data["url"] = http_url
    else:
        data["error"] = "Article introuvable"
    return data, driver

