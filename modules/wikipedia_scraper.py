# -*- coding: utf-8 -*-
"""Utilitaire pour extraire quelques sections des pages Wikipédia de communes françaises.

Ce module reprend le script fourni et l'adapte sous forme de fonction
facilement réutilisable par l'application.
"""

from __future__ import annotations

import logging
import re
import time
import unicodedata
from typing import Dict, Tuple, List, Optional

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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

# --- Nouveaux paramètres pour le scraping Wikipédia ---
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


# ---------------------------------------------------------------------------
# Nouvelle implémentation adaptée aux besoins de l'application
# ---------------------------------------------------------------------------

_SHARED_DRIVER: Optional[webdriver.Chrome] = None


def get_shared_driver() -> webdriver.Chrome:
    """Retourne une instance unique de ``webdriver.Chrome``."""
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
        _SHARED_DRIVER.maximize_window()
    return _SHARED_DRIVER


def wait_css(driver: webdriver.Chrome, selector: str, timeout: int = TIMEOUT):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
    )


def wait_any(driver: webdriver.Chrome, selectors: List[str], timeout: int = TIMEOUT):
    conds = [EC.presence_of_element_located((By.CSS_SELECTOR, s)) for s in selectors]
    WebDriverWait(driver, timeout).until(EC.any_of(*conds))


def ensure_fr_wikipedia_home(driver: webdriver.Chrome) -> None:
    if not driver.current_url.startswith("https://fr.wikipedia.org"):
        driver.get("https://fr.wikipedia.org")
    try:
        wait_css(driver, "#searchInput", TIMEOUT)
    except TimeoutException:
        driver.get("https://fr.wikipedia.org")
        wait_css(driver, "#searchInput", TIMEOUT)


def resolve_search_results_if_needed(driver: webdriver.Chrome, commune_label: str) -> None:
    """Résout les pages de résultats ou d'homonymie."""
    try:
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.mw-search-result-heading a"))
        )
    except TimeoutException:
        return  # déjà sur l'article

    links = driver.find_elements(By.CSS_SELECTOR, "div.mw-search-result-heading a")
    target = None
    lower_label = commune_label.lower()
    for link in links:
        title = (link.get_attribute("title") or "").lower()
        if lower_label in title:
            target = link
            break
    if target is None and links:
        target = links[0]

    if target is None:
        return

    href = target.get_attribute("href")
    driver.get(href)

    # Vérifier l'infobox "Commune" si nécessaire
    if lower_label not in (target.get_attribute("title") or "").lower():
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.infobox"))
            )
            soup = BeautifulSoup(driver.page_source, "lxml")
            infobox = soup.find("table", class_=re.compile("infobox"))
            if not infobox or "Commune" not in infobox.get_text():
                driver.back()
                if links:
                    driver.get(links[0].get_attribute("href"))
        except TimeoutException:
            pass


def extract_paragraph_by_prefixes(
    soup: BeautifulSoup,
    section_hint,
    prefixes: List[str],
    contains_all: List[str],
) -> str:
    heading = None
    if isinstance(section_hint, tuple):
        h2 = _find_section_heading(soup, section_hint[0])
        if h2:
            for sib in h2.find_all_next("h3"):
                if sib.find("span", class_="mw-headline") and section_hint[1].lower() in sib.get_text(strip=True).lower():
                    heading = sib
                    break
    else:
        heading = _find_section_heading(soup, section_hint)
    if not heading:
        return ""
    candidate = ""
    for el in heading.find_all_next():
        if el.name and el.name.startswith("h") and el.name <= heading.name:
            break
        if el.name != "p":
            continue
        text = el.get_text(" ", strip=True)
        for pref in prefixes:
            if re.match(pref, text, flags=re.IGNORECASE):
                return text
        if all(re.search(tok, text, flags=re.IGNORECASE) for tok in contains_all) and not candidate:
            candidate = text
    return candidate


def clean_text(txt: str) -> str:
    if not txt:
        return ""
    txt = unicodedata.normalize("NFC", txt)
    txt = re.sub(r"\[\d+\]", "", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def run_wikipedia_scrape(commune_label: str) -> Dict[str, str]:
    """Scrape Wikipédia pour ``commune_label``."""
    data = {"climat": "", "corine": "", "url": "", "commune": commune_label}
    try:
        driver = get_shared_driver()
        ensure_fr_wikipedia_home(driver)
        box = wait_css(driver, "#searchInput", TIMEOUT)
        box.clear()
        time.sleep(TYPE_DELAY)
        box.send_keys(commune_label + Keys.ENTER)
        resolve_search_results_if_needed(driver, commune_label)
        wait_any(driver, ["h1#firstHeading", "h1 .mw-page-title-main"], TIMEOUT)
        data["url"] = driver.current_url
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
        climat = clean_text(climat)
        corine = clean_text(corine)
        if not climat:
            climat = "Donnée non disponible"
        if not corine:
            corine = "Donnée non disponible"
        data["climat"] = climat
        data["corine"] = corine
        logger.info("Scraping Wikipédia réussi pour %s", commune_label)
    except (TimeoutException, NoSuchElementException, StaleElementReferenceException) as e:
        logger.error("Erreur lors du scraping Wikipédia : %s", e)
        if not data["climat"]:
            data["climat"] = "Donnée non disponible"
        if not data["corine"]:
            data["corine"] = "Donnée non disponible"
    return data


