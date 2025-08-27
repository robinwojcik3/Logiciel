import re
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Départements courants (ajoutez si besoin)
DEP = {
    "01": "Ain", "03": "Allier", "04": "Alpes-de-Haute-Provence", "05": "Hautes-Alpes", "06": "Alpes-Maritimes",
    "07": "Ardèche", "09": "Ariège", "11": "Aude", "13": "Bouches-du-Rhône", "15": "Cantal", "21": "Côte-d'Or",
    "26": "Drôme", "30": "Gard", "31": "Haute-Garonne", "34": "Hérault", "38": "Isère", "39": "Jura",
    "42": "Loire", "43": "Haute-Loire", "63": "Puy-de-Dôme", "69": "Rhône", "73": "Savoie", "74": "Haute-Savoie",
    "75": "Paris", "83": "Var", "84": "Vaucluse", "90": "Territoire de Belfort"
}


def find_section_heading(soup, heading_text):
    span = soup.find('span', class_='mw-headline',
                     string=lambda t: t and heading_text.lower() in t.lower())
    return span.find_parent(['h2', 'h3']) if span else None


def scrape_wikipedia_sections(driver):
    out = {"climat_p1": "Non trouvé", "climat_p2": "Non trouvé", "occupation_p1": "Non trouvé"}
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    h = find_section_heading(soup, "Climat")
    if h:
        start = None
        for p in h.find_all_next('p'):
            t = p.get_text(strip=True)
            if t.startswith("En 2010, le climat de la commune est de type") or "climat de la commune est de type" in t:
                start = p
                break
        if start:
            fol = start.find_next_siblings('p', limit=2)
            if len(fol) >= 1:
                out["climat_p1"] = fol[0].get_text(strip=True)
            if len(fol) >= 2:
                out["climat_p2"] = fol[1].get_text(strip=True)

    h = find_section_heading(soup, "Occupation des sols")
    if h:
        for p in h.find_all_next('p'):
            t = p.get_text(strip=True)
            if t.startswith("L'occupation des sols de la commune, telle qu'elle") or "L'occupation des sols de la commune, telle qu'elle ressort" in t:
                out["occupation_p1"] = t
                break
    return out


def normalize_query(s: str) -> str:
    s = s.strip()
    m = re.match(r"^(.*?)[\s,;_-]*\(?([0-9]{2})\)?$", s)
    if m and m.group(2) in DEP:
        base = m.group(1).strip()
        return f"{base} ({DEP[m.group(2)]})"
    return s


def open_wikipedia_article(driver, query: str, wait: WebDriverWait) -> bool:
    driver.get("https://fr.wikipedia.org/")
    try:
        btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(.,'Accepter') or contains(.,'Tout accepter') or contains(.,\"J'ai compris\") or contains(.,'J’ai compris')]")
        ))
        btn.click()
    except TimeoutException:
        pass
    box = wait.until(EC.element_to_be_clickable((By.ID, "searchInput")))
    box.clear()
    box.send_keys(query)
    box.send_keys(Keys.ENTER)
    try:
        wait.until(EC.presence_of_element_located((By.ID, "firstHeading")))
        if "Spécial:Recherche" in driver.current_url or "Special:Search" in driver.current_url:
            link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.mw-search-result-heading a")))
            link.click()
            wait.until(EC.presence_of_element_located((By.ID, "firstHeading")))
        return True
    except TimeoutException:
        return False


def fetch_commune_info(commune: str):
    query = normalize_query(commune)
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--log-level=3")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--headless=new")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 10)
    try:
        ok = open_wikipedia_article(driver, query, wait)
        if not ok:
            alt = f"{query} (commune)"
            ok = open_wikipedia_article(driver, alt, wait)
        if not ok:
            raise RuntimeError("Article Wikipédia introuvable")
        data = scrape_wikipedia_sections(driver)
        url = driver.current_url
        return url, data
    finally:
        try:
            driver.quit()
        except Exception:
            pass
