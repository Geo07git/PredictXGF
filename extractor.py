import time
import random
import pandas as pd
from bs4 import BeautifulSoup


def extract_table(driver, url: str, log):
    """
    Întoarce un DataFrame cu rândurile principale din tabelul xG
    sau None dacă nu găsește un tabel valid.
    """

    try:
        log(f"🌐 GET {url}")
        driver.get(url)
    except Exception as e:
        log(f"GET ERROR: {type(e).__name__}: {e}")
        return None

    # scurt delay ca să se încarce pagina
    time.sleep(random.uniform(3, 5))

    # detectare simplă Cloudflare / captcha
    if "captcha" in driver.current_url.lower() or "cloudflare" in driver.page_source.lower():
        log("🚫 Cloudflare / CAPTCHA detectat")
        return None

    # mic scroll ca să forțăm lazy-load
    try:
        driver.execute_script("window.scrollTo(0, 800);")
    except Exception as e:
        log(f"Scroll error: {e}")
    time.sleep(2)

    soup = BeautifulSoup(driver.page_source, "html.parser")

    # tabelul standard xG
    table = soup.select_one("table.full-league-table.table-sort.xg-all")
    
    headers = [th.get_text(" ", strip=True) for th in table.find_all("th")]
    headers = [h.replace("Matches Played", "").strip() for h in headers]
    #log(f"HEADERS: {headers}")

    if not table:
        log("❌ Tabel xG nu a fost găsit")
        return None

    tbody = table.find("tbody")
    if not tbody:
        log("❌ Tabel fără <tbody>")
        return None

    rows = []

    for tr in tbody.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)

        # sărim peste rânduri goale / expandate
        if len(tds) < 5:
            continue

        first = tds[0].get_text(strip=True)
        if not first.isdigit():
            continue

        #original
        #row = [td.get_text(" ", strip=True) for td in tds]
        #rows.append(row)

        row = []

        for td, header in zip(tds, headers):

            text = td.get_text(" ", strip=True)

            # 👉 curățăm doar TEAM
            if "Team" in header:
                a = td.find("a")

                if a:
                    # doar text direct, fără div/modal
                    team = ""

                    for child in a.children:
                        if isinstance(child, str):
                            team = child.strip()
                            if team:
                                break

                    text = team

            row.append(text)

        rows.append(row)

    if len(rows) < 5:
        log("⚠️ Prea puține rânduri valide în tabel")
        return None

    df = pd.DataFrame(rows)
    log(f"✅ Tabel extras: {len(df)} rânduri x {len(df.columns)} coloane")
    return df