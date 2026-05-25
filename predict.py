import io
import itertools
from datetime import datetime
import re

import numpy as np
import pandas as pd
import streamlit as st
from scipy.stats import poisson

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

import os

# ============== CONFIGURARE DE BAZĂ ==============

st.set_page_config(page_title="xG Poisson Analyzer", layout="wide", page_icon="⚽")
st.title("⚽ xG Poisson Analyzer")
st.markdown(
    "Analiză meci individual cu distribuție Poisson bazată pe date xG din repo. "
    "Selectezi liga și meciul, calculezi probabilități, cote fair și EV."
)

# ============== HOW TO USE — popup ==============

if "app_started" not in st.session_state:
    st.session_state.app_started = False

if not st.session_state.app_started:
    with st.expander("ℹ️ Cum se folosește aplicația", expanded=True):

        st.markdown("""
### 🗂️ Structura aplicației
Aplicația are **3 tab-uri** accesibile din bara de sus:

| Tab | Rol |
|-----|-----|
| 🎯 **Single Match** | Analizezi un meci individual: alegi liga, echipele, piața și cota |
| 🎯 **Calibrare Model** | Validezi acuratețea modelului pe baza unui fișier cu predicții istorice |
| 📚 **Istoric** | Încarci, vizualizezi și actualizezi rezultatele predicțiilor anterioare |

---

### ⚙️ Sidebar — Setări globale
> Accesibil din stânga (click **>** dacă e ascuns)

- **Încarcă xGData.xlsx** *(opțional)* — dacă fișierul e deja în repo, se încarcă automat. Upload manual doar dacă vrei să îl suprascrii.
- **Prag EV BET** — valoarea minimă Expected Value pentru verdict **BET** *(default: 0.03)*
- **Prag Prob BET** — probabilitatea minimă pentru verdict **BET** *(default: 0.50)*
- **Prag EV LEAN** — EV minim pentru verdict **LEAN** *(default: 0.00)*
- **Max goals Poisson** — numărul maxim de goluri simulat în matricea Poisson *(default: 7)*

---

### 🎯 Tab: Single Match — Cum funcționează

**Pasul 1** — Alege **liga** din dropdown (ligile vin din foile Excel ale fișierului xGData.xlsx)

**Pasul 2** — Alege **Gazde** și **Oaspeți** (echipe din liga selectată)

**Pasul 3** — Alege **Piața** pe care vrei să o analizezi *(ex: Over 2.5 FT, 1X2 - 1, BTTS Yes)*

**Pasul 4** *(opțional)* — Introdu **cota bookmakerului** pentru a calcula EV și Verdict

**Pasul 5** — Click **🔍 Calculează**

**Rezultate afișate:**
- λ home / λ away — forța de atac estimată prin xG
- Probabilitate estimată de model
- Cotă fair (1/probabilitate)
- EV (Expected Value) și Verdict: **BET / LEAN / NO BET**
- Matricea scorurilor cu top 3 scoruri probabile
- Shortlist automat cu toate piețele BET/LEAN pentru cota introdusă

**Verdictele:**
| Verdict | Condiție |
|---------|----------|
| 🟢 **BET** | EV ≥ prag EV și Prob ≥ prag Prob |
| 🟡 **LEAN** | EV > prag LEAN (interes marginal) |
| ⚫ **NO BET** | Sub praguri |

---

### 🎯 Tab: Calibrare Model — Cum funcționează

**Pasul 1** — Încarcă un fișier Excel cu predicțiile tale istorice  
*(sheet obligatoriu: **Istoric**, coloane: `prob`, `rezultat`, `market`, `section`, `odds`, `profit`)*

**Pasul 2** — Fișierul trebuie să conțină rezultate reale: **W** (câștigat), **L** (pierdut), **V** (void/push)

**Rezultate afișate (4 sub-tab-uri):**
- 📊 **Calibrare generală** — prob medie vs win rate real, eroare calibrare, tabel pe intervale
- 🪣 **Bucket Analysis** — grupare pe quintile, grafic calibrare vizuală
- 📈 **Brier & Log Loss** — metrici de scoring probabilistic cu skill score față de baseline
- 🔬 **Per piață** — performanță per secțiune și piață (minim 3 rezultate)

---

### 📚 Tab: Istoric — Cum funcționează

**Încărcare:**
- Click **Încarcă Predictii_xG_Poisson.xlsx** → alege fișierul Excel cu sheet *Istoric*
- Datele se încarcă în sesiune *(nu se salvează pe server)*

**Introducere rezultate reale:**
1. Selectează meciul din dropdown-ul „Meciuri fără rezultat"
2. Introdu golurile (Gazde / Oaspeți)
3. Click **💾 Salvează rezultatele în sesiune**
4. Click **📥 Descarcă Excel actualizat** pentru a păstra modificările

**Statistici afișate automat** (după ce ai W/L/V):
- Total / Câștigate / Pierdute / Win Rate / ROI
- Performanță per piață (Win %, ROI %)
- Grafic evoluție profit cumulat

---

### 📁 Formatul fișierului xGData.xlsx

Fișierul trebuie să aibă **câte o foaie per ligă**, cu coloanele:

| Coloană | Descriere |
|---------|-----------|
| `Team` / `Squad` / `Club` | Numele echipei |
| `xG` | Expected Goals marcate per echipă |
| `xGA` | Expected Goals primite per echipă |

Celelalte coloane sunt ignorate. Ordinea nu contează.

---

### 📁 Formatul fișierului Predictii_xG_Poisson.xlsx

Fișierul trebuie să aibă sheet-ul **Istoric** cu coloanele standard:
`ts`, `league`, `home`, `away`, `market`, `section`, `line`, `prob`, `push_prob`, `fair_odds`, `odds`, `ev`, `decision`, `selected_for_play`, `proposal_rank`, `proposal_group`, `lambda_home`, `lambda_away`, `notes`, `rezultat`, `profit`

> ℹ️ Rezultatele `rezultat` trebuie să fie: **W**, **L**, **V**, sau **pending**

---
        """)
        if st.button("▶️ Pornește analiza", use_container_width=True, type="primary"):
            st.session_state.app_started = True
            st.rerun()

    st.stop()

HISTORY_PATH = "Predictii_xG_Poisson.xlsx"
MAX_GOALS = 7
DEFAULT_MIN_EV = 0.03
DEFAULT_MIN_PROB = 0.50
DEFAULT_MIN_EV_LEAN = 0.00

MATCH_TOTAL_LINES = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5]
TEAM_TOTAL_LINES = [0.5, 1.5, 2.5, 3.5]

# ============== FUNCȚII UTILE ==============

@st.cache_data(show_spinner=False)
def load_xg_workbook(file_bytes):
    xls = pd.ExcelFile(file_bytes)
    leagues = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = df.columns.str.strip()
        leagues[sheet] = df
    return leagues

def fmt2(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "-"
    return f"{x:.2f}"

def fmt3(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "-"
    return f"{x:.3f}"

def fmt1pct(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "-"
    return f"{x:.1f}%"

def infer_columns(df: pd.DataFrame):
    cols = list(df.columns)
    team_col = None
    for c in cols:
        name = str(c).strip().lower()
        if name in ("team", "squad", "club"):
            team_col = c
            break
    if team_col is None:
        team_col = cols[1] if len(cols) > 1 else cols[0]

    xgf_col = None
    for c in cols:
        if str(c).strip().lower() == "xg":
            xgf_col = c
            break

    xga_col = None
    for c in cols:
        if str(c).strip().lower() == "xga":
            xga_col = c
            break

    return team_col, xgf_col, xga_col

def build_team_strengths(df: pd.DataFrame):
    team_col, xgf_col, xga_col = infer_columns(df)
    if xgf_col is None or xga_col is None:
        raise ValueError(f"Lipsesc coloanele xG / xGA. Coloane detectate: {list(df.columns)}")

    df = df.copy().dropna(subset=[team_col])
    df[xgf_col] = pd.to_numeric(df[xgf_col], errors="coerce")
    df[xga_col] = pd.to_numeric(df[xga_col], errors="coerce")
    df = df.dropna(subset=[xgf_col, xga_col])

    mean_xg = df[xgf_col].mean() or 1.0
    mean_xga = df[xga_col].mean() or 1.0

    df["att"] = df[xgf_col] / mean_xg
    df["def"] = df[xga_col] / mean_xga

    return {
        "teams": df[team_col].astype(str).unique().tolist(),
        "team_col": team_col,
        "xgf_col": xgf_col,
        "xga_col": xga_col,
        "avg_xg": mean_xg,
        "avg_xga": mean_xga,
        "att": dict(zip(df[team_col].astype(str), df["att"])),
        "def": dict(zip(df[team_col].astype(str), df["def"])),
    }

def get_lambdas(strengths, home, away):
    avg_xg = strengths["avg_xg"]
    avg_xga = strengths["avg_xga"]
    att = strengths["att"]
    deff = strengths["def"]
    lam_home = att.get(home, 1.0) * deff.get(away, 1.0) * avg_xg
    lam_away = att.get(away, 1.0) * deff.get(home, 1.0) * avg_xga
    return lam_home, lam_away

def poisson_matrix(lam_home, lam_away, max_goals=7):
    goals = np.arange(0, max_goals + 1)
    return np.outer(poisson.pmf(goals, lam_home), poisson.pmf(goals, lam_away))

def prob_1x2(p_matrix):
    h = np.arange(p_matrix.shape[0])[:, None]
    a = np.arange(p_matrix.shape[1])[None, :]
    return (
        p_matrix[h > a].sum(),
        p_matrix[h == a].sum(),
        p_matrix[h < a].sum(),
    )

def prob_btts_yes(p_matrix):
    h = np.arange(p_matrix.shape[0])[:, None]
    a = np.arange(p_matrix.shape[1])[None, :]
    return p_matrix[(h > 0) & (a > 0)].sum()

def prob_total_over(p_matrix, line):
    total = np.add.outer(np.arange(p_matrix.shape[0]), np.arange(p_matrix.shape[1]))
    return p_matrix[total > line].sum()

def prob_total_under(p_matrix, line):
    total = np.add.outer(np.arange(p_matrix.shape[0]), np.arange(p_matrix.shape[1]))
    return p_matrix[total < line].sum()

def prob_total_push(p_matrix, line):
    total = np.add.outer(np.arange(p_matrix.shape[0]), np.arange(p_matrix.shape[1]))
    return p_matrix[total == line].sum() if float(line).is_integer() else 0.0

def prob_team_over(p_matrix, side, line):
    h = np.arange(p_matrix.shape[0])[:, None]
    a = np.arange(p_matrix.shape[1])[None, :]
    mask = (h > line) if side == "home" else (a > line)
    return np.sum(p_matrix * mask)

def prob_team_under(p_matrix, side, line):
    h = np.arange(p_matrix.shape[0])[:, None]
    a = np.arange(p_matrix.shape[1])[None, :]
    mask = (h < line) if side == "home" else (a < line)
    return np.sum(p_matrix * mask)

def prob_team_push(p_matrix, side, line):
    if not float(line).is_integer():
        return 0.0
    h = np.arange(p_matrix.shape[0])[:, None]
    a = np.arange(p_matrix.shape[1])[None, :]
    mask = (h == line) if side == "home" else (a == line)
    return np.sum(p_matrix * mask)

def prob_double_chance(p1, px, p2, option):
    if option == "1X":
        return p1 + px
    if option == "12":
        return p1 + p2
    if option == "X2":
        return px + p2
    return np.nan

def prob_dnb_home(p_home, p_draw):
    denom = 1 - p_draw
    return p_home / denom if denom > 0 else np.nan

def prob_dnb_away(p_away, p_draw):
    denom = 1 - p_draw
    return p_away / denom if denom > 0 else np.nan

def prob_win_to_nil(p_matrix, side):
    h = np.arange(p_matrix.shape[0])[:, None]
    a = np.arange(p_matrix.shape[1])[None, :]
    if side == "home":
        return p_matrix[(a == 0) & (h > a)].sum()
    return p_matrix[(h == 0) & (a > h)].sum()

def fair_odds(prob):
    return 1.0 / prob if prob is not None and prob > 0 else np.nan

def calc_ev(prob, odds):
    if odds is None or odds <= 1.0 or prob is None or pd.isna(prob):
        return None
    return prob * odds - 1

def classify_bet(ev, prob, min_ev, min_prob, min_ev_lean):
    if ev is None or prob is None or pd.isna(prob):
        return "NO BET"
    if ev >= min_ev and prob >= min_prob:
        return "BET"
    if ev > min_ev_lean:
        return "LEAN"
    return "NO BET"

def build_market_definitions():
    defs = []
    defs += [
        {"market": "1X2 - 1", "section": "Mize", "family": "1X2", "side": "home", "line": None},
        {"market": "1X2 - X", "section": "Mize", "family": "1X2", "side": "draw", "line": None},
        {"market": "1X2 - 2", "section": "Mize", "family": "1X2", "side": "away", "line": None},
        {"market": "Double Chance - 1X", "section": "Meci - sansa dubla", "family": "DC", "side": "1X", "line": None},
        {"market": "Double Chance - 12", "section": "Meci - sansa dubla", "family": "DC", "side": "12", "line": None},
        {"market": "Double Chance - X2", "section": "Meci - sansa dubla", "family": "DC", "side": "X2", "line": None},
        {"market": "BTTS Yes", "section": "Ambele marcheaza", "family": "BTTS", "side": "yes", "line": None},
        {"market": "BTTS No", "section": "Ambele marcheaza", "family": "BTTS", "side": "no", "line": None},
        {"market": "Draw No Bet - Home", "section": "Victorie fara egal", "family": "DNB", "side": "home", "line": None},
        {"market": "Draw No Bet - Away", "section": "Victorie fara egal", "family": "DNB", "side": "away", "line": None},
        {"market": "Win To Nil - Home", "section": "Win To Nil", "family": "WTN", "side": "home", "line": None},
        {"market": "Win To Nil - Away", "section": "Win To Nil", "family": "WTN", "side": "away", "line": None},
    ]
    for line in MATCH_TOTAL_LINES:
        defs.append({"market": f"Under {line} FT", "section": "Total goluri", "family": "MATCH_TOTAL", "side": "under", "line": line})
        defs.append({"market": f"Over {line} FT", "section": "Total goluri", "family": "MATCH_TOTAL", "side": "over", "line": line})

    for line in TEAM_TOTAL_LINES:
        defs.append({"market": f"Home Under {line} Goals", "section": "Echipa 1 Total goluri", "family": "TEAM_TOTAL", "team_side": "home", "side": "under", "line": line})
        defs.append({"market": f"Home Over {line} Goals", "section": "Echipa 1 Total goluri", "family": "TEAM_TOTAL", "team_side": "home", "side": "over", "line": line})
        defs.append({"market": f"Away Under {line} Goals", "section": "Echipa 2 Total goluri", "family": "TEAM_TOTAL", "team_side": "away", "side": "under", "line": line})
        defs.append({"market": f"Away Over {line} Goals", "section": "Echipa 2 Total goluri", "family": "TEAM_TOTAL", "team_side": "away", "side": "over", "line": line})

    return defs

MARKET_DEFS = build_market_definitions()
MARKET_MAP = {m["market"]: m for m in MARKET_DEFS}
MARKETS = [m["market"] for m in MARKET_DEFS]

def get_market_probability(market_sel, p_matrix, lam_home, lam_away):
    meta = MARKET_MAP.get(market_sel)
    if meta is None:
        return np.nan, None, 0.0

    family = meta["family"]
    line = meta.get("line")

    if family == "1X2":
        p1, px, p2 = prob_1x2(p_matrix)
        if meta["side"] == "home":
            return p1, line, 0.0
        if meta["side"] == "draw":
            return px, line, 0.0
        return p2, line, 0.0

    if family == "DC":
        p1, px, p2 = prob_1x2(p_matrix)
        return prob_double_chance(p1, px, p2, meta["side"]), line, 0.0

    if family == "DNB":
        p1, px, p2 = prob_1x2(p_matrix)
        if meta["side"] == "home":
            return prob_dnb_home(p1, px), line, px
        return prob_dnb_away(p2, px), line, px

    if family == "BTTS":
        p_yes = prob_btts_yes(p_matrix)
        return (p_yes if meta["side"] == "yes" else 1 - p_yes), line, 0.0

    if family == "WTN":
        return prob_win_to_nil(p_matrix, meta["side"]), line, 0.0

    if family == "MATCH_TOTAL":
        push_prob = prob_total_push(p_matrix, line)
        if meta["side"] == "over":
            return prob_total_over(p_matrix, line), line, push_prob
        return prob_total_under(p_matrix, line), line, push_prob

    if family == "TEAM_TOTAL":
        team_side = meta["team_side"]
        push_prob = prob_team_push(p_matrix, team_side, line)
        if meta["side"] == "over":
            return prob_team_over(p_matrix, team_side, line), line, push_prob
        return prob_team_under(p_matrix, team_side, line), line, push_prob

    return np.nan, line, 0.0

def resolve_result(market: str, line, score_home: int, score_away: int) -> str:
    try:
        gh, ga = int(score_home), int(score_away)
        total = gh + ga
        meta = MARKET_MAP.get(str(market).strip())
        if meta is None:
            return "?"

        family = meta["family"]
        ln = meta.get("line") if line is None or pd.isna(line) else float(line)

        if family == "1X2":
            if meta["side"] == "home":
                return "W" if gh > ga else "L"
            if meta["side"] == "draw":
                return "W" if gh == ga else "L"
            return "W" if ga > gh else "L"

        if family == "DC":
            if meta["side"] == "1X":
                return "W" if gh >= ga else "L"
            if meta["side"] == "12":
                return "W" if gh != ga else "L"
            return "W" if ga >= gh else "L"

        if family == "DNB":
            if gh == ga:
                return "V"
            if meta["side"] == "home":
                return "W" if gh > ga else "L"
            return "W" if ga > gh else "L"

        if family == "BTTS":
            yes = gh > 0 and ga > 0
            if meta["side"] == "yes":
                return "W" if yes else "L"
            return "W" if not yes else "L"

        if family == "WTN":
            if meta["side"] == "home":
                return "W" if gh > ga and ga == 0 else "L"
            return "W" if ga > gh and gh == 0 else "L"

        if family == "MATCH_TOTAL":
            if float(ln).is_integer() and total == ln:
                return "V"
            if meta["side"] == "over":
                return "W" if total > ln else "L"
            return "W" if total < ln else "L"

        if family == "TEAM_TOTAL":
            team_goals = gh if meta["team_side"] == "home" else ga
            if float(ln).is_integer() and team_goals == ln:
                return "V"
            if meta["side"] == "over":
                return "W" if team_goals > ln else "L"
            return "W" if team_goals < ln else "L"

        return "?"
    except Exception:
        return "?"

def calc_profit(rezultat: str, odds):
    try:
        if rezultat == "W":
            return round(float(odds) - 1, 4) if odds is not None and pd.notna(odds) else None
        if rezultat == "L":
            return -1.0
        if rezultat == "V":
            return 0.0
        return None
    except Exception:
        return None

def compute_market_row(market_name, p_matrix, lam_home, lam_away, odds=None):
    prob, line, push_prob = get_market_probability(market_name, p_matrix, lam_home, lam_away)
    fair = fair_odds(prob) if not np.isnan(prob) else np.nan
    ev = calc_ev(prob, odds)
    decision = classify_bet(ev, prob, min_ev, min_prob, min_ev_lean)
    meta = MARKET_MAP[market_name]
    return {
        "Piață": market_name,
        "Secțiune": meta["section"],
        "Prob %": f"{prob * 100:.1f}%" if not np.isnan(prob) else "-",
        "Cotă fair": f"{fair:.2f}" if not np.isnan(fair) else "-",
        "Cotă BK": f"{odds:.2f}" if odds is not None else "-",
        "EV": f"{ev:.3f}" if ev is not None else "-",
        "Verdict": decision,
        "_prob": float(prob) if not np.isnan(prob) else None,
        "_fair": float(fair) if not np.isnan(fair) else None,
        "_odds": odds,
        "_ev": ev,
        "_line": line,
        "_push": push_prob,
        "_family": meta.get("family"),
        "_team_side": meta.get("team_side"),
        "_side": meta.get("side"),
    }

def style_verdict(val):
    if val == "BET":
        return "background-color:#1a7a3a;color:white;font-weight:700"
    if val == "LEAN":
        return "background-color:#b8860b;color:white;font-weight:700"
    if val == "NO BET":
        return "background-color:#444;color:#bbb"
    return ""

def render_score_matrix(p_matrix, max_goals):
    df_scores = pd.DataFrame(
        [{"Home": h, "Away": a, "Prob": p_matrix[h, a]}
         for h in range(int(max_goals) + 1)
         for a in range(int(max_goals) + 1)]
    )
    score_matrix = df_scores.pivot(index="Home", columns="Away", values="Prob")

    def highlight_top3_global(df):
        styles = pd.DataFrame("", index=df.index, columns=df.columns)
        flat = df.stack().sort_values(ascending=False).head(3)
        colors = [
            "background-color: #FFD700; color: black; font-weight: 700;",
            "background-color: #C0C0C0; color: black; font-weight: 700;",
            "background-color: #CD7F32; color: white; font-weight: 700;",
        ]
        for i, ((r, c), _) in enumerate(flat.items()):
            styles.loc[r, c] = colors[i]
        return styles

    st.dataframe(
        score_matrix.style.format("{:.3%}").background_gradient(cmap="YlGnBu", axis=None).apply(highlight_top3_global, axis=None),
        use_container_width=True,
    )
    top3 = score_matrix.stack().sort_values(ascending=False).head(3)
    st.markdown("#### 🏅 Top 3 scoruri probabile")
    for (h, a), p in top3.items():
        st.write(f"**{h}-{a}** → {p:.3%}")

def normalize_offer_text(raw: str) -> str:
    if raw is None:
        return ""
    txt = str(raw)
    txt = txt.replace("\xa0", " ")
    txt = txt.replace("\r", "\n")
    txt = txt.replace("–", "-").replace("—", "-").replace("−", "-")
    txt = txt.replace(",", ".")
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r"\n+", "\n", txt)
    return txt.strip()

def _normalize_title_key(txt: str) -> str:
    txt = normalize_offer_text(txt).lower()
    repl = {"ă": "a", "â": "a", "î": "i", "ș": "s", "ş": "s", "ț": "t", "ţ": "t"}
    for k, v in repl.items():
        txt = txt.replace(k, v)
    return txt

def extract_section(text: str, start_title: str, stop_titles: list) -> str:
    txt = normalize_offer_text(text)
    txt_key = _normalize_title_key(txt)
    start_key = _normalize_title_key(start_title)
    start = txt_key.find(start_key)
    if start == -1:
        return ""
    start += len(start_key)
    tail = txt[start:]
    tail_key = _normalize_title_key(tail)
    stop_pos = len(tail)
    for stop_title in stop_titles:
        stop_key = _normalize_title_key(stop_title)
        p = tail_key.find(stop_key)
        if p != -1:
            stop_pos = min(stop_pos, p)
    return tail[:stop_pos].strip()

def parse_total_ladder_section(section_text: str, prefix_under: str, prefix_over: str, suffix: str = "FT") -> dict:
    out = {}
    s = normalize_offer_text(section_text)
    s = re.sub(r"(?<=\d\.\d)(?=\d\.\d{2})", " ", s)
    s = re.sub(r"(?<=sub\s\d)(?=\d\.\d{2})", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"(?<=peste\s\d)(?=\d\.\d{2})", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"(?<=-\s\d)(?=\d\.\d{2})", " ", s)
    s = re.sub(r"(?<=\+\s\d)(?=\d\.\d{2})", " ", s)

    patterns = [
        re.compile(r"(?:sub|-)\\s*(\\d+(?:\\.\\d+)?)\\s+([0-9]+(?:\\.\\d+)?)\\s*(?:peste|\\+)\\s*\\s+([0-9]+(?:\\.\\d+)?)", re.IGNORECASE),
        re.compile(r"(?:sub|-)\\s*(\\d+(?:\\.\\d+)?)\\s+([0-9]+(?:\\.\\d+)?)\\s*(?:peste|\\+)\\s*(\\d+(?:\\.\\d+)?)\\s+([0-9]+(?:\\.\\d+)?)", re.IGNORECASE),
    ]

    for pattern in patterns:
        for m in pattern.finditer(s):
            line1 = float(m.group(1))
            odd_under = float(m.group(2))
            if len(m.groups()) == 3:
                line2 = line1
                odd_over = float(m.group(3))
            else:
                line2 = float(m.group(3))
                odd_over = float(m.group(4))
            if abs(line1 - line2) < 1e-9:
                out[f"{prefix_under} {line1:.1f} {suffix}"] = odd_under
                out[f"{prefix_over} {line1:.1f} {suffix}"] = odd_over

    return out

def extract_decimal_odds(raw: str) -> list:
    txt = normalize_offer_text(raw)
    txt = re.sub(r"\d+[.,]\d+\s*%", " ", txt)
    txt = re.sub(r"\d+\s*%", " ", txt)
    vals = re.findall(r"(\b[1-9]\d*\.\d{2}\b)", txt)
    result = []
    for v in vals:
        try:
            f = float(v)
            if 1.01 <= f <= 50.0:
                result.append(f)
        except Exception:
            pass
    return result

def parse_bookmaker_offer_skip_htft(raw_text: str) -> dict:
    txt = normalize_offer_text(raw_text)
    odds_map = {}

    sec_meci = extract_section(txt, "Meci", ["Mize", "Meci - sansa dubla", "Pauza sau final"])
    vals = extract_decimal_odds(sec_meci)
    if len(vals) >= 3:
        odds_map["1X2 - 1"] = vals[0]
        odds_map["1X2 - X"] = vals[1]
        odds_map["1X2 - 2"] = vals[2]

    sec_dc = extract_section(txt, "Meci - sansa dubla", ["Pauza sau final", "Victorie fara egal", "Ambele marcheaza"])
    vals = extract_decimal_odds(sec_dc)
    if len(vals) >= 3:
        odds_map["Double Chance - 1X"] = vals[0]
        odds_map["Double Chance - 12"] = vals[1]
        odds_map["Double Chance - X2"] = vals[2]

    sec_dnb = extract_section(txt, "Victorie fara egal", ["Ambele marcheaza", "Total goluri"])
    vals = extract_decimal_odds(sec_dnb)
    if len(vals) >= 2:
        odds_map["Draw No Bet - Home"] = vals[0]
        odds_map["Draw No Bet - Away"] = vals[1]

    sec_btts = extract_section(txt, "Ambele marcheaza", ["Total goluri", "Echipa 1 Total goluri", "Echipa 2 Total goluri"])
    vals = extract_decimal_odds(sec_btts)
    if len(vals) >= 2:
        odds_map["BTTS Yes"] = vals[0]
        odds_map["BTTS No"] = vals[1]

    sec_total = extract_section(txt, "Total goluri", ["Echipa 1 Total goluri", "Echipa 2 Total goluri", "Win To Nil"])
    odds_map.update(parse_total_ladder_section(sec_total, "Under", "Over", suffix="FT"))

    sec_home = extract_section(txt, "Echipa 1 Total goluri", ["Echipa 2 Total goluri", "Win To Nil"])
    odds_map.update(parse_total_ladder_section(sec_home, "Home Under", "Home Over", suffix="Goals"))

    sec_away = extract_section(txt, "Echipa 2 Total goluri", ["Win To Nil"])
    odds_map.update(parse_total_ladder_section(sec_away, "Away Under", "Away Over", suffix="Goals"))

    return odds_map

# ============== HISTORIC COLUMNS (folosit doar pentru import/calibrare) ==============

HIST_COLUMNS = [
    "ts", "league", "home", "away", "market", "section", "line",
    "prob", "push_prob", "fair_odds", "odds", "ev", "decision",
    "selected_for_play", "proposal_rank", "proposal_group",
    "lambda_home", "lambda_away", "notes", "rezultat", "profit"
]

def resolve_result_conflict(market: str, line, score_home: int, score_away: int) -> str:
    return resolve_result(market, line, score_home, score_away)

# ============== HELPERS CALIBRARE REALISTĂ ==============

def safe_clip_prob(s):
    return pd.to_numeric(s, errors="coerce").clip(1e-6, 1 - 1e-6)

def infer_family_from_market(market_name: str) -> str:
    meta = MARKET_MAP.get(str(market_name), {})
    return meta.get("family", "UNKNOWN")

def infer_side_from_market(market_name: str) -> str:
    meta = MARKET_MAP.get(str(market_name), {})
    team_side = meta.get("team_side")
    side = meta.get("side")
    if team_side in ["home", "away"] and side in ["over", "under"]:
        return f"{team_side}_{side}"
    return str(side) if side is not None else "unknown"

def infer_line_bucket(line_val):
    try:
        if pd.isna(line_val):
            return "no_line"
        x = float(line_val)
    except Exception:
        return "no_line"
    if x <= 0.5:
        return "<=0.5"
    if x <= 1.5:
        return "1.0-1.5"
    if x <= 2.5:
        return "2.0-2.5"
    if x <= 3.5:
        return "3.0-3.5"
    return ">=4.0"

def implied_prob_from_odds(odds):
    odds = pd.to_numeric(odds, errors="coerce")
    out = pd.Series(np.nan, index=odds.index, dtype=float)
    mask = odds > 1.0
    out.loc[mask] = 1.0 / odds.loc[mask]
    return out.clip(1e-6, 1 - 1e-6)

def add_evaluation_groups(df):
    df = df.copy()
    df["family"] = df["market"].apply(infer_family_from_market)
    df["side_eval"] = df["market"].apply(infer_side_from_market)
    df["line_bucket"] = df["line"].apply(infer_line_bucket)
    df["league_group"] = df["league"].fillna("UNKNOWN").astype(str)
    home_tokens = df["home"].fillna("UNKNOWN").astype(str).str.upper().str[:1]
    away_tokens = df["away"].fillna("UNKNOWN").astype(str).str.upper().str[:1]
    df["home_away_split"] = np.where((home_tokens <= "M") & (away_tokens <= "M"), "A-M vs A-M",
                               np.where((home_tokens > "M") & (away_tokens > "M"), "N-Z vs N-Z", "mixed"))
    return df

def compute_binary_metrics(y, p):
    y = np.asarray(y, dtype=float)
    p = np.asarray(p, dtype=float)
    p = np.clip(p, 1e-6, 1 - 1e-6)
    brier = np.mean((p - y) ** 2)
    log_loss = -np.mean(y * np.log(p) + (1 - y) * np.log(1 - p))
    return float(brier), float(log_loss)

def fit_calibration_slope_intercept(y, p):
    y = np.asarray(y, dtype=float)
    p = np.clip(np.asarray(p, dtype=float), 1e-6, 1 - 1e-6)
    if len(y) < 20 or len(np.unique(y)) < 2:
        return np.nan, np.nan
    x = np.log(p / (1 - p))
    X = np.column_stack([np.ones(len(x)), x])
    beta = np.zeros(X.shape[1])
    for _ in range(50):
        eta = X @ beta
        mu = 1 / (1 + np.exp(-eta))
        w = np.clip(mu * (1 - mu), 1e-8, None)
        z = eta + (y - mu) / w
        XtW = X.T * w
        try:
            beta_new = np.linalg.solve(XtW @ X, XtW @ z)
        except np.linalg.LinAlgError:
            return np.nan, np.nan
        if np.max(np.abs(beta_new - beta)) < 1e-8:
            beta = beta_new
            break
        beta = beta_new
    return float(beta[1]), float(beta[0])

def grouped_eval_table(df, group_cols, min_n=8):
    rows = []
    for keys, g in df.groupby(group_cols, dropna=False):
        if len(g) < min_n:
            continue
        model_brier, model_logloss = compute_binary_metrics(g["outcome"], g["prob"])
        row = {"N": len(g), "Model Brier": model_brier, "Model LogLoss": model_logloss,
               "Model Prob %": g["prob"].mean() * 100, "Win Rate %": g["outcome"].mean() * 100,
               "Gap %": (g["prob"].mean() - g["outcome"].mean()) * 100}
        if isinstance(keys, tuple):
            for col, val in zip(group_cols, keys):
                row[col] = val
        else:
            row[group_cols[0]] = keys
        if g["implied_prob"].notna().sum() >= min_n:
            mask = g["implied_prob"].notna()
            base_brier, base_logloss = compute_binary_metrics(g.loc[mask, "outcome"], g.loc[mask, "implied_prob"])
            row["Odds Brier"] = base_brier
            row["Odds LogLoss"] = base_logloss
            row["Brier Skill vs Odds %"] = (1 - model_brier / base_brier) * 100 if base_brier > 0 else np.nan
            row["LogLoss Skill vs Odds %"] = (1 - model_logloss / base_logloss) * 100 if base_logloss > 0 else np.nan
        rows.append(row)
    return pd.DataFrame(rows)


def prepare_walkforward_df(df):
    df = df.copy()
    df["ts"] = pd.to_datetime(df["ts"], errors="coerce")
    df = df.dropna(subset=["ts", "prob", "outcome"]).sort_values("ts").reset_index(drop=True)
    df["group_key"] = (
        df["family"].fillna("UNKNOWN").astype(str) + "|" +
        df["line_bucket"].fillna("UNKNOWN").astype(str) + "|" +
        df["side_eval"].fillna("UNKNOWN").astype(str)
    )
    return df

def compute_drawdown(equity_series):
    roll_max = equity_series.cummax()
    drawdown = equity_series - roll_max
    return drawdown

def walkforward_backtest(df, min_train=30, retrain_step=10, prob_col="prob"):
    df = prepare_walkforward_df(df)
    if len(df) < min_train + 10:
        return pd.DataFrame(), {}

    preds = []
    last_fit_idx = -1
    group_stats = None
    global_mean = None

    for i in range(len(df)):
        if i < min_train:
            continue
        if group_stats is None or (i - last_fit_idx) >= retrain_step:
            train = df.iloc[:i].copy()
            group_stats = train.groupby("group_key")["outcome"].agg(["mean", "count"]).reset_index()
            group_stats = group_stats.rename(columns={"mean": "wf_prob_raw", "count": "train_n"})
            global_mean = float(train["outcome"].mean())
            last_fit_idx = i

        row = df.iloc[[i]].copy()
        row = row.merge(group_stats, on="group_key", how="left")
        raw = row["wf_prob_raw"].iloc[0] if pd.notna(row["wf_prob_raw"].iloc[0]) else global_mean
        train_n = row["train_n"].iloc[0] if pd.notna(row["train_n"].iloc[0]) else 0
        shrink = min(float(train_n) / 20.0, 1.0)
        wf_prob = shrink * float(raw) + (1 - shrink) * float(global_mean)
        wf_prob = float(np.clip(wf_prob, 1e-6, 1 - 1e-6))

        stake_ev = np.nan
        if pd.notna(row["odds"].iloc[0]) and float(row["odds"].iloc[0]) > 1.0:
            stake_ev = wf_prob * float(row["odds"].iloc[0]) - 1

        do_bet = pd.notna(stake_ev) and stake_ev > 0
        realized_profit = np.nan
        if do_bet:
            rez = row["rezultat"].iloc[0]
            odds = row["odds"].iloc[0]
            realized_profit = calc_profit(rez, odds)

        preds.append({
            "ts": row["ts"].iloc[0],
            "league": row["league"].iloc[0],
            "family": row["family"].iloc[0],
            "line_bucket": row["line_bucket"].iloc[0],
            "side_eval": row["side_eval"].iloc[0],
            "outcome": int(row["outcome"].iloc[0]),
            "model_prob": float(row[prob_col].iloc[0]),
            "wf_prob": wf_prob,
            "implied_prob": row["implied_prob"].iloc[0],
            "odds": row["odds"].iloc[0],
            "ev_wf": stake_ev,
            "bet_wf": int(bool(do_bet)),
            "profit_wf": realized_profit,
            "train_size": i
        })

    pred_df = pd.DataFrame(preds)
    if pred_df.empty:
        return pred_df, {}

    pred_df["profit_wf_filled"] = pred_df["profit_wf"].fillna(0.0)
    pred_df["equity_wf"] = pred_df["profit_wf_filled"].cumsum()
    pred_df["drawdown_wf"] = compute_drawdown(pred_df["equity_wf"])

    summary = {
        "n_test": int(len(pred_df)),
        "n_bets": int(pred_df["bet_wf"].sum()),
        "avg_ev": float(pred_df.loc[pred_df["bet_wf"] == 1, "ev_wf"].mean()) if (pred_df["bet_wf"] == 1).any() else np.nan,
        "profit_total": float(pred_df["profit_wf_filled"].sum()),
        "roi": float(pred_df.loc[pred_df["bet_wf"] == 1, "profit_wf"].sum() / pred_df["bet_wf"].sum()) if pred_df["bet_wf"].sum() > 0 else np.nan,
        "max_drawdown": float(pred_df["drawdown_wf"].min()) if len(pred_df) else np.nan,
    }
    return pred_df, summary

# ============== SIDEBAR ==============

st.sidebar.header("⚙️ Setări")
st.sidebar.info("Modul cloud: date din fișierul xGData.xlsx din repo.")

leagues = {}
league_names = []

DEFAULT_FILE = "xGData.xlsx"

uploaded_xg = st.sidebar.file_uploader("Încarcă xGData.xlsx (opțional override)", type=["xlsx"])

if uploaded_xg is not None:
    file_buffer = io.BytesIO(uploaded_xg.getvalue())
    leagues = load_xg_workbook(file_buffer)
    league_names = list(leagues.keys())
    st.sidebar.success("Folosesc fișierul încărcat.")
elif os.path.exists(DEFAULT_FILE):
    leagues = load_xg_workbook(DEFAULT_FILE)
    league_names = list(leagues.keys())
    st.sidebar.success("Date încărcate din repo (xGData.xlsx).")
else:
    st.sidebar.warning("⚠️ Fișierul xGData.xlsx nu a fost găsit în repo. Încarcă manual din sidebar.")

col_min_ev, col_min_prob = st.sidebar.columns(2)
with col_min_ev:
    min_ev = st.number_input("Prag EV BET", value=DEFAULT_MIN_EV, format="%.3f")
with col_min_prob:
    min_prob = st.number_input("Prag Prob BET", value=DEFAULT_MIN_PROB, format="%.2f")
min_ev_lean = st.sidebar.number_input("Prag EV LEAN (>0)", value=DEFAULT_MIN_EV_LEAN, format="%.3f")
max_goals = st.sidebar.number_input("Max goals Poisson", value=MAX_GOALS, min_value=5, max_value=12, step=1)

st.sidebar.markdown("---")
st.sidebar.caption("xG Poisson Analyzer — Cloud Edition")

# ============== TABURI ==============

tab_single, tab_calib, tab_history = st.tabs(
    ["🎯 Single Match", "🎯 Calibrare Model", "📚 Istoric"]
)

# ──────────────── TAB: SINGLE MATCH ────────────────
with tab_single:
    if len(leagues) == 0:
        st.warning("⚠️ Încarcă fișierul xGData.xlsx din sidebar pentru a folosi acest tab.")
    else:
        st.subheader("🎯 Analiză meci individual (Poisson)")
        col_lg, col_home, col_away = st.columns(3)
        with col_lg:
            league_sel = st.selectbox("Alege liga", league_names)
        df_lg = leagues[league_sel]
        strengths = build_team_strengths(df_lg)
        teams = strengths["teams"]
        with col_home:
            home_team = st.selectbox("Gazde", teams, key="single_home")
        with col_away:
            away_team = st.selectbox("Oaspeți", teams, key="single_away", index=min(1, len(teams) - 1))

        if home_team == away_team:
            st.warning("⚠️ Alege echipe diferite pentru gazde și oaspeți.")
            st.stop()

        market_sel = st.selectbox("Piață", MARKETS, key="single_market")
        odds_input = st.text_input("Cota (opțional)", value="")
        notes = st.text_input("Note (opțional)", value="")

        if st.button("🔍 Calculează"):
            lam_home, lam_away = get_lambdas(strengths, home_team, away_team)
            p_matrix = poisson_matrix(lam_home, lam_away, max_goals=int(max_goals))
            try:
                odds_val = float(odds_input.replace(",", ".")) if odds_input.strip() else None
            except ValueError:
                odds_val = None

            row = compute_market_row(market_sel, p_matrix, lam_home, lam_away, odds_val)
            col_res1, col_res2 = st.columns(2)
            with col_res1:
                st.metric("λ home", f"{lam_home:.2f}")
                st.metric("λ away", f"{lam_away:.2f}")
            with col_res2:
                st.metric("Probabilitate", row["Prob %"])
                st.metric("Cotă corectă (fair)", row["Cotă fair"])
            if odds_val is not None:
                st.write(f"**Cota introdusă**: {odds_val:.2f}")
                st.write(f"**EV**: {row['EV']}")
            st.markdown(f"### Verdict: **{row['Verdict']}**")
            st.markdown(f"### 🔢 Distribuția scorurilor (0..{int(max_goals)} goluri)")
            render_score_matrix(p_matrix, max_goals)

            # Toate piețele — shortlist instant
            st.markdown("---")
            st.markdown("### 📋 Shortlist toate piețele (pentru cotă introdusă)")
            if odds_val is not None:
                all_rows = []
                for mkt in MARKETS:
                    r2 = compute_market_row(mkt, p_matrix, lam_home, lam_away, odds_val)
                    all_rows.append({
                        "Piață": r2["Piață"],
                        "Secțiune": r2["Secțiune"],
                        "Prob %": r2["Prob %"],
                        "Cotă fair": r2["Cotă fair"],
                        "EV": r2["EV"],
                        "Verdict": r2["Verdict"],
                    })
                df_all = pd.DataFrame(all_rows)
                df_bet = df_all[df_all["Verdict"].isin(["BET", "LEAN"])]
                if not df_bet.empty:
                    st.dataframe(df_bet.style.applymap(style_verdict, subset=["Verdict"]), use_container_width=True, hide_index=True)
                else:
                    st.info("Nicio piață BET/LEAN pentru această cotă.")
            else:
                st.caption("Introdu o cotă pentru a vedea shortlistul pe toate piețele.")

# ──────────────── TAB: CALIBRARE MODEL ────────────────
with tab_calib:
    st.subheader("🎯 Calibrare & Validare Model Poisson")
    st.markdown(
        "Încarcă un fișier Excel cu istoricul predicțiilor (format standard: sheet **Istoric**, "
        "coloane: `prob`, `rezultat`, `market`, `section`, `odds`, `profit`) "
        "pentru a valida acuratețea modelului Poisson."
    )

    df_calib_raw = pd.DataFrame(columns=HIST_COLUMNS)
    if os.path.exists(HISTORY_PATH):
        try:
            df_calib_raw = pd.read_excel(HISTORY_PATH, sheet_name="Istoric")
            for col in HIST_COLUMNS:
                if col not in df_calib_raw.columns:
                    df_calib_raw[col] = None
            df_calib_raw = df_calib_raw[HIST_COLUMNS].reset_index(drop=True)
        except Exception:
            df_calib_raw = pd.DataFrame(columns=HIST_COLUMNS)

    with st.expander("📂 Încarcă fișier istoric pentru calibrare", expanded=True):
        if os.path.exists(HISTORY_PATH) and len(df_calib_raw) > 0:
            st.success(f"Fișier implicit încărcat din repo: {HISTORY_PATH} ({len(df_calib_raw)} intrări).")
        elif os.path.exists(HISTORY_PATH):
            st.info(f"Fișierul {HISTORY_PATH} există în repo, dar nu a putut fi citit corect din sheet-ul 'Istoric'.")
        else:
            st.info(f"Fișierul implicit {HISTORY_PATH} nu a fost găsit în repo. Poți încărca unul manual.")

        uploaded_calib = st.file_uploader(
            "Încarcă fișierul Excel cu predicții (opțional override, sheet 'Istoric')",
            type=["xlsx"],
            key="calib_upload"
        )

        if uploaded_calib is not None:
            try:
                df_calib_raw = pd.read_excel(io.BytesIO(uploaded_calib.getvalue()), sheet_name="Istoric")
                for col in HIST_COLUMNS:
                    if col not in df_calib_raw.columns:
                        df_calib_raw[col] = None
                df_calib_raw = df_calib_raw[HIST_COLUMNS].reset_index(drop=True)
                st.success(f"✅ Fișier încărcat manual: {len(df_calib_raw)} intrări.")
            except Exception as e:
                st.error(f"❌ Eroare la citire fișier: {e}")
    if df_calib_raw.empty:
        st.info("Încarcă un fișier cu predicții și rezultate W/L/V pentru a vedea calibrarea modelului.")
    else:
        df_calib = df_calib_raw.copy()
        df_calib = df_calib[df_calib["rezultat"].isin(["W", "L", "V"])].copy()
        if df_calib.empty:
            st.info("Fișierul nu conține rezultate finalizate (W/L/V). Adaugă rezultate reale în fișierul Excel.")
        else:
            df_calib["prob"] = safe_clip_prob(df_calib["prob"])
            df_calib["push_prob"] = pd.to_numeric(df_calib["push_prob"], errors="coerce")
            df_calib["odds"] = pd.to_numeric(df_calib["odds"], errors="coerce")
            df_calib["profit"] = pd.to_numeric(df_calib["profit"], errors="coerce")
            df_calib = df_calib.dropna(subset=["prob"])
            df_calib["outcome"] = (df_calib["rezultat"] == "W").astype(int)
            df_calib["implied_prob"] = implied_prob_from_odds(df_calib["odds"])
            df_calib = add_evaluation_groups(df_calib)
            n_total = len(df_calib)
            st.markdown(f"**Pariuri analizate:** {n_total} (W/L/V)")
            st.caption("Evaluarea de mai jos este intenționat conservatoare: compară modelul cu probabilitatea implicită din cote, separă grupurile de piață și raportează calibration slope.")
            if n_total < 20:
                st.warning("Sub 20 rezultate, concluziile sunt foarte instabile. Interpretarea trebuie făcută cu prudență.")

            cal1, cal2, cal3, cal4, cal5, cal6 = st.tabs([
                "📊 Calibrare generală",
                "🪣 Grupuri realiste",
                "📈 Skill vs odds",
                "📐 Calibration slope",
                "🔬 Split-uri evaluare",
                "🚶 Walk-forward"
            ])

            with cal1:
                avg_prob = df_calib["prob"].mean()
                actual_wr = df_calib["outcome"].mean()
                calib_error = avg_prob - actual_wr
                k1, k2, k3 = st.columns(3)
                k1.metric("Prob medie model", f"{avg_prob*100:.1f}%")
                k2.metric("Win rate real", f"{actual_wr*100:.1f}%")
                k3.metric("Gap agregat", f"{calib_error*100:.1f} pp")

                prob_bins = pd.cut(df_calib["prob"], bins=np.linspace(0, 1, 11), include_lowest=True)
                bin_counts = df_calib.groupby(prob_bins, observed=True).agg(
                    Nr_pariuri=("outcome", "count"),
                    Win_rate_real=("outcome", "mean"),
                    Prob_medie_bin=("prob", "mean")
                ).reset_index()
                bin_counts["Interval prob"] = bin_counts["prob"].astype(str) if "prob" in bin_counts.columns else bin_counts.iloc[:, 0].astype(str)
                if "prob" in bin_counts.columns:
                    bin_counts = bin_counts.drop(columns=["prob"])
                bin_counts["Win rate real %"] = (bin_counts["Win_rate_real"] * 100).round(1)
                bin_counts["Prob medie %"] = (bin_counts["Prob_medie_bin"] * 100).round(1)
                bin_counts["Diferență %"] = (bin_counts["Prob_medie_bin"] - bin_counts["Win_rate_real"]).mul(100).round(1)
                st.dataframe(bin_counts[["Interval prob", "Nr_pariuri", "Prob medie %", "Win rate real %", "Diferență %"]], use_container_width=True, hide_index=True)

            with cal2:
                st.markdown("#### Regrupare pe family + line bucket + side")
                grouped = grouped_eval_table(df_calib, ["family", "line_bucket", "side_eval"], min_n=8)
                if grouped.empty:
                    st.info("Nu există suficiente date pentru gruparea family + line bucket + side (minim 8 rezultate pe grup).")
                else:
                    grouped = grouped.sort_values(["N", "Gap %"], ascending=[False, True], key=lambda s: abs(s) if s.name == "Gap %" else s)
                    for col in ["Model Brier", "Model LogLoss", "Odds Brier", "Odds LogLoss", "Brier Skill vs Odds %", "LogLoss Skill vs Odds %", "Model Prob %", "Win Rate %", "Gap %"]:
                        if col in grouped.columns:
                            grouped[col] = pd.to_numeric(grouped[col], errors="coerce").round(3)
                    st.dataframe(grouped, use_container_width=True, hide_index=True)

            with cal3:
                outcomes = df_calib["outcome"].values
                probs = df_calib["prob"].values
                model_brier, model_log_loss = compute_binary_metrics(outcomes, probs)
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Model Brier", f"{model_brier:.4f}")
                m3.metric("Model Log Loss", f"{model_log_loss:.4f}")

                if df_calib["implied_prob"].notna().sum() >= 20:
                    mask = df_calib["implied_prob"].notna()
                    odds_brier, odds_log_loss = compute_binary_metrics(df_calib.loc[mask, "outcome"], df_calib.loc[mask, "implied_prob"])
                    brier_skill = 1 - model_brier / odds_brier if odds_brier > 0 else np.nan
                    logloss_skill = 1 - model_log_loss / odds_log_loss if odds_log_loss > 0 else np.nan
                    m2.metric("Skill Brier vs odds", f"{brier_skill*100:.1f}%")
                    m4.metric("Skill LogLoss vs odds", f"{logloss_skill*100:.1f}%")
                    st.caption("Skill pozitiv = modelul bate baseline-ul din cotele bookmakerului. Skill negativ = modelul este mai slab decât baseline-ul.")
                    st.write(f"Baseline odds — Brier: {odds_brier:.4f} | Log Loss: {odds_log_loss:.4f}")
                else:
                    m2.metric("Skill Brier vs odds", "n/a")
                    m4.metric("Skill LogLoss vs odds", "n/a")
                    st.warning("Prea puține cote valide pentru un baseline realist bazat pe implied probability.")

            with cal4:
                slope, intercept = fit_calibration_slope_intercept(df_calib["outcome"], df_calib["prob"])
                c1, c2 = st.columns(2)
                c1.metric("Calibration slope", "n/a" if pd.isna(slope) else f"{slope:.3f}")
                c2.metric("Calibration intercept", "n/a" if pd.isna(intercept) else f"{intercept:.3f}")
                st.markdown("""
- **Slope < 1**: predicțiile sunt prea extreme, deci evaluarea optimistă dinainte te-ar fi putut păcăli.
- **Slope > 1**: predicțiile sunt prea conservatoare.
- **Intercept < 0**: modelul supraestimează win rate-ul.
- **Intercept > 0**: modelul subestimează win rate-ul.
""")

            with cal5:
                st.markdown("#### Split pe ligi")
                league_eval = grouped_eval_table(df_calib, ["league_group"], min_n=8)
                if league_eval.empty:
                    st.info("Nu există suficiente rezultate per ligă pentru evaluare separată.")
                else:
                    keep_cols = [c for c in ["league_group", "N", "Model Prob %", "Win Rate %", "Gap %", "Model Brier", "Odds Brier", "Brier Skill vs Odds %"] if c in league_eval.columns]
                    st.dataframe(league_eval[keep_cols].sort_values("N", ascending=False), use_container_width=True, hide_index=True)

                st.markdown("#### Split home/away proxy")
                ha_eval = grouped_eval_table(df_calib, ["home_away_split"], min_n=8)
                if ha_eval.empty:
                    st.info("Nu există suficiente rezultate pentru split-ul home/away proxy.")
                else:
                    keep_cols = [c for c in ["home_away_split", "N", "Model Prob %", "Win Rate %", "Gap %", "Model Brier", "Odds Brier", "Brier Skill vs Odds %"] if c in ha_eval.columns]
                    st.dataframe(ha_eval[keep_cols].sort_values("N", ascending=False), use_container_width=True, hide_index=True)

            with cal6:
                st.markdown("#### Walk-forward = test de realitate")
                st.caption("Split temporal, retraining incremental și tracking EV / ROI / drawdown pe ferestre out-of-sample.")

                if df_calib["ts"].isna().all():
                    st.warning("Fișierul nu conține coloana ts validă. Walk-forward are nevoie de timestamp pentru split temporal.")
                else:
                    wf_col1, wf_col2 = st.columns(2)
                    with wf_col1:
                        min_train = st.number_input("Minim observații train", min_value=20, max_value=500, value=30, step=5, key="wf_min_train")
                    with wf_col2:
                        retrain_step = st.number_input("Retrain incremental la fiecare N meciuri", min_value=1, max_value=100, value=10, step=1, key="wf_retrain_step")

                    wf_df, wf_summary = walkforward_backtest(df_calib, min_train=int(min_train), retrain_step=int(retrain_step), prob_col="prob")
                    if wf_df.empty:
                        st.info("Nu sunt suficiente date pentru walk-forward cu setările actuale.")
                    else:
                        a1, a2, a3, a4, a5 = st.columns(5)
                        a1.metric("Meciuri test", wf_summary.get("n_test", 0))
                        a2.metric("Beturi plasate", wf_summary.get("n_bets", 0))
                        a3.metric("EV mediu", "n/a" if pd.isna(wf_summary.get("avg_ev")) else f"{wf_summary['avg_ev']:.3f}")
                        a4.metric("ROI", "n/a" if pd.isna(wf_summary.get("roi")) else f"{wf_summary['roi']*100:.1f}%")
                        a5.metric("Max drawdown", "n/a" if pd.isna(wf_summary.get("max_drawdown")) else f"{wf_summary['max_drawdown']:.2f}u")
                        st.write(f"Profit total walk-forward: {wf_summary.get('profit_total', 0):.2f}u")

                        wf_model_brier, wf_model_logloss = compute_binary_metrics(wf_df["outcome"], wf_df["wf_prob"])
                        st.write(f"Walk-forward Brier: {wf_model_brier:.4f} | Walk-forward Log Loss: {wf_model_logloss:.4f}")
                        if wf_df["implied_prob"].notna().sum() >= 20:
                            mask = wf_df["implied_prob"].notna()
                            wf_odds_brier, wf_odds_logloss = compute_binary_metrics(wf_df.loc[mask, "outcome"], wf_df.loc[mask, "implied_prob"])
                            st.write(
                                f"Baseline odds pe setul walk-forward — Brier: {wf_odds_brier:.4f} | "
                                f"Log Loss: {wf_odds_logloss:.4f} | "
                                f"Skill Brier: {(1 - wf_model_brier / wf_odds_brier)*100:.1f}%"
                            )

                        st.markdown("#### Equity curve")
                        st.line_chart(wf_df.set_index("ts")[["equity_wf"]], use_container_width=True)
                        st.markdown("#### Drawdown")
                        st.line_chart(wf_df.set_index("ts")[["drawdown_wf"]], use_container_width=True)

                        st.markdown("#### Test pe fiecare segment")
                        seg = wf_df.groupby(["family", "line_bucket", "side_eval"]).agg(
                            N=("outcome", "count"),
                            Bets=("bet_wf", "sum"),
                            Avg_EV=("ev_wf", "mean"),
                            Profit=("profit_wf", "sum")
                        ).reset_index()
                        seg = seg[seg["N"] >= 5].copy()
                        if seg.empty:
                            st.info("Nu există suficiente observații per segment în walk-forward.")
                        else:
                            seg["ROI %"] = np.where(seg["Bets"] > 0, seg["Profit"] / seg["Bets"] * 100, np.nan)
                            st.dataframe(seg.sort_values(["N", "ROI %"], ascending=[False, False]), use_container_width=True, hide_index=True)

# ──────────────── TAB: ISTORIC ────────────────
with tab_history:
    st.subheader("📚 Istoric predicții + 📈 Tracker rezultate")
    st.markdown(
        "Încarcă un fișier Excel cu istoricul predicțiilor exportat anterior. "
        "Poți introduce rezultate reale (goluri) și urmări performanța modelului."
    )

    default_hist_df = pd.DataFrame(columns=HIST_COLUMNS)
    if os.path.exists(HISTORY_PATH):
        try:
            default_hist_df = pd.read_excel(HISTORY_PATH, sheet_name="Istoric")
            for col in HIST_COLUMNS:
                if col not in default_hist_df.columns:
                    default_hist_df[col] = None
            default_hist_df = default_hist_df[HIST_COLUMNS].reset_index(drop=True)
        except Exception:
            default_hist_df = pd.DataFrame(columns=HIST_COLUMNS)

    if "hist_df" not in st.session_state:
        st.session_state["hist_df"] = default_hist_df.copy()
        st.session_state["hist_source"] = "repo" if len(default_hist_df) > 0 else "empty"

    with st.expander("📂 Încarcă fișier istoric", expanded=True):
        if os.path.exists(HISTORY_PATH) and len(default_hist_df) > 0:
            st.success(f"Fișier implicit încărcat din repo: {HISTORY_PATH} ({len(default_hist_df)} intrări).")
        elif os.path.exists(HISTORY_PATH):
            st.info(f"Fișierul {HISTORY_PATH} există în repo, dar nu a putut fi citit corect din sheet-ul 'Istoric'.")
        else:
            st.info(f"Fișierul implicit {HISTORY_PATH} nu a fost găsit în repo. Poți încărca unul manual.")

        uploaded_hist = st.file_uploader(
            "Încarcă Predictii_xG_Poisson.xlsx (opțional override, sheet 'Istoric')",
            type=["xlsx"],
            key="upload_hist"
        )
        if uploaded_hist is not None:
            try:
                df_imp = pd.read_excel(io.BytesIO(uploaded_hist.getvalue()), sheet_name="Istoric")
                for col in HIST_COLUMNS:
                    if col not in df_imp.columns:
                        df_imp[col] = None
                df_imp = df_imp[HIST_COLUMNS].reset_index(drop=True)
                st.session_state["hist_df"] = df_imp
                st.session_state["hist_source"] = "upload"
                st.success(f"✅ Importat manual {len(df_imp)} intrări.")
            except Exception as e:
                st.error(f"❌ Eroare import: {e}")

    df = st.session_state["hist_df"]
    n = len(df)
    col_r, col_dl = st.columns([1, 1])
    with col_r:
        if st.button("🔄 Resetează sesiunea"):
            st.session_state["hist_df"] = pd.DataFrame(columns=HIST_COLUMNS)
            st.rerun()
    with col_dl:
        if n > 0:
            buf_dl = io.BytesIO()
            with pd.ExcelWriter(buf_dl, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Istoric")
                workbook = writer.book
                worksheet = writer.sheets["Istoric"]
                header_fmt = workbook.add_format({"bold": True, "bg_color": "#2E86AB", "font_color": "white"})
                for col_num, col_name in enumerate(df.columns):
                    worksheet.write(0, col_num, col_name, header_fmt)
                    worksheet.set_column(col_num, col_num, max(15, len(str(col_name)) + 2))
            buf_dl.seek(0)
            st.download_button(
                "📥 Descarcă Excel actualizat",
                data=buf_dl,
                file_name="Predictii_xG_Poisson.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.markdown(f"### 📋 Total intrări: **{n}**")
    if n == 0:
        st.info("Nu există predicții încărcate. Folosește butonul de mai sus pentru a importa un fișier.")
    else:
        df_show = df.reset_index(drop=True)
        if "ts" in df_show.columns:
            try:
                df_show["ts"] = pd.to_datetime(df_show["ts"])
                df_show = df_show.sort_values("ts", ascending=False).reset_index(drop=True)
            except Exception:
                pass
        df_display = df_show.copy()
        history_order = ["ts", "league", "home", "away", "section", "market", "line", "prob", "push_prob",
                         "fair_odds", "odds", "ev", "decision", "selected_for_play", "proposal_rank",
                         "proposal_group", "rezultat", "profit", "lambda_home", "lambda_away", "notes"]
        df_display = df_display[[c for c in history_order if c in df_display.columns]]
        for c in ["prob", "push_prob"]:
            if c in df_display.columns:
                df_display[c] = df_display[c].apply(lambda x: f"{float(x)*100:.1f}%" if pd.notna(x) else "-")
        for c in ["fair_odds", "odds", "line"]:
            if c in df_display.columns:
                df_display[c] = df_display[c].apply(lambda x: f"{float(x):.2f}" if pd.notna(x) and str(x) != '' else "-")
        for c in ["ev", "lambda_home", "lambda_away", "profit"]:
            if c in df_display.columns:
                df_display[c] = df_display[c].apply(lambda x: f"{float(x):.3f}" if pd.notna(x) else "-")
        st.dataframe(df_display, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### Introducere rezultate reale")
        if "rezultat" in df.columns:
            df_pend = df[df["rezultat"].isin(["pending", None, "", "?"])].copy()
        else:
            df_pend = df.copy()
        if df_pend.empty:
            st.success("Toate intrările au rezultat introdus!")
        else:
            df_pend["key"] = df_pend["home"].astype(str) + " vs " + df_pend["away"].astype(str) + " | " + df_pend["league"].astype(str)
            meciuri = df_pend["key"].unique().tolist()
            mecisel = st.selectbox(f"Meciuri fără rezultat ({len(meciuri)})", meciuri, key="sel_meci_tracker")
            df_m = df_pend[df_pend["key"] == mecisel]
            st.dataframe(
                df_m[[c for c in ["market", "section", "line", "prob", "odds", "decision", "selected_for_play", "proposal_rank"] if c in df_m.columns]],
                use_container_width=True, hide_index=True
            )
            c1, c2 = st.columns(2)
            with c1:
                gh = st.number_input("Goluri Gazde", min_value=0, max_value=20, value=0, key="gh_in")
            with c2:
                ga = st.number_input("Goluri Oaspeți", min_value=0, max_value=20, value=0, key="ga_in")
            if st.button("💾 Salvează rezultatele în sesiune", type="primary", key="btn_save_rez"):
                df_upd = st.session_state["hist_df"].copy()
                for idx in df_m.index:
                    mkt = df_upd.loc[idx, "market"] if "market" in df_upd.columns else None
                    line = df_upd.loc[idx, "line"] if "line" in df_upd.columns else None
                    odds_r = df_upd.loc[idx, "odds"] if "odds" in df_upd.columns else None
                    rez = resolve_result(mkt, line, gh, ga)
                    profit = calc_profit(rez, odds_r)
                    df_upd.loc[idx, "rezultat"] = rez
                    df_upd.loc[idx, "profit"] = profit
                st.session_state["hist_df"] = df_upd
                st.success("✅ Rezultate salvate în sesiune! Descarcă Excel pentru a păstra modificările.")
                st.rerun()

        if "rezultat" in df.columns:
            df_s = df[df["rezultat"].isin(["W", "L", "V"])].copy()
            if not df_s.empty:
                df_s["profit"] = pd.to_numeric(df_s["profit"], errors="coerce")
                total = len(df_s)
                wins = (df_s["rezultat"] == "W").sum()
                losses = (df_s["rezultat"] == "L").sum()
                voids = (df_s["rezultat"] == "V").sum()
                profit_total = df_s["profit"].sum()
                settled = total - voids
                wr = wins / settled * 100 if settled > 0 else 0
                roi = profit_total / settled * 100 if settled > 0 else 0
                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("Total pariuri", total)
                k2.metric("Câștigate", wins)
                k3.metric("Pierdute", losses)
                k4.metric("Win Rate", f"{wr:.1f}%")
                k5.metric("ROI", f"{roi:.1f}%", delta=f"{profit_total:.2f}u")

                grp = df_s.groupby(["section", "market"]).agg(
                    Pariuri=("rezultat", "count"),
                    W=("rezultat", lambda x: (x == "W").sum()),
                    L=("rezultat", lambda x: (x == "L").sum()),
                    V=("rezultat", lambda x: (x == "V").sum()),
                    Profit=("profit", "sum")
                ).reset_index()
                grp["Win %"] = np.where((grp["Pariuri"] - grp["V"]) > 0, grp["W"] / (grp["Pariuri"] - grp["V"]) * 100, 0)
                grp["ROI %"] = np.where((grp["Pariuri"] - grp["V"]) > 0, grp["Profit"] / (grp["Pariuri"] - grp["V"]) * 100, 0)
                st.markdown("### Per piață")
                st.dataframe(grp.sort_values("ROI %", ascending=False), use_container_width=True, hide_index=True)

                if "ts" in df_s.columns:
                    try:
                        df_s["ts"] = pd.to_datetime(df_s["ts"])
                        df_s = df_s.sort_values("ts").reset_index(drop=True)
                        df_s["profit_cumulat"] = df_s["profit"].cumsum()
                        st.markdown("### Evoluție profit cumulat")
                        st.line_chart(df_s.set_index("ts")["profit_cumulat"], use_container_width=True)
                    except Exception:
                        pass