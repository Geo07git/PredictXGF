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
import queue

RUN_MODE = os.getenv("RUN_MODE", "cloud").strip().lower()
ALLOW_SCRAPING = RUN_MODE == "local"
SCRAPING_IMPORT_ERROR = None

if ALLOW_SCRAPING:
    try:
        from driver_factory import get_driver
        from extractor import extract_table
    except Exception as e:
        ALLOW_SCRAPING = False
        SCRAPING_IMPORT_ERROR = str(e)

# ============== CONFIGURARE DE BAZĂ ==============

st.set_page_config(page_title="xG Poisson Scanner", layout="wide")
st.title("⚽ xG Poisson Scanner + Shortlist + Propuneri")
st.markdown(
    "App pentru ligile cu xG: selectezi liga/meciul sau scanezi toată liga, "
    "calculezi probabilități cu Poisson, EV, salvezi toate piețele și generezi propuneri fără conflicte majore."
)

HISTORY_PATH = "Predictii_xG_Poisson.xlsx"
MAX_GOALS = 7
DEFAULT_MIN_EV = 0.03
DEFAULT_MIN_PROB = 0.50
DEFAULT_MIN_EV_LEAN = 0.00
DEFAULT_MAX_PROPOSALS = 3

MATCH_TOTAL_LINES = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5]
TEAM_TOTAL_LINES = [0.5, 1.5, 2.5, 3.5]


# ============== LIVE SCRAPE HELPERS ==============
def clean_team_name_live(name: str):
    if not isinstance(name, str):
        return name
    stop_words = [
        "League Pos.",
        "Statistics Overall",
        "Want to see",
    ]
    for stop in stop_words:
        if stop in name:
            name = name.split(stop)[0]
    name = name.strip()
    words = name.split()
    half = len(words) // 2
    if len(words) % 2 == 0 and words[:half] == words[half:]:
        name = " ".join(words[:half])
    return name

def normalize_live_league_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if df.shape[1] >= 10:
        df = df.iloc[:, :10]
        df.columns = ["#", "_drop", "Team", "MP", "xG", "xGA", "xGD", "GF", "GA", "xG vs Actual"]
        df = df.drop(columns=["_drop"])
    elif df.shape[1] >= 9:
        df = df.iloc[:, :9]
        df.columns = ["#", "Team", "MP", "xG", "xGA", "xGD", "GF", "GA", "xG vs Actual"]
    else:
        return df
    df["#"] = pd.to_numeric(df["#"], errors="coerce")
    df["Team"] = df["Team"].apply(clean_team_name_live)
    df["MP"] = pd.to_numeric(df["MP"], errors="coerce")
    for col in ["xG", "xGA", "xGD", "GF", "GA", "xG vs Actual"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["#", "Team"], how="any")
    df["#"] = df["#"].astype(int)
    return df

def detect_catalog_league_col(df_links: pd.DataFrame) -> str:
    cols = [str(c).strip().lower() for c in df_links.columns]
    df_links.columns = cols
    if "countryleague" in cols:
        return "countryleague"
    if "league" in cols:
        return "league"
    raise ValueError("Nu găsesc coloana 'league' sau 'countryleague' în catalog.")

def _live_log(msg):
    try:
        st.session_state.setdefault("live_scrape_logs", []).append(str(msg))
    except Exception:
        pass

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_live_league_df(league_name: str, xg_url: str, headless: bool = False) -> pd.DataFrame:
    if not ALLOW_SCRAPING:
        msg = "Live scraping este dezactivat în acest mediu. Folosește Excel local."
        if SCRAPING_IMPORT_ERROR:
            msg += f" Detalii import: {SCRAPING_IMPORT_ERROR}"
        raise RuntimeError(msg)
    driver = get_driver(headless=headless)
    try:
        raw_df = extract_table(driver, xg_url, _live_log)
        return normalize_live_league_df(raw_df)
    finally:
        try:
            driver.quit()
        except Exception:
            pass


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


def extract_section(text: str, start_title: str, stop_titles: list[str]) -> str:
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
        re.compile(r"(?:sub|-)\s*(\d+(?:\.\d+)?)\s+([0-9]+(?:\.\d+)?)\s*(?:peste|\+)\s*\s+([0-9]+(?:\.\d+)?)", re.IGNORECASE),
        re.compile(r"(?:sub|-)\s*(\d+(?:\.\d+)?)\s+([0-9]+(?:\.\d+)?)\s*(?:peste|\+)\s*(\d+(?:\.\d+)?)\s+([0-9]+(?:\.\d+)?)", re.IGNORECASE),
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


def extract_decimal_odds(raw: str) -> list[float]:
    txt = normalize_offer_text(raw)
    txt = re.sub(r"\d+[.,]\d+\s*%", " ", txt)
    txt = re.sub(r"\d+\s*%", " ", txt)
    vals = re.findall(r"(?<!\d)(\d{1,2}\.\d{1,2})(?!\d)", txt)
    out = []
    for v in vals:
        try:
            out.append(float(v.replace(",", ".")))
        except Exception:
            pass
    return out


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


def bookmaker_offer_preview_df(raw_text: str) -> pd.DataFrame:
    parsed = parse_bookmaker_offer_skip_htft(raw_text)
    rows = [{"Piață": k, "Cotă extrasă": v} for k, v in parsed.items()]
    return pd.DataFrame(rows)


HIST_COLUMNS = [
    "ts", "league", "home", "away", "market", "section", "line",
    "prob", "push_prob", "fair_odds", "odds", "ev", "decision",
    "selected_for_play", "proposal_rank", "proposal_group",
    "lambda_home", "lambda_away", "notes", "rezultat", "profit"
]


def ensure_history_file(path: str):
    try:
        wb = load_workbook(path)
        if "Istoric" not in wb.sheetnames:
            ws = wb.create_sheet("Istoric")
            ws.append(HIST_COLUMNS)
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True)
            wb.save(path)
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "Istoric"
        ws.append(HIST_COLUMNS)
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = Font(bold=True)
        wb.save(path)


def load_history_df(path: str) -> pd.DataFrame:
    ensure_history_file(path)
    try:
        df = pd.read_excel(path, sheet_name="Istoric")
    except Exception:
        return pd.DataFrame(columns=HIST_COLUMNS)
    for col in HIST_COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df[HIST_COLUMNS].copy()


def save_history_df(path: str, df: pd.DataFrame):
    ensure_history_file(path)
    df = df.copy()
    for col in HIST_COLUMNS:
        if col not in df.columns:
            df[col] = None
    df = df[HIST_COLUMNS].reset_index(drop=True)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False, sheet_name="Istoric")
    wb = load_workbook(path)
    ws = wb["Istoric"]
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(path)


def append_history_row(path: str, row: dict):
    df = load_history_df(path)
    new_row = {col: row.get(col, None) for col in HIST_COLUMNS}
    if new_row.get("rezultat") in [None, ""]:
        new_row["rezultat"] = "pending"
    if "profit" not in new_row:
        new_row["profit"] = None
    new_df = pd.DataFrame([new_row], columns=HIST_COLUMNS)
    if df.empty:
        df = new_df.copy()
    else:
        df = pd.concat([df, new_df], ignore_index=True)
    save_history_df(path, df)


def append_history_rows(path: str, rows: list[dict]):
    for row in rows:
        append_history_row(path, row)


def history_to_excel_bytes(path: str) -> bytes:
    df = load_history_df(path)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Istoric")
        workbook = writer.book
        worksheet = writer.sheets["Istoric"]
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#2E86AB", "font_color": "white"})
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, max(15, len(str(col_name)) + 2))
    buf.seek(0)
    return buf.getvalue()


def get_history_count(path: str) -> int:
    return len(load_history_df(path))


def market_constraints_from_name(market_name: str):
    meta = MARKET_MAP.get(market_name, {})
    family = meta.get("family")
    side = meta.get("side")
    line = meta.get("line")
    team_side = meta.get("team_side")
    c = {
        "home_min": 0,
        "away_min": 0,
        "home_max": 99,
        "away_max": 99,
        "total_min": 0,
        "total_max": 99,
        "home_gt_away": False,
        "away_gt_home": False,
        "home_ge_away": False,
        "away_ge_home": False,
        "draw_only": False,
        "btts_yes": False,
        "btts_no": False,
        "home_zero": False,
        "away_zero": False,
    }
    if family == "1X2":
        if side == "home":
            c["home_gt_away"] = True
        elif side == "away":
            c["away_gt_home"] = True
        elif side == "draw":
            c["draw_only"] = True
    elif family == "DC":
        if side == "1X":
            c["home_ge_away"] = True
        elif side == "X2":
            c["away_ge_home"] = True
    elif family == "DNB":
        if side == "home":
            c["home_gt_away"] = True
        elif side == "away":
            c["away_gt_home"] = True
    elif family == "BTTS":
        if side == "yes":
            c["btts_yes"] = True
            c["home_min"] = 1
            c["away_min"] = 1
        else:
            c["btts_no"] = True
    elif family == "WTN":
        if side == "home":
            c["home_gt_away"] = True
            c["away_zero"] = True
            c["away_max"] = 0
        else:
            c["away_gt_home"] = True
            c["home_zero"] = True
            c["home_max"] = 0
    elif family == "MATCH_TOTAL":
        if side == "over":
            c["total_min"] = int(np.floor(line) + 1 if not float(line).is_integer() else line + 1)
        else:
            c["total_max"] = int(np.ceil(line) - 1 if not float(line).is_integer() else line - 1)
    elif family == "TEAM_TOTAL":
        if side == "over":
            threshold = int(np.floor(line) + 1 if not float(line).is_integer() else line + 1)
            if team_side == "home":
                c["home_min"] = threshold
            else:
                c["away_min"] = threshold
        else:
            threshold = int(np.ceil(line) - 1 if not float(line).is_integer() else line - 1)
            if team_side == "home":
                c["home_max"] = threshold
            else:
                c["away_max"] = threshold
    return c


def has_hard_conflict_names(market_a: str, market_b: str) -> bool:
    ca = market_constraints_from_name(market_a)
    cb = market_constraints_from_name(market_b)
    home_min = max(ca["home_min"], cb["home_min"])
    away_min = max(ca["away_min"], cb["away_min"])
    home_max = min(ca["home_max"], cb["home_max"])
    away_max = min(ca["away_max"], cb["away_max"])
    total_min = max(ca["total_min"], cb["total_min"], home_min + away_min)
    total_max = min(ca["total_max"], cb["total_max"], home_max + away_max)
    if home_min > home_max or away_min > away_max or total_min > total_max:
        return True
    if ca["draw_only"] and (cb["home_gt_away"] or cb["away_gt_home"]):
        return True
    if cb["draw_only"] and (ca["home_gt_away"] or ca["away_gt_home"]):
        return True
    if ca["home_gt_away"] and cb["away_ge_home"]:
        return True
    if cb["home_gt_away"] and ca["away_ge_home"]:
        return True
    if ca["away_gt_home"] and cb["home_ge_away"]:
        return True
    if cb["away_gt_home"] and ca["home_ge_away"]:
        return True
    if ca["btts_yes"] and (cb["btts_no"] or cb["home_zero"] or cb["away_zero"]):
        return True
    if cb["btts_yes"] and (ca["btts_no"] or ca["home_zero"] or ca["away_zero"]):
        return True
    return False


def soft_conflict_penalty_names(market_a: str, market_b: str) -> float:
    ma = MARKET_MAP.get(market_a, {})
    mb = MARKET_MAP.get(market_b, {})
    pair = [ma, mb]
    names = {market_a, market_b}
    penalty = 0.0
    if any(m.get("family") == "DC" and m.get("side") == "X2" for m in pair) and any(m.get("family") == "TEAM_TOTAL" and m.get("team_side") == "home" and m.get("side") == "over" and (m.get("line") or 0) >= 2.5 for m in pair):
        penalty += 0.12
    if any(m.get("family") == "DC" and m.get("side") == "1X" for m in pair) and any(m.get("family") == "TEAM_TOTAL" and m.get("team_side") == "away" and m.get("side") == "over" and (m.get("line") or 0) >= 2.5 for m in pair):
        penalty += 0.12
    if any(m.get("family") == "MATCH_TOTAL" and m.get("side") == "under" and (m.get("line") or 99) <= 2.5 for m in pair) and any(m.get("family") == "TEAM_TOTAL" and m.get("side") == "over" and (m.get("line") or 0) >= 1.5 for m in pair):
        penalty += 0.08
    if names in ({"BTTS No", "Over 3.5 FT"}, {"BTTS No", "Over 4.0 FT"}, {"BTTS No", "Over 4.5 FT"}):
        penalty += 0.25
    return penalty


def proposal_score_row(row: pd.Series) -> float:
    prob = row.get("prob") if pd.notna(row.get("prob")) else 0.0
    ev = row.get("ev") if pd.notna(row.get("ev")) else -1.0
    odds = row.get("odds") if pd.notna(row.get("odds")) else 1.0
    bonus = 0.03 if row.get("Verdict") == "BET" else 0.0
    odds_penalty = max(0.0, min((odds - 2.8) * 0.01, 0.05)) if odds is not None else 0.0
    return ev * 0.65 + prob * 0.35 + bonus - odds_penalty


def select_best_proposals(df_results: pd.DataFrame, max_picks: int = DEFAULT_MAX_PROPOSALS):
    if df_results.empty:
        return pd.DataFrame(), []
    cand = df_results[df_results["Verdict"].isin(["BET", "LEAN"])].copy()
    if cand.empty:
        return pd.DataFrame(), []
    cand["proposal_score"] = cand.apply(proposal_score_row, axis=1)
    cand = cand.sort_values(["proposal_score", "ev", "prob"], ascending=[False, False, False]).reset_index(drop=True)

    selected = []
    notes = []
    used_groups = set()

    def group_key(market_name: str):
        meta = MARKET_MAP.get(market_name, {})
        fam = meta.get("family")
        if fam in ["1X2", "DC", "DNB"]:
            return "RESULT"
        if fam == "MATCH_TOTAL":
            return "MATCH_TOTAL"
        if fam == "TEAM_TOTAL":
            return f"TEAM_TOTAL_{meta.get('team_side')}"
        return fam

    for _, row in cand.iterrows():
        market_name = row["Piață"]
        gk = group_key(market_name)
        if gk in used_groups:
            continue
        if any(has_hard_conflict_names(market_name, s["Piață"]) for s in selected):
            continue
        soft_pen = sum(soft_conflict_penalty_names(market_name, s["Piață"]) for s in selected)
        eff = row["proposal_score"] - soft_pen
        if eff < 0.02:
            continue
        rr = row.to_dict()
        rr["proposal_score_effective"] = eff
        selected.append(rr)
        used_groups.add(gk)
        if soft_pen > 0:
            notes.append(f"{market_name}: penalizare soft {soft_pen:.2f}")
        if len(selected) >= max_picks:
            break

    if not selected:
        return pd.DataFrame(), notes
    out = pd.DataFrame(selected).reset_index(drop=True)
    out["Rank"] = np.arange(1, len(out) + 1)
    out["Motiv"] = out.apply(lambda x: f"Prob {x['prob']*100:.1f}% | EV {x['ev']:.3f} | scor {x['proposal_score_effective']:.3f}", axis=1)
    return out, notes


def render_analysis_sections(df_results_full, home_full, away_full):
    results_map = {row["Piață"]: row for row in df_results_full.to_dict("records")}

    st.markdown("### Meci")
    st.write(f"**{home_full}** vs **{away_full}**")

    st.markdown("### Meci 1X2")
    cols = st.columns(3)
    for idx, mkt in enumerate(["1X2 - 1", "1X2 - X", "1X2 - 2"]):
        row = results_map.get(mkt, {})
        label = mkt.split(" - ")[-1]
        cols[idx].metric(label, row.get("Cotă fair", "-"), row.get("Prob %", "-"))

    st.markdown("### Meci - sansa dubla")
    cols = st.columns(3)
    for idx, mkt in enumerate(["Double Chance - 1X", "Double Chance - 12", "Double Chance - X2"]):
        row = results_map.get(mkt, {})
        label = mkt.split(" - ")[-1]
        cols[idx].metric(label, row.get("Cotă fair", "-"), row.get("Prob %", "-"))

    st.markdown("### Ambele marcheaza")
    cols = st.columns(2)
    for idx, mkt in enumerate(["BTTS Yes", "BTTS No"]):
        row = results_map.get(mkt, {})
        label = "Da" if mkt.endswith("Yes") else "Nu"
        cols[idx].metric(label, row.get("Cotă fair", "-"), row.get("Prob %", "-"))

    st.markdown("### Victorie fara egal")
    cols = st.columns(2)
    for idx, mkt in enumerate(["Draw No Bet - Home", "Draw No Bet - Away"]):
        row = results_map.get(mkt, {})
        label = home_full if mkt.endswith("Home") else away_full
        cols[idx].metric(label, row.get("Cotă fair", "-"), row.get("Prob %", "-"))

    st.markdown("### Win To Nil")
    cols = st.columns(2)
    for idx, mkt in enumerate(["Win To Nil - Home", "Win To Nil - Away"]):
        row = results_map.get(mkt, {})
        label = home_full if mkt.endswith("Home") else away_full
        cols[idx].metric(label, row.get("Cotă fair", "-"), row.get("Prob %", "-"))

    st.markdown("### Total goluri")
    total_rows = []
    for line in MATCH_TOTAL_LINES:
        under_key = f"Under {line} FT"
        over_key = f"Over {line} FT"
        total_rows.append({
            "Linie": line,
            "Sub": results_map.get(under_key, {}).get("Cotă fair", "-"),
            "Peste": results_map.get(over_key, {}).get("Cotă fair", "-"),
            "Prob Sub": results_map.get(under_key, {}).get("Prob %", "-"),
            "Prob Peste": results_map.get(over_key, {}).get("Prob %", "-"),
        })
    st.dataframe(pd.DataFrame(total_rows), use_container_width=True, hide_index=True)

    st.markdown("### Echipa 1 Total goluri")
    home_rows = []
    for line in TEAM_TOTAL_LINES:
        under_key = f"Home Under {line} Goals"
        over_key = f"Home Over {line} Goals"
        home_rows.append({
            "Linie": line,
            "Sub": results_map.get(under_key, {}).get("Cotă fair", "-"),
            "Peste": results_map.get(over_key, {}).get("Cotă fair", "-"),
            "Prob Sub": results_map.get(under_key, {}).get("Prob %", "-"),
            "Prob Peste": results_map.get(over_key, {}).get("Prob %", "-"),
        })
    st.dataframe(pd.DataFrame(home_rows), use_container_width=True, hide_index=True)

    st.markdown("### Echipa 2 Total goluri")
    away_rows = []
    for line in TEAM_TOTAL_LINES:
        under_key = f"Away Under {line} Goals"
        over_key = f"Away Over {line} Goals"
        away_rows.append({
            "Linie": line,
            "Sub": results_map.get(under_key, {}).get("Cotă fair", "-"),
            "Peste": results_map.get(over_key, {}).get("Cotă fair", "-"),
            "Prob Sub": results_map.get(under_key, {}).get("Prob %", "-"),
            "Prob Peste": results_map.get(over_key, {}).get("Prob %", "-"),
        })
    st.dataframe(pd.DataFrame(away_rows), use_container_width=True, hide_index=True)

xg_loaded = False
leagues = {}
league_names = []

st.sidebar.header("⚙️ Setări")
if ALLOW_SCRAPING:
    data_mode = st.sidebar.radio("Sursă date ligi", ["Excel local", "Live scrape din catalog URL"], index=0)
else:
    data_mode = "Excel local"
    st.sidebar.info("În Streamlit Cloud este activ doar modul Excel local.")
    if SCRAPING_IMPORT_ERROR:
        st.sidebar.caption(f"Motiv dezactivare scraping: {SCRAPING_IMPORT_ERROR}")

leagues = {}
league_names = []

if data_mode == "Excel local":
    uploaded_xg = st.sidebar.file_uploader("Încarcă xGDATA.xlsx", type=["xlsx"])

    DEFAULT_FILE = "xgDATA2105.xlsx"

    if uploaded_xg is not None:
        file_buffer = io.BytesIO(uploaded_xg.getvalue())
        leagues = load_xg_workbook(file_buffer)
        league_names = list(leagues.keys())
        st.sidebar.success("Folosesc fișierul încărcat.")
    else:
        if os.path.exists(DEFAULT_FILE):
            st.sidebar.success("Folosesc automat fișierul din repo.")
            leagues = load_xg_workbook(DEFAULT_FILE)
            league_names = list(leagues.keys())
        else:
            st.sidebar.info("🔼 Încarcă fișierul cu xG (Excel cu foaie per ligă).")
            leagues = {}
            league_names = []
else:
    uploaded_catalog = st.sidebar.file_uploader("Încarcă catalog ligi+URL (Excel/CSV)", type=["xlsx", "csv"], key="live_catalog_file")
    live_headless = st.sidebar.checkbox("Live scrape headless", value=False, key="live_headless_chk")
    if st.sidebar.button("♻️ Clear live cache", key="clear_live_cache_btn"):
        fetch_live_league_df.clear()
        st.sidebar.success("Cache live golit.")
    xg_loaded = False
    if uploaded_catalog is not None:
        try:
            if uploaded_catalog.name.lower().endswith('.csv'):
                df_catalog = pd.read_csv(uploaded_catalog)
            else:
                try:
                    df_catalog = pd.read_excel(uploaded_catalog, sheet_name="Leagues")
                except Exception:
                    df_catalog = pd.read_excel(uploaded_catalog)
            league_col_catalog = detect_catalog_league_col(df_catalog)
            if "xg_url" not in df_catalog.columns:
                st.sidebar.error("Catalogul trebuie să conțină coloana xg_url.")
            else:
                catalog_names = df_catalog[league_col_catalog].astype(str).tolist()
                if catalog_names:
                    selected_live_league = st.sidebar.selectbox("Liga live", catalog_names, key="selected_live_league")
                    selected_live_url = df_catalog.loc[df_catalog[league_col_catalog].astype(str) == str(selected_live_league), "xg_url"].iloc[0]
                    st.sidebar.caption(f"URL: {selected_live_url}")
                    if st.sidebar.button("📡 Încarcă liga live", key="load_live_league_btn"):
                        with st.spinner(f"Se încarcă live {selected_live_league}..."):
                            df_live = fetch_live_league_df(str(selected_live_league), str(selected_live_url), bool(live_headless))
                            leagues = {str(selected_live_league): df_live}
                            league_names = list(leagues.keys())
                            st.session_state["live_leagues_data"] = leagues
                            st.session_state["live_league_names"] = league_names
                            xg_loaded = True
                    if "live_leagues_data" in st.session_state and "live_league_names" in st.session_state:
                        leagues = st.session_state["live_leagues_data"]
                        league_names = st.session_state["live_league_names"]
                        xg_loaded = bool(league_names)
                else:
                    st.sidebar.warning("Catalogul este gol.")
        except Exception as e:
            st.sidebar.error(f"Eroare catalog/live scrape: {e}")
    else:
        st.sidebar.info("🔼 Încarcă un catalog cu coloanele 'league'/'countryleague' și 'xg_url'.")

col_min_ev, col_min_prob = st.sidebar.columns(2)
with col_min_ev:
    min_ev = st.number_input("Prag EV BET", value=DEFAULT_MIN_EV, format="%.3f")
with col_min_prob:
    min_prob = st.number_input("Prag Prob BET", value=DEFAULT_MIN_PROB, format="%.2f")
min_ev_lean = st.sidebar.number_input("Prag EV LEAN (>0)", value=DEFAULT_MIN_EV_LEAN, format="%.3f")
max_goals = st.sidebar.number_input("Max goals Poisson", value=MAX_GOALS, min_value=5, max_value=12, step=1)
max_proposals = st.sidebar.number_input("Max propuneri", value=DEFAULT_MAX_PROPOSALS, min_value=1, max_value=5, step=1)
st.sidebar.markdown("---")
st.sidebar.caption("Istoric salvat local în Predictii_xG_Poisson.xlsx")
st.sidebar.caption("RUN_MODE=" + RUN_MODE + " | Scraping activ: " + ("Da" if ALLOW_SCRAPING else "Nu"))

tab_single, tab_full, tab_prop, tab_multi, tab_scanner, tab_calib, tab_history = st.tabs(
    ["🎯 Single Match", "🔍 Full Analysis", "🧠 Propuneri", "⚡ Multi-Meci", "📊 League Scanner", "🎯 Calibrare Model", "📚 Istoric"]
)

with tab_single:
    if not xg_loaded:
        st.warning("⚠️ Încarcă fișierul xGDATA.xlsx din sidebar pentru a folosi acest tab.")
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
            st.session_state["last_result"] = {
                "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "league": league_sel,
                "home": home_team,
                "away": away_team,
                "market": market_sel,
                "section": MARKET_MAP[market_sel]["section"],
                "line": row["_line"],
                "prob": row["_prob"],
                "push_prob": row["_push"],
                "fair_odds": row["_fair"],
                "odds": odds_val,
                "ev": row["_ev"],
                "decision": row["Verdict"],
                "selected_for_play": True,
                "proposal_rank": 1,
                "proposal_group": "single_match",
                "lambda_home": float(lam_home),
                "lambda_away": float(lam_away),
                "notes": notes,
                "rezultat": "pending",
                "profit": None,
            }

        if "last_result" in st.session_state:
            r = st.session_state["last_result"]
            prob_text = fmt1pct(r["prob"] * 100 / 100) if r["prob"] is not None else "-"
            st.info(f"Rezultat pregătit: **{r['home']} vs {r['away']}** | Piață: {r['market']} | Prob: {prob_text}")
            if st.button("💾 Salvează în Istoric"):
                append_history_row(HISTORY_PATH, r)
                del st.session_state["last_result"]
                st.success(f"✅ Salvat în istoric! Total intrări: {get_history_count(HISTORY_PATH)}")

with tab_full:
    if not xg_loaded:
        st.warning("⚠️ Încarcă fișierul xGDATA.xlsx din sidebar pentru a folosi acest tab.")
    else:
        st.subheader("🔍 Full Match Analysis — toate piețele")
        st.markdown(
            "Selectezi meciul, lipești oferta bookmakerului, iar app-ul mapează automat cotele pe piețe."
            "Secțiunea «Pauza sau final» este ignorată, iar Win To Nil rămâne manual."
        )

        col_flg, col_fh, col_fa = st.columns(3)
        with col_flg:
            league_full = st.selectbox("Ligă", league_names, key="full_league")
            df_full_lg = leagues[league_full]
            strengths_full = build_team_strengths(df_full_lg)
            teams_full = strengths_full["teams"]
        with col_fh:
            home_full = st.selectbox("Gazde", teams_full, key="full_home")
        with col_fa:
            away_full = st.selectbox("Oaspeți", teams_full, key="full_away", index=min(1, len(teams_full) - 1))

        if home_full == away_full:
            st.warning("Alege echipe diferite.")
            st.stop()

        st.markdown("---")
        offer_text = st.text_area("Lipește oferta în ordinea bookmakerului", value="", height=320, key="full_offer_paste")

        markets_auto = [m for m in MARKETS if not m.startswith("Win To Nil")]
        markets_manual_only = ["Win To Nil - Home", "Win To Nil - Away"]
        parsed_offer_map = {}

        if "full_offer_last_applied" not in st.session_state:
            st.session_state["full_offer_last_applied"] = ""

        if offer_text.strip():
            parsed_offer_map = parse_bookmaker_offer_skip_htft(offer_text)
            if offer_text != st.session_state["full_offer_last_applied"]:
                all_markets_for_auto = markets_auto + markets_manual_only
                for idx, mkt_name in enumerate(all_markets_for_auto):
                    key_name = f"odds_full_{idx}"
                    auto_val = parsed_offer_map.get(mkt_name, None)
                    if mkt_name in markets_manual_only:
                        continue
                    st.session_state[key_name] = f"{auto_val:.2f}" if auto_val is not None else ""
                st.session_state["full_offer_last_applied"] = offer_text
                st.info("Secțiunea «Pauza sau final» a fost ignorată automat. Win To Nil se completează manual.")
        else:
            st.caption("Poți lăsa gol și introduce manual doar cotele dorite.")

        st.markdown("---")
        st.markdown("### Cote folosite la calcul")
        odds_inputs = {}
        all_input_markets = markets_auto + markets_manual_only
        with st.expander("Input cote / override manual", expanded=False):
            cols_odds = st.columns(4)
            for idx, mkt_name in enumerate(all_input_markets):
                placeholder = "Win To Nil manual" if mkt_name in markets_manual_only else "ex: 1.85"
                with cols_odds[idx % 4]:
                    val = st.text_input(mkt_name, key=f"odds_full_{idx}", placeholder=placeholder)
                    odds_inputs[mkt_name] = val

        notes_full = st.text_input("Note opțional", value="", key="full_notes")

        if st.button("🔍 Calculează toate piețele", type="primary"):
            lam_h, lam_a = get_lambdas(strengths_full, home_full, away_full)
            p_mat = poisson_matrix(lam_h, lam_a, max_goals=int(max_goals))
            results_full = []

            for mkt_name in all_input_markets:
                raw_odds = odds_inputs.get(mkt_name, "").strip()
                try:
                    odds_f = float(raw_odds.replace(",", ".")) if raw_odds else None
                except ValueError:
                    odds_f = None
                row = compute_market_row(mkt_name, p_mat, lam_h, lam_a, odds_f)
                results_full.append({
                    "Piață": row["Piață"],
                    "Secțiune": row["Secțiune"],
                    "Prob %": row["Prob %"],
                    "Cotă fair": row["Cotă fair"],
                    "Cotă BK": row["Cotă BK"],
                    "EV": row["EV"],
                    "Verdict": row["Verdict"],
                    "prob": row["_prob"],
                    "fair": row["_fair"],
                    "odds": row["_odds"],
                    "ev": row["_ev"],
                    "line": row["_line"],
                    "push_prob": row["_push"],
                    "family": row["_family"],
                    "team_side": row["_team_side"],
                    "side": row["_side"],
                })

            df_results_full = pd.DataFrame(results_full)
            proposals_df, proposal_notes = select_best_proposals(df_results_full, int(max_proposals))

            st.markdown(f"### Rezultate {home_full} vs {away_full}")
            col_lam1, col_lam2 = st.columns(2)
            col_lam1.metric(home_full, f"{lam_h:.2f}")
            col_lam2.metric(away_full, f"{lam_a:.2f}")

            render_analysis_sections(df_results_full, home_full, away_full)

            display_cols = ["Piață", "Secțiune", "Prob %", "Cotă fair", "Cotă BK", "EV", "Verdict"]
            with st.expander("Tabel complet piețe"):
                st.dataframe(df_results_full[display_cols].style.applymap(style_verdict, subset=["Verdict"]), use_container_width=True, hide_index=True)

            df_short_full = df_results_full[df_results_full["Verdict"].isin(["BET", "LEAN"])].copy()
            if not df_short_full.empty:
                st.markdown("### Shortlist BET / LEAN")
                st.dataframe(df_short_full[display_cols].style.applymap(style_verdict, subset=["Verdict"]), use_container_width=True, hide_index=True)

            if not proposals_df.empty:
                st.markdown("### 🧠 Propunere rapidă")
                show_prop = proposals_df[["Rank", "Piață", "Secțiune", "Prob %", "Cotă fair", "Cotă BK", "EV", "Verdict", "Motiv"]]
                st.dataframe(show_prop.style.applymap(style_verdict, subset=["Verdict"]), use_container_width=True, hide_index=True)
                if proposal_notes:
                    with st.expander("Note penalizare / overlap"):
                        for note in proposal_notes:
                            st.write(f"- {note}")

            with st.expander("Distribuția scorurilor"):
                render_score_matrix(p_mat, max_goals)

            normalized_rows_full = []
            for rr in results_full:
                rr2 = dict(rr)
                if "Piață" not in rr2 and "market" in rr2:
                    rr2["Piață"] = rr2.get("market")
                if "Secțiune" not in rr2 and "section" in rr2:
                    rr2["Secțiune"] = rr2.get("section")
                normalized_rows_full.append(rr2)

            st.session_state["full_results"] = {
                "rows": normalized_rows_full,
                "home": home_full,
                "away": away_full,
                "league": league_full,
                "lam_h": lam_h,
                "lam_a": lam_a,
                "notes": notes_full,
            }
            proposal_rows = proposals_df.to_dict("records") if not proposals_df.empty else []
            normalized_prop_rows = []
            for rr in proposal_rows:
                rr2 = dict(rr)
                if "Piață" not in rr2 and "market" in rr2:
                    rr2["Piață"] = rr2.get("market")
                if "Secțiune" not in rr2 and "section" in rr2:
                    rr2["Secțiune"] = rr2.get("section")
                normalized_prop_rows.append(rr2)

            st.session_state["proposal_results"] = {
                "rows": normalized_prop_rows,
                "notes": proposal_notes,
                "home": home_full,
                "away": away_full,
                "league": league_full,
                "lam_h": lam_h,
                "lam_a": lam_a,
                "notes_match": notes_full,
            }

        if "full_results" in st.session_state:
            fr = st.session_state["full_results"]
            st.markdown("---")
            st.markdown(f"### Salvare în istoric — {fr['home']} vs {fr['away']}")
            piete_disponibile = [r.get("Piață", r.get("market", "")) for r in fr["rows"] if r.get("Piață", r.get("market", ""))]
            piete_selectate = st.multiselect(
                "Selectează piețele de salvat",
                options=piete_disponibile,
                default=[r.get("Piață", r.get("market", "")) for r in fr["rows"] if r.get("Verdict") in ["BET", "LEAN"] and r.get("Piață", r.get("market", ""))],
                key="full_save_markets",
            )
            save_all_markets = st.checkbox("Salvează toate piețele calculate", value=True, key="save_all_markets_chk")
            mark_proposals = st.checkbox("Marchează propunerile ca selected_for_play", value=True, key="mark_proposals_chk")
            if st.button("💾 Salvează piețele în Istoric", key="full_save_btn"):
                if not save_all_markets and not piete_selectate:
                    st.warning("Selectează cel puțin o piață sau bifează salvarea tuturor piețelor.")
                else:
                    ts_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    rows_to_save = []
                    proposal_rank_map = {}
                    if "proposal_results" in st.session_state:
                        for rr in st.session_state["proposal_results"].get("rows", []):
                            rr_market = rr.get("Piață", rr.get("market"))
                            if rr_market:
                                proposal_rank_map[rr_market] = rr.get("Rank")
                    for r in fr["rows"]:
                        row_market = r.get("Piață", r.get("market"))
                        row_section = r.get("Secțiune", r.get("section"))
                        if save_all_markets or row_market in piete_selectate:
                            rows_to_save.append({
                                "ts": ts_now,
                                "league": fr["league"],
                                "home": fr["home"],
                                "away": fr["away"],
                                "market": row_market,
                                "section": row_section if row_section else MARKET_MAP.get(row_market, {}).get("section"),
                                "line": r["line"],
                                "prob": r["prob"],
                                "push_prob": r["push_prob"],
                                "fair_odds": r["fair"],
                                "odds": r["odds"],
                                "ev": r["ev"],
                                "decision": r["Verdict"],
                                "selected_for_play": bool(mark_proposals and row_market in proposal_rank_map),
                                "proposal_rank": proposal_rank_map.get(row_market, None),
                                "proposal_group": "full_analysis",
                                "lambda_home": fr["lam_h"],
                                "lambda_away": fr["lam_a"],
                                "notes": fr["notes"],
                                "rezultat": "pending",
                                "profit": None,
                            })
                    append_history_rows(HISTORY_PATH, rows_to_save)
                    st.success(f"✅ {len(rows_to_save)} piețe salvate! Total în istoric: {get_history_count(HISTORY_PATH)}")

with tab_prop:
    st.subheader("🧠 Propunerea / Propunerile")
    st.markdown("Aici vezi ce e cel mai corect de jucat: maximum 3 piețe, fără conflicte majore, pornind din BET/LEAN.")
    if "proposal_results" not in st.session_state:
        st.info("Rulează întâi Full Analysis pentru un meci.")
    else:
        pr = st.session_state["proposal_results"]
        st.write(f"**{pr['home']} vs {pr['away']}** | {pr['league']}")
        if not pr["rows"]:
            st.warning("Nu există propuneri curate pe baza pragurilor curente.")
        else:
            df_prop = pd.DataFrame(pr["rows"])
            show_cols = ["Rank", "Piață", "Secțiune", "Prob %", "Cotă fair", "Cotă BK", "EV", "Verdict", "Motiv"]
            st.dataframe(df_prop[show_cols].style.applymap(style_verdict, subset=["Verdict"]), use_container_width=True, hide_index=True)
            if pr.get("notes"):
                with st.expander("Note conflict / penalizare"):
                    for note in pr["notes"]:
                        st.write(f"- {note}")
            if st.button("💾 Salvează doar propunerile în Istoric", key="save_prop_only_btn"):
                ts_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                rows_to_save = []
                for r in pr["rows"]:
                    rows_to_save.append({
                        "ts": ts_now,
                        "league": pr["league"],
                        "home": pr["home"],
                        "away": pr["away"],
                        "market": r.get("Piață", r.get("market")),
                        "section": r.get("Secțiune", r.get("section")) or MARKET_MAP.get(r.get("Piață", r.get("market")), {}).get("section"),
                        "line": r["line"],
                        "prob": r["prob"],
                        "push_prob": r["push_prob"],
                        "fair_odds": r["fair"],
                        "odds": r["odds"],
                        "ev": r["ev"],
                        "decision": r["Verdict"],
                        "selected_for_play": True,
                        "proposal_rank": r.get("Rank"),
                        "proposal_group": "proposal_tab",
                        "lambda_home": pr["lam_h"],
                        "lambda_away": pr["lam_a"],
                        "notes": pr["notes_match"],
                        "rezultat": "pending",
                        "profit": None,
                    })
                append_history_rows(HISTORY_PATH, rows_to_save)
                st.success(f"✅ {len(rows_to_save)} propuneri salvate! Total în istoric: {get_history_count(HISTORY_PATH)}")

with tab_multi:
    if not xg_loaded:
        st.warning("⚠️ Încarcă fișierul xGDATA.xlsx din sidebar pentru a folosi acest tab.")
    else:
        st.subheader("⚡ Multi-Meci — Analiză rundă întreagă")
        league_multi = st.selectbox("Ligă", league_names, key="multi_league")
        df_multi_lg = leagues[league_multi]
        strengths_multi = build_team_strengths(df_multi_lg)
        teams_multi = strengths_multi["teams"]
        n_matches = st.number_input("Câte meciuri vrei să analizezi?", min_value=1, max_value=20, value=5, step=1, key="multi_n")
        markets_multi_sel = st.multiselect("Piețe de analizat:", options=MARKETS, default=["Over 1.5 FT", "Over 2.5 FT", "BTTS Yes", "1X2 - 1", "1X2 - 2"], key="multi_markets")
        if not markets_multi_sel:
            st.warning("Selectează cel puțin o piață.")
            st.stop()

        match_inputs = []
        for i in range(int(n_matches)):
            c1, c2, c3 = st.columns([2, 2, 2])
            with c1:
                h = st.selectbox(f"Gazde #{i+1}", teams_multi, key=f"multi_h_{i}")
            with c2:
                a = st.selectbox(f"Oaspeți #{i+1}", teams_multi, key=f"multi_a_{i}", index=min(1, len(teams_multi) - 1))
            with c3:
                odds_str = st.text_input(f"Cotă globală #{i+1} (opț.)", value="", key=f"multi_odds_{i}", placeholder="ex: 1.85")
            match_inputs.append({"home": h, "away": a, "odds_str": odds_str})

        if st.button("⚡ Calculează toată runda", type="primary"):
            all_rows = []
            for mi, match in enumerate(match_inputs):
                if match["home"] == match["away"]:
                    st.warning(f"Meciul #{mi+1}: echipe identice ({match['home']}) — ignorat.")
                    continue
                try:
                    odds_m = float(match["odds_str"].replace(",", ".")) if match["odds_str"].strip() else None
                except ValueError:
                    odds_m = None
                lam_h_m, lam_a_m = get_lambdas(strengths_multi, match["home"], match["away"])
                p_mat_m = poisson_matrix(lam_h_m, lam_a_m, max_goals=int(max_goals))
                for mkt in markets_multi_sel:
                    row = compute_market_row(mkt, p_mat_m, lam_h_m, lam_a_m, odds_m)
                    all_rows.append({
                        "Meci": f"{match['home']} vs {match['away']}",
                        "home": match["home"],
                        "away": match["away"],
                        "Piață": mkt,
                        "Secțiune": row["Secțiune"],
                        "λH": f"{lam_h_m:.2f}",
                        "λA": f"{lam_a_m:.2f}",
                        "Prob %": row["Prob %"],
                        "Cotă fair": row["Cotă fair"],
                        "Cotă BK": row["Cotă BK"],
                        "EV": row["EV"],
                        "Verdict": row["Verdict"],
                        "_prob": row["_prob"],
                        "_fair": row["_fair"],
                        "_odds": row["_odds"],
                        "_ev": row["_ev"],
                        "_line": row["_line"],
                        "_push": row["_push"],
                        "_lam_h": lam_h_m,
                        "_lam_a": lam_a_m,
                    })
            if not all_rows:
                st.error("Niciun meci valid de calculat.")
                st.stop()
            df_multi_res = pd.DataFrame(all_rows)
            display_cols = ["Meci", "Secțiune", "Piață", "λH", "λA", "Prob %", "Cotă fair", "Cotă BK", "EV", "Verdict"]
            st.dataframe(df_multi_res[display_cols].style.applymap(style_verdict, subset=["Verdict"]), use_container_width=True, hide_index=True)
            df_short_multi = df_multi_res[df_multi_res["Verdict"].isin(["BET", "LEAN"])].copy()
            if not df_short_multi.empty:
                st.markdown(f"### ✅ Shortlist global rundă ({len(df_short_multi)} pariuri)")
                st.dataframe(df_short_multi[display_cols].style.applymap(style_verdict, subset=["Verdict"]), use_container_width=True, hide_index=True)
            buf_multi = io.BytesIO()
            with pd.ExcelWriter(buf_multi, engine="xlsxwriter") as writer:
                df_multi_res[display_cols].to_excel(writer, index=False, sheet_name="Toate_pietele")
                if not df_short_multi.empty:
                    df_short_multi[display_cols].to_excel(writer, index=False, sheet_name="Shortlist")
            buf_multi.seek(0)
            st.download_button("📥 Descarcă analiza rundei (Excel)", data=buf_multi, file_name=f"Runda_{league_multi}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.session_state["multi_results"] = {"rows": all_rows, "league": league_multi}

        if "multi_results" in st.session_state:
            mr = st.session_state["multi_results"]
            optiuni = [f"{r['Meci']} | {r['Piață']} | {r['Verdict']}" for r in mr["rows"]]
            default_sel = [f"{r['Meci']} | {r['Piață']} | {r['Verdict']}" for r in mr["rows"] if r["Verdict"] in ("BET", "LEAN")]
            selectate = st.multiselect("Selectează ce salvezi:", options=optiuni, default=default_sel, key="multi_save_sel")
            if st.button("💾 Salvează selecția în Istoric", key="multi_save_btn"):
                ts_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                saved_count = 0
                for r in mr["rows"]:
                    key_r = f"{r['Meci']} | {r['Piață']} | {r['Verdict']}"
                    if key_r in selectate:
                        append_history_row(HISTORY_PATH, {
                            "ts": ts_now,
                            "league": mr["league"],
                            "home": r["home"],
                            "away": r["away"],
                            "market": r["Piață"],
                            "section": r["Secțiune"],
                            "line": r["_line"],
                            "prob": r["_prob"],
                            "push_prob": r["_push"],
                            "fair_odds": r["_fair"],
                            "odds": r["_odds"],
                            "ev": r["_ev"],
                            "decision": r["Verdict"],
                            "selected_for_play": r["Verdict"] in ("BET", "LEAN"),
                            "proposal_rank": None,
                            "proposal_group": "multi_match",
                            "lambda_home": r["_lam_h"],
                            "lambda_away": r["_lam_a"],
                            "notes": "",
                            "rezultat": "pending",
                            "profit": None,
                        })
                        saved_count += 1
                st.success(f"✅ {saved_count} pariuri salvate! Total în istoric: {get_history_count(HISTORY_PATH)}")
                del st.session_state["multi_results"]

with tab_scanner:
    if not xg_loaded:
        st.warning("⚠️ Încarcă fișierul xGDATA.xlsx din sidebar pentru a folosi acest tab.")
    else:
        st.subheader("📊 Scanner ligă (Poisson Shortlist)")
        col_lg2, col_mkt2 = st.columns(2)
        with col_lg2:
            league_scan = st.selectbox("Liga pentru scanare", league_names, key="scan_league")
            df_scan = leagues[league_scan]
            strengths_scan = build_team_strengths(df_scan)
            teams_scan = strengths_scan["teams"]
        with col_mkt2:
            market_scan = st.selectbox("Piață de scanat", MARKETS, index=1, key="scan_market")
        odds_global_str = st.text_input("Cota standard pentru scanner (opțional, ex: 1.85)", value="")
        if st.button("🚀 Rulează scanner"):
            try:
                odds_global = float(odds_global_str.replace(",", ".")) if odds_global_str.strip() else None
            except ValueError:
                odds_global = None
            rows = []
            pairs = list(itertools.permutations(teams_scan, 2))
            prog = st.progress(0.0)
            total_pairs = len(pairs)
            for i, (home, away) in enumerate(pairs, start=1):
                lam_home, lam_away = get_lambdas(strengths_scan, home, away)
                p_matrix = poisson_matrix(lam_home, lam_away, max_goals=int(max_goals))
                row = compute_market_row(market_scan, p_matrix, lam_home, lam_away, odds_global)
                rows.append({
                    "league": league_scan,
                    "home": home,
                    "away": away,
                    "section": row["Secțiune"],
                    "market": market_scan,
                    "line": row["_line"],
                    "prob": row["_prob"],
                    "push_prob": row["_push"],
                    "prob_%": row["Prob %"],
                    "fair_odds": row["_fair"],
                    "odds": odds_global,
                    "ev": row["_ev"],
                    "decision": row["Verdict"],
                    "lambda_home": lam_home,
                    "lambda_away": lam_away,
                })
                prog.progress(i / total_pairs)
            df_scan_res = pd.DataFrame(rows).sort_values(["decision", "ev", "prob"], ascending=[True, False, False])
            scan_display_cols = ["league", "home", "away", "section", "market", "line", "prob_%", "fair_odds", "odds", "ev", "decision", "lambda_home", "lambda_away"]
            st.dataframe(df_scan_res[scan_display_cols], use_container_width=True, hide_index=True)
            df_short = df_scan_res[df_scan_res["decision"].isin(["BET", "LEAN"])].copy()
            if not df_short.empty:
                st.markdown("### ✅ Shortlist (BET / LEAN)")
                st.dataframe(df_short[scan_display_cols], use_container_width=True, hide_index=True)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                df_scan_res.to_excel(writer, index=False, sheet_name="Scanner")
                if not df_short.empty:
                    df_short.to_excel(writer, index=False, sheet_name="Shortlist")
            buf.seek(0)
            st.download_button("📥 Descarcă rezultate scanner (Excel)", data=buf, file_name=f"Scanner_{league_scan}_{market_scan.replace(' ', '')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_calib:
    st.subheader("🎯 Calibrare & Validare Model Poisson")
    df_calib_raw = load_history_df(HISTORY_PATH)
    if df_calib_raw.empty:
        st.info("Nu există date în istoric. Salvează predicții și introduce rezultate reale mai întâi.")
    else:
        df_calib = df_calib_raw.copy()
        df_calib = df_calib[df_calib["rezultat"].isin(["W", "L", "V"])].copy()
        if df_calib.empty:
            st.info("Nu există încă rezultate W/L/V introduse. Mergi la tab-ul Istoric.")
        else:
            df_calib["prob"] = pd.to_numeric(df_calib["prob"], errors="coerce")
            df_calib["push_prob"] = pd.to_numeric(df_calib["push_prob"], errors="coerce")
            df_calib["odds"] = pd.to_numeric(df_calib["odds"], errors="coerce")
            df_calib["profit"] = pd.to_numeric(df_calib["profit"], errors="coerce")
            df_calib = df_calib.dropna(subset=["prob"])
            df_calib["outcome"] = (df_calib["rezultat"] == "W").astype(int)
            n_total = len(df_calib)
            st.markdown(f"**Pariuri analizate:** {n_total} (W/L/V)")
            if n_total < 10:
                st.warning("Sunt necesare cel puțin 10 rezultate pentru o calibrare relevantă statistic.")
            cal1, cal2, cal3, cal4 = st.tabs(["📊 Calibrare generală", "🪣 Bucket Analysis", "📈 Brier & Log Loss", "🔬 Per piață"])
            with cal1:
                avg_prob = df_calib["prob"].mean()
                actual_wr = df_calib["outcome"].mean()
                calib_error = avg_prob - actual_wr
                k1, k2, k3 = st.columns(3)
                k1.metric("Prob medie estimată", f"{avg_prob*100:.1f}%")
                k2.metric("Win rate real", f"{actual_wr*100:.1f}%")
                k3.metric("Eroare calibrare", f"{calib_error*100:.1f} pp")
                prob_bins = pd.cut(df_calib["prob"], bins=10)
                bin_counts = df_calib.groupby(prob_bins, observed=True)["outcome"].agg(["count", "mean"]).reset_index()
                bin_counts.columns = ["Interval prob", "Nr. pariuri", "Win rate real"]
                bin_counts["Prob medie bin"] = df_calib.groupby(prob_bins, observed=True)["prob"].mean().values
                bin_counts["Interval prob"] = bin_counts["Interval prob"].astype(str)
                bin_counts["Win rate real %"] = (bin_counts["Win rate real"] * 100).round(1)
                bin_counts["Prob medie %"] = (bin_counts["Prob medie bin"] * 100).round(1)
                bin_counts["Diferență %"] = (bin_counts["Prob medie bin"] - bin_counts["Win rate real"]).mul(100).round(1)
                st.dataframe(bin_counts[["Interval prob", "Nr. pariuri", "Prob medie %", "Win rate real %", "Diferență %"]], use_container_width=True, hide_index=True)
            with cal2:
                n_buckets = st.slider("Număr de grupe (buckets)", min_value=3, max_value=10, value=5, key="calib_buckets")
                df_calib["bucket"] = pd.qcut(df_calib["prob"], q=n_buckets, duplicates="drop")
                bucket_df = df_calib.groupby("bucket", observed=True).agg(Nr_pariuri=("outcome", "count"), Prob_medie=("prob", "mean"), Win_rate_real=("outcome", "mean")).reset_index()
                bucket_df["Bucket"] = bucket_df["bucket"].astype(str)
                bucket_df["Prob medie %"] = (bucket_df["Prob_medie"] * 100).round(1)
                bucket_df["Win rate real %"] = (bucket_df["Win_rate_real"] * 100).round(1)
                bucket_df["Gap %"] = (bucket_df["Prob_medie"] - bucket_df["Win_rate_real"]).mul(100).round(1)
                st.dataframe(bucket_df[["Bucket", "Nr_pariuri", "Prob medie %", "Win rate real %", "Gap %"]], use_container_width=True, hide_index=True)
                chart_cal = bucket_df[["Prob medie %", "Win rate real %"]].copy()
                chart_cal["Linie perfectă"] = chart_cal["Prob medie %"]
                chart_cal = chart_cal.set_index("Prob medie %")
                st.line_chart(chart_cal, use_container_width=True)
            with cal3:
                outcomes = df_calib["outcome"].values
                probs = df_calib["prob"].values
                brier = np.mean((probs - outcomes) ** 2)
                eps = 1e-7
                probs_clip = np.clip(probs, eps, 1 - eps)
                log_loss = -np.mean(outcomes * np.log(probs_clip) + (1 - outcomes) * np.log(1 - probs_clip))
                base_prob = outcomes.mean()
                brier_base = np.mean((np.full_like(probs, base_prob) - outcomes) ** 2)
                logloss_base = -np.mean(outcomes * np.log(base_prob + eps) + (1 - outcomes) * np.log(1 - base_prob + eps))
                brier_skill = 1 - brier / brier_base if brier_base > 0 else 0
                logloss_skill = 1 - log_loss / logloss_base if logloss_base > 0 else 0
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Brier Score", f"{brier:.4f}")
                m2.metric("Brier Skill Score", f"{brier_skill*100:.1f}%")
                m3.metric("Log Loss", f"{log_loss:.4f}")
                m4.metric("Log Loss Skill", f"{logloss_skill*100:.1f}%")
            with cal4:
                market_cal = df_calib.groupby(["section", "market"]).agg(Nr=("outcome", "count"), Prob_medie=("prob", "mean"), Win_rate=("outcome", "mean"), Profit=("profit", "sum")).reset_index()
                market_cal = market_cal[market_cal["Nr"] >= 3]
                if market_cal.empty:
                    st.info("Nicio piață nu are suficiente date (minim 3 rezultate per piață).")
                else:
                    market_cal["Gap %"] = (market_cal["Prob_medie"] - market_cal["Win_rate"]).mul(100).round(1)
                    market_cal["Prob medie %"] = (market_cal["Prob_medie"] * 100).round(1)
                    market_cal["Win rate %"] = (market_cal["Win_rate"] * 100).round(1)
                    market_cal = market_cal.sort_values("Gap %", key=abs)
                    st.dataframe(market_cal[["section", "market", "Nr", "Prob medie %", "Win rate %", "Gap %", "Profit"]].rename(columns={"section": "Secțiune", "market": "Piață"}), use_container_width=True, hide_index=True)

with tab_history:
    st.subheader("📚 Istoric predicții + 📈 Tracker rezultate")
    df = load_history_df(HISTORY_PATH)
    n = len(df)
    with st.expander("📂 Importă istoric dintr-un Excel salvat anterior"):
        uploaded_hist = st.file_uploader("Încarcă Predictii_xG_Poisson.xlsx", type=["xlsx"], key="upload_hist")
        if uploaded_hist is not None:
            try:
                df_imp = pd.read_excel(io.BytesIO(uploaded_hist.getvalue()), sheet_name="Istoric")
                for col in HIST_COLUMNS:
                    if col not in df_imp.columns:
                        df_imp[col] = None
                df_imp = df_imp[HIST_COLUMNS].reset_index(drop=True)
                save_history_df(HISTORY_PATH, df_imp)
                st.success(f"✅ Importat {len(df_imp)} intrări în fișierul local!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Eroare import: {e}")

    col_r, col_dl, col_del = st.columns([1, 1, 1])
    with col_r:
        if st.button("🔄 Reîncarcă"):
            st.rerun()
    with col_dl:
        if n > 0:
            buf = history_to_excel_bytes(HISTORY_PATH)
            st.download_button("📥 Descarcă Excel", data=buf, file_name="Predictii_xG_Poisson.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.button("📥 Descarcă Excel", disabled=True)
    with col_del:
        with st.expander("⚠️ Șterge tot istoricul"):
            confirm_del = st.text_input("Scrie CONFIRM pentru a șterge tot:", value="", key="del_conf")
            if st.button("🔥 Șterge tot", type="primary", key="btn_del_all"):
                if confirm_del.strip().upper() == "CONFIRM":
                    save_history_df(HISTORY_PATH, pd.DataFrame(columns=HIST_COLUMNS))
                    st.success("✅ Istoricul a fost șters complet!")
                    st.rerun()
                else:
                    st.error("Scrie CONFIRM pentru a confirma ștergerea.")

    st.markdown(f"### 📋 Total intrări: **{n}**")
    if n == 0:
        st.info("Nu există predicții salvate.")
    else:
        df_show = df.reset_index(drop=True)
        if "ts" in df_show.columns:
            try:
                df_show["ts"] = pd.to_datetime(df_show["ts"])
                df_show = df_show.sort_values("ts", ascending=False).reset_index(drop=True)
            except Exception:
                pass
        df_display = df_show.copy()
        history_order = ["ts", "league", "home", "away", "section", "market", "line", "prob", "push_prob", "fair_odds", "odds", "ev", "decision", "selected_for_play", "proposal_rank", "proposal_group", "rezultat", "profit", "lambda_home", "lambda_away", "notes"]
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
            st.dataframe(df_m[[c for c in ["market", "section", "line", "prob", "odds", "decision", "selected_for_play", "proposal_rank"] if c in df_m.columns]], use_container_width=True, hide_index=True)
            c1, c2 = st.columns(2)
            with c1:
                gh = st.number_input("Goluri Gazde", min_value=0, max_value=20, value=0, key="gh_in")
            with c2:
                ga = st.number_input("Goluri Oaspeți", min_value=0, max_value=20, value=0, key="ga_in")
            if st.button("💾 Salvează rezultatele", type="primary", key="btn_save_rez"):
                df_upd = df.copy()
                for idx in df_m.index:
                    mkt = df_upd.loc[idx, "market"] if "market" in df_upd.columns else None
                    line = df_upd.loc[idx, "line"] if "line" in df_upd.columns else None
                    odds = df_upd.loc[idx, "odds"] if "odds" in df_upd.columns else None
                    rez = resolve_result(mkt, line, gh, ga)
                    profit = calc_profit(rez, odds)
                    df_upd.loc[idx, "rezultat"] = rez
                    df_upd.loc[idx, "profit"] = profit
                save_history_df(HISTORY_PATH, df_upd)
                st.success("✅ Rezultate salvate în fișierul local!")
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
                grp = df_s.groupby(["section", "market"]).agg(Pariuri=("rezultat", "count"), W=("rezultat", lambda x: (x == "W").sum()), L=("rezultat", lambda x: (x == "L").sum()), V=("rezultat", lambda x: (x == "V").sum()), Profit=("profit", "sum")).reset_index()
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