import io
import re
import queue
import time

import pandas as pd
import streamlit as st

from engine_core import ScraperEngine
from driver_factory import get_driver
from extractor import extract_table

# =============== HELPERS PENTRU EXCEL ===============

def clean_sheet_name(name: str) -> str:
    """Curăță numele de ligă pentru a fi folosit ca sheet name în Excel."""
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(name))
    return name[:31] if name else "Sheet"


def clean_team_name(name: str) -> str:
    if not isinstance(name, str):
        return name

    # tăiem partea cu statistici
    stop_words = [
        "League Pos.",
        "Statistics Overall",
        "Want to see",
    ]

    for stop in stop_words:
        if stop in name:
            name = name.split(stop)[0]

    name = name.strip()

    # eliminăm duplicarea completă
    words = name.split()
    half = len(words) // 2

    if len(words) % 2 == 0 and words[:half] == words[half:]:
        name = " ".join(words[:half])

    return name

def clean_sheet_name(name: str) -> str:
    """Curăță numele de ligă pentru a fi folosit ca sheet name în Excel."""
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(name))
    return name[:31] if name else "Sheet"


def normalize_league_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mapează coloanele brute (0,1,2,3...) pe schema:
    #, Team, MP, xG, xGA, xGD, GF, GA, xG vs Actual
    """
    df = df.copy()

    if df.shape[1] >= 10:
        df = df.iloc[:, :10]
        df.columns = [
            "#",
            "_drop",
            "Team",
            "MP",
            "xG",
            "xGA",
            "xGD",
            "GF",
            "GA",
            "xG vs Actual",
        ]
        df = df.drop(columns=["_drop"])
    elif df.shape[1] >= 9:
        df = df.iloc[:, :9]
        df.columns = [
            "#",
            "Team",
            "MP",
            "xG",
            "xGA",
            "xGD",
            "GF",
            "GA",
            "xG vs Actual",
        ]
    else:
        # lăsăm cum este dacă nu are suficiente coloane
        return df

    # tipuri numerice
    df["#"] = pd.to_numeric(df["#"], errors="coerce")
    df["Team"] = df["Team"].apply(clean_team_name)
    df["MP"] = pd.to_numeric(df["MP"], errors="coerce")

    for col in ["xG", "xGA", "xGD", "GF", "GA", "xG vs Actual"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # scoatem rânduri goale
    df = df.dropna(subset=["#", "Team"], how="any")
    df["#"] = df["#"].astype(int)

    return df


def build_excel_file(results_dict: dict) -> bytes:
    """
    Creează un fișier XLSX în memorie, cu un sheet per ligă.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for league_name, df in results_dict.items():
            clean_df = normalize_league_df(df)
            sheet_name = clean_sheet_name(league_name)

            clean_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # setăm lățimi de coloane pentru citibilitate
            ws = writer.sheets[sheet_name]
            widths = {
                "A": 6,   # #
                "B": 28,  # Team
                "C": 12,  # MP
                "D": 10,  # xG
                "E": 10,  # xGA
                "F": 10,  # xGD
                "G": 10,  # GF
                "H": 10,  # GA
                "I": 14,  # xG vs Actual
            }
            for col_letter, width in widths.items():
                if col_letter in ws.column_dimensions:
                    ws.column_dimensions[col_letter].width = width

    output.seek(0)
    return output.getvalue()


# =============== UI PRINCIPALĂ ===============

st.set_page_config(page_title="xG ENGINE", layout="wide")
st.title("⚽ xG SCRAPER FULL ENGINE")

# =============== STATE INIT ===============
if "logs" not in st.session_state:
    st.session_state.logs = []

if "log_queue" not in st.session_state:
    st.session_state.log_queue = queue.Queue()

if "progress" not in st.session_state:
    st.session_state.progress = 0.0

if "status" not in st.session_state:
    st.session_state.status = "idle"

if "results" not in st.session_state:
    st.session_state.results = {}

if "engine" not in st.session_state:
    st.session_state.engine = None

if "selected_leagues" not in st.session_state:
    st.session_state.selected_leagues = []

# =============== FILE UPLOAD ===============
uploaded_file = st.file_uploader("Upload leagues file (Excel/CSV)")

engine = st.session_state.engine

if uploaded_file:
    # ---- load data ----
    if uploaded_file.name.endswith(".csv"):
        df_links = pd.read_csv(uploaded_file)
    else:
        try:
            df_links = pd.read_excel(uploaded_file, sheet_name="Leagues")
        except Exception:
            df_links = pd.read_excel(uploaded_file)

    df_links.columns = df_links.columns.str.strip().str.lower()

    if "countryleague" in df_links.columns:
        league_col = "countryleague"
    elif "league" in df_links.columns:
        league_col = "league"
    else:
        st.error("Nu găsesc coloana 'league' sau 'countryleague' în fișier.")
        st.stop()

    if "xg_url" not in df_links.columns:
        st.error("Lipsește coloana obligatorie 'xg_url'.")
        st.stop()

    st.subheader("Preview")
    st.dataframe(df_links[[league_col, "xg_url"]].head())

    # =============== LEAGUES UI ===============
    all_leagues = df_links[league_col].astype(str).tolist()

    # curățăm selecția să fie subset valid din all_leagues
    st.session_state.selected_leagues = [
        x for x in st.session_state.selected_leagues if x in all_leagues
    ] or all_leagues.copy()

    col1, col2 = st.columns(2)

    with col1:
        if st.button("✅ Select all"):
            st.session_state.selected_leagues = all_leagues.copy()

    with col2:
        if st.button("❌ Select none"):
            st.session_state.selected_leagues = []

    selected = st.multiselect(
        "Select leagues",
        options=all_leagues,
        key="selected_leagues",
    )

    st.write(f"Selected: {len(selected)}")

    # =============== ENGINE INIT ===============
    if st.session_state.engine is None:
        st.session_state.engine = ScraperEngine(
            driver_factory=get_driver,
            extractor=extract_table,
            message_queue=st.session_state.log_queue,
        )

    engine = st.session_state.engine

    # =============== CONTROLS ===============
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        headless_mode = st.checkbox("Headless mode", value=False)

    with col2:
        if st.button("▶ START"):
            if not selected:
                st.warning("Nu ai selectat nicio ligă.")
            else:
                # reset UI state
                st.session_state.logs = []
                st.session_state.progress = 0.0
                st.session_state.status = "running"
                st.session_state.results = {}

                # golim coada veche
                while not st.session_state.log_queue.empty():
                    st.session_state.log_queue.get()

                df_selected = df_links[df_links[league_col].isin(selected)]
                pairs = list(
                    df_selected[[league_col, "xg_url"]]
                    .itertuples(index=False, name=None)
                )

                def driver_factory():
                    return get_driver(headless=headless_mode)

                st.session_state.engine = ScraperEngine(
                    driver_factory=driver_factory,
                    extractor=extract_table,
                    message_queue=st.session_state.log_queue,
                )

                engine = st.session_state.engine
                engine.load_queue(pairs)
                engine.start()

    with col3:
        if st.button("⛔ STOP"):
            if engine:
                engine.stop()
            st.session_state.status = "stopped"

    with col4:
        if st.button("🧹 CLEAR LOGS"):
            st.session_state.logs = []
            st.session_state.progress = 0.0
            st.session_state.status = "idle"
            st.session_state.results = {}
            st.session_state.engine = None

            while not st.session_state.log_queue.empty():
                st.session_state.log_queue.get()

            st.rerun()

    with col5:
        st.metric("Status", st.session_state.status.upper())
    # =============== QUEUE → UI SYNC ===============
    while not st.session_state.log_queue.empty():
        msg = st.session_state.log_queue.get()

        if isinstance(msg, tuple) and msg[0] == "__PROGRESS__":
            st.session_state.progress = float(msg[1])
        else:
            st.session_state.logs.append(str(msg))

    # =============== PROGRESS ===============
    st.progress(st.session_state.progress)

    # =============== STATUS UPDATE ===============
    if engine and engine.running:
        st.session_state.status = "running"
    elif st.session_state.progress >= 1.0 and st.session_state.status != "stopped":
        st.session_state.status = "done"

    st.write(f"Status: {st.session_state.status}")

    # =============== RESULTS ===============
    if engine and engine.results:
        st.session_state.results = engine.results

    st.subheader("Results")
    st.write(len(st.session_state.results), "leagues scraped")

    if st.session_state.results:
        excel_bytes = build_excel_file(st.session_state.results)

        st.download_button(
            "📥 Download XLSX",
            data=excel_bytes,
            file_name="xgDATA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # =============== LOGS ===============
    st.subheader("Logs")
    st.text("\n".join(st.session_state.logs[-200:]))
    current_engine = st.session_state.engine

    if (
        current_engine
        and current_engine.running
        and st.session_state.status == "running"
    ):
        time.sleep(1)
        st.rerun()

    if not current_engine.running:
        st.session_state.status = "Done"

else:
    st.info(
        "Încarcă un fișier cu coloanele 'league' sau 'countryleague' + 'xg_url'."
    )