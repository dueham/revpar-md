"""
revpar_app.py  —  RevPar MD Multi-Hotel Platform
=================================================
Run with:  streamlit run revpar_app.py

Requirements:
  pip install streamlit pandas plotly openpyxl pyxlsb reportlab

File structure:
  revpar_app.py       ← this file
  hotel_config.py     ← hotel registry + file detection
  data_loader.py      ← all data parsing logic
"""

import streamlit as st
import streamlit.components.v1 as components
import os
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import math

import hotel_config as cfg
import data_loader as dl

# ── Version guard: ensure data_loader.py has IHG support ──────────────────────
if not hasattr(dl, "IHG_HOTELS"):
    st.error(
        "⚠️ **data_loader.py is out of date.** "
        "The running data_loader.py does not have IHG hotel support. "
        "Please replace data_loader.py with the latest version and restart both Streamlit processes.",
        icon="🚨"
    )

# ══════════════════════════════════════════════════════════════════════════════
# TOKEN GATE  —  validates launcher session before anything else runs
# ══════════════════════════════════════════════════════════════════════════════

from auth_utils import authenticate, create_cal_token, validate_cal_token

# ── Cal-date URL params (survive WebSocket reconnect on calendar click) ────────
_params      = st.query_params
_cal_sel_url = _params.get("cal_sel", None)
_cal_hot_url = _params.get("hotel", None)
_cal_tok_url = _params.get("cal_token", None)

if _cal_sel_url:
    try:
        import datetime as _dt_mod
        _cal_sel_date = _dt_mod.datetime.strptime(_cal_sel_url, "%Y-%m-%d").date()
        if _cal_hot_url:
            st.session_state[f"cal_sel_{_cal_hot_url}"] = _cal_sel_date
            st.session_state[f"_cal_clicked_{_cal_hot_url}"] = True
            st.session_state["_pending_cal_hotel"] = _cal_hot_url
    except Exception:
        pass

# Session restore via cal_token (30-min TTL — survives page reload)
if _cal_tok_url and "revpar_user" not in st.session_state:
    try:
        _u = validate_cal_token(_cal_tok_url)
    except Exception:
        _u = None
    if _u:
        st.session_state.revpar_user           = _u
        st.session_state.revpar_allowed_hotels = _u["hotels"]
        st.session_state.revpar_cal_token      = _cal_tok_url

# ── Inline login gate (replaces two-app launcher system for AWS) ──────────────
if "revpar_user" not in st.session_state:
    import base64
    from hotel_config import download_file_bytes

    st.set_page_config(
        page_title="RevPar MD | Sign In",
        page_icon="🏨",
        layout="centered",
        initial_sidebar_state="collapsed",
    )

    # Load logo from S3
    _logo_b64 = None
    try:
        _logo_buf = download_file_bytes("photos/New_Revparmd_Logo.png")
        if _logo_buf:
            _logo_b64 = base64.b64encode(_logo_buf.read()).decode()
    except Exception:
        pass

    _logo_html = (
        f'<img src="data:image/png;base64,{_logo_b64}" style="width:220px;height:auto;" alt="RevPar MD"/>'
        if _logo_b64 else
        '<div style="font-size:32px;font-weight:700;color:#1e2d35;">RevPar MD</div>'
    )

    st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
*, *::before, *::after {{ box-sizing: border-box; }}
html, body, [class*="css"] {{ font-family: 'Outfit', sans-serif; margin: 0; padding: 0; }}
#MainMenu, footer, header {{ visibility: hidden; }}
.stApp {{ background: #eef2f3; min-height: 100vh; }}
.stApp::before {{
  content: '';
  position: fixed; inset: 0;
  background-image:
    linear-gradient(rgba(27,112,167,0.04) 1px, transparent 1px),
    linear-gradient(90deg, rgba(27,112,167,0.04) 1px, transparent 1px);
  background-size: 40px 40px;
  pointer-events: none; z-index: 0;
}}
.block-container {{ padding: 0 1rem 2rem !important; max-width: 960px !important; margin: 0 auto !important; }}
.login-outer {{ min-height:100vh; display:flex; align-items:center; justify-content:center; padding:40px 16px; position:relative; z-index:1; }}
.login-card {{ background:#ffffff; border:1px solid #c4cfd4; border-radius:20px; padding:48px 44px 40px; width:100%; max-width:420px; box-shadow:0 0 0 1px rgba(109,161,75,0.06),0 32px 64px rgba(0,0,0,0.5); }}
.logo-wrap {{ text-align:center; margin-bottom:32px; }}
.stTextInput > label {{ color:#3a5260 !important; font-size:12px !important; font-weight:500 !important; letter-spacing:0.06em !important; text-transform:uppercase !important; font-family:'DM Mono',monospace !important; margin-bottom:6px !important; }}
.stTextInput > div > div > input {{ background:#f5f8f9 !important; border:1px solid #b4c0c6 !important; border-radius:10px !important; color:#1e2d35 !important; font-family:'Outfit',sans-serif !important; font-size:15px !important; padding:12px 16px !important; }}
.stButton > button {{ width:100%; background:linear-gradient(135deg,#6A924D,#556848) !important; color:#fff !important; border:none !important; border-radius:10px !important; padding:13px !important; font-size:15px !important; font-weight:600 !important; font-family:'Outfit',sans-serif !important; cursor:pointer !important; }}
</style>
<div class="login-outer">
  <div class="login-card">
    <div class="logo-wrap">{_logo_html}</div>
    <div style="font-size:11px;font-weight:600;letter-spacing:0.14em;text-transform:uppercase;color:#6a8090;margin-bottom:20px;text-align:center;font-family:'DM Mono',monospace;">Revenue Intelligence Platform</div>
""", unsafe_allow_html=True)

    if "login_error" not in st.session_state:
        st.session_state.login_error = ""

    _uname = st.text_input("Username", key="_login_uname")
    _pw    = st.text_input("Password", type="password", key="_login_pw")

    if st.session_state.login_error:
        st.markdown(
            f'<div style="background:#fdf1f0;border:1px solid #e8c0bc;border-radius:8px;'
            f'padding:10px 14px;color:#a03020;font-size:13px;margin-bottom:12px;">'
            f'⚠ {st.session_state.login_error}</div>',
            unsafe_allow_html=True,
        )

    if st.button("Sign In", key="_login_btn"):
        _user = authenticate(_uname.strip(), _pw)
        if _user:
            st.session_state.revpar_user           = _user
            st.session_state.revpar_allowed_hotels = _user["hotels"]
            st.session_state.revpar_cal_token      = create_cal_token(_user)
            st.session_state.login_error           = ""
            st.rerun()
        else:
            st.session_state.login_error = "Invalid username or password."
            st.rerun()

    st.markdown("</div></div>", unsafe_allow_html=True)
    st.stop()

# Gate passed — allowed hotels for this session:
#   st.session_state.revpar_allowed_hotels  →  list of hotel names
#   st.session_state.revpar_user            →  dict: username, display_name, role
#
# The hotel picker below (visible_hotels) is already filtered automatically.
# Admin users receive ALL_HOTELS from auth_utils so they see every property.

_allowed_hotels = st.session_state.revpar_allowed_hotels

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG & GLOBAL STYLES
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="RevPar MD | Hotel Portfolio",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Force Streamlit light mode base theme ──────────────────────────────────
# Injected before any other st.markdown to establish the base color scheme
st.markdown("""
<style>
/* Force Streamlit base to light mode - overrides any dark theme detection */
html[data-theme="dark"] { filter: none !important; }
body { background-color: #eef2f3 !important; color: #1e2d35 !important; }
.stApp, #root, #stDecoration { background: #eef2f3 !important; }

/* Override Streamlit's CSS variables that drive its dark mode */
:root {
  --primary-color: #2E618D;
  --background-color: #eef2f3;
  --secondary-background-color: #e4eaec;
  --text-color: #1e2d35;
  --font: "Source Sans Pro", sans-serif;
}

/* Nuke Streamlit's internal dark theme — background only, NOT text color */
[class*="st-emotion-cache"] {
    background-color: transparent !important;
}

/* Tab radio: aggressive override for ALL Streamlit versions */
div[data-testid="stRadio"],
div[data-testid="stRadio"] > div,
div[data-testid="stRadio"] > label {
    background: transparent !important;
}
div[data-testid="stRadio"] label,
div[data-testid="stRadio"] label > div,
div[data-testid="stRadio"] label p {
    color: #6a8090 !important;
    background: transparent !important;
}
div[role="radiogroup"] {
    background: #eef2f3 !important;
}
div[role="radiogroup"] label p,
div[role="radiogroup"] label span,
div[role="radiogroup"] label div p {
    color: #6a8090 !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 11px !important;
    text-transform: uppercase !important;
    letter-spacing: 0.07em !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>/* v3 — single corporate theme */
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&family=DM+Mono:wght@400;500&display=swap');

/* ── CSS VARIABLE THEME SYSTEM ─────────────────────────────────────────────
   Single corporate theme — muted, professional, light gray-blue base.
   Palette derived from Oliver Companies brand colors.
────────────────────────────────────────────────────────────────────────── */
:root {
  /* backgrounds */
  --bg:        #eef2f3;   --card:      #ffffff;   --card2:     #e4eaec;
  /* headers / panel surfaces */
  --hdr:       #314B63;   --hdr2:      #dce3e6;
  /* borders */
  --border:    #c4cfd4;   --border2:   #b4c0c6;
  /* text */
  --txt-pri:   #1e2d35;   --txt-sec:   #3a5260;   --txt-mut:   #4e6878;
  --txt-dim:   #6a8090;   --txt-label: #5a7080;   --txt-faint: #9ab0ba;
  /* snapshot table */
  --today-bg:  #d4e4f0;   --wknd-bg:   #eaeef0;
  --snap-th:   #314B63;   --snap-td:   #1e2d35;   --snap-td2:  #d8e2e6;
  --tbl-hdr:   #314B63;   --tbl-alt:   #e8eeF0;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Hide default Streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 0.5rem 1rem 0 1rem !important; max-width: 100% !important; }

/* ── Force Streamlit to light mode ── */
/* These override Streamlit's internal dark-theme variables */
:root, [data-theme="dark"], [data-theme="light"] {
    --background-color: #eef2f3 !important;
    --secondary-background-color: #e4eaec !important;
    --text-color: #1e2d35 !important;
    --font: 'DM Sans', sans-serif !important;
}

/* Force all Streamlit markdown wrappers to light bg */
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
[data-testid="stMainBlockContainer"],
section[data-testid="stSidebar"],
.main .block-container,
div[data-testid="column"],
div[data-testid="stVerticalBlock"],
div[data-testid="stHorizontalBlock"] {
    background-color: transparent !important;
}

/* Streamlit's own dark theme injects styles — override them all */
.stApp { background: #eef2f3 !important; }
[data-testid="stMarkdownContainer"] { color: #1e2d35; }
[data-testid="stMarkdownContainer"] > p,
[data-testid="stMarkdownContainer"] > span {
    color: #1e2d35;
}
/* Do NOT set color on div — allows inner divs with custom colors to work */
[data-testid="stMarkdownContainer"] div { background: transparent; }

/* CRITICAL: Override Streamlit's column/block wrappers that add dark bg */
div.stColumn > div,
div[data-testid="stColumn"] > div[data-testid="stVerticalBlockBorderWrapper"],
div[data-testid="stVerticalBlockBorderWrapper"] {
    background: transparent !important;
    border: none !important;
}

/* ── Theme toggle — fixed top-right ── */
div[data-testid="stButton"]:has(button[key="_rm_toggle_theme_landing"]),
div[data-testid="stButton"]:has(button[key="_rm_toggle_theme_dash"]) {
    position: fixed !important;
    top: 16px !important;
    right: 24px !important;
    z-index: 99999 !important;
    width: auto !important;
}
div[data-testid="stButton"]:has(button[key="_rm_toggle_theme_landing"]) button,
div[data-testid="stButton"]:has(button[key="_rm_toggle_theme_dash"]) button {
    background: rgba(255,255,255,0.92) !important;
    border: 1px solid #b4c0c6 !important;
    border-radius: 40px !important;
    padding: 7px 20px 7px 16px !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 11px !important;
    letter-spacing: 0.08em !important;
    color: #3a5260 !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.10) !important;
    white-space: nowrap !important;
    min-width: 110px !important;
    backdrop-filter: blur(8px) !important;
    transition: border-color 0.2s, box-shadow 0.2s !important;
}
div[data-testid="stButton"]:has(button[key="_rm_toggle_theme_landing"]) button:hover,
div[data-testid="stButton"]:has(button[key="_rm_toggle_theme_dash"]) button:hover {
    border-color: #2E618D !important;
    box-shadow: 0 2px 12px rgba(46,97,141,0.18) !important;
    color: #1e2d35 !important;
}

/* ── Platform header ── */
.platform-header {
    background: var(--card);
    border-bottom: 1px solid var(--border);
    padding: 10px 40px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.platform-title {
    display: flex;
    align-items: center;
}
.platform-title img {
    height: 120px;
    width: auto;
    display: block;
}
.platform-meta {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    color: var(--txt-dim);
    letter-spacing: 0.1em;
}

/* ── Landing page ── */
.landing-wrap {
    padding: 40px 40px 60px;
}
/* Tighten the gap between hotel cards */
.landing-wrap [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
    padding-left: 6px !important;
    padding-right: 6px !important;
}
.landing-headline {
    font-family: 'Syne', sans-serif;
    font-size: clamp(28px, 4vw, 48px);
    font-weight: 800;
    color: var(--txt-pri);
    margin-bottom: 6px;
    line-height: 1.1;
}
.landing-sub {
    font-size: 14px;
    color: var(--txt-dim);
    margin-bottom: 40px;
    font-weight: 300;
    letter-spacing: 0.03em;
}

/* ── Hotel cards ── */
.hotel-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 0;
    cursor: pointer;
    transition: all 0.2s ease;
    position: relative;
    overflow: hidden;
    aspect-ratio: 1 / 1;
    max-width: 240px;
    margin: 0 auto;
    display: flex;
    flex-direction: column;
    text-decoration: none;
}
.hotel-card:hover {
    border-color: #2E618D;
    transform: translateY(-2px);
    box-shadow: 0 8px 32px rgba(46,97,141,0.18);
}
.hotel-card-accent {
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    z-index: 2;
}
.hotel-card-brand {
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--txt-label);
    margin-bottom: 6px;
    padding: 0 14px;
}
.hotel-card-name {
    font-family: 'Syne', sans-serif;
    font-size: 13px;
    font-weight: 700;
    color: var(--txt-pri);
    margin-bottom: 0;
    line-height: 1.25;
    padding: 0 14px;
}
.hotel-card-stat {
    display: flex;
    align-items: baseline;
    gap: 6px;
    padding: 8px 14px 4px;
}
.hotel-card-stat-value {
    font-family: 'Syne', sans-serif;
    font-size: 26px;
    font-weight: 800;
    color: #2E618D;
    line-height: 1;
}
.hotel-card-stat-label {
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--txt-label);
}
.hotel-card-event {
    font-size: 10px;
    color: #7a5020;
    padding: 2px 14px 0;
    font-style: italic;
}
.hotel-card-rooms {
    position: absolute;
    bottom: 10px;
    right: 12px;
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    color: var(--txt-faint);
}
.hotel-card-files {
    position: absolute;
    bottom: 10px;
    left: 14px;
    display: flex;
    gap: 3px;
}
.file-dot {
    width: 5px; height: 5px;
    border-radius: 50%;
    display: inline-block;
}
.hotel-card-photo {
    width: 100%;
    height: 110px;
    overflow: hidden;
    flex-shrink: 0;
}
.hotel-card-photo img {
    width: 100%;
    height: 100%;
    object-fit: cover;
    object-position: center;
    display: block;
}
.hotel-card-body {
    padding: 10px 0 28px;
    flex: 1;
    display: flex;
    flex-direction: column;
    justify-content: flex-start;
}

/* ── Card click wrapper ── */
.card-click-wrap {
    display: contents;
}
/* ── Anchor wrapper for clickable hotel card ── */
a.hotel-card-link {
    text-decoration: none !important;
    color: inherit !important;
    display: block;
    max-width: 240px;
    margin: 0 auto;
}
a.hotel-card-link:hover .hotel-card {
    border-color: #2E618D !important;
    transform: translateY(-2px);
    box-shadow: 0 8px 32px rgba(46,97,141,0.18);
}

/* ── Coming soon card ── */
.hotel-card-coming {
    background: var(--card2);
    border: 1px dashed var(--border);
    border-radius: 14px;
    padding: 24px;
    min-height: 200px;
    position: relative;
    opacity: 0.7;
}
.coming-badge {
    position: absolute;
    top: 16px;
    right: 16px;
    background: var(--border);
    color: var(--txt-dim);
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.15em;
    padding: 4px 10px;
    border-radius: 100px;
    text-transform: uppercase;
}

/* ── Dashboard header ── */
.dash-header {
    background: var(--card);
    border-bottom: 1px solid var(--border);
    padding: 16px 40px;
    display: flex;
    align-items: center;
    gap: 16px;
}
.back-btn {
    background: var(--border);
    color: var(--txt-pri);
    border: none;
    padding: 6px 14px;
    border-radius: 6px;
    font-size: 13px;
    cursor: pointer;
    font-family: 'DM Sans', sans-serif;
}
.dash-hotel-name {
    font-family: 'Syne', sans-serif;
    font-size: 18px;
    font-weight: 700;
    color: var(--txt-pri);
}
.dash-hotel-sub {
    font-size: 12px;
    color: var(--txt-dim);
}
.dash-last-refresh {
    font-family: 'DM Mono', monospace;
    font-size: 10px;
    color: var(--txt-dim);
    margin-left: auto;
}

/* ── KPI cards ── */
.kpi-grid {
    display: flex;
    gap: 12px;
    padding: 20px 40px;
    flex-wrap: wrap;
}
.kpi-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 16px 20px;
    min-width: 140px;
    flex: 1;
}
.kpi-label {
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: var(--txt-dim);
    margin-bottom: 8px;
}
.kpi-value {
    font-family: 'Syne', sans-serif;
    font-size: 26px;
    font-weight: 800;
    color: var(--txt-pri);
    line-height: 1;
}
.kpi-delta {
    font-size: 11px;
    margin-top: 4px;
}
.kpi-delta.pos { color: #6A924D; }
.kpi-delta.neg { color: #a03020; }
.kpi-delta.neu { color: var(--txt-dim); }

/* ── Tab nav ── */
.tab-nav-wrap {
    position: relative;
}
.tab-nav-wrap::after {
    content: '';
    position: absolute;
    top: 0; right: 0; bottom: 1px;
    width: 60px;
    pointer-events: none;
    background: linear-gradient(to right, transparent, var(--bg));
}
.tab-nav {
    display: flex;
    gap: 2px;
    padding: 0 40px;
    border-bottom: 1px solid var(--border);
    background: var(--bg);
    overflow-x: auto;
    scrollbar-width: none;
}
.tab-nav::-webkit-scrollbar { display: none; }
.tab-btn {
    padding: 12px 20px;
    font-size: 12px;
    font-family: 'DM Mono', monospace;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--txt-dim);
    border: none;
    background: none;
    cursor: pointer;
    border-bottom: 2px solid transparent;
    white-space: nowrap;
    transition: all 0.15s;
}
.tab-btn:hover { color: var(--txt-pri); }
.tab-btn.active {
    color: #2E618D;
    border-bottom-color: #2E618D;
}

/* ── Content area ── */
.content-area { padding: 24px clamp(16px, 2.5vw, 40px); }

/* ── Section header ── */
.section-head {
    font-family: 'Syne', sans-serif;
    font-size: 18px;
    font-weight: 700;
    color: var(--txt-pri);
    margin-bottom: 4px;
}
.section-sub {
    font-size: 12px;
    color: var(--txt-dim);
    margin-bottom: 20px;
}

/* ── Alert / insight box ── */
.insight-box {
    background: rgba(46,97,141,0.06);
    border: 1px solid rgba(46,97,141,0.22);
    border-left: 3px solid #2E618D;
    border-radius: 8px;
    padding: 12px 16px;
    font-size: 13px;
    color: var(--txt-sec);
    margin-bottom: 16px;
    line-height: 1.5;
}
.alert-box {
    background: rgba(180,80,30,0.06);
    border: 1px solid rgba(180,80,30,0.20);
    border-left: 3px solid #b44820;
    border-radius: 8px;
    padding: 12px 16px;
    font-size: 13px;
    color: var(--txt-sec);
    margin-bottom: 16px;
    line-height: 1.5;
}

/* ═══════════════════════════════════════
   SNAPSHOT TAB
═══════════════════════════════════════ */
.snap-wrap {
    display: grid;
    grid-template-columns: 230px 1fr 200px;
    gap: 10px;
    align-items: stretch;
    margin-bottom: 12px;
}
.snap-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
    font-family: 'DM Sans', sans-serif;
}
.snap-table th {
    background: var(--snap-th);
    color: var(--txt-mut);
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 5px 8px;
    text-align: center;
    border-bottom: 1px solid var(--border);
}
.snap-table th:first-child { text-align: left; }
.snap-table td {
    padding: 4px 8px;
    text-align: center;
    color: var(--snap-td);
    border-bottom: 1px solid var(--snap-td2);
    white-space: nowrap;
}
.snap-table td:first-child { text-align: left; color: var(--txt-pri); }
.snap-table tr:hover td { background: rgba(46,97,141,0.05); }
.snap-table tr.snap-highlight td { background: rgba(46,97,141,0.10); color: #1e2d35; font-weight: 500; }
.snap-table tr.snap-weekend td { color: #6a4820; }
.snap-table tr.snap-avg td { background: var(--card2); color: var(--txt-mut); font-style: italic; border-top: 1px solid #c4cfd4; }
.snap-table .snap-header-row th {
    background: #314B63 !important;
    color: #ffffff !important;
    font-size: 11px;
    padding: 7px 8px;
    text-align: center;
    letter-spacing: 0.05em;
}
.snap-section-title {
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    color: #ffffff !important;
    background: #314B63 !important;
    padding: 6px 10px;
    border-radius: 6px 6px 0 0;
    margin-bottom: 0;
    display: block;
}
.snap-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 8px;
    overflow: hidden;
    height: 100%;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
}
.snap-card .snap-table-wrap {
    flex: 1;
    overflow: auto;
}
.snap-cal-dow {
    font-family: 'DM Mono', monospace;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-align: center;
}
.snap-cal-date {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    opacity: 0.85;
    text-align: center;
}
.snap-pickup-row {
    display: grid;
    align-items: center;
    background: var(--hdr2);
    border-top: 1px solid var(--border);
    font-size: 11px;
    font-family: 'DM Mono', monospace;
}
.snap-pickup-label {
    background: #314B63;
    color: #ffffff;
    font-size: 9px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    padding: 5px 8px;
    text-align: center;
    border-right: 1px solid #c4cfd4;
}
.snap-pickup-val {
    text-align: center;
    color: var(--txt-sec);
    padding: 4px 2px;
    border-right: 1px solid var(--border2);
}
.snap-pickup-val.pos { color: #6A924D; }
.snap-pickup-val.neg { color: #a03020; }
.snap-rate-row {
    display: grid;
    background: var(--hdr2);
    border-top: 1px solid var(--border);
    font-size: 11px;
    font-family: 'DM Mono', monospace;
}
.snap-rate-val {
    text-align: center;
    color: #2E618D;
    padding: 4px 2px;
    border-right: 1px solid #c4cfd4;
    font-size: 11px;
}
.snap-rate-label {
    background: #314B63;
    color: #c8d8e4;
    font-size: 9px;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 5px 6px;
    text-align: center;
    border-right: 1px solid #c4cfd4;
}
.snap-bottom-wrap {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 10px;
    margin-top: 10px;
}
.snap-comp-grid {
    display: grid;
    gap: 0;
}
.snap-comp-hotel {
    text-align: center;
    padding: 5px 6px;
    border-right: 1px solid var(--border);
    font-size: 11px;
    color: var(--txt-sec);
}
.snap-comp-hotel-name {
    font-size: 9px;
    color: var(--txt-mut);
    font-family: 'DM Mono', monospace;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
.snap-comp-rate {
    font-size: 14px;
    color: #2E618D;
    font-family: 'Syne', sans-serif;
    font-weight: 700;
}
.snap-variance-pos { color: #6A924D; font-weight: 600; }
.snap-variance-neg { color: #a03020; font-weight: 600; }

/* ── Streamlit native widget overrides for corporate theme ─────────────── */

/* ══ TAB NAV: st.radio horizontal — maximum specificity overrides ══ */
/* Target the radiogroup container at every possible DOM depth */
div[role="radiogroup"],
div[data-testid="stRadio"] div[role="radiogroup"],
div[data-testid="stRadio"] > div > div[role="radiogroup"] {
    gap: 0 !important;
    border-bottom: 2px solid #c4cfd4 !important;
    padding: 0 !important;
    background: #eef2f3 !important;
    flex-wrap: nowrap !important;
    overflow-x: auto !important;
    scrollbar-width: none !important;
    margin: 0 !important;
}
div[role="radiogroup"]::-webkit-scrollbar { display: none !important; }

/* Every label in the radiogroup */
div[role="radiogroup"] label {
    padding: 10px 18px !important;
    margin: 0 !important;
    border-radius: 0 !important;
    border-bottom: 3px solid transparent !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 11px !important;
    letter-spacing: 0.07em !important;
    text-transform: uppercase !important;
    color: #4a6070 !important;
    background: transparent !important;
    cursor: pointer !important;
    white-space: nowrap !important;
    transition: color 0.15s, border-color 0.15s !important;
    display: flex !important;
    align-items: center !important;
}
div[role="radiogroup"] label:hover { color: #1e2d35 !important; }

/* Force ALL text inside labels to the right color */
div[role="radiogroup"] label *,
div[role="radiogroup"] label p,
div[role="radiogroup"] label span,
div[role="radiogroup"] label div {
    color: #4a6070 !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 11px !important;
    text-transform: uppercase !important;
    letter-spacing: 0.07em !important;
    background: transparent !important;
}

/* Hide the radio circle dot completely */
div[role="radiogroup"] label > div:first-child { display: none !important; }
div[role="radiogroup"] input[type="radio"] {
    position: absolute !important;
    opacity: 0 !important;
    width: 0 !important; height: 0 !important;
}

/* Selected (active) tab — blue underline */
div[role="radiogroup"] label:has(input:checked) {
    color: #2E618D !important;
    border-bottom-color: #2E618D !important;
    font-weight: 700 !important;
}
div[role="radiogroup"] label:has(input:checked) *,
div[role="radiogroup"] label:has(input:checked) p,
div[role="radiogroup"] label:has(input:checked) span {
    color: #2E618D !important;
    font-weight: 700 !important;
}

/* Force snap section titles and table headers to always show white on navy */
.snap-section-title,
[data-testid="stMarkdownContainer"] .snap-section-title,
[data-testid="stHtml"] .snap-section-title {
    color: #ffffff !important;
    background: #314B63 !important;
}
.snap-table .snap-header-row th,
[data-testid="stMarkdownContainer"] .snap-table .snap-header-row th {
    background: #314B63 !important;
    color: #ffffff !important;
}

/* Enforce inline color styles on spans — allows red/teal change deltas to show */
[data-testid="stMarkdownContainer"] span[style*="color:#cc2200"],
[data-testid="stMarkdownContainer"] span[style*="color: #cc2200"] {
    color: #cc2200 !important;
}
[data-testid="stMarkdownContainer"] span[style*="color:#2E618D"],
[data-testid="stMarkdownContainer"] span[style*="color: #2E618D"] {
    color: #2E618D !important;
}
[data-testid="stMarkdownContainer"] span[style*="color:#6A924D"],
[data-testid="stMarkdownContainer"] span[style*="color: #6A924D"] {
    color: #6A924D !important;
}
/* Enforce all inline colors — must override any cascade */
[data-testid="stMarkdownContainer"] span[style*="color:#cc2200"],
[data-testid="stMarkdownContainer"] span[style*="color: #cc2200"] {
    color: #cc2200 !important;
}
[data-testid="stMarkdownContainer"] span[style*="color:#2E618D"],
[data-testid="stMarkdownContainer"] span[style*="color: #2E618D"] {
    color: #2E618D !important;
}
[data-testid="stMarkdownContainer"] span[style*="color:#3a5260"],
[data-testid="stMarkdownContainer"] span[style*="color: #3a5260"] {
    color: #3a5260 !important;
}

/* SRP Pace total row — target by data attribute for maximum specificity */
[data-testid="stMarkdownContainer"] td[data-tot="1"],
[data-testid="stMarkdownContainer"] td[data-tot="1"] b,
[data-testid="stMarkdownContainer"] td[data-tot="1"] * {
    color: #ffffff !important;
}

/* Force inline background styles in st.markdown HTML to not be overridden */
/* This targets any HTML div rendered inside st.markdown() */
[data-testid="stMarkdownContainer"] div[style*="background:#314B63"],
[data-testid="stMarkdownContainer"] div[style*="background: #314B63"],
[data-testid="stHtml"] div[style*="background:#314B63"],
[data-testid="stHtml"] div[style*="background: #314B63"] {
    background: #314B63 !important;
}
[data-testid="stMarkdownContainer"] div[style*="background:#2E618D"],
[data-testid="stMarkdownContainer"] div[style*="background: #2E618D"] {
    background: #2E618D !important;
}
[data-testid="stMarkdownContainer"] *[style*="color:#ffffff"],
[data-testid="stMarkdownContainer"] *[style*="color: #ffffff"],
[data-testid="stMarkdownContainer"] div[style*="#ffffff"],
[data-testid="stMarkdownContainer"] *[style*="color:#ffffff !important"],
[data-testid="stMarkdownContainer"] span[style*="color:#ffffff"] {
    color: #ffffff !important;
}
body [data-testid="stMarkdownContainer"] div[style*="background:#314B63"] span,
.stApp [data-testid="stMarkdownContainer"] div[style*="background:#314B63"] * {
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
}
/* Protect event/group/note dot colors */
[data-testid="stMarkdownContainer"] span[style*="color:#e74c3c"] {
    color: #e74c3c !important;
    -webkit-text-fill-color: #e74c3c !important;
}
[data-testid="stMarkdownContainer"] span[style*="color:#f1c40f"] {
    color: #f1c40f !important;
    -webkit-text-fill-color: #f1c40f !important;
}
[data-testid="stMarkdownContainer"] span[style*="color:#888"] {
    color: #888888 !important;
    -webkit-text-fill-color: #888888 !important;
}
[data-testid="stMarkdownContainer"] *[style*="background:#ffffff"],
[data-testid="stMarkdownContainer"] *[style*="background: #ffffff"],
[data-testid="stMarkdownContainer"] *[style*="background:#f5f8f9"],
[data-testid="stMarkdownContainer"] *[style*="background:#eef2f3"],
[data-testid="stMarkdownContainer"] *[style*="background:#e4eaed"],
[data-testid="stMarkdownContainer"] *[style*="background:#e8eef1"],
[data-testid="stMarkdownContainer"] *[style*="background:#eaeef1"] {
    background: inherit !important;
}

/* Monthly Intelligence toggle buttons + all primary/secondary buttons */
div[data-testid="stButton"] > button[kind="primary"],
div[data-testid="stButton"] > button[data-testid="baseButton-primary"],
[data-testid="stButton"] button[kind="primary"] {
    color: #ffffff !important;
    background-color: #2E618D !important;
}
div[data-testid="stButton"] > button[kind="primary"] p,
div[data-testid="stButton"] > button[kind="primary"] span,
div[data-testid="stButton"] > button[data-testid="baseButton-primary"] p {
    color: #ffffff !important;
}
div[data-testid="stButton"] > button[kind="secondary"],
div[data-testid="stButton"] > button[data-testid="baseButton-secondary"] {
    color: #1e2d35 !important;
    background: #ffffff !important;
}
div[data-testid="stButton"] > button[kind="secondary"] p,
div[data-testid="stButton"] > button[kind="secondary"] span {
    color: #1e2d35 !important;
}

/* ← Portfolio and Refresh Data buttons */
div[data-testid="stButton"] > button {
    background: #ffffff !important;
    color: #1e2d35 !important;
    border: 1px solid #c4cfd4 !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    padding: 8px 16px !important;
    transition: background 0.15s, border-color 0.15s !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06) !important;
}
div[data-testid="stButton"] > button:hover {
    background: #f0f4f6 !important;
    border-color: #2E618D !important;
    color: #1e2d35 !important;
}
div[data-testid="stButton"] > button[kind="primary"] {
    background: #2E618D !important;
    color: #ffffff !important;
    border-color: #2E618D !important;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
    background: #245070 !important;
    border-color: #245070 !important;
}

/* Streamlit expanders */
[data-testid="stExpander"] summary {
    background: #f5f8f9 !important;
    color: var(--txt-pri) !important;
    border: 1px solid var(--border) !important;
}
[data-testid="stExpander"] summary:hover {
    background: #edf1f3 !important;
}

/* Streamlit selectbox / multiselect */
div[data-baseweb="select"] > div {
    background: #ffffff !important;
    border-color: #c4cfd4 !important;
    color: #1e2d35 !important;
}

/* Streamlit text inputs */
div[data-baseweb="input"] > div,
div[data-baseweb="textarea"] > div {
    background: #ffffff !important;
    border-color: #c4cfd4 !important;
}
div[data-baseweb="input"] input,
div[data-baseweb="textarea"] textarea {
    color: #1e2d35 !important;
    background: #ffffff !important;
}

/* Streamlit number input */
div[data-testid="stNumberInput"] input {
    color: #1e2d35 !important;
    background: #ffffff !important;
}

/* Streamlit markdown / p tags */
.stMarkdown p, .stMarkdown li, .stMarkdown span {
    color: var(--txt-pri) !important;
}

/* Streamlit dataframe */
[data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th {
    background: #ffffff !important;
    color: #1e2d35 !important;
}

/* Streamlit spinner */
[data-testid="stSpinner"] > div {
    color: #3a5260 !important;
}

/* Streamlit warning / info / error boxes */
div[data-testid="stAlert"] {
    background: #f5f8f9 !important;
    color: #1e2d35 !important;
    border-color: #c4cfd4 !important;
}

/* Sidebar / collapse */
[data-testid="stSidebar"] {
    background: #eef2f3 !important;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CHART HELPERS
# ══════════════════════════════════════════════════════════════════════════════

BG      = "#eef2f3"
CARD    = "#ffffff"
BORDER  = "#c4cfd4"
TEAL    = "#2E618D"   # primary accent: OTB, positive, active
PURPLE  = "#556848"   # secondary accent: forecast, comparison lines
ORANGE  = "#b44820"   # alert / warning
GOLD    = "#6A924D"   # positive delta / budget beat
RED     = "#a03020"   # negative / below threshold

def chart_layout():
    """Return Plotly layout dict — single corporate theme."""
    return dict(
        template="plotly_white",
        plot_bgcolor="#f5f8f9", paper_bgcolor="#ffffff",
        font=dict(family="DM Sans", color="#3a5260", size=11),
        margin=dict(t=40, b=40, l=40, r=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(size=11, color="#3a5260")),
        xaxis=dict(gridcolor="#dce5e8", showgrid=True, color="#4e6878", linecolor="#c4cfd4"),
        yaxis=dict(gridcolor="#dce5e8", color="#4e6878", linecolor="#c4cfd4"),
        height=340,
    )

CHART_LAYOUT = None  # kept for grep-safety; all code uses chart_layout()


def fig_line(df, x, ys, colors=None, title="", y_fmt=None):
    fig = go.Figure()
    colors = colors or [TEAL, PURPLE, ORANGE, GOLD]
    for i, (y, label) in enumerate(ys):
        fig.add_trace(go.Scatter(
            x=df[x], y=df[y], name=label,
            mode="lines+markers",
            line=dict(color=colors[i % len(colors)], width=2),
            marker=dict(size=4),
        ))
    fig.update_layout(**chart_layout(), title=dict(text=title, font=dict(size=13, color="#1e2d35")))
    if y_fmt == "pct":
        fig.update_yaxes(ticksuffix="%")
    elif y_fmt == "dollar":
        fig.update_yaxes(tickprefix="$")
    return fig


def fig_bar(df, x, ys, colors=None, title="", barmode="group", y_fmt=None):
    fig = go.Figure()
    colors = colors or [TEAL, PURPLE, ORANGE]
    for i, (y, label) in enumerate(ys):
        fig.add_trace(go.Bar(
            x=df[x], y=df[y], name=label,
            marker_color=colors[i % len(colors)],
        ))
    fig.update_layout(**chart_layout(), barmode=barmode,
                      title=dict(text=title, font=dict(size=13, color="#1e2d35")))
    if y_fmt == "pct":
        fig.update_yaxes(ticksuffix="%")
    elif y_fmt == "dollar":
        fig.update_yaxes(tickprefix="$")
    return fig


def fig_bar_index(df, x, y_mine, y_comp, y_idx, title, y_fmt=None):
    """Grouped bar with index score as secondary axis line."""
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=df[x], y=df[y_mine], name="My Property",
                         marker_color=TEAL), secondary_y=False)
    fig.add_trace(go.Bar(x=df[x], y=df[y_comp], name="Comp Set",
                         marker_color=PURPLE), secondary_y=False)
    IDX_COLOR = "#e6b800"   # yellow for index line
    fig.add_trace(go.Scatter(x=df[x], y=df[y_idx], name="Index",
                             mode="lines+markers",
                             line=dict(color=IDX_COLOR, width=2),
                             marker=dict(size=5, color=IDX_COLOR)), secondary_y=True)
    fig.add_hline(y=100, line_dash="dot", line_color="#1e2d35", line_width=1.5,
                  secondary_y=True)
    layout = chart_layout()
    fig.update_layout(**layout, barmode="group",
                      title=dict(text=title, font=dict(size=13, color="#1e2d35")))
    TICK_COL = "#1e2d35"  # dark for readable tick labels
    fig.update_xaxes(tickfont=dict(color=TICK_COL, size=11))
    fig.update_yaxes(secondary_y=False,
                     tickfont=dict(color=TICK_COL, size=11),
                     gridcolor=BORDER)
    fig.update_yaxes(secondary_y=True, gridcolor=BORDER,
                     tickcolor=IDX_COLOR,
                     tickfont=dict(color=IDX_COLOR, size=11),
                     title_text="Index",
                     title_font=dict(color=IDX_COLOR),
                     range=[60, 200])
    if y_fmt == "pct":
        fig.update_yaxes(secondary_y=False, ticksuffix="%",
                         tickfont=dict(color=TICK_COL, size=11))
    elif y_fmt == "dollar":
        fig.update_yaxes(secondary_y=False, tickprefix="$",
                         tickfont=dict(color=TICK_COL, size=11))
    return fig


def kpi_card(label, value, delta=None, suffix="", prefix=""):
    delta_html = ""
    if delta is not None:
        cls = "pos" if delta > 0 else ("neg" if delta < 0 else "neu")
        arrow = "▲" if delta > 0 else ("▼" if delta < 0 else "—")
        delta_html = f'<div class="kpi-delta {cls}">{arrow} {abs(delta):.1f}% vs LY</div>'
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{prefix}{value}{suffix}</div>
        {delta_html}
    </div>"""


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════

if "hotel_id"         not in st.session_state: st.session_state.hotel_id         = None
if "active_tab"       not in st.session_state: st.session_state.active_tab       = "Biweekly Snapshot"
if "hotel_data"       not in st.session_state: st.session_state.hotel_data       = {}
if "hotel_tab_memory" not in st.session_state: st.session_state.hotel_tab_memory = {}
if "srp_drill"   not in st.session_state: st.session_state.srp_drill   = None
# ── Single corporate theme — no light/dark toggle ───────────────────────────
if "rm_theme" not in st.session_state: st.session_state.rm_theme = "corporate"

# ── Restore hotel navigation from calendar date click (cal_sel URL param) ───
# cal_sel + hotel were stashed in session_state before the auth gate.
# Now that auth has passed, navigate to the correct hotel + Dashboard tab.
_pending_cal_hotel = st.session_state.pop("_pending_cal_hotel", None)
if _pending_cal_hotel and st.session_state.hotel_id is None:
    _matched_cal = next((h for h in cfg.HOTELS if h["id"] == _pending_cal_hotel and h.get("active", True)), None)
    if _matched_cal:
        st.session_state.hotel_id   = _pending_cal_hotel
        st.session_state.active_tab = "Calendar View"
        st.session_state.hotel_data = {}
        st.query_params.clear()


def theme_toggle_btn(key):
    """No-op: single theme, toggle removed."""
    pass

TABS = ["Biweekly Snapshot", "Calendar View", "Monthly Performance", "Segment Analysis", "Booking Pace by SRP", "Groups", "Rate Insights", "STR",
        "Demand Nights", "Market Events", "Call Recap"]


# ══════════════════════════════════════════════════════════════════════════════
# PLATFORM HEADER (always visible)
# ══════════════════════════════════════════════════════════════════════════════

now_str = datetime.now().strftime("%A, %B %d %Y")
_clock_html = f"""
<div class="platform-header">
    <div class="platform-title"><img src="data:image/png;base64,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" alt="RevPar MD" /></div>
    <div class="platform-meta" id="rm-live-clock">{now_str}&nbsp;&middot;&nbsp;--:-- --</div>
</div>
"""
st.markdown(_clock_html, unsafe_allow_html=True)

# ── Live clock: inject via components.html so JS executes reliably ──
components.html("""
<script>
(function() {
  function updateClock() {
    // Walk up through iframes to find the element in the parent document
    var el = null;
    try { el = window.parent.document.getElementById('rm-live-clock'); } catch(e) {}
    if (!el) { try { el = window.top.document.getElementById('rm-live-clock'); } catch(e) {} }
    if (!el) return;
    var now = new Date();
    var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var dow = days[now.getDay()];
    var mon = months[now.getMonth()];
    var dd  = String(now.getDate()).padStart(2,'0');
    var yr  = now.getFullYear();
    var hh  = now.getHours();
    var mm  = String(now.getMinutes()).padStart(2,'0');
    var ampm = hh >= 12 ? 'PM' : 'AM';
    hh = hh % 12 || 12;
    el.textContent = dow + ', ' + mon + ' ' + dd + ' ' + yr + '  \u00b7  ' + hh + ':' + mm + ' ' + ampm;
  }
  updateClock();
  setInterval(updateClock, 15000);
})();
</script>
""", height=0)


# ══════════════════════════════════════════════════════════════════════════════
# LANDING PAGE
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def cached_tonight(hotel_id, year_path_str, total_rooms, file_mtime: str = ""):
    """Cache tonight's OTB per hotel. _mtime busts cache when file changes."""
    try:
        if hotel_id in dl.IHG_HOTELS:
            year_df = dl.load_ihg_year(year_path_str)
        else:
            year_df = dl.load_year(year_path_str)
        return dl.get_tonight_otb(year_df, total_rooms)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def cached_landing_stats(hotel_id, year_path_str, budget_path_str, total_rooms, file_mtime: str = ""):
    """Compute card stats: tonight forecast occ% and current-month forecast vs budget revenue delta."""
    try:
        today = pd.Timestamp(datetime.now().date())

        # IHG: run full pipeline to get Forecast_Rev populated via DSS scaling
        if hotel_id in dl.IHG_HOTELS:
            import hotel_config as _cfg
            hotel = next((h for h in _cfg.HOTELS if h["id"] == hotel_id), None)
            if hotel:
                _files = _cfg.detect_files(hotel)
                _data  = dl.load_all(_files, total_rooms, hotel_id=hotel_id)
                year_df = _data.get("year")
            else:
                year_df = dl.load_ihg_year(year_path_str)
        else:
            year_df = dl.load_year(year_path_str)

        if year_df is None or year_df.empty:
            return {}

        # Tonight forecast occ%
        tonight_row = year_df[year_df["Date"] == today]
        if tonight_row.empty:
            tonight_row = year_df[year_df["Date"] == today + pd.Timedelta(days=1)]
        fcst_occ = None
        if not tonight_row.empty:
            r = tonight_row.iloc[0]
            fcst = r.get("Forecast_Rooms") or r.get("OTB") or 0
            cap  = r.get("Capacity") or total_rooms
            fcst_occ = round(fcst / cap * 100, 0) if cap else None

        # Current-month forecast vs budget revenue
        mo    = today.to_period("M")
        mo_df = year_df[year_df["Date"].dt.to_period("M") == mo]

        if "Revenue_Forecast" in mo_df.columns and mo_df["Revenue_Forecast"].notna().any():
            fcst_rev = float(mo_df["Revenue_Forecast"].sum())
        elif "Forecast_Rev" in mo_df.columns and mo_df["Forecast_Rev"].notna().any():
            fut_rev  = float(mo_df["Forecast_Rev"].fillna(0).sum())
            past_rev = float(mo_df.loc[mo_df["Forecast_Rev"].isna(), "Revenue_OTB"].fillna(0).sum())                        if "Revenue_OTB" in mo_df.columns else 0.0
            fcst_rev = fut_rev + past_rev
        elif "Revenue_OTB" in mo_df.columns and mo_df["Revenue_OTB"].notna().any():
            fcst_rev = float(mo_df["Revenue_OTB"].sum())
        else:
            fcst_rev = None

        bud_rev = None
        if budget_path_str and budget_path_str != "None":
            try:
                bud_df = dl.load_ihg_budget(budget_path_str) if hotel_id in dl.IHG_HOTELS                          else dl.load_budget(budget_path_str)
                bud_mo = bud_df[bud_df["Date"].dt.to_period("M") == mo]
                if not bud_mo.empty and "Revenue" in bud_mo.columns:
                    bud_rev = float(bud_mo["Revenue"].sum())
            except Exception:
                pass

        mo_delta = (fcst_rev - bud_rev) if (fcst_rev and bud_rev) else None
        mo_label = today.strftime("%b") + " Forecast vs Budget"

        return {
            "fcst_occ":  fcst_occ,
            "mo_delta":  mo_delta,
            "mo_label":  mo_label,
            "fcst_rev":  fcst_rev,
            "bud_rev":   bud_rev,
        }
    except Exception:
        return {}


def render_landing():
    theme_toggle_btn("theme_landing")
    st.markdown("""
    <div class="landing-wrap">
        <div class="landing-headline">Your Hotel Portfolio</div>
        <div class="landing-sub">Select a property to open its revenue dashboard</div>
    </div>
    """, unsafe_allow_html=True)

    _role = st.session_state.revpar_user.get("role", "")
    if _role == "admin":
        hotels = cfg.HOTELS
    else:
        hotels = [
            h for h in cfg.HOTELS
            if h.get("display_name") in _allowed_hotels
            or h.get("id") in _allowed_hotels
            or h.get("name") in _allowed_hotels
        ]
    # Sort alphabetically by city (first part of subtitle before the comma)
    hotels = sorted(hotels, key=lambda h: h.get("subtitle", "").split(",")[0].strip().lower())
    cols_per_row = 4
    rows_needed  = math.ceil(len(hotels) / cols_per_row)

    for row_i in range(rows_needed):
        row_hotels = hotels[row_i * cols_per_row : (row_i + 1) * cols_per_row]
        cols = st.columns(cols_per_row, gap="small")

        for col_i, hotel in enumerate(row_hotels):
            with cols[col_i]:
                status = cfg.file_status(hotel)
                files  = status["files"]

                if not hotel.get("active", True):
                    # ── Coming soon card — photo + details, no metrics ──
                    import base64 as _b64, os as _os
                    _cs_photo_path = hotel.get("photo_path", "")
                    _cs_photo_html = ""
                    if _cs_photo_path and _os.path.exists(str(_cs_photo_path)):
                        _cs_ext  = str(_cs_photo_path).rsplit(".", 1)[-1].lower()
                        _cs_mime = "image/jpeg" if _cs_ext in ("jpg", "jpeg") else "image/png"
                        try:
                            with open(_cs_photo_path, "rb") as _csf:
                                _cs_b64 = _b64.b64encode(_csf.read()).decode()
                            _cs_photo_html = f'<img src="data:{_cs_mime};base64,{_cs_b64}" style="width:100%;height:160px;object-fit:cover;display:block;filter:brightness(0.75);" />'
                        except Exception:
                            _cs_photo_html = ""
                    if not _cs_photo_html:
                        _cs_bc = hotel.get("brand_color", "#003087")
                        _cs_photo_html = f'<div style="width:100%;height:160px;background:linear-gradient(135deg,{_cs_bc}22,#eef2f3);display:flex;align-items:center;justify-content:center;"><span style="font-size:32px;">🏨</span></div>'
                    _cs_bc  = hotel.get("brand_color",  "#003087")
                    _cs_ac  = hotel.get("accent_color", "#0066CC")
                    st.markdown(f"""
<div style="background:#ffffff;border:1px solid {_cs_bc}44;border-radius:10px;overflow:hidden;font-family:DM Sans,sans-serif;width:100%;opacity:0.82;">
  <div style="width:100%;height:160px;overflow:hidden;position:relative;">
    {_cs_photo_html}
    <div style="position:absolute;top:10px;right:10px;background:#ffffff;border:1px solid {_cs_bc}88;border-radius:12px;padding:3px 10px;font-size:10px;font-weight:700;color:{_cs_ac};font-family:DM Mono,monospace;letter-spacing:0.08em;text-transform:uppercase;">Coming Soon</div>
  </div>
  <div style="height:3px;background:linear-gradient(90deg,{_cs_bc},{_cs_ac});"></div>
  <div style="padding:12px 14px 14px;text-align:center;">
    <div style="font-family:Syne,sans-serif;font-size:15px;font-weight:700;color:#1e2d35;line-height:1.2;margin-bottom:2px;">{hotel['display_name']}</div>
    <div style="font-family:Syne,sans-serif;font-size:13px;font-weight:600;color:#1e2d35;line-height:1.2;">{hotel['subtitle']}</div>
    <div style="font-size:10px;color:#4e6878;margin-top:3px;">{hotel['brand']} &middot; {hotel.get('total_rooms','?')} Rooms</div>
  </div>
</div>""", unsafe_allow_html=True)
                    continue

                # ── Load tonight + month stats ──
                import base64 as _b64, os as _os
                _yr_path  = str(files["year"])  if files.get("year")   else "None"
                _bud_path = str(files["budget"]) if files.get("budget") else "None"
                _yr_mtime = f"{_os.path.getmtime(_yr_path):.0f}" if _yr_path != "None" and _os.path.exists(_yr_path) else ""

                tonight = cached_tonight(hotel["id"], _yr_path, hotel.get("total_rooms", 112), file_mtime=_yr_mtime)
                lstats  = cached_landing_stats(hotel["id"], _yr_path, _bud_path, hotel.get("total_rooms", 112), file_mtime=_yr_mtime)

                occ_pct   = tonight["occ_pct"]   if tonight else None
                event_txt = tonight["event"]      if tonight else ""
                otb_rooms = tonight["otb_rooms"]  if tonight else None
                fcst_occ  = lstats.get("fcst_occ")
                mo_delta  = lstats.get("mo_delta")
                mo_label  = lstats.get("mo_label", "Month vs Budget")

                # Sanitize event_txt
                if not isinstance(event_txt, str) or str(event_txt).lower() in ("nan","none",""):
                    event_txt = ""

                # ── Metric colors & display ──
                occ_color  = TEAL if fcst_occ and fcst_occ >= 70 else (GOLD if fcst_occ and fcst_occ >= 50 else (ORANGE if fcst_occ else "#6a8090"))
                delta_color = TEAL if mo_delta and mo_delta >= 0 else ORANGE
                # Clean number formats: occ as "74%", delta as "+$48,505" or "-$12,300"
                fcst_disp   = f"{round(fcst_occ)}%" if fcst_occ is not None else "—"
                if mo_delta is not None:
                    _sign = "+" if mo_delta >= 0 else "-"
                    _amt  = abs(mo_delta)
                    delta_disp = f"{_sign}${_amt:,.0f}"
                else:
                    delta_disp = "—"
                event_html  = f"<div style='font-size:11px;color:#2a5e2a;margin-top:6px;padding:3px 10px;background:#c8eac8;border-radius:12px;display:inline-block;font-weight:600;border:1px solid #96cc96;'>📅 {event_txt}</div>" if event_txt else ""

                # ── File readiness ──
                if status["ready"]:
                    readiness_html = ""
                elif status["present"] == 0:
                    readiness_html = "<div style='font-size:10px;color:#b44820;margin-top:8px;'>⚠ No data files detected</div>"
                else:
                    readiness_html = f"<div style='font-size:10px;color:#4e6878;margin-top:8px;'>📁 {status['present']}/{status['total']} files</div>"

                # ── Hotel photo ──
                photo_html = ""
                _photo_path = hotel.get("photo_path", "")
                if _photo_path and _os.path.exists(str(_photo_path)):
                    _ext  = str(_photo_path).rsplit(".", 1)[-1].lower()
                    _mime = "image/jpeg" if _ext in ("jpg","jpeg") else "image/png"
                    try:
                        with open(_photo_path, "rb") as _f:
                            _b64d = _b64.b64encode(_f.read()).decode()
                        photo_html = f'<img src="data:{_mime};base64,{_b64d}" style="width:100%;height:160px;object-fit:cover;display:block;" />'
                    except Exception as _photo_err:
                        photo_html = f"<div style='width:100%;height:150px;background:#fdf0ec;display:flex;align-items:center;justify-content:center;font-size:10px;color:#b44820;'>⚠ Photo error: {_photo_err}</div>"
                elif _photo_path:
                    photo_html = f"<div style='width:100%;height:150px;background:#fdf0ec;display:flex;align-items:center;justify-content:center;font-size:10px;color:#b44820;text-align:center;padding:8px;'>⚠ Not found: {_photo_path}</div>"

                # Pre-resolve photo section to avoid nested f-string issues
                if photo_html:
                    _photo_section = photo_html
                else:
                    _bc = hotel['brand_color']
                    _photo_section = f'<div style="width:100%;height:150px;background:linear-gradient(135deg,{_bc}22,#eef2f3);display:flex;align-items:center;justify-content:center;"><span style="font-size:32px;">🏨</span></div>'

                card_html = f"""
<div style="background:#ffffff;border:1px solid {hotel['brand_color']}44;border-radius:10px;
            overflow:hidden;cursor:pointer;transition:box-shadow 0.2s;font-family:DM Sans,sans-serif;width:100%;">
  <!-- Photo -->
  <div style="width:100%;height:150px;background:#eef2f3;overflow:hidden;">
    {_photo_section}
  </div>
  <!-- Accent bar -->
  <div style="height:3px;background:linear-gradient(90deg,{hotel['brand_color']},{hotel['accent_color']});"></div>
  <!-- Body -->
  <div style="padding:12px 14px 14px;text-align:center;">
    <div style="font-family:Syne,sans-serif;font-size:15px;font-weight:700;color:#1e2d35;line-height:1.2;margin-bottom:2px;">{hotel['display_name']}</div>
    <div style="font-family:Syne,sans-serif;font-size:13px;font-weight:600;color:#1e2d35;line-height:1.2;">{hotel['subtitle']}</div>
    <div style="font-size:10px;color:#4e6878;margin-top:3px;">{hotel['brand']} &middot; {hotel.get('total_rooms','?')} Rooms</div>
    <div style="border-top:1px solid #c4cfd4;margin:10px 0;"></div>
    <div style="display:flex;gap:0;">
      <div style="flex:0 0 40%;text-align:center;padding-right:8px;">
        <div style="font-size:8px;color:#4e6878;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:4px;">Forecast Tonight</div>
        <div style="font-family:Syne,sans-serif;font-size:24px;font-weight:800;color:{occ_color};line-height:1;">{fcst_disp}</div>
      </div>
      <div style="flex:0 0 60%;border-left:1px solid #c4cfd4;text-align:center;padding-left:8px;">
        <div style="font-size:8px;color:#4e6878;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:4px;">{mo_label}</div>
        <div style="font-family:Syne,sans-serif;font-size:20px;font-weight:800;color:{delta_color};line-height:1;">{delta_disp}</div>
      </div>
    </div>
    {event_html}
  </div>
</div>"""

                # Render card HTML then a Streamlit button to navigate.
                # NEVER use <a href="?open=..."> — that triggers a full page reload
                # which wipes st.session_state and fires the Access Denied gate.
                st.markdown(card_html, unsafe_allow_html=True)
                if st.button("Open →", key=f"open_hotel_{hotel['id']}", use_container_width=True):
                    st.session_state.hotel_id   = hotel["id"]
                    st.session_state.active_tab = "Biweekly Snapshot"
                    st.session_state.hotel_data = {}
                    st.rerun()
                if readiness_html:
                    st.markdown(readiness_html, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD — load data for selected hotel
# ══════════════════════════════════════════════════════════════════════════════

def _file_mtimes(hotel_id: str) -> str:
    """Return a string of all source file modification times — used as cache key."""
    hotel = next((h for h in cfg.HOTELS if h["id"] == hotel_id), None)
    if not hotel:
        return ""
    files = cfg.detect_files(hotel)
    parts = []
    for role, path in sorted(files.items()):
        if path and os.path.exists(str(path)):
            parts.append(f"{role}:{os.path.getmtime(str(path)):.0f}")
    return "|".join(parts)


@st.cache_data(show_spinner=False)
def load_hotel_data(hotel_id: str, file_mtimes: str = "", dl_version: str = "", cache_date: str = ""):
    """Load all hotel data. file_mtimes + dl_version bust cache when files or loader change."""
    hotel = next(h for h in cfg.HOTELS if h["id"] == hotel_id)
    files = cfg.detect_files(hotel)
    return dl.load_all(files, hotel.get("total_rooms", 112), hotel_id=hotel_id)


def _get_fresh_hotel_data(hotel_id: str) -> dict:
    """
    Cache-aware loader that auto-clears stale cache when dl_version changes.
    Stores last-seen dl_version in session_state and clears cache on mismatch.
    """
    current_version = getattr(dl, "_DL_VERSION", "")
    seen_key = f"_seen_dl_version_{hotel_id}"
    if st.session_state.get(seen_key) != current_version:
        load_hotel_data.clear()
        st.session_state[seen_key] = current_version
    return load_hotel_data(
        hotel_id,
        file_mtimes=_file_mtimes(hotel_id),
        dl_version=current_version,
        cache_date=datetime.now().strftime("%Y-%m-%d"),
    )


def render_snapshot_tab(data, hotel):
    """14-day outlook snapshot matching the classic RM view."""
    import base64 as _b64, os as _os

    year_df   = data.get("year")
    budget_df = data.get("budget")
    pickup1   = data.get("pickup1")
    pickup7   = data.get("pickup7")
    rates_ov  = data.get("rates_overview")
    rates_comp= data.get("rates_comp")
    rates_vs7 = data.get("rates_vs7")
    tonight_d = data.get("tonight") or {}
    groups_df = data.get("groups")
    total_rooms = hotel.get("total_rooms", 112)

    today   = pd.Timestamp(datetime.now().date())

    # ── Theme-aware color palette for snapshot ──────────────────────────────
    _slt = False
    _S = {
        # backgrounds
        "bg_main":    "#ffffff",
        "bg_card":    "#ffffff",
        "bg_header":  "#314B63",
        "bg_today":   "#314B63",
        "bg_wknd":    "#eaeef1",
        "bg_row_alt": "#e8eef1",
        "bg_row_std": "#ffffff",
        "bg_bar_std": "#eaeef1",
        "bg_bar_wknd":"#e4eaed",
        "bg_bar_today":"#314B63",
        "bg_lbl_cell":"#314B63",
        "bg_our_row": "#e4eaed",
        "bg_comp_name":"#e4eaed",
        "bg_monthly_hdr":"#e4eaed",
        "bg_monthly_hdr2":"#314B63",
        # text
        "txt_pri":    "#1e2d35",
        "txt_sec":    "#3a5260",
        "txt_dim":    "#6a8090",
        "txt_header": "#ffffff",
        "txt_col_hdr":"#314B63",
        "txt_today":  "#ffffff",
        "txt_wknd":   "#6A924D",
        "txt_rate":   "#6A924D",
        "txt_accent": "#ffffff",
        "txt_teal":   "#2E618D",
        "txt_orange": "#b44820",
        "txt_red":    "#cc2200",
        "txt_muted":  "#5a7080",
        "txt_our":    "#6A924D",
        "txt_comp":   "#3a5260",
        # borders
        "bdr":        "#c4cfd4",
        "bdr2":       "#d4dce0",
        # bar chart empty area
        "bar_empty":  "rgba(0,0,0,0.06)",
        "bar_empty_bdr":"rgba(0,0,0,0.12)",
        # forecast stripe
        "fcst_stripe_bg": "#e0edf8",
        # bar chart fill colors — blue theme in light, teal in dark
        "bar_hi":  "#2E618D",
        "bar_mid": "#1D70B8",
        "bar_lo":  "#6A924D",
        "bar_fcst":"#556848",
    }

    # ── Build 14-day forecast slice ──────────────────────────────────────────
    if year_df is not None:
        year_df["Date"] = pd.to_datetime(year_df["Date"]).dt.normalize()
        fwd14 = year_df[year_df["Date"] >= today].head(16).copy()
        fwd14["Date"] = pd.to_datetime(fwd14["Date"]).dt.normalize()
    else:
        fwd14 = pd.DataFrame()

    if budget_df is not None and not fwd14.empty:
        fwd14 = fwd14.merge(
            budget_df[["Date","Occ_Rooms","ADR","Revenue","RevPAR"]].rename(
                columns={"Occ_Rooms":"Bud_Rooms","ADR":"Bud_ADR","Revenue":"Bud_Rev","RevPAR":"Bud_RevPAR"}),
            on="Date", how="left")

    # Pickup rows aligned to the same 14 dates — normalize keys for consistent matching
    pickup1_map, pickup7_map = {}, {}
    if pickup1 is not None:
        for _, r in pickup1.iterrows():
            if pd.notna(r.get("Date")):
                # Hilton: OTB_Change column  |  IHG: Pickup column (renamed by _build_pickup)
                val = r.get("OTB_Change") if pd.notna(r.get("OTB_Change", None)) else r.get("Pickup", 0)
                try:
                    pickup1_map[pd.Timestamp(r["Date"]).normalize()] = int(float(val)) if pd.notna(val) else 0
                except (ValueError, TypeError):
                    pickup1_map[pd.Timestamp(r["Date"]).normalize()] = 0
    if pickup7 is not None:
        for _, r in pickup7.iterrows():
            if pd.notna(r.get("Date")):
                # Hilton: OTB_Change column  |  IHG: Pickup column (renamed by _build_pickup)
                val = r.get("OTB_Change") if pd.notna(r.get("OTB_Change", None)) else r.get("Pickup", 0)
                try:
                    pickup7_map[pd.Timestamp(r["Date"]).normalize()] = int(float(val)) if pd.notna(val) else 0
                except (ValueError, TypeError):
                    pickup7_map[pd.Timestamp(r["Date"]).normalize()] = 0

    # ── Past 1 week actuals ──────────────────────────────────────────────────
    if year_df is not None:
        past14 = year_df[(year_df["Date"] >= today - timedelta(days=8)) &
                         (year_df["Date"] < today)].copy()
        past14 = past14.sort_values("Date", ascending=False)
    else:
        past14 = pd.DataFrame()

    # ── Comp rates tonight ────────────────────────────────────────────────────
    comp_tonight = []
    if rates_comp is not None:
        tonight_rates = rates_comp[rates_comp["Date"] == today]
        if tonight_rates.empty and not rates_comp.empty:
            tonight_rates = rates_comp[rates_comp["Date"] == rates_comp["Date"].min()]
        for _, r in tonight_rates.iterrows():
            comp_tonight.append({"hotel": r.get("Hotel","—"), "rate": r.get("Rate", None)})

    our_rate_tonight = None
    if rates_ov is not None:
        t_rate = rates_ov[rates_ov["Date"] == today]
        if not t_rate.empty:
            our_rate_tonight = t_rate.iloc[0].get("Our_Rate")


    # ── Monthly Forecast vs Budget table ─────────────────────────────────────
    monthly_var = []
    if year_df is not None:
        yr = year_df.copy()
        yr["Month"] = yr["Date"].dt.to_period("M")
        # Use forecast rooms if available, fallback to OTB
        rooms_col = "Forecast_Rooms" if "Forecast_Rooms" in yr.columns else "OTB"
        # ADR column: Hilton uses ADR_OTB, IHG uses plain ADR
        adr_col   = "ADR_OTB" if "ADR_OTB" in yr.columns else ("ADR" if "ADR" in yr.columns else None)

        bd = budget_df.copy() if budget_df is not None else None
        if bd is not None:
            bd["Month"] = bd["Date"].dt.to_period("M")

        # Start from current month, show next 3 months
        current_mo = pd.Period(today, "M")
        all_months = sorted(yr["Month"].unique())
        fwd_months = [m for m in all_months if m >= current_mo][:3]

        for mo in fwd_months:
            y_mo = yr[yr["Month"] == mo]

            # Revenue priority:
            # 1. Revenue_Forecast (Hilton name — full-month column, rooms unchanged)
            # 2. Forecast_Rev (IHG name — future days only; blend with past Revenue_OTB)
            #    CRITICAL: when blending revenue, ty_rooms must use the same day-split
            #    (past=OTB, future=Forecast_Rooms) so ADR = ty_rev / ty_rooms is correct.
            #    Using full-month Forecast_Rooms with only 1-2 days of Forecast_Rev
            #    collapses ADR to near-zero (e.g. $67 instead of $871).
            # 3. Revenue_OTB fallback
            # 4. Rooms × ADR last resort
            if "Revenue_Forecast" in y_mo.columns and y_mo["Revenue_Forecast"].notna().any():
                ty_rooms = float(y_mo[rooms_col].sum()) if rooms_col in y_mo.columns else 0.0
                ty_rev   = float(y_mo["Revenue_Forecast"].sum())
                ty_adr   = (ty_rev / ty_rooms) if ty_rooms > 0 else None
            elif "Forecast_Rev" in y_mo.columns and y_mo["Forecast_Rev"].notna().any():
                # IHG blend: align rooms split to match revenue split
                _past_mask = y_mo["Forecast_Rev"].isna()
                _fut_mask  = y_mo["Forecast_Rev"].notna()
                past_rooms = float(y_mo.loc[_past_mask, "OTB"].fillna(0).sum()) if "OTB" in y_mo.columns else 0.0
                fut_rooms  = float(y_mo.loc[_fut_mask,  rooms_col].fillna(0).sum()) if rooms_col in y_mo.columns else 0.0
                ty_rooms   = past_rooms + fut_rooms
                fut_rev    = float(y_mo["Forecast_Rev"].fillna(0).sum())
                past_rev   = float(y_mo.loc[_past_mask, "Revenue_OTB"].fillna(0).sum()) if "Revenue_OTB" in y_mo.columns else 0.0
                ty_rev     = fut_rev + past_rev
                ty_adr     = (ty_rev / ty_rooms) if ty_rooms > 0 else None
            elif "Revenue_OTB" in y_mo.columns and y_mo["Revenue_OTB"].notna().any():
                ty_rooms = float(y_mo[rooms_col].sum()) if rooms_col in y_mo.columns else 0.0
                ty_rev   = float(y_mo["Revenue_OTB"].sum())
                ty_adr   = (ty_rev / ty_rooms) if ty_rooms > 0 else None
            else:
                ty_rooms = float(y_mo[rooms_col].sum()) if rooms_col in y_mo.columns else 0.0
                # Last resort: rooms * ADR
                adr_col2 = "ADR_Forecast" if "ADR_Forecast" in y_mo.columns else adr_col
                if adr_col2 and adr_col2 in y_mo.columns and ty_rooms > 0:
                    rev_sum = (y_mo[rooms_col] * y_mo[adr_col2]).sum()
                    ty_adr  = float(rev_sum / ty_rooms)
                    ty_rev  = float(rev_sum)
                else:
                    ty_adr = ty_rev = None

            bu_rooms = bu_adr = bu_rev = None
            if bd is not None:
                b_mo = bd[bd["Month"] == mo]
                if not b_mo.empty:
                    # Budget has: Occ_Rooms, ADR, Revenue columns
                    bu_rooms = float(b_mo["Occ_Rooms"].sum()) if "Occ_Rooms" in b_mo.columns else None
                    if bu_rooms and bu_rooms > 0 and "ADR" in b_mo.columns:
                        # Revenue-weighted budget ADR
                        bu_rev_sum = (b_mo["Occ_Rooms"] * b_mo["ADR"]).sum()
                        bu_adr = float(bu_rev_sum / bu_rooms) if bu_rooms else None
                    bu_rev = float(b_mo["Revenue"].sum()) if "Revenue" in b_mo.columns and b_mo["Revenue"].notna().any() else (
                             (bu_rooms * bu_adr) if (bu_rooms and bu_adr) else None)

            var_rev = (ty_rev - bu_rev) if (ty_rev is not None and bu_rev is not None) else None
            monthly_var.append({
                "Month": str(mo),
                "TY_Rooms": ty_rooms,
                "TY_ADR": ty_adr,
                "TY_Rev": ty_rev,
                "Bu_Rooms": bu_rooms,
                "Bu_ADR": bu_adr,
                "Bu_Rev": bu_rev,
                "Var_Rev": var_rev,
            })

    # ════════════════════════════════════════════════════════════════════════
    # RENDER — Build HTML for the three-column layout
    # ════════════════════════════════════════════════════════════════════════

    # ── Helpers ──────────────────────────────────────────────────────────────
    def _d(v, fmt="dollar"):
        if v is None or (isinstance(v, float) and math.isnan(v)): return "—"
        if fmt == "dollar": return f"${v:,.0f}"
        if fmt == "pct":    return f"{v:.1f}%"
        return f"{v:,.0f}"

    def _row_cls(dt):
        dow = dt.strftime("%a")
        if dt.date() == today.date(): return "snap-highlight"
        if dow in ("Fri", "Sat"):     return "snap-weekend"
        return ""

    # ── LEFT: 14-day forecast table ──────────────────────────────────────────
    left_rows = ""
    avg_rooms = avg_adr = avg_rev = cnt = 0
    # Build a lookup from pickup1 for Forecast_Current per date
    p1_fcst_map = {}
    if pickup1 is not None:
        for _, pr in pickup1.iterrows():
            if pd.notna(pr.get("Date")):
                p1_fcst_map[pr["Date"]] = pr.get("Forecast_Current")

    for _, r in fwd14.iterrows():
        dt  = r["Date"]
        lbl = "Tonight" if dt.date() == today.date() else dt.strftime("%a, %b %d").lstrip("0")
        # Rooms: use year_df Forecast_Rooms first (DSS-scaled for IHG, most accurate)
        # fallback to pickup1 Forecast_Current, then OTB
        yr_fc = r.get("Forecast_Rooms")
        if yr_fc is not None and pd.notna(yr_fc):
            rooms = int(round(float(yr_fc)))
        else:
            p1_fcst = p1_fcst_map.get(dt)
            rooms = int(round(p1_fcst)) if p1_fcst is not None and pd.notna(p1_fcst) else (
                    int(r["OTB"]) if "OTB" in r and pd.notna(r["OTB"]) else None)
        # ADR: try Hilton names first, then IHG "ADR" column
        adr   = r.get("ADR_Forecast") if pd.notna(r.get("ADR_Forecast", None)) else (
                r.get("ADR_OTB")      if pd.notna(r.get("ADR_OTB",      None)) else (
                r.get("ADR")          if pd.notna(r.get("ADR",          None)) else None))
        # Revenue: try Hilton names, then IHG Forecast_Rev, then Revenue_OTB, then rooms×ADR
        rev   = r.get("Revenue_Forecast") if pd.notna(r.get("Revenue_Forecast", None)) else (
                r.get("Forecast_Rev")      if pd.notna(r.get("Forecast_Rev",     None)) else (
                r.get("Revenue_OTB")       if pd.notna(r.get("Revenue_OTB",      None)) else (
                (rooms * adr)              if (rooms and adr)                            else None)))
        ooo   = int(r["OOO"]) if "OOO" in r and pd.notna(r.get("OOO")) and r.get("OOO") else 0
        occ_pct = (rooms / total_rooms * 100) if rooms else 0
        cls   = _row_cls(dt)
        if rooms:
            avg_rooms += rooms
            avg_adr    = (avg_adr or 0) + (adr or 0 if adr else 0)
            avg_rev    = (avg_rev or 0) + (rev or 0)
            cnt       += 1
        left_rows += f"""<tr class="{cls}">
            <td>{lbl}</td>
            <td style="text-align:center">{rooms if rooms else "—"}</td>
            <td style="text-align:center">{_d(adr)}</td>
            <td style="text-align:center">{_d(rev)}</td>
            <td style="text-align:center;color:#6a8090">{ooo}</td>
        </tr>"""
    # avg / budget rows
    # ADR = simple average of daily ADR values (what the RM sees per night)
    # Revenue = avg_rooms × avg_adr so the three columns are always self-consistent
    a_rooms = round(avg_rooms / cnt) if cnt else None
    a_adr   = (avg_adr   / cnt) if cnt else None
    a_rev   = (a_rooms * a_adr) if (a_rooms and a_adr) else None
    left_rows += f"""<tr class="snap-avg"><td>Average</td><td style="text-align:center">{_d(a_rooms,'num')}</td><td style="text-align:center">{_d(a_adr)}</td><td style="text-align:center">{_d(a_rev)}</td><td></td></tr>"""
    if budget_df is not None and not fwd14.empty:
        b14 = fwd14[["Bud_Rooms","Bud_ADR","Bud_Rev"]].head(14)
        bav_r = b14["Bud_Rooms"].mean() if "Bud_Rooms" in b14 else None
        bav_a = b14["Bud_ADR"].mean()   if "Bud_ADR"   in b14 else None
        bav_v = b14["Bud_Rev"].mean()   if "Bud_Rev"   in b14 else None
        left_rows += f"""<tr class="snap-avg"><td>Budgeted</td><td style="text-align:center">{_d(bav_r,'num')}</td><td style="text-align:center">{_d(bav_a)}</td><td style="text-align:center">{_d(bav_v)}</td><td></td></tr>"""

    left_html = f"""
    <div class="snap-card">
      <span class="snap-section-title">Forecast — 14 Day Outlook</span>
      <table class="snap-table">
        <tr class="snap-header-row"><th></th><th>Rooms</th><th>ADR</th><th>Revenue</th><th>OOO</th></tr>
        {left_rows}
      </table>
    </div>"""

    # ── RIGHT: Past 2 weeks actuals ──────────────────────────────────────────
    right_rows = ""
    pa_rooms = pa_adr = pa_rev = pa_cnt = 0
    for _, r in past14.iterrows():
        dt  = r["Date"]
        lbl = "Last Night" if dt.date() == (today - timedelta(days=1)).date() else dt.strftime("%a, %b %d").lstrip("0")
        # Occupancy On Books This Year, ADR On Books This Year, Booked Room Revenue This Year
        rooms = int(r["OTB"]) if "OTB" in r and pd.notna(r.get("OTB")) else None
        # ADR: Hilton uses ADR_OTB, IHG uses ADR
        adr   = (float(r["ADR_OTB"])  if "ADR_OTB" in r and pd.notna(r.get("ADR_OTB"))  else (
                 float(r["ADR"])       if "ADR"     in r and pd.notna(r.get("ADR"))       else None))
        rev   = (float(r["Revenue_OTB"]) if "Revenue_OTB" in r and pd.notna(r.get("Revenue_OTB")) else (
                 (rooms * adr)            if (rooms and adr) else None))
        cls   = "snap-weekend" if dt.strftime("%a") in ("Fri","Sat") else ""
        if rooms: pa_rooms += rooms; pa_adr = (pa_adr or 0)+(adr or 0); pa_rev = (pa_rev or 0)+(rev or 0); pa_cnt+=1
        right_rows += f"""<tr class="{cls}">
            <td>{lbl}</td>
            <td style="text-align:center">{rooms if rooms else "—"}</td>
            <td style="text-align:center">{_d(adr)}</td>
            <td style="text-align:center">{_d(rev)}</td>
        </tr>"""
    pr = round(pa_rooms/pa_cnt) if pa_cnt else None
    pa = (pa_adr/pa_cnt) if pa_cnt else None
    pv = (pa_rev/pa_cnt) if pa_cnt else None
    right_rows += f"""<tr class="snap-avg"><td>Average</td><td style="text-align:center">{_d(pr,'num')}</td><td style="text-align:center">{_d(pa)}</td><td style="text-align:center">{_d(pv)}</td></tr>"""
    if budget_df is not None:
        p2w = budget_df[(budget_df["Date"] >= today - timedelta(days=8)) & (budget_df["Date"] < today)]
        bpr = p2w["Occ_Rooms"].mean() if "Occ_Rooms" in p2w else None
        bpa = p2w["ADR"].mean()       if "ADR"       in p2w else None
        bpv = (bpr * bpa)             if (bpr and bpa) else None
        right_rows += f"""<tr class="snap-avg"><td>Budget</td><td style="text-align:center">{_d(bpr,'num')}</td><td style="text-align:center">{_d(bpa)}</td><td style="text-align:center">{_d(bpv)}</td></tr>"""

    # ── Hotel photo + info block for right panel ─────────────────────────────
    _photo_path_r = hotel.get("photo_path", "")
    _photo_html_r = ""
    if _photo_path_r and _os.path.exists(str(_photo_path_r)):
        _ext_r  = str(_photo_path_r).rsplit(".", 1)[-1].lower()
        _mime_r = "image/jpeg" if _ext_r in ("jpg","jpeg") else "image/png"
        try:
            with open(_photo_path_r, "rb") as _fr:
                _b64d_r = _b64.b64encode(_fr.read()).decode()
            _photo_html_r = f'<img src="data:{_mime_r};base64,{_b64d_r}" style="width:100%;height:140px;object-fit:cover;display:block;border-radius:8px 8px 0 0;" />'
        except Exception:
            _photo_html_r = ""
    if not _photo_html_r:
        _bc_r = hotel.get("brand_color","#314B63")
        _photo_html_r = f'<div style="width:100%;height:140px;background:linear-gradient(135deg,{_bc_r}33,#eef2f3);display:flex;align-items:center;justify-content:center;border-radius:8px 8px 0 0;"><span style="font-size:36px;">🏨</span></div>'

    _hotel_info_html = (
        f'<div style="background:#ffffff;border:1px solid #c4cfd4;border-radius:8px;overflow:hidden;margin-bottom:8px;">'
        f'{_photo_html_r}'
        f'<div style="height:3px;background:linear-gradient(90deg,{hotel.get("brand_color","#314B63")},{hotel.get("accent_color","#0066CC")});"></div>'
        f'<div style="padding:8px 10px;text-align:center;">'
        f'<div style="font-family:Syne,sans-serif;font-size:14px;font-weight:700;color:#1e2d35;line-height:1.3;">{hotel.get("display_name","")}</div>'
        f'<div style="font-family:Syne,sans-serif;font-size:12px;font-weight:500;color:#3a5260;line-height:1.3;">{hotel.get("subtitle","")}</div>'
        f'<div style="font-size:10px;color:#6a8090;margin-top:2px;">{hotel.get("brand","")} &middot; {hotel.get("total_rooms","?")} Rooms</div>'
        f'</div></div>'
    )

    right_html = f"""
    {_hotel_info_html}
    <div class="snap-card">
      <span class="snap-section-title">Actuals — Past Week</span>
      <table class="snap-table">
        <tr class="snap-header-row"><th></th><th>Rooms</th><th>ADR</th><th>Revenue</th></tr>
        {right_rows}
      </table>
    </div>"""

    # ── CENTER: HTML grid — bars on top, pickup/rate rows below (pixel-perfect alignment) ─
    dates14 = list(fwd14["Date"]) if not fwd14.empty else []

    def _fmt_bar_rate(rate):
        """Safely format a BAR rate value — handles numeric, Sold Out strings, and None."""
        if rate is None or (isinstance(rate, float) and math.isnan(rate)):
            return "\u2014"
        s = str(rate).strip()
        if not s or s.lower() in ("nan", "none", "—", "-"):
            return "\u2014"
        if "sold" in s.lower():
            return "SOLD"
        try:
            return f"${int(float(s.replace(',', '').replace('$', '')))}"
        except (ValueError, TypeError):
            return s[:6]  # truncate unknown strings

    center_html = ""
    cols_data = []   # always defined — populated below if fwd14 has data
    if not fwd14.empty:
        cols_data = []
        # Build pickup1 lookups for bar chart
        p1_otb_map  = {}
        p1_fcst_map2 = {}
        if pickup1 is not None:
            for _, pr in pickup1.iterrows():
                if pd.notna(pr.get("Date")):
                    # Hilton: OTB_Current / Forecast_Current columns
                    # IHG:    OTB / Forecast_Rooms columns (from _load_all_ihg pickup builder)
                    otb_val  = pr.get("OTB_Current")
                    if otb_val is None or pd.isna(otb_val):
                        otb_val = pr.get("OTB")
                    fcst_val = pr.get("Forecast_Current")
                    if fcst_val is None or pd.isna(fcst_val):
                        fcst_val = pr.get("Forecast_Rooms")
                    p1_otb_map[pd.Timestamp(pr["Date"]).normalize()]   = otb_val
                    p1_fcst_map2[pd.Timestamp(pr["Date"]).normalize()] = fcst_val

        # Build group rate lookup: date → highest group rate for that day
        grp_snap_lookup = {}
        if groups_df is not None and "Occ_Date" in groups_df.columns and "Rate" in groups_df.columns:
            for _, gr in groups_df.iterrows():
                ocd = gr.get("Occ_Date")
                if ocd is None or pd.isna(ocd):
                    continue
                rt = gr.get("Rate")
                if rt is None or pd.isna(rt):
                    continue
                try:
                    rt_f = float(rt)
                except (TypeError, ValueError):
                    continue
                if rt_f <= 0:
                    continue
                d = pd.Timestamp(ocd).normalize()
                if d not in grp_snap_lookup or rt_f > grp_snap_lookup[d]:
                    grp_snap_lookup[d] = rt_f

        for _, r in fwd14.iterrows():
            dt    = r["Date"]
            # OTB: use year_df OTB directly (most accurate — Data_Glance for IHG, 1.xlsx for Hilton)
            # p1_otb_map fallback only if year_df OTB is missing
            yr_otb = r.get("OTB")
            if yr_otb is not None and pd.notna(yr_otb):
                rooms = int(round(float(yr_otb)))
            else:
                otb_raw = p1_otb_map.get(dt)
                rooms = int(round(otb_raw)) if otb_raw is not None and pd.notna(otb_raw) else 0
            # Forecast: use year_df Forecast_Rooms (DSS-scaled for IHG, pickup-based for Hilton)
            # This ensures left panel and center bar use the same forecast number
            yr_fcst = r.get("Forecast_Rooms")
            if yr_fcst is not None and pd.notna(yr_fcst):
                fcst = int(round(float(yr_fcst)))
            else:
                fcst_raw = p1_fcst_map2.get(dt)
                fcst = int(round(fcst_raw)) if fcst_raw is not None and pd.notna(fcst_raw) else rooms
            fcst_pickup = max(0, fcst - rooms)   # remaining pickup = Forecast minus OTB
            occ_otb  = min(rooms / total_rooms, 1.0)
            occ_fcst = min(fcst  / total_rooms, 1.0)
            lbl1  = "Tonight" if dt.date() == today.date() else dt.strftime("%a")
            lbl2  = ""        if dt.date() == today.date() else dt.strftime("%b") + " " + str(dt.day)
            rate  = None
            if rates_ov is not None:
                rr = rates_ov[rates_ov["Date"] == dt]
                if not rr.empty:
                    rate = rr.iloc[0].get("Our_Rate")
            p1 = pickup1_map.get(dt, 0)
            p7 = pickup7_map.get(dt, 0)
            # Color OTB bar by forecasted occupancy level
            bar_color_otb  = _S["bar_hi"] if occ_fcst >= 0.80 else (_S["bar_mid"] if occ_fcst >= 0.50 else _S["bar_lo"])
            bar_color_fcst = _S["bar_fcst"]  # darker segment for forecast pickup
            is_today  = dt.date() == today.date()
            is_wknd   = dt.strftime("%a") in ("Fri", "Sat")
            cols_data.append({
                "lbl1": lbl1, "lbl2": lbl2,
                "rooms": rooms, "fcst": fcst, "fcst_pickup": fcst_pickup,
                "occ_otb": occ_otb, "occ_fcst": occ_fcst,
                "bar_color_otb": bar_color_otb, "bar_color_fcst": bar_color_fcst,
                "rate": _fmt_bar_rate(rate),
                "p1": p1, "p7": p7,
                "is_today": is_today, "is_wknd": is_wknd,
                "grp_rate": grp_snap_lookup.get(dt.normalize() if hasattr(dt, "normalize") else pd.Timestamp(dt).normalize()),
            })

        def _pu_fmt(v):
            return f"+{v}" if v > 0 else str(v)
        def _pu_color(v):
            return "#2E618D" if v > 0 else ("#a03020" if v < 0 else "#6a8090")

        n = len(cols_data)
        grid_cols = f"56px repeat({n}, 1fr)"
        max_bar_h = 140

        # ── Date header row ──
        date_cells = ""
        for c in cols_data:
            bg    = f"background:{_S['bg_today']};" if c["is_today"] else (f"background:{_S['bg_wknd']};" if c["is_wknd"] else f"background:{_S['bg_main']};")
            color = _S["txt_today"] if c["is_today"] else (_S["txt_wknd"] if c["is_wknd"] else _S["txt_sec"])
            date_cells += (
                f'<div style="text-align:center;padding:5px 2px 4px;{bg}border-right:1px solid {_S["bdr"]};">' +
                f'<div class="snap-cal-dow" style="color:{color};">{c["lbl1"]}</div>' +
                f'<div class="snap-cal-date" style="color:{color};">{c["lbl2"]}</div>' +
                '</div>'
            )

        # ── Bar chart area — stacked bars: OTB (solid bottom) + forecast (lighter top) ──
        bar_cells = ""
        for c in cols_data:
            otb_h   = max(2, int(c["occ_otb"]  * max_bar_h))
            fcst_h  = max(0, int(c["occ_fcst"] * max_bar_h)) - otb_h
            fcst_h  = max(0, fcst_h)
            empty_h = max_bar_h - otb_h - fcst_h
            bg_col  = c["bar_color_otb"]
            bg      = (f"background:{_S['bg_bar_wknd']};" if c["is_wknd"] else f"background:{_S['bg_bar_std']};")

            fcst_total = c["fcst"] if c["fcst"] else 0

            # Number label INSIDE bar — forecast total, white, bold
            # Shows fcst_total (forecast rooms) centered near top of filled area
            lbl_val = fcst_total if fcst_total else (c["rooms"] if c["rooms"] else 0)
            # Pickup annotation shown to the right of main number (above bar)
            if c["fcst_pickup"] > 0:
                above_lbl = (
                    f'<span style="font-family:DM Mono,monospace;font-size:11px !important;'
                    f'font-weight:600;color:{_S["txt_sec"]};">{c["rooms"]}</span>'
                    f'<span style="color:#2E618D !important;font-size:10px !important;'
                    f'font-weight:600;margin-left:2px;">+{c["fcst_pickup"]}</span>'
                )
            else:
                above_lbl = ''

            otb_color  = c["bar_color_otb"]
            fcst_color = c["bar_color_fcst"]
            bar_top_h  = otb_h + fcst_h  # total filled height

            # Forecast segment (lighter, top) — no label here
            fcst_div = (
                f'<div style="height:{fcst_h}px;background:{otb_color};opacity:0.42;'
                f'border-radius:3px 3px 0 0;margin:0 4px 0;"></div>'
            ) if fcst_h > 0 else ''

            # OTB bar — number label centered inside it near the top
            # Only show inside label if bar is tall enough (≥24px)
            if bar_top_h >= 24 and lbl_val:
                inner_lbl = (
                    f'<div style="position:absolute;top:6px;left:0;right:0;'
                    f'text-align:center;line-height:1;pointer-events:none;">'
                    f'<b style="font-family:DM Mono,monospace;font-size:13px;'
                    f'font-weight:800;color:white;filter:brightness(10);'
                    f'-webkit-text-fill-color:white;">{lbl_val}</b></div>'
                )
            else:
                inner_lbl = ''

            otb_div = (
                f'<div style="height:{otb_h}px;background:{otb_color};'
                f'border-radius:{"3px 3px 0 0" if fcst_h == 0 else "0"};'
                f'margin:0 4px;position:relative;">'
                + inner_lbl +
                '</div>'
            )

            bar_cells += (
                f'<div style="{bg}border-right:1px solid {_S["bdr"]};padding:0 2px;">'
                # ── Row 1: pickup annotation above bar (small, only if pickup > 0) ──
                f'<div style="height:18px;display:flex;align-items:center;'
                f'justify-content:center;text-align:center;">{above_lbl}</div>'
                # ── Row 2: bar area (max_bar_h px, bars bottom-anchored) ──
                f'<div style="height:{max_bar_h}px;display:flex;flex-direction:column;'
                f'justify-content:flex-end;">'
                + fcst_div + otb_div +
                '</div>'
                '</div>'
            )

        # ── Data rows helper ──
        def data_row(label, label_color, values_fn):
            cells = (
                f'<div style="background:#314B63 !important;color:#ffffff !important;font-family:DM Mono,monospace;font-size:9px;' +
                f'letter-spacing:0.1em;text-transform:uppercase;padding:5px 4px;border-right:1px solid {_S["bdr"]};' +
                'display:flex;align-items:center;justify-content:center;text-align:center;line-height:1.4;">' +
                f'{label}</div>'
            )
            for c in cols_data:
                val, color = values_fn(c)
                bg = f"background:{_S['bg_bar_wknd']};" if c["is_wknd"] else ""
                cells += (
                    f'<div style="{bg}text-align:center;font-family:DM Mono,monospace;font-size:12px;' +
                    f'color:{color};padding:5px 2px;border-right:1px solid {_S["bdr2"]};font-weight:700;">{val}</div>'
                )
            return f'<div style="display:grid;grid-template-columns:{grid_cols};border-top:1px solid {_S["bdr"]};">{cells}</div>'

        def _pu_color_t(v):
            return _S["txt_teal"] if v > 0 else (_S["txt_red"] if v < 0 else _S["txt_dim"])
        p1_row   = data_row("OVNT<br>PKUP", _S["txt_accent"], lambda c: (_pu_fmt(c["p1"]), _pu_color_t(c["p1"])))
        p7_row   = data_row("7-DAY<br>PKUP", _S["txt_accent"], lambda c: (_pu_fmt(c["p7"]), _pu_color_t(c["p7"])))
        rate_row = data_row("BAR<br>RATE",  _S["txt_accent"], lambda c: (c["rate"], _S["txt_rate"]))

        # ── Group Rate row — highest group rate for that day, blank if no group ──
        grp_label = (
            f'<div style="background:#2a4a5e !important;color:#ffffff !important;-webkit-text-fill-color:#ffffff !important;'
            f'font-family:DM Mono,monospace;font-size:8px;letter-spacing:0.1em;text-transform:uppercase;'
            f'padding:5px 4px;border-right:1px solid {_S["bdr"]};'
            f'display:flex;align-items:center;justify-content:center;text-align:center;line-height:1.4;">'
            f'GRP<br>RATE</div>'
        )
        grp_cells = ""
        for c in cols_data:
            bg  = f"background:{_S['bg_bar_wknd']};" if c["is_wknd"] else f"background:{_S['bg_main']};"
            grp_rt = c.get("grp_rate")
            if grp_rt is not None:
                disp  = f"${grp_rt:,.0f}"
                color = "#2E618D"
                weight = "700"
            else:
                disp  = "—"
                color = _S["txt_dim"]
                weight = "400"
            grp_cells += (
                f'<div style="{bg}padding:5px 2px;border-right:1px solid {_S["bdr2"]};'
                f'text-align:center;font-family:DM Mono,monospace;font-size:12px;'
                f'font-weight:{weight};color:{color};">{disp}</div>'
            )
        grp_rate_row = (
            f'<div style="display:grid;grid-template-columns:{grid_cols};border-top:1px solid {_S["bdr"]};">'
            f'{grp_label}{grp_cells}</div>'
        )

        # ── Rate Change row — persists for the current day, auto-resets next day ──
        import json as _json
        _rc_path  = cfg.get_hotel_folder(hotel) / "rate_change_snap.json"
        _today_s  = today.strftime("%Y-%m-%d")
        _rc_hotel = hotel.get("id","hotel")

        def _rc_load():
            try:
                if _rc_path.exists():
                    d = _json.loads(_rc_path.read_text(encoding="utf-8"))
                    if d.get("date") == _today_s and d.get("hotel") == _rc_hotel:
                        return d.get("values", {})
            except Exception:
                pass
            return {}

        def _rc_save(vals):
            try:
                _rc_path.write_text(_json.dumps(
                    {"date": _today_s, "hotel": _rc_hotel, "values": vals},
                    ensure_ascii=False), encoding="utf-8")
            except Exception:
                pass

        # Load today's saved values (stale entries from yesterday are ignored)
        _rc_vals = _rc_load()

        # Build the rate change row with saved values displayed as static HTML
        rc_label = (
            f'<div style="background:#314B63 !important;color:#ffffff !important;font-family:DM Mono,monospace;font-size:8px;'
            f'letter-spacing:0.1em;text-transform:uppercase;padding:5px 4px;border-right:1px solid {_S["bdr"]};'
            f'display:flex;align-items:center;justify-content:center;text-align:center;line-height:1.4;">'
            f'RATE<br>CHANGE</div>'
        )
        rc_cells = ""
        for ci, c in enumerate(cols_data):
            bg  = f"background:{_S['bg_bar_wknd']};" if c["is_wknd"] else f"background:{_S['bg_main']};"
            val = _rc_vals.get(str(ci), "")
            disp = val if val else "—"
            color = _S["txt_orange"] if val else _S["txt_dim"]
            rc_cells += (
                f'<div style="{bg}padding:5px 2px;border-right:1px solid {_S["bdr2"]};'
                f'text-align:center;font-family:DM Mono,monospace;font-size:12px;'
                f'font-weight:700;color:{color};">{disp}</div>'
            )
        rate_change_row = (
            f'<div style="display:grid;grid-template-columns:{grid_cols};border-top:1px solid {_S["bdr"]};">'
            f'{rc_label}{rc_cells}</div>'
        )

        center_html = (
            f'<div style="background:{_S["bg_card"]};border:1px solid {_S["bdr"]};border-radius:8px;overflow:hidden;">' +
            f'<div style="display:grid;grid-template-columns:{grid_cols};border-bottom:1px solid {_S["bdr"]};">' +
            f'<div style="background:{_S["bg_header"]};"></div>' +
            date_cells +
            '</div>' +
            f'<div style="display:grid;grid-template-columns:{grid_cols};">' +
            f'<div style="background:{_S["bg_bar_std"]};border-right:1px solid {_S["bdr"]};display:flex;align-items:flex-end;justify-content:center;padding-bottom:6px;height:'+str(max_bar_h+20)+'px;">' +
            '<span style="font-family:DM Mono,monospace;font-size:8px;color:'+_S["txt_muted"]+';text-transform:uppercase;writing-mode:vertical-rl;transform:rotate(180deg);">OTB</span>' +
            '</div>' +
            bar_cells +
            '</div>' +
            p1_row + p7_row + rate_row + grp_rate_row + rate_change_row +
            '</div>'
        )



    # ── Rate Changes Today — full width, above the 3-col layout ────────────
    if cols_data:
        _hid = hotel.get("id","hotel")
        st.markdown('<style>.rc-input-zone [data-testid="column"] { padding: 0 2px !important; } .rc-input-zone div[data-testid="stTextInput"] input { text-align:center !important; height:26px !important; padding:3px 4px !important; font-size:11px !important; font-family:"DM Mono",monospace !important; } .rc-input-zone div[data-testid="stTextInput"] { margin-top:0 !important; padding-top:0 !important; } .rc-input-zone { margin-top:0.8rem !important; margin-bottom:0.8rem !important; }</style>', unsafe_allow_html=True)
        st.markdown('<div class="rc-input-zone">', unsafe_allow_html=True)
        input_cols = st.columns(len(cols_data))
        changed = False
        for ci2, (ecol, c2) in enumerate(zip(input_cols, cols_data)):
            is_wknd = c2["is_wknd"]
            color   = _S["txt_wknd"] if is_wknd else "#1e2d35"
            lbl_html = (
                f'<div style="text-align:center;font-family:DM Mono,monospace;font-size:13px;' +
                f'font-weight:700;color:{color};letter-spacing:0.05em;line-height:1.1;margin-bottom:4px;">' +
                f'{c2["lbl1"]}<br><span style="font-size:11px;font-weight:400;">{c2["lbl2"] if c2["lbl2"] else "&nbsp;"}</span></div>'
            )
            ecol.markdown(lbl_html, unsafe_allow_html=True)
            cur = _rc_vals.get(str(ci2), "")
            nv  = ecol.text_input("", value=cur,
                                  placeholder="—",
                                  key=f"rc_{_hid}_{ci2}",
                                  label_visibility="collapsed",
                                  max_chars=8)
            if nv != cur:
                _rc_vals[str(ci2)] = nv
                changed = True
        st.markdown('</div>', unsafe_allow_html=True)
        if changed:
            _rc_save(_rc_vals)
            st.rerun()

        # ── Render 3-col layout ──────────────────────────────────────────────────
    col_l, col_c, col_r = st.columns([2.2, 5.5, 2.0])

    with col_l:
        st.markdown(left_html, unsafe_allow_html=True)

    with col_c:
        if center_html:
            st.markdown(center_html, unsafe_allow_html=True)
        else:
            st.markdown("<p style='color:#6a8090;padding:20px'>No forecast data.</p>",
                        unsafe_allow_html=True)

        # ── Monthly Forecast vs Budget sits directly under center chart ──────
        st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)

        def _safe(v, fmt="dollar"):
            if v is None or (isinstance(v, float) and math.isnan(v)): return "—"
            if fmt == "dollar": return f"${v:,.0f}"
            if fmt == "num":    return f"{v:,.0f}"
            return str(v)

        if monthly_var:
            header_cells = "".join(
                f'<th style="padding:4px 10px;text-align:{"left" if h=="Month" else "center"};color:{_S["txt_col_hdr"]};font-family:DM Mono,monospace;font-size:11px;letter-spacing:0.1em;text-transform:uppercase;white-space:nowrap;border-right:1px solid {_S["bdr"]};">{h}</th>'
                for h in ["Month", "Fcst Rms", "Fcst ADR", "Fcst Rev", "Bud Rms", "Bud ADR", "Bud Rev", "Variance"]
            )
            col_keys = ["Month","Fcst Rms","Fcst ADR","Fcst Rev","Bud Rms","Bud ADR","Bud Rev"]
            tbody = ""
            for i, mv in enumerate(monthly_var):
                mo_label = pd.Period(mv["Month"]).strftime("%B %Y")
                var = mv["Var_Rev"]
                if var is not None and not (isinstance(var, float) and math.isnan(var)):
                    var_color = "#2E618D" if var >= 0 else "#a03020"
                    var_cell  = f'<td style="padding:4px 10px;text-align:center;font-size:14px;border-right:1px solid #d0dde2;"><span style="color:{var_color};font-weight:600">${var:+,.0f}</span></td>'
                else:
                    var_cell  = '<td style="padding:4px 10px;text-align:center;font-size:14px;color:#5a7080;border-right:1px solid #d0dde2;">—</td>'
                row_data = {
                    "Month":     mo_label,
                    "Fcst Rms":  _safe(mv["TY_Rooms"], "num"),
                    "Fcst ADR":  _safe(mv["TY_ADR"]),
                    "Fcst Rev":  _safe(mv["TY_Rev"]),
                    "Bud Rms":   _safe(mv["Bu_Rooms"], "num"),
                    "Bud ADR":   _safe(mv["Bu_ADR"]),
                    "Bud Rev":   _safe(mv["Bu_Rev"]),
                }
                bg = _S["bg_row_alt"] if i % 2 == 0 else _S["bg_row_std"]
                cells = ""
                for k in col_keys:
                    align = "left" if k == "Month" else "center"
                    color = _S["txt_pri"] if k == "Month" else _S["txt_sec"]
                    cells += f'<td style="padding:4px 10px;text-align:{align};color:{color};font-size:14px;border-right:1px solid {_S["bdr2"]};white-space:nowrap;">{row_data[k]}</td>'
                tbody += f'<tr style="background:{bg};">{cells}{var_cell}</tr>'

            mv_html = (
                f'<div style="background:{_S["bg_card"]};border:1px solid {_S["bdr"]};border-radius:8px;overflow:hidden;">'
                f'<div style="background:{_S["bg_monthly_hdr2"]} !important;padding:6px 10px;">'
                f'<span style="font-family:DM Mono,monospace;font-size:11px;letter-spacing:0.15em;text-transform:uppercase;color:{_S["txt_header"]};">Monthly Forecast vs Budget</span>'
                '</div>'
                '<table style="width:100%;border-collapse:collapse;">'
                f'<thead><tr style="background:{_S["bg_monthly_hdr"]};border-bottom:1px solid {_S["bdr"]};">{header_cells}</tr></thead>'
                f'<tbody>{tbody}</tbody>'
                '</table></div>'
            )
            st.markdown(mv_html, unsafe_allow_html=True)
        else:
            st.markdown('<div style="background:#ffffff;border:1px solid #c4cfd4;padding:10px;font-size:12px;color:#6a8090;border-radius:8px">No forecast/budget data available.</div>', unsafe_allow_html=True)

    with col_r:
        st.markdown(right_html, unsafe_allow_html=True)

    # ── Row 2: Comp Set Rates — next 15 nights, full width ───────────────────
    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    if rates_vs7 is not None and not rates_vs7.empty:
        # Build list of hotels from column names (format: HotelName__rate / HotelName__change)
        rate_cols   = {c.replace("__rate","") for c in rates_vs7.columns if c.endswith("__rate")}
        change_cols = {c.replace("__change","") for c in rates_vs7.columns if c.endswith("__change")}
        hotels_vs7  = sorted(rate_cols & change_cols)

        # Also get our rate from rates_overview
        our_rates_map = {}
        our_change_map = {}
        if rates_ov is not None:
            # our rate from overview
            for _, rr in rates_ov.iterrows():
                our_rates_map[pd.Timestamp(rr["Date"])] = rr.get("Our_Rate")

        # Get "Hampton (Us)" change from vs7 if column exists (our hotel column)
        our_col = None
        for col in rates_vs7.columns:
            if col.endswith("__rate") and "(us)" in col.lower():
                our_col = col.replace("__rate","")
                break
        # Fallback: any column with "(us" in it
        if our_col is None:
            for col in rates_vs7.columns:
                if col.endswith("__rate") and "(us" in col.lower():
                    our_col = col.replace("__rate","")
                    break

        # Build 15-night window from vs7 dates
        rate_dates_fwd = sorted([d for d in rates_vs7["Date"].unique() if pd.Timestamp(d) >= today])[:15]

        # Lookup helpers
        def get_vs7(dt, hotel, field):
            col = f"{hotel}__{field}"
            if col not in rates_vs7.columns: return None
            row = rates_vs7[rates_vs7["Date"] == pd.Timestamp(dt)]
            if row.empty: return None
            v = row.iloc[0][col]
            return v if pd.notna(v) else None

        def _day_hdr(dt):
            t = pd.Timestamp(dt)
            is_td   = t.date() == today.date()
            is_wknd = t.strftime("%a") in ("Fri","Sat")
            color   = _S["txt_today"] if is_td else (_S["txt_wknd"] if is_wknd else _S["txt_sec"])
            bg      = f"background:{_S['bg_today']};" if is_td else (f"background:{_S['bg_wknd']};" if is_wknd else "")
            return (
                f'<th style="padding:5px 4px;text-align:center;{bg}border-right:1px solid {_S["bdr"]};min-width:52px;">'
                f'<div style="font-family:DM Mono,monospace;font-size:11px;color:{color};font-weight:700;">{t.strftime("%a")}</div>'
                f'<div style="font-family:DM Mono,monospace;font-size:11px;color:{color};opacity:0.85;">{t.strftime("%b") + " " + str(t.day)}</div>'
                f'</th>'
            )

        def _rate_cell(rate_val, change_val, is_our=False, dt=None):
            t = pd.Timestamp(dt) if dt is not None else None
            is_td   = t is not None and t.date() == today.date()
            is_wknd = t is not None and t.strftime("%a") in ("Fri","Sat")
            bg      = f"background:{_S['bg_bar_wknd']};" if is_wknd else ""
            bold    = "font-weight:700;" if is_our else "font-weight:400;"
            font_sz = "15px" if is_our else "14px"
            # Rate display
            if rate_val is None or (isinstance(rate_val, float) and math.isnan(rate_val)):
                rate_disp  = "—"
                rate_color = _S["txt_dim"]
                chg_html   = ""
            else:
                try:
                    rate_disp  = f"${int(float(rate_val)):,}"
                except (ValueError, TypeError):
                    rate_str = str(rate_val).strip()
                    rate_disp = "SOLD" if "sold" in rate_str.lower() else rate_str[:6]
                rate_color = _S["txt_our"] if is_our else _S["txt_comp"]
                # Inline change delta — teal=up, red=down
                chg_html = ""
                if change_val is not None and not (isinstance(change_val, float) and math.isnan(change_val)):
                    try:
                        chg = int(float(str(change_val).split("[")[0].strip()))
                        if chg > 0:
                            chg_html = (f'<span style="color:{_S["txt_teal"]} !important;font-size:12px;'
                                        f'font-weight:600;margin-left:3px;">+{chg}</span>')
                        elif chg < 0:
                            chg_html = (f'<span style="color:#cc2200 !important;font-size:12px;'
                                        f'font-weight:600;margin-left:3px;">{chg}</span>')
                    except (ValueError, TypeError):
                        pass
            cell_color = rate_color
            # Build final content: rate in cell_color, change in its own color
            # Use explicit font color on the rate span to prevent inheritance issues
            if chg_html:
                cell_content = (f'<span style="color:{cell_color} !important;">{rate_disp}</span>'
                                + chg_html)
            else:
                cell_content = rate_disp
            return (
                f'<td style="{bg}{bold}padding:6px 8px;text-align:center;'
                f'font-family:DM Mono,monospace;font-size:{font_sz};'
                f'border-right:1px solid {_S["bdr2"]};white-space:nowrap;">'
                f'{cell_content}</td>'
            )

        date_hdrs = "".join(_day_hdr(d) for d in rate_dates_fwd)

        # Our hotel row (rate from overview, change from vs7 if available)
        _our_label = "Our Hotel"
        our_cells = (f'<td style="padding:6px 12px;background:{_S["bg_our_row"]};font-family:DM Mono,monospace;font-size:13px;color:{_S["txt_our"]};font-weight:700;border-right:2px solid {_S["bdr"]};white-space:nowrap;letter-spacing:0.03em;text-align:center;">{_our_label}</td>')
        for d in rate_dates_fwd:
            our_r = our_rates_map.get(pd.Timestamp(d))
            our_c = get_vs7(d, our_col, "change") if our_col else None
            our_cells += _rate_cell(our_r, our_c, is_our=True, dt=d)
        our_row = f'<tr style="background:{_S["bg_our_row"]};border-bottom:2px solid {_S["bdr"]};">{our_cells}</tr>'

        # Comp rows
        # Exclude our own hotel column from comp rows
        comp_only = [h for h in hotels_vs7 if our_col is None or h != our_col]
        comp_rows = ""
        for i, h in enumerate(comp_only):
            bg_row = _S["bg_row_alt"] if i % 2 == 0 else _S["bg_row_std"]
            cells  = f'<td style="padding:5px 10px;background:{_S["bg_comp_name"]};font-family:DM Mono,monospace;font-size:12px;color:{_S["txt_comp"]};font-weight:600;border-right:1px solid {_S["bdr"]};white-space:nowrap;text-align:center;">{h}</td>'
            for d in rate_dates_fwd:
                rv = get_vs7(d, h, "rate")
                cv = get_vs7(d, h, "change")
                cells += _rate_cell(rv, cv, is_our=False, dt=d)
            comp_rows += f'<tr style="background:{bg_row};">{cells}</tr>'

        comp_html = (
            f'<div style="background:{_S["bg_card"]};border:1px solid {_S["bdr"]};border-radius:8px;overflow:hidden;">'
            f'<div style="background:{_S["bg_header"]} !important;padding:7px 12px;">'
            f'<span style="font-family:DM Mono,monospace;font-size:11px;letter-spacing:0.15em;text-transform:uppercase;color:{_S["txt_header"]};">Comp Set Rates — Next 15 Nights</span>'
            f'<span style="font-family:DM Mono,monospace;font-size:10px;color:{_S["txt_muted"]};float:right;padding-top:1px;">rate change vs 7 days ago</span>'
            '</div>'
            '<div style="overflow-x:auto;">'
            '<table style="width:100%;border-collapse:collapse;min-width:900px;">'
            f'<thead><tr style="border-bottom:1px solid {_S["bdr"]};">'
            f'<th style="padding:5px 10px;text-align:center;color:{_S["txt_header"]};font-family:DM Mono,monospace;font-size:9px;letter-spacing:0.1em;text-transform:uppercase;border-right:1px solid {_S["bdr"]};white-space:nowrap;background:{_S["bg_comp_name"]};min-width:120px;"></th>'
            f'{date_hdrs}'
            '</tr></thead>'
            f'<tbody>{our_row}{comp_rows}</tbody>'
            '</table></div></div>'
        )
        st.markdown(comp_html, unsafe_allow_html=True)

    elif rates_comp is not None and not rates_comp.empty:
        # Fallback: use rates_comp without change indicators
        rate_dates_fwd = sorted([d for d in rates_comp["Date"].unique() if pd.Timestamp(d) >= today])[:15]
        all_hotels = sorted(rates_comp["Hotel"].unique())
        our_rates_map = {}
        if rates_ov is not None:
            for _, rr in rates_ov.iterrows():
                our_rates_map[pd.Timestamp(rr["Date"])] = rr.get("Our_Rate")

        def _day_hdr2(dt):
            t = pd.Timestamp(dt)
            is_td = t.date() == today.date()
            is_wknd = t.strftime("%a") in ("Fri","Sat")
            color = _S["txt_today"] if is_td else (_S["txt_wknd"] if is_wknd else _S["txt_sec"])
            bg = f"background:{_S['bg_today']};" if is_td else (f"background:{_S['bg_wknd']};" if is_wknd else "")
            return (f'<th style="padding:5px 4px;text-align:center;{bg}border-right:1px solid {_S["bdr"]};min-width:52px;">'
                    f'<div style="font-family:DM Mono,monospace;font-size:11px;color:{color};font-weight:700;">{t.strftime("%a")}</div>'
                    f'<div style="font-family:DM Mono,monospace;font-size:11px;color:{color};opacity:0.85;">{t.strftime("%b") + " " + str(t.day)}</div></th>')

        def _rc2(rv, is_our=False, dt=None):
            t = pd.Timestamp(dt) if dt else None
            is_td = t is not None and t.date() == today.date()
            is_wknd = t is not None and t.strftime("%a") in ("Fri","Sat")
            bg = f"background:{_S['bg_bar_wknd']};" if is_wknd else ""
            bold = "font-weight:700;" if is_our else ""
            if rv is None or (isinstance(rv, float) and math.isnan(rv)):
                disp = "—"; color = _S["txt_dim"]
            elif isinstance(rv, str):
                disp = "SOLD" if "sold" in rv.lower() else rv[:6]
                color = _S["txt_our"] if is_our else _S["txt_comp"]
            else:
                try:
                    disp = f"${int(float(rv)):,}"
                except (ValueError, TypeError):
                    disp = "—"
                color = _S["txt_our"] if is_our else _S["txt_comp"]
            return f'<td style="{bg}{bold}padding:5px 6px;text-align:center;font-family:DM Mono,monospace;font-size:12px;color:{color};border-right:1px solid {_S["bdr2"]};">{disp}</td>'

        date_hdrs2 = "".join(_day_hdr2(d) for d in rate_dates_fwd)
        _fb_our_col = next((h for h in all_hotels if "(us)" in h.lower()), "Our Hotel")
        our_cells2 = f'<td style="padding:5px 10px;background:{_S["bg_our_row"]};font-family:DM Mono,monospace;font-size:11px;color:{_S["txt_our"]};font-weight:700;border-right:1px solid {_S["bdr"]};white-space:nowrap;text-align:center;">{_fb_our_col}</td>'
        for d in rate_dates_fwd:
            our_cells2 += _rc2(our_rates_map.get(pd.Timestamp(d)), is_our=True, dt=d)
        our_row2 = f'<tr style="background:{_S["bg_our_row"]};border-bottom:2px solid {_S["bdr"]};">{our_cells2}</tr>'
        comp_rows2 = ""
        for i, h in enumerate(all_hotels):
            bg_row = _S["bg_row_alt"] if i % 2 == 0 else _S["bg_row_std"]
            cells  = f'<td style="padding:5px 10px;background:{_S["bg_comp_name"]};font-family:DM Mono,monospace;font-size:12px;color:{_S["txt_comp"]};font-weight:600;border-right:1px solid {_S["bdr"]};white-space:nowrap;text-align:center;">{h}</td>'
            for d in rate_dates_fwd:
                row_match = rates_comp[(rates_comp["Date"] == pd.Timestamp(d)) & (rates_comp["Hotel"] == h)]
                rv = row_match.iloc[0]["Rate"] if not row_match.empty else None
                cells += _rc2(rv, dt=d)
            comp_rows2 += f'<tr style="background:{bg_row};">{cells}</tr>'
        comp_html2 = (
            f'<div style="background:{_S["bg_card"]};border:1px solid {_S["bdr"]};border-radius:8px;overflow:hidden;">'
            f'<div style="background:{_S["bg_header"]} !important;padding:7px 12px;">'
            f'<span style="font-family:DM Mono,monospace;font-size:11px;letter-spacing:0.15em;text-transform:uppercase;color:{_S["txt_header"]};">Comp Set Rates — Next 15 Nights</span>'
            '</div><div style="overflow-x:auto;">'
            '<table style="width:100%;border-collapse:collapse;min-width:900px;">'
            f'<thead><tr style="border-bottom:1px solid {_S["bdr"]};">'
            f'<th style="padding:5px 10px;text-align:center;color:{_S["txt_header"]};font-family:DM Mono,monospace;font-size:9px;border-right:1px solid {_S["bdr"]};white-space:nowrap;background:{_S["bg_comp_name"]};min-width:120px;"></th>'
            f'{date_hdrs2}</tr></thead>'
            f'<tbody>{our_row2}{comp_rows2}</tbody>'
            '</table></div></div>'
        )
        st.markdown(comp_html2, unsafe_allow_html=True)


def render_dashboard_tab(data, hotel):
    import calendar as _cal, sqlite3 as _sq3
    year_df     = data.get("year")
    budget_df   = data.get("budget")
    pickup7     = data.get("pickup7")
    groups_df   = data.get("groups")
    tonight     = data.get("tonight") or {}
    rates_ov    = data.get("rates_overview")
    rates_comp  = data.get("rates_comp")
    total_rooms = hotel.get("total_rooms", 112)
    hotel_id    = hotel.get("id", "hotel")

    # ── Notes DB (SQLite, WAL mode = concurrent-read safe, trivial Postgres swap for web) ──
    _db_path = str(cfg.get_hotel_folder(hotel) / "revpar_notes.db")

    def _db_conn():
        con = _sq3.connect(_db_path, timeout=5)
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("""CREATE TABLE IF NOT EXISTS cal_notes (
                hotel_id TEXT NOT NULL, date TEXT NOT NULL,
                note TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL,
                PRIMARY KEY (hotel_id, date))""")
        con.commit()
        return con

    def _load_note(ds):
        try:
            con = _db_conn(); row = con.execute(
                "SELECT note FROM cal_notes WHERE hotel_id=? AND date=?", (hotel_id, ds)).fetchone()
            con.close(); return row[0] if row else ""
        except: return ""

    def _save_note(ds, txt):
        try:
            from datetime import datetime as _dt
            con = _db_conn()
            con.execute("""INSERT INTO cal_notes (hotel_id,date,note,updated_at) VALUES(?,?,?,?)
                ON CONFLICT(hotel_id,date) DO UPDATE SET note=excluded.note,updated_at=excluded.updated_at""",
                (hotel_id, ds, txt, _dt.now().isoformat(timespec="seconds")))
            con.commit(); con.close()
        except: pass

    def _all_notes():
        try:
            con = _db_conn()
            rows = con.execute("SELECT date,note FROM cal_notes WHERE hotel_id=? AND note!=''",
                               (hotel_id,)).fetchall()
            con.close(); return {r[0]: r[1] for r in rows}
        except: return {}

    # ── Session state ─────────────────────────────────────────────────────────
    _sk = f"cal_sel_{hotel_id}"
    _clicked_flag = f"_cal_clicked_{hotel_id}"
    today_date = datetime.now().date()
    # Always show today unless the user has explicitly clicked a calendar tile
    if not st.session_state.get(_clicked_flag, False):
        st.session_state[_sk] = today_date
    elif _sk not in st.session_state:
        st.session_state[_sk] = today_date

    # ── Build per-date data lookup ────────────────────────────────────────────
    day_data = {}
    if year_df is not None:
        for _, r in year_df.iterrows():
            dt = r.get("Date")
            if dt is None or pd.isna(dt): continue
            cap = r.get("Capacity") or total_rooms
            otb_r = r.get("OTB")
            # ADR: Hilton uses ADR_OTB, IHG uses ADR
            adr_o = r.get("ADR_OTB") if pd.notna(r.get("ADR_OTB", None)) else r.get("ADR")
            rev_o = r.get("Revenue_OTB")
            if rev_o is None or pd.isna(rev_o):
                rev_o = (float(otb_r)*float(adr_o)) if (otb_r and adr_o and pd.notna(otb_r) and pd.notna(adr_o)) else None
            fc_r = r.get("Forecast_Rooms")
            # ADR for forecast: Hilton uses ADR_Forecast, IHG uses ADR
            adr_f = r.get("ADR_Forecast") if pd.notna(r.get("ADR_Forecast", None)) else r.get("ADR")
            # Revenue for forecast: Hilton uses Revenue_Forecast, IHG uses Forecast_Rev
            rev_f = r.get("Revenue_Forecast") if pd.notna(r.get("Revenue_Forecast", None)) else r.get("Forecast_Rev")
            if rev_f is None or pd.isna(rev_f):
                rev_f = (float(fc_r)*float(adr_f)) if (fc_r and adr_f and pd.notna(fc_r) and pd.notna(adr_f)) else None
            rv = fc_r if (fc_r is not None and pd.notna(fc_r)) else otb_r
            occ = (float(rv)/float(cap)*100) if (rv and cap and pd.notna(rv)) else None
            ad  = adr_f if (adr_f is not None and pd.notna(adr_f)) else adr_o
            ly_r = r.get("OTB_LY"); rev_ly = r.get("Revenue_LY")
            adr_ly = (float(rev_ly)/float(ly_r)) if (ly_r and rev_ly and pd.notna(ly_r) and pd.notna(rev_ly) and float(ly_r)>0) else None
            grp_o = r.get("Group_OTB")
            evt = r.get("Event",""); evt = "" if (evt is None or str(evt).lower() in ("nan","none")) else str(evt)
            day_data[pd.Timestamp(dt).date()] = {
                "occ_pct":   round(float(occ),1)   if occ is not None else None,
                "adr_disp":  round(float(ad),0)    if (ad is not None and pd.notna(ad)) else None,
                "event": evt,
                "otb_rooms": int(round(float(otb_r))) if (otb_r is not None and pd.notna(otb_r)) else None,
                "adr_otb":   round(float(adr_o),2) if (adr_o is not None and pd.notna(adr_o)) else None,
                "rev_otb":   round(float(rev_o),0) if (rev_o is not None and pd.notna(rev_o)) else None,
                "fcst_rooms":int(round(float(fc_r))) if (fc_r is not None and pd.notna(fc_r)) else None,
                "adr_fcst":  round(float(adr_f),2) if (adr_f is not None and pd.notna(adr_f)) else None,
                "rev_fcst":  round(float(rev_f),0) if (rev_f is not None and pd.notna(rev_f)) else None,
                "otb_ly":    int(round(float(ly_r))) if (ly_r is not None and pd.notna(ly_r)) else None,
                "adr_ly":    round(float(adr_ly),2) if adr_ly is not None else None,
                "rev_ly":    round(float(rev_ly),0) if (rev_ly is not None and pd.notna(rev_ly)) else None,
                "grp_otb":   int(round(float(grp_o))) if (grp_o is not None and pd.notna(grp_o) and float(grp_o)>0) else 0,
                "capacity":  int(cap),
            }

    # ── Group lookup ──────────────────────────────────────────────────────────
    grp_lookup = {}
    if groups_df is not None and "Occ_Date" in groups_df.columns:
        for _, gr in groups_df.iterrows():
            ocd = gr.get("Occ_Date")
            if ocd is None or pd.isna(ocd): continue
            d = pd.Timestamp(ocd).date()
            blk=gr.get("Block"); pku=gr.get("Pickup"); rt=gr.get("Rate"); nm=gr.get("Group_Name","—")
            if blk and pd.notna(blk) and float(blk)>0:
                grp_lookup.setdefault(d,[]).append({
                    "name": str(nm), "block": int(float(blk)),
                    "pickup": int(float(pku)) if (pku and pd.notna(pku)) else 0,
                    "rate": float(rt) if (rt and pd.notna(rt)) else None})

    # ── Rate lookups ──────────────────────────────────────────────────────────
    our_rate_map  = {}   # date -> our rate float
    comp_rate_map = {}   # date -> {hotel_short: rate}
    if rates_ov is not None:
        for _, r in rates_ov.iterrows():
            dt = r.get("Date")
            if dt is not None and pd.notna(dt):
                try: our_rate_map[pd.Timestamp(dt).date()] = float(r.get("Our_Rate") or 0) or None
                except: pass
    if rates_comp is not None:
        for _, r in rates_comp.iterrows():
            dt = r.get("Date")
            if dt is not None and pd.notna(dt):
                d = pd.Timestamp(dt).date(); h = str(r.get("Hotel","—")); rv = r.get("Rate")
                if rv is not None and pd.notna(rv):
                    comp_rate_map.setdefault(d, {})[h] = rv
    _all_comp = sorted({h for v in comp_rate_map.values() for h in v})
    _us = next((h for h in _all_comp if "(us)" in h.lower()), "Our Hotel")
    _comp_order = [_us] + [h for h in _all_comp if h != _us]

    # ── Load notes ────────────────────────────────────────────────────────────
    notes_map = _all_notes()

    # ── Color helpers ─────────────────────────────────────────────────────────
    def _occ_col(occ):
        if occ is None:  return "#f4f6f7","#8aa0ad"
        if occ >= 85:    return "#1e2d35","#ffffff"  # darkest slate — white text
        if occ >= 70:    return "#314B63","#ffffff"  # mid-dark slate — white text
        if occ >= 50:    return "#8aa8bc","#ffffff"  # light slate — white text
        return "#c0392b","#ffffff"                   # red only for <50%

    # ── Month list ────────────────────────────────────────────────────────────
    months_list = []
    for off in range(3):
        mo = today_date.month + off
        yr = today_date.year + (mo-1)//12; mo = ((mo-1)%12)+1
        months_list.append((yr,mo))
    max_wk = max(len(_cal.Calendar(firstweekday=6).monthdayscalendar(y,m)) for y,m in months_list)

    sel_date = st.session_state[_sk]

    # ── Page header ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-head">Forecast Calendar</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Click any date to view full detail and add notes. '
                '<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#e74c3c;"></span> event&nbsp;&nbsp;'
                '<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#f1c40f;"></span> group&nbsp;&nbsp;'
                '<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#aaaaaa;border:1px solid #888;"></span> note</div>',
                unsafe_allow_html=True)

    ev = tonight.get("event","")
    if ev and str(ev).lower() not in ("nan","none",""):
        st.markdown(f'<div class="insight-box">📅 <strong>Tonight\'s Event:</strong> {ev}</div>',
                    unsafe_allow_html=True)

    # ── Hidden input FIRST — must exist in DOM before calendar iframe renders ──
    # The calendar iframe's JS tags it, then pickDate() writes to it.
    # Stable placeholder — doesn't change when widget key rotates
    _placeholder = f"__cpick_{hotel_id}__"
    # Visually-hidden pattern — keeps React fiber fully mounted and tracking the input.
    # display:none causes React to ignore synthetic events dispatched to the input;
    # position:absolute + opacity:0 + pointer-events:none hides it visually while
    # leaving it interactive from React's perspective.
    st.markdown(f"""<style>
div[data-testid="stTextInput"]:has(input[placeholder="{_placeholder}"]) {{
    position: absolute !important;
    width: 1px !important;
    height: 1px !important;
    padding: 0 !important;
    margin: -1px !important;
    overflow: hidden !important;
    clip: rect(0,0,0,0) !important;
    white-space: nowrap !important;
    border: 0 !important;
    opacity: 0 !important;
    pointer-events: none !important;
}}
</style>""", unsafe_allow_html=True)
    # Key rotation pattern: incrementing a counter changes the widget key,
    # which forces Streamlit to remount it fresh (empty value). This is the
    # correct way to reset a text_input without the "cannot modify after
    # instantiation" error.
    _counter_key = f"_cpick_n_{hotel_id}"
    if _counter_key not in st.session_state:
        st.session_state[_counter_key] = 0
    _widget_key = f"_cpick_{hotel_id}_{st.session_state[_counter_key]}"

    # Do NOT pass value="" — that would override session_state on every render,
    # wiping any value JS just wrote. Let Streamlit read from session_state via key only.
    # Key rotation (counter) guarantees the widget starts empty on each new key.
    _cal_pick = st.text_input(
        label="date_picker",
        key=_widget_key,
        placeholder=_placeholder,
        label_visibility="collapsed",
    )
    if _cal_pick:
        try:
            st.session_state[_sk] = datetime.strptime(_cal_pick, "%Y-%m-%d").date()
            st.session_state[_clicked_flag] = True
            sel_date = st.session_state[_sk]
        except Exception:
            pass
        # Rotate the key so the widget remounts empty on next run
        st.session_state[_counter_key] += 1
        st.rerun()

    # ── Calendar iframe — tagger runs inside here, same context as pickDate ──
    _cal_height = max_wk * 70 + 120

    def _month_html(yr, mo, n_wk):
        cal  = _cal.Calendar(firstweekday=6).monthdayscalendar(yr, mo)
        mnam = datetime(yr, mo, 1).strftime("%B %Y")
        DOW  = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"]
        dow_row = "".join(
            f'<th style="padding:6px 2px;text-align:center;font-family:DM Mono,monospace;'
            f'font-size:11px;letter-spacing:0.10em;text-transform:uppercase;color:#4e6878;'
            f'font-weight:600;border-bottom:1px solid #c4cfd4;">{d}</th>' for d in DOW)
        body = ""
        for week in cal:
            body += "<tr>"
            for dn in week:
                if dn == 0:
                    body += '<td style="background:#eef2f3;border:1px solid #c4cfd4;height:68px;width:14.28%;"></td>'
                else:
                    cd   = datetime(yr, mo, dn).date()
                    info = day_data.get(cd, {})
                    occ  = info.get("occ_pct"); adr = info.get("adr_disp")
                    evt  = info.get("event","")
                    hgrp = bool(grp_lookup.get(cd))
                    hnot = bool(notes_map.get(str(cd),"").strip())
                    is_td  = (cd==today_date); is_sel = (cd==sel_date)
                    is_past= (cd<today_date)
                    bg,tc  = _occ_col(occ)
                    if is_past: bg="#f0f4f6"; tc="#9ab0ba"
                    if is_td:   bg="#1a3a52"; tc="#ffffff"
                    if is_sel and is_td:    bdr="border:3px solid #7ec8e3;"
                    elif is_sel:            bdr="border:2px solid #ffffffaa;"
                    elif is_td:             bdr="border:3px solid #7ec8e3;"
                    else:                   bdr="border:1px solid #c4cfd4;"
                    dc = "#ffffff" if not is_past else "#9ab0ba"
                    od = f"{occ:.0f}%" if occ is not None else "—"
                    ad = f"${adr:.0f}"  if adr is not None else "—"
                    dots = ""
                    if evt:  dots+='<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:#e74c3c;"></span>'
                    if hgrp: dots+='<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:#f1c40f;margin-left:2px;"></span>'
                    if hnot: dots+='<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:#ffffff;margin-left:2px;border:1px solid #aaa;"></span>'
                    dd = f'<div style="min-height:10px;text-align:center;margin-top:2px;">{dots}</div>'
                    inner = (
                        f'<div onclick="pickDate(\'{cd}\')" '
                        f'style="display:block;padding:5px 4px;height:100%;box-sizing:border-box;cursor:pointer;">'
                        f'<div style="font-family:DM Mono,monospace;font-size:10px;color:{dc};font-weight:700;line-height:1;">{dn}</div>'
                        f'<div style="font-family:Syne,sans-serif;font-size:14px;font-weight:800;color:{tc};line-height:1.2;margin-top:3px;">{od}</div>'
                        f'<div style="font-family:DM Mono,monospace;font-size:12px;color:{tc};opacity:0.85;line-height:1.1;">{ad}</div>'
                        f'{dd}</div>')
                    body += (f'<td style="background:{bg};{bdr}height:68px;width:14.28%;vertical-align:top;padding:0;"'
                             f' onmouseover="this.style.filter=\'brightness(1.3)\'"'
                             f' onmouseout="this.style.filter=\'brightness(1)\'">'
                             f'{inner}</td>')
            body += "</tr>"
        blank = '<td style="background:#eef2f3;border:1px solid #c4cfd4;height:68px;"></td>'
        for _ in range(n_wk - len(cal)):
            body += f"<tr>{blank*7}</tr>"
        legend = (
            '<div style="padding:7px 10px;border-top:1px solid #c4cfd4;display:flex;gap:14px;flex-wrap:wrap;align-items:center;">'
            '<span style="font-family:DM Mono,monospace;font-size:11px;color:#6a8090;display:flex;align-items:center;gap:6px;">'
            '<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#e74c3c;"></span> event&nbsp;'
            '<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#f1c40f;"></span> group&nbsp;'
            '<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#ffffff;border:1px solid #aaa;"></span> note'
            '</span></div>')
        return (
            f'<div style="background:#ffffff;border:1px solid #c4cfd4;border-radius:10px;overflow:hidden;">'
            f'<div style="background:#314B63;padding:9px 14px;border-bottom:1px solid #c4cfd4;">'
            f'<span style="font-family:Syne,sans-serif;font-size:15px;font-weight:700;color:#ffffff !important;-webkit-text-fill-color:#ffffff;">{mnam}</span></div>'
            f'<table style="width:100%;border-collapse:collapse;table-layout:fixed;">'
            f'<thead><tr>{dow_row}</tr></thead><tbody>{body}</tbody></table>{legend}</div>')
    _months_body = (
        _month_html(*months_list[0], max_wk) +
        _month_html(*months_list[1], max_wk) +
        _month_html(*months_list[2], max_wk)
    )

    _cal_iframe = f"""<!DOCTYPE html><html><head><style>
body{{margin:0;padding:0;background:#eef2f3;font-family:'DM Sans',sans-serif;}}
</style></head><body>
<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;padding:2px;">
{_months_body}
</div>
<script>
// pickDate: uses window.parent.document (same origin, confirmed working via live clock).
// A small setTimeout ensures the input is fully mounted before events fire.
var _cachedInp = null;

function findInput() {{
  if (_cachedInp && window.parent.document.body.contains(_cachedInp)) return _cachedInp;
  try {{
    var inp = window.parent.document.querySelector('input[placeholder="__cpick_{hotel_id}__"]');
    if (inp) {{ _cachedInp = inp; return inp; }}
  }} catch(e) {{}}
  return null;
}}

function pickDate(d) {{
  // Small delay so Streamlit's React fiber is ready after any prior rerender
  setTimeout(function() {{
    var inp = findInput();
    if (!inp) return;
    inp.focus();
    var nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    nativeSetter.call(inp, d);
    inp.dispatchEvent(new Event('input',  {{bubbles: true}}));
    inp.dispatchEvent(new Event('change', {{bubbles: true}}));
    setTimeout(function() {{ inp.blur(); }}, 50);
  }}, 30);
}}
</script>
</body></html>"""

    components.html(_cal_iframe, height=_cal_height, scrolling=False)

    # ════════════════════════════════════════════════════════════════════════════
    # DAY DETAIL PANEL
    # ════════════════════════════════════════════════════════════════════════════
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    info     = day_data.get(sel_date, {})
    date_str = str(sel_date)

    def _fp(v):  return f"{v:.1f}%"   if v is not None else "—"
    def _fr(v):  return str(v)          if v is not None else "—"
    def _fa(v):  return f"${v:,.2f}"   if v is not None else "—"
    def _fv(v):  return f"${v:,.0f}"   if v is not None else "—"
    def _dlt(a,b): return (a-b) if (a is not None and b is not None) else None

    sel_lbl = sel_date.strftime("%A, %B %d, %Y")
    sel_dow = sel_date.strftime("%A")
    sel_evt = info.get("event","")
    sel_cap = info.get("capacity", total_rooms)

    otb_r=info.get("otb_rooms"); otb_a=info.get("adr_otb"); otb_v=info.get("rev_otb")
    fc_r =info.get("fcst_rooms");fc_a =info.get("adr_fcst");fc_v =info.get("rev_fcst")
    ly_r =info.get("otb_ly");    ly_a =info.get("adr_ly");  ly_v =info.get("rev_ly")
    otb_occ=(otb_r/sel_cap*100) if (otb_r and sel_cap) else None
    fc_occ =(fc_r /sel_cap*100) if (fc_r  and sel_cap) else None
    ly_occ =(ly_r /sel_cap*100) if (ly_r  and sel_cap) else None

    grp_rows=grp_lookup.get(sel_date,[])
    t_blk=sum(g["block"] for g in grp_rows); t_pku=sum(g["pickup"] for g in grp_rows)
    mx_rt=max((g["rate"] for g in grp_rows if g["rate"] is not None),default=None)
    occ_cv=_occ_col(info.get("occ_pct"))[1] if info else "#6a8090"

    # ── Date header ───────────────────────────────────────────────────────────
    ebadge=(f'<span style="background:#b4482022;border:1px solid #b4482066;color:#b44820;'
            f'font-family:DM Mono,monospace;font-size:10px;padding:3px 10px;border-radius:100px;">'
            f'📅 {sel_evt}</span>') if sel_evt else ""
    st.markdown(
        '<div style="background:#e4eaed;border:1px solid #c4cfd4;border-radius:10px;'
        'padding:12px 18px;margin-bottom:12px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;">'
        f'<span style="font-family:Syne,sans-serif;font-size:18px;font-weight:700;color:#1e2d35;">{sel_lbl}</span>'
        f'<span style="font-family:DM Mono,monospace;font-size:11px;color:#6a8090;letter-spacing:0.12em;">{sel_dow.upper()}</span>'
        f'{ebadge}'
        f'<span style="margin-left:auto;font-family:Syne,sans-serif;font-size:22px;font-weight:800;color:{occ_cv};">'
        f'{_fp(info.get("occ_pct"))} Fcst Occ</span></div>', unsafe_allow_html=True)

    # ── Rate snapshot strip ───────────────────────────────────────────────────
    our_rt = our_rate_map.get(sel_date)
    comp_d = comp_rate_map.get(sel_date, {})
    comp_vals = []
    for v in comp_d.values():
        if v is None:
            continue
        try:
            comp_vals.append(float(v))
        except (TypeError, ValueError):
            pass  # skip "LOS2", "SOLD", or any non-numeric string
    med_comp  = sorted(comp_vals)[len(comp_vals)//2] if comp_vals else None

    if our_rt is not None or comp_d:
        if our_rt is not None and pd.notna(our_rt) and med_comp:
            df2 = our_rt - med_comp
            our_rc = "#2E618D" if df2 < -5 else ("#6A924D" if abs(df2)<=5 else "#b44820")
        else:
            our_rc = "#6A924D"
        our_disp = f"${int(our_rt):,}" if (our_rt is not None and pd.notna(our_rt)) else "—"

        def _rcell(label, val_html, bg="#f8fafb", label_color="#6a8090", min_w="90px"):
            return (
                f'<td style="padding:8px 12px;background:{bg};border-right:1px solid #dce6ea;'
                f'white-space:nowrap;width:{min_w};max-width:{min_w};overflow:hidden;vertical-align:middle;">'
                f'<div style="font-family:DM Mono,monospace;font-size:9px;color:{label_color} !important;-webkit-text-fill-color:{label_color};'
                f'letter-spacing:0.12em;text-transform:uppercase;margin-bottom:3px;'
                f'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{label}">{label}</div>'
                f'{val_html}</td>')

        cells = _rcell(_us,
            f'<b style="font-family:Syne,sans-serif;font-size:22px;font-weight:800;color:#ffffff !important;-webkit-text-fill-color:#ffffff;">{our_disp}</b>',
            bg="#1e2d35", label_color="#7ec8e3", min_w="130px")

        cells += _rcell("OTB · Fcst",
            f'<b style="font-family:DM Mono,monospace;font-size:15px;font-weight:700;color:#7ec8e3 !important;-webkit-text-fill-color:#7ec8e3;">'
            f'{_fr(otb_r)}</b>'
            f'<span style="font-family:DM Mono,monospace;font-size:13px;color:#6a9ab0 !important;-webkit-text-fill-color:#6a9ab0;"> · </span>'
            f'<b style="font-family:DM Mono,monospace;font-size:15px;font-weight:700;color:#90e090 !important;-webkit-text-fill-color:#90e090;">'
            f'{_fr(fc_r)}</b>',
            bg="#1e2d35", min_w="110px")

        # Separator
        cells += '<td style="border-right:2px solid #4a6070;padding:0;width:0;"></td>'

        for hn in _comp_order:
            if hn == _us: continue
            rv = comp_d.get(hn)
            if rv is None:
                rv_disp = "—"; rv_c = "#5a7080"
            else:
                try:
                    ri = int(float(str(rv).split("[")[0].strip()))
                    rv_disp = f"${ri:,}"
                    if our_rt:
                        d3 = our_rt - float(ri)
                        rv_c = "#2E618D" if d3 > 5 else ("#3a5260" if abs(d3)<=5 else "#a03020")
                    else: rv_c = "#3a5260"
                except:
                    s = str(rv).strip()
                    rv_disp = "SOLD" if "sold" in s.lower() else s[:6]
                    rv_c = "#a03020" if "sold" in s.lower() else "#3a5260"
            cells += _rcell(hn,
                f'<b style="font-family:DM Mono,monospace;font-size:18px;font-weight:800;color:{rv_c} !important;-webkit-text-fill-color:{rv_c};">{rv_disp}</b>')

        if med_comp:
            cells += _rcell("Median Comp",
                f'<b style="font-family:DM Mono,monospace;font-size:18px;font-weight:800;color:#1e2d35 !important;-webkit-text-fill-color:#1e2d35;">${int(med_comp):,}</b>',
                min_w="110px")

        st.markdown(
            '<div style="background:#ffffff;border:1px solid #c4cfd4;border-radius:10px;overflow:hidden;margin-bottom:12px;">'
            '<div style="background:#314B63;padding:8px 14px;border-bottom:1px solid #243d52;display:flex;align-items:center;">'
            '<span style="font-family:DM Mono,monospace;font-size:10px;letter-spacing:0.15em;text-transform:uppercase;color:#ffffff !important;-webkit-text-fill-color:#ffffff;font-weight:700;">Rate Snapshot</span>'
            '</div>'
            '<div style="overflow-x:auto;">'
            f'<table style="border-collapse:collapse;width:100%;table-layout:fixed;border-top:1px solid #dce6ea;"><tr>{cells}</tr></table>'
            '</div></div>', unsafe_allow_html=True)

    # ── Three metric panels ───────────────────────────────────────────────────
    def _panel(title, accent, rows):
        th = (f'<div style="background:{accent};border-bottom:1px solid {accent};padding:10px 16px;">'
              f'<span style="font-family:DM Mono,monospace;font-size:10px;letter-spacing:0.15em;'
              f'text-transform:uppercase;color:#ffffff !important;-webkit-text-fill-color:#ffffff;font-weight:700;">{title}</span></div>')
        tb=""
        for i,(lbl,val,vc,delta,delta_s) in enumerate(rows):
            dh=""
            if delta is not None and delta_s is not None:
                dc="#2E618D" if delta>=0 else "#c0392b"; arr="▲" if delta>=0 else "▼"
                dh=f'<span style="-webkit-text-fill-color:{dc};font-size:11px;margin-left:6px;font-weight:600;">{arr} {delta_s}</span>'
            row_bg = "#f4f7f9" if i%2==0 else "#ffffff"
            tb+=(f'<tr style="background:{row_bg};border-bottom:1px solid #e4eaed;">'
                 f'<td style="padding:6px 10px;color:#6a8090 !important;-webkit-text-fill-color:#6a8090;font-family:DM Mono,monospace;'
                 f'font-size:12px;letter-spacing:0.08em;text-transform:uppercase;width:38%;text-align:center;">{lbl}</td>'
                 f'<td style="padding:6px 10px;color:{vc} !important;-webkit-text-fill-color:{vc};font-family:DM Mono,monospace;'
                 f'font-size:13px;font-weight:800;text-align:center;">{val}{dh}</td></tr>')
        return (f'<div style="background:#ffffff;border:1px solid #dce6ea;border-radius:10px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.06);">'
                f'{th}<table style="width:100%;border-collapse:collapse;">{tb}</table></div>')

    c1,c2,c3 = st.columns(3)
    def _dlt_fmt(a, b, fmt="num"):
        if a is None or b is None: return None, None
        d = a - b
        if fmt == "pct":  s = f"{abs(d):.1f}%"
        elif fmt == "adr": s = f"${abs(d):.2f}"
        elif fmt == "rev": s = f"${abs(int(d)):,}"
        else:              s = str(abs(int(d)))
        return d, s
    with c1: st.markdown(_panel("On Books","#1e2d35",[
        ("Occ %",_fp(otb_occ),"#1e2d35",None,None),("Rooms",_fr(otb_r),"#1e2d35",None,None),
        ("ADR",_fa(otb_a),"#6A924D",None,None),("Revenue",_fv(otb_v),"#1e2d35",None,None)]),unsafe_allow_html=True)
    with c2: st.markdown(_panel("Forecast","#314B63",[
        ("Occ %",_fp(fc_occ),"#1e2d35",*_dlt_fmt(fc_occ,ly_occ,"pct")),
        ("Rooms",_fr(fc_r),"#1e2d35",*_dlt_fmt(fc_r,ly_r,"num")),
        ("ADR",_fa(fc_a),"#6A924D",*_dlt_fmt(fc_a,ly_a,"adr")),
        ("Revenue",_fv(fc_v),"#1e2d35",*_dlt_fmt(fc_v,ly_v,"rev"))]),unsafe_allow_html=True)
    with c3: st.markdown(_panel("Last Year Actual","#4a6a84",[
        ("Occ %",_fp(ly_occ),"#1e2d35",None,None),("Rooms",_fr(ly_r),"#1e2d35",None,None),
        ("ADR",_fa(ly_a),"#6A924D",None,None),("Revenue",_fv(ly_v),"#1e2d35",None,None)]),unsafe_allow_html=True)

    # ── Group panel ───────────────────────────────────────────────────────────
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    if grp_rows:
        gh=(
            '<div style="background:#ffffff;border:1px solid #c4cfd4;border-radius:10px;overflow:hidden;">'
            '<div style="background:#d0e8f5;border-bottom:1px solid #a8cfe0;padding:9px 14px;display:flex;align-items:center;gap:16px;">'
            '<span style="font-family:DM Mono,monospace;font-size:10px;letter-spacing:0.15em;text-transform:uppercase;color:#1e2d35 !important;-webkit-text-fill-color:#1e2d35;font-weight:700;">Group Business</span>'
            f'<span style="font-family:DM Mono,monospace;font-size:11px;color:#3a5260;">{t_blk} rooms blocked &nbsp;·&nbsp; {t_pku} picked up</span>')
        if mx_rt: gh+=f'<span style="margin-left:auto;font-family:DM Mono,monospace;font-size:13px;color:#6A924D;font-weight:700;">Highest rate: ${mx_rt:,.0f}</span>'
        gh+='</div><div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr style="border-bottom:1px solid #c4cfd4;">'
        for hd,al in [("Group Name","left"),("Block","center"),("Pickup","center"),("Avail","center"),("Rate","center")]:
            gh+=f'<th style="padding:7px 14px;text-align:{al};font-family:DM Mono,monospace;font-size:10px;letter-spacing:0.1em;text-transform:uppercase;color:#4e6878;">{hd}</th>'
        gh+='</tr></thead><tbody>'
        for i,g in enumerate(sorted(grp_rows,key=lambda x:-(x["rate"] or 0))):
            bg2="#e8eef1" if i%2==0 else "#ffffff"; av=g["block"]-g["pickup"]
            rd=f'${g["rate"]:,.0f}' if g["rate"] is not None else "—"
            rc="#6A924D" if g["rate"] is not None else "#5a7080"
            gh+=(f'<tr style="background:{bg2};border-bottom:1px solid #c4cfd4;">'
                 f'<td style="padding:7px 14px;color:#1e2d35;font-family:DM Mono,monospace;">{g["name"]}</td>'
                 f'<td style="padding:7px 14px;text-align:center;color:#3a5260;font-family:DM Mono,monospace;">{g["block"]}</td>'
                 f'<td style="padding:7px 14px;text-align:center;color:#2E618D;font-family:DM Mono,monospace;">{g["pickup"]}</td>'
                 f'<td style="padding:7px 14px;text-align:center;color:#6a8090;font-family:DM Mono,monospace;">{av}</td>'
                 f'<td style="padding:7px 14px;text-align:center;color:{rc};font-family:DM Mono,monospace;font-weight:700;">{rd}</td></tr>')
        gh+='</tbody></table></div></div>'
        st.markdown(gh,unsafe_allow_html=True)
    else:
        _db_lt2 = False
        _ng_bg = "#ffffff"
        _ng_bdr = "#c4cfd4"
        _ng_col = "#5a7080"
        st.markdown(f'<div style="background:{_ng_bg};border:1px solid {_ng_bdr};border-radius:10px;'
                    f'padding:12px 18px;font-family:DM Mono,monospace;font-size:12px;color:{_ng_col};">'
                    f'No group blocks on this date.</div>',unsafe_allow_html=True)

    # ── Notes panel ───────────────────────────────────────────────────────────
    st.markdown("<div style='height:10px'></div>",unsafe_allow_html=True)
    _db_lt3 = False
    _dn_outer_bg = "#ffffff"
    _dn_outer_bdr= "#c4cfd4"
    _dn_hdr_bg   = "#1e2d35"
    _dn_hdr_bdr  = "#2a3f52"
    _dn_title    = "#ffffff"
    _dn_sub      = "#5a7080"
    st.markdown(
        f'<div style="background:{_dn_outer_bg};border:1px solid {_dn_outer_bdr};border-radius:10px;overflow:hidden;">'
        f'<div style="background:{_dn_hdr_bg};border-bottom:1px solid {_dn_hdr_bdr};padding:9px 14px;">'
        f'<span style="font-family:DM Mono,monospace;font-size:10px;letter-spacing:0.15em;'
        f'text-transform:uppercase;color:{_dn_title} !important;-webkit-text-fill-color:{_dn_title};font-weight:700;">📝 Day Notes</span>'
        f'<span style="font-family:DM Mono,monospace;font-size:10px;color:{_dn_sub};margin-left:12px;">'
        f'Saved across sessions · shared with all users</span></div></div>',unsafe_allow_html=True)

    existing = _load_note(date_str)
    note_in  = st.text_area("Notes",value=existing,height=110,
        placeholder=f"Add notes for {sel_lbl} — rate strategy, demand drivers, action items...",
        key=f"note_{hotel_id}_{date_str}",label_visibility="collapsed")

    sc,cc,_ = st.columns([1,1,6])
    with sc:
        if st.button("💾  Save",key=f"ns_{hotel_id}_{date_str}",type="primary"):
            _save_note(date_str,note_in.strip()); notes_map[date_str]=note_in.strip()
            st.success("Note saved.")
    with cc:
        if st.button("🗑  Clear",key=f"nc_{hotel_id}_{date_str}"):
            _save_note(date_str,""); notes_map.pop(date_str,None); st.rerun()

    # Upcoming notes strip
    upcoming=sorted(((d,n) for d,n in notes_map.items() if n.strip() and d>=str(today_date)),key=lambda x:x[0])[:10]
    if upcoming:
        nr=""
        for ds2,nt in upcoming:
            try:    dl=datetime.strptime(ds2,"%Y-%m-%d").strftime("%a, %b %d")
            except: dl=ds2
            pv=nt[:140]+("…" if len(nt)>140 else "")
            nr+=(f'<tr style="border-bottom:1px solid #c4cfd4;">'
                 f'<td style="padding:6px 14px;white-space:nowrap;font-family:DM Mono,monospace;font-size:11px;color:#4e6878;width:120px;">{dl}</td>'
                 f'<td style="padding:6px 14px;font-family:DM Sans,sans-serif;font-size:12px;color:#3a5260;">{pv}</td></tr>')
        st.markdown(
            '<div style="background:#ffffff;border:1px solid #c4cfd4;border-radius:10px;overflow:hidden;margin-top:10px;">'
            '<div style="background:#314B63;border-bottom:1px solid #253e52;padding:7px 14px;">'
            '<span style="font-family:DM Mono,monospace;font-size:10px;letter-spacing:0.12em;text-transform:uppercase;color:#4e6878;">Upcoming Notes</span></div>'
            f'<table style="width:100%;border-collapse:collapse;"><tbody>{nr}</tbody></table></div>',
            unsafe_allow_html=True)

    # ── SRP Pace for Selected Date ────────────────────────────────────────────
    pace_df = data.get("srp_pace")
    st.markdown("---")
    _pace_date_lbl = sel_date.strftime("%A, %B %d, %Y")
    st.markdown(
        f'<div class="section-head">SRP Pace &nbsp;<span style="font-size:14px;font-weight:400;color:#6a8090;">&middot; {_pace_date_lbl}</span></div>',
        unsafe_allow_html=True)

    if pace_df is None or pace_df.empty:
        st.info("No SRP Pace file found for this hotel.")
    elif data.get("_ihg_hotel"):
        st.info("SRP Pace detail view is not available for IHG hotels on the dashboard.")
    else:
        _sel_ts = pd.Timestamp(sel_date)
        _day_pace = pace_df[pace_df["Date"].dt.normalize() == _sel_ts].copy()

        if _day_pace.empty:
            st.info(f"No SRP Pace data found for {_pace_date_lbl}.")
        else:
            SEGMENTS = ["BAR","CMP","CMTG","CNR","CONS","CONV","DISC","GOV","GT","IT","LNR","MKT","SMRF"]

            _sp_seg_agg = []
            for _seg in SEGMENTS:
                _grp = _day_pace[_day_pace["Segment"] == _seg]
                if _grp.empty:
                    continue
                _otb      = _grp["OTB"].sum()
                _stly_otb = _grp["STLY_OTB"].sum()
                _rev      = _grp["Revenue"].sum()
                _stly_rev = _grp["STLY_Revenue"].sum()
                _adr      = _rev      / _otb      if _otb      > 0 else 0.0
                _stly_adr = _stly_rev / _stly_otb if _stly_otb > 0 else 0.0
                _sp_seg_agg.append({
                    "Segment":      _seg,
                    "OTB":          _otb,
                    "STLY_OTB":     _stly_otb,
                    "Var_OTB":      _otb - _stly_otb,
                    "Revenue":      _rev,
                    "STLY_Revenue": _stly_rev,
                    "Var_Revenue":  _rev - _stly_rev,
                    "ADR":          _adr,
                    "STLY_ADR":     _stly_adr,
                    "Var_ADR":      _adr - _stly_adr,
                })

            _tot_otb      = _day_pace["OTB"].sum()
            _tot_stly_otb = _day_pace["STLY_OTB"].sum()
            _tot_rev      = _day_pace["Revenue"].sum()
            _tot_stly_rev = _day_pace["STLY_Revenue"].sum()
            _tot_adr      = _tot_rev      / _tot_otb      if _tot_otb      > 0 else 0.0
            _tot_stly_adr = _tot_stly_rev / _tot_stly_otb if _tot_stly_otb > 0 else 0.0
            _sp_total = {
                "Segment":      "_TOTAL",
                "OTB":          _tot_otb,      "STLY_OTB":     _tot_stly_otb,
                "Var_OTB":      _tot_otb - _tot_stly_otb,
                "Revenue":      _tot_rev,      "STLY_Revenue": _tot_stly_rev,
                "Var_Revenue":  _tot_rev - _tot_stly_rev,
                "ADR":          _tot_adr,      "STLY_ADR":     _tot_stly_adr,
                "Var_ADR":      _tot_adr - _tot_stly_adr,
            }

            # ── Theme (matches SRP Pace tab) ──────────────────────────────────
            _sp_hdr_bg   = "#dce5e8"; _sp_hdr_col  = "#4e6878"
            _sp_hdr_bdr  = "#c4cfd4"; _sp_cell_bdr = "#d0dde2"
            _sp_tbl_bg   = "#ffffff"; _sp_row_alt  = "#dce5e8"
            _sp_lbl_col  = "#4e6878"; _sp_val_col  = "#1e2d35"
            _sp_tot_bg   = "#2E618D"; _sp_tot_col  = "#ffffff"
            _sp_stly_col = "#2E618D"; _sp_var_pos  = "#556848"
            _sp_var_neg  = "#a03020"; _sp_zero_col = "#6a8090"

            def _sp_fi(v):   return "0" if v == 0 else f"{int(round(v)):,}"
            def _sp_fa(v):   return "$0.00" if v == 0 else f"${v:,.2f}"
            def _sp_fr(v):   return "$0" if v == 0 else f"${v:,.0f}"
            def _sp_fvi(v):
                if v == 0: return "0"
                return ("+%s" if v > 0 else "%s") % f"{int(round(v)):,}"
            def _sp_fva(v):
                if v == 0: return "$0.00"
                return ("+$%s" if v > 0 else "-$%s") % f"{abs(v):,.2f}"
            def _sp_fvr(v):
                if v == 0: return "$0"
                return ("+$%s" if v > 0 else "-$%s") % f"{abs(v):,.0f}"
            def _sp_vc(v):
                return _sp_zero_col if v == 0 else (_sp_var_pos if v > 0 else _sp_var_neg)

            _SP_BASE = (f"padding:7px 10px;font-family:DM Mono,monospace;font-size:13px;"
                        f"text-align:center;border-bottom:1px solid {_sp_cell_bdr};white-space:nowrap;")

            def _sp_td_lbl(text, bg, bold=False, color=None):
                fw = "font-weight:700;" if bold else ""
                col = color or _sp_lbl_col
                return (f'<td style="padding:7px 14px;text-align:center;font-size:13px;{fw}'
                        f'color:{col};background:{bg};border-bottom:1px solid {_sp_cell_bdr};white-space:nowrap;">{text}</td>')
            def _sp_td_val(text, bg, color=None):
                return f'<td style="{_SP_BASE}color:{color or _sp_val_col};background:{bg};">{text}</td>'
            def _sp_td_stly(text, bg):
                return f'<td style="{_SP_BASE}color:{_sp_stly_col};background:{bg};">{text}</td>'
            def _sp_td_var(text, v, bg):
                return f'<td style="{_SP_BASE}color:{_sp_vc(v)};background:{bg};">{text}</td>'

            def _sp_build_row(row, bg, hotel_name="", is_total=False):
                lbl_text = hotel_name if is_total else row["Segment"]
                lbl = _sp_td_lbl(lbl_text, bg, bold=is_total, color=_sp_tot_col if is_total else None)
                if is_total:
                    def _td(text):
                        return (f'<td data-tot="1" style="padding:7px 10px;font-family:DM Mono,monospace;font-size:13px;'
                                f'text-align:center;border-bottom:1px solid {_sp_cell_bdr};'
                                f'white-space:nowrap;background:{bg};color:#ffffff !important;">'
                                f'<b style="color:#ffffff !important;font-weight:700;">{text}</b></td>')
                    return (f"<tr>{lbl}"
                            + _td(_sp_fi(row["OTB"]))        + _td(_sp_fi(row["STLY_OTB"]))   + _td(_sp_fvi(row["Var_OTB"]))
                            + _td(_sp_fa(row["ADR"]))        + _td(_sp_fa(row["STLY_ADR"]))   + _td(_sp_fva(row["Var_ADR"]))
                            + _td(_sp_fr(row["Revenue"]))    + _td(_sp_fr(row["STLY_Revenue"])) + _td(_sp_fvr(row["Var_Revenue"]))
                            + "</tr>")
                return (f"<tr>{lbl}"
                        + _sp_td_val(_sp_fi(row["OTB"]), bg)
                        + _sp_td_stly(_sp_fi(row["STLY_OTB"]), bg)
                        + _sp_td_var(_sp_fvi(row["Var_OTB"]), row["Var_OTB"], bg)
                        + _sp_td_val(_sp_fa(row["ADR"]), bg)
                        + _sp_td_stly(_sp_fa(row["STLY_ADR"]), bg)
                        + _sp_td_var(_sp_fva(row["Var_ADR"]), row["Var_ADR"], bg)
                        + _sp_td_val(_sp_fr(row["Revenue"]), bg)
                        + _sp_td_stly(_sp_fr(row["STLY_Revenue"]), bg)
                        + _sp_td_var(_sp_fvr(row["Var_Revenue"]), row["Var_Revenue"], bg)
                        + "</tr>")

            _TH  = (f"padding:8px 10px;text-align:center;font-size:11px;font-weight:700;"
                    f"letter-spacing:.05em;text-transform:uppercase;color:{_sp_hdr_col};"
                    f"background:{_sp_hdr_bg};border-bottom:2px solid {_sp_hdr_bdr};white-space:nowrap;")
            _THL = (f"padding:8px 14px;text-align:center;font-size:11px;font-weight:700;"
                    f"letter-spacing:.05em;text-transform:uppercase;color:{_sp_hdr_col};"
                    f"background:{_sp_hdr_bg};border-bottom:2px solid {_sp_hdr_bdr};white-space:nowrap;")

            _sp_header = (
                f'<tr><th style="{_THL}" rowspan="2">Segment</th>'
                f'<th style="{_TH}" colspan="3">Occupancy On Books</th>'
                f'<th style="{_TH}" colspan="3">ADR On Books</th>'
                f'<th style="{_TH}" colspan="3">Revenue On Books</th></tr><tr>'
                + "".join(f'<th style="{_TH}">{lbl}</th>' for lbl in
                          ["Current","STLY","Variance","Current","STLY","Variance","Current","STLY","Variance"])
                + "</tr>"
            )

            _sp_hotel_name = hotel.get("display_name", "Hotel Total")
            _sp_total_html = _sp_build_row(_sp_total, _sp_tot_bg, hotel_name=_sp_hotel_name, is_total=True)
            _sp_rows_html  = "".join(
                _sp_build_row(row, _sp_row_alt if i % 2 == 0 else _sp_tbl_bg)
                for i, row in enumerate(_sp_seg_agg)
            )

            _sp_html = f"""
<div style="margin-top:12px;overflow-x:auto;">
  <table style="width:100%;border-collapse:collapse;font-size:13px;background:{_sp_tbl_bg};
                border:1px solid {_sp_hdr_bdr};border-radius:8px;overflow:hidden;">
    <thead>{_sp_header}</thead>
    <tbody>{_sp_total_html}{_sp_rows_html}</tbody>
  </table>
</div>
"""
            st.markdown(_sp_html, unsafe_allow_html=True)






def render_biweekly_tab(data, hotel):
    """Monthly Intelligence — monthly summary with toggle: deltas vs LY or vs Budget."""
    year_df     = data.get("year")
    budget_df   = data.get("budget")
    total_rooms = hotel.get("total_rooms", 112)
    hotel_id    = hotel.get("id", "hotel")
    today       = pd.Timestamp(datetime.now().date())

    # ── Theme palette — must be defined before any HTML building ─────────────
    _mi_lt = False
    _mi_card    = "#ffffff"
    _mi_hdr     = "#314B63"
    _mi_cur     = "#dde8f5"
    _mi_alt     = "#f0f4f6"
    _mi_std     = "#f0f4f8" if _mi_lt else ""
    _mi_fy      = "#e4eaed"
    _mi_qhdr    = "#314B63"
    _mi_qgap    = "#e8eef1"
    _mi_bdr     = "#c4cfd4"
    _mi_bdr2    = "#d0dde2"
    _mi_bdr3    = "#d0dde2"
    _mi_sep     = "#314B63"
    _mi_qlbl    = "#ffffff"
    _mi_note    = "#ffffff"
    _mi_cur_bar = "#2E618D"
    _mi_sec_txt = "#3a5260"

    st.markdown('<div class="section-head">Monthly Intelligence</div>', unsafe_allow_html=True)

    if year_df is None:
        st.markdown('<div class="insight-box">No year data loaded.</div>', unsafe_allow_html=True)
        return

    # ── Toggle: compare vs LY or vs Budget ────────────────────────────────────
    _tog_key = f"mi_toggle_{hotel_id}"
    if _tog_key not in st.session_state:
        st.session_state[_tog_key] = "vs_ly"

    tc1, tc2, tc3 = st.columns([1.2, 1.2, 8])
    with tc1:
        if st.button("Δ vs Last Year",
                     type="primary" if st.session_state[_tog_key] == "vs_ly" else "secondary",
                     use_container_width=True, key=f"mi_btn_ly_{hotel_id}"):
            st.session_state[_tog_key] = "vs_ly"; st.rerun()
    with tc2:
        if st.button("Δ vs Budget",
                     type="primary" if st.session_state[_tog_key] == "vs_bud" else "secondary",
                     use_container_width=True, key=f"mi_btn_bud_{hotel_id}"):
            st.session_state[_tog_key] = "vs_bud"; st.rerun()

    compare_vs = st.session_state[_tog_key]   # "vs_ly" or "vs_bud"

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── Build monthly data ────────────────────────────────────────────────────
    yr = year_df.copy()
    yr["Month"] = yr["Date"].dt.to_period("M")

    bd = None
    if budget_df is not None:
        bd = budget_df.copy()
        bd["Month"] = bd["Date"].dt.to_period("M")

    MONTHS = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]
    QUARTERS = {1:"Q1",2:"Q1",3:"Q1",4:"Q2",5:"Q2",6:"Q2",
                7:"Q3",8:"Q3",9:"Q3",10:"Q4",11:"Q4",12:"Q4"}

    def _s(v, fmt="dollar"):
        if v is None or (isinstance(v, float) and v != v): return "—"
        if fmt == "dollar": return f"${v:,.0f}"
        if fmt == "num":    return f"{int(round(v)):,}"
        if fmt == "pct":    return f"{v:.1f}%"
        return str(v)

    def _dh(ty, comp, fmt="dollar"):
        """Inline delta HTML: TY vs comparison value."""
        if ty is None or comp is None: return ""
        d = ty - comp
        if abs(d) < 0.005: return ""
        color = "#2E618D" if d > 0 else "#a03020"
        arrow = "▲" if d > 0 else "▼"
        if fmt == "dollar": ds = f"${abs(d):,.0f}"
        elif fmt == "pct":  ds = f"{abs(d):.1f}pp"
        else:               ds = f"{int(round(abs(d))):,}"
        return f'<span style="color:{color};font-size:10px;margin-left:4px;">{arrow}{ds}</span>'

    current_mo = pd.Period(today, "M")
    rows_data  = []

    for mn in range(1, 13):
        # Pick the 2026 period for this month (year_df spans 2025-2027)
        target_period = pd.Period(f"2026-{mn:02d}", "M")
        y_mo = yr[yr["Month"] == target_period]
        if y_mo.empty:
            rows_data.append(None); continue

        is_past    = target_period < current_mo
        is_current = target_period == current_mo
        is_future  = target_period > current_mo

        # Days in this month from the data (actual row count)
        mo_days = len(y_mo)
        mo_cap  = total_rooms * mo_days

        # ── TY: past = OTB actuals, current/future = Forecast ────────────────
        if is_past:
            rc = "OTB";            vc = "Revenue_OTB";    ac = "ADR_OTB"
        else:
            rc = "Forecast_Rooms" if "Forecast_Rooms" in y_mo.columns else "OTB"
            vc = "Revenue_Forecast" if "Revenue_Forecast" in y_mo.columns else "Revenue_OTB"
            ac = "ADR_Forecast"    if "ADR_Forecast"    in y_mo.columns else "ADR_OTB"

        # For IHG: neither ADR_OTB nor ADR_Forecast exist — IHG uses plain "ADR"
        if ac not in y_mo.columns and "ADR" in y_mo.columns:
            ac = "ADR"

        # IHG current-month rooms blend: past days use OTB actuals, future days use
        # Forecast_Rooms — mirroring the revenue blend so ADR = ty_rev / ty_rooms
        # is computed from matched numerator and denominator. Without this alignment,
        # ty_rooms = full-month Forecast_Rooms while ty_rev covers only a day or two
        # of Forecast_Rev, collapsing ADR to near-zero (e.g. $67 instead of $871).
        _is_ihg_blend = (
            not is_past
            and "Revenue_Forecast" not in y_mo.columns
            and "Forecast_Rev" in y_mo.columns
            and y_mo["Forecast_Rev"].notna().any()
        )
        if _is_ihg_blend:
            _past_mask  = y_mo["Forecast_Rev"].isna()
            _fut_mask   = y_mo["Forecast_Rev"].notna()
            past_rooms  = float(y_mo.loc[_past_mask, "OTB"].fillna(0).sum()) if "OTB" in y_mo.columns else 0.0
            fut_rooms   = float(y_mo.loc[_fut_mask,  rc].fillna(0).sum())    if rc    in y_mo.columns else 0.0
            ty_rooms    = past_rooms + fut_rooms
            fut_rev     = float(y_mo["Forecast_Rev"].fillna(0).sum())
            past_rev    = float(y_mo.loc[_past_mask, "Revenue_OTB"].fillna(0).sum()) if "Revenue_OTB" in y_mo.columns else 0.0
            ty_rev      = fut_rev + past_rev
        else:
            ty_rooms = float(y_mo[rc].sum()) if (rc in y_mo.columns and y_mo[rc].notna().any()) else 0.0
            ty_rev   = float(y_mo[vc].sum()) if (vc in y_mo.columns and y_mo[vc].notna().any()) else None

        # ADR = revenue / rooms (most accurate); fallback to daily ADR mean
        if ty_rev and ty_rooms > 0:
            ty_adr = ty_rev / ty_rooms
        elif ac in y_mo.columns and y_mo[ac].notna().any():
            ty_adr = float(y_mo[ac].mean())
            ty_rev = ty_adr * ty_rooms if ty_rooms > 0 else None
        else:
            ty_adr = ty_rev = None

        ty_occ    = (ty_rooms / mo_cap * 100) if mo_cap > 0 and ty_rooms else None
        ty_revpar = (ty_rev   / mo_cap)       if mo_cap > 0 and ty_rev   else None

        # ── LY Actuals ────────────────────────────────────────────────────────
        ly_r = float(y_mo["OTB_LY"].sum())     if ("OTB_LY"     in y_mo.columns and y_mo["OTB_LY"].notna().any())     else None
        ly_v = float(y_mo["Revenue_LY"].sum()) if ("Revenue_LY" in y_mo.columns and y_mo["Revenue_LY"].notna().any()) else None
        ly_a = (ly_v / ly_r)                   if (ly_r and ly_v and ly_r > 0)                                        else None
        ly_o = (ly_r / mo_cap * 100)           if (ly_r and mo_cap > 0)                                               else None
        ly_rp= (ly_v / mo_cap)                 if (ly_v and mo_cap > 0)                                               else None

        # ── Budget ────────────────────────────────────────────────────────────
        bu_r = bu_a = bu_v = bu_o = bu_rp = None
        if bd is not None:
            b_mo = bd[bd["Month"] == target_period]
            if not b_mo.empty:
                bu_r = float(b_mo["Occ_Rooms"].sum()) if ("Occ_Rooms" in b_mo.columns and b_mo["Occ_Rooms"].notna().any()) else None
                bu_v = float(b_mo["Revenue"].sum())   if ("Revenue"   in b_mo.columns and b_mo["Revenue"].notna().any())   else None
                # Use RevPAR directly from budget if available
                bu_rp_direct = float(b_mo["RevPAR"].sum()) / mo_days if ("RevPAR" in b_mo.columns and b_mo["RevPAR"].notna().any()) else None
                if bu_r and bu_r > 0:
                    if bu_v:
                        bu_a = bu_v / bu_r
                    elif "ADR" in b_mo.columns and b_mo["ADR"].notna().any():
                        bu_a = float((b_mo["Occ_Rooms"] * b_mo["ADR"]).sum() / bu_r)
                        bu_v = bu_a * bu_r
                    bu_o  = (bu_r / mo_cap * 100) if mo_cap > 0 else None
                    bu_rp = bu_rp_direct if bu_rp_direct else ((bu_v / mo_cap) if bu_v and mo_cap > 0 else None)

        rows_data.append({
            "mn": mn, "name": MONTHS[mn-1], "period": target_period,
            "is_past": is_past, "is_current": is_current, "is_future": is_future,
            "mo_cap": mo_cap,
            "ty_r": ty_rooms if ty_rooms > 0 else None,
            "ty_a": ty_adr, "ty_v": ty_rev, "ty_o": ty_occ, "ty_rp": ty_revpar,
            "ly_r": ly_r,   "ly_a": ly_a,   "ly_v": ly_v,   "ly_o": ly_o,  "ly_rp": ly_rp,
            "bu_r": bu_r,   "bu_a": bu_a,   "bu_v": bu_v,   "bu_o": bu_o,  "bu_rp": bu_rp,
        })

    # ── Full Year totals ──────────────────────────────────────────────────────
    def _tot(key):
        return sum(r[key] for r in rows_data if r and r[key] is not None) or None

    fy_cap  = sum(r["mo_cap"] for r in rows_data if r)
    fy_ty_r = _tot("ty_r"); fy_ty_v = _tot("ty_v")
    fy_ty_a = (fy_ty_v/fy_ty_r) if (fy_ty_v and fy_ty_r) else None
    fy_ty_o = (fy_ty_r/fy_cap*100) if (fy_ty_r and fy_cap) else None
    fy_ty_rp= (fy_ty_v/fy_cap)   if (fy_ty_v and fy_cap) else None
    fy_ly_r = _tot("ly_r"); fy_ly_v = _tot("ly_v")
    fy_ly_a = (fy_ly_v/fy_ly_r) if (fy_ly_v and fy_ly_r) else None
    fy_ly_o = (fy_ly_r/fy_cap*100) if (fy_ly_r and fy_cap) else None
    fy_ly_rp= (fy_ly_v/fy_cap)   if (fy_ly_v and fy_cap) else None
    fy_bu_r = _tot("bu_r"); fy_bu_v = _tot("bu_v")
    fy_bu_a = (fy_bu_v/fy_bu_r) if (fy_bu_v and fy_bu_r) else None
    fy_bu_o = (fy_bu_r/fy_cap*100) if (fy_bu_r and fy_cap) else None
    fy_bu_rp= (fy_bu_v/fy_cap)   if (fy_bu_v and fy_cap) else None

    # ════════════════════════════════════════════════════════════════════════════
    # HTML TABLE
    # Column order: Month | TY (5) | sep | Budget (5) | sep | LY (5)
    # ════════════════════════════════════════════════════════════════════════════
    T  = "#2E618D"
    O  = "#b44820"
    G  = "#6A924D"
    PU = "#556848"
    MU = "#4e6878"
    DI = "#5a7080"
    WH = "#1e2d35"

    # Delta comparison source label for header note
    delta_lbl = "vs Last Year" if compare_vs == "vs_ly" else "vs Budget"
    delta_accent = O if compare_vs == "vs_ly" else PU

    _mi_th_default = "#c8d8e4"  # white on dark navy header in light mode
    def _th(label, color=None, align="center", colspan=1, border_left=False):
        c   = color if color is not None else _mi_th_default
        sp  = f' colspan="{colspan}"' if colspan > 1 else ""
        bl  = f"border-left:2px solid {_mi_sep};" if border_left else ""
        return (f'<th{sp} style="padding:7px 8px;text-align:{align};color:{c};'
                f'background:{_mi_hdr};'
                f'font-family:DM Mono,monospace;font-size:10px;letter-spacing:0.1em;'
                f'text-transform:uppercase;white-space:nowrap;{bl}'
                f'border-right:1px solid {_mi_bdr};border-bottom:1px solid {_mi_bdr};">{label}</th>')

    def _td(val, color=WH, bold=False, delta="", bg="", border_left=False):
        fw = "font-weight:700;" if bold else ""
        bg_s = f"background:{bg};" if bg else (f"background:{_mi_std};" if _mi_lt else "")
        bl   = f"border-left:2px solid {_mi_sep};" if border_left else ""
        return (f'<td style="padding:8px 6px;text-align:center;color:{color};'
                f'font-family:DM Mono,monospace;font-size:13px;{fw}{bg_s}{bl}'
                f'border-right:1px solid {_mi_bdr2};'
                f'border-bottom:1px solid {_mi_bdr3};">{val}{delta}</td>')

    # Header rows
    thead = (
        '<thead>'
        '<tr>'
        + _th("Month", align="left")
        + _th("Actuals / Forecast TY", color="#deeeff",  colspan=5)
        + _th("Budget",                color="#d8f0e0", colspan=5, border_left=True)
        + _th("Last Year Actuals",     color="#ffe0cc", colspan=5, border_left=True)
        + '</tr><tr>'
        + _th("", align="left")
        + _th("Occ %",   color="#b8ccd8") + _th("Rooms",   color="#b8ccd8")
        + _th("ADR",     color="#b8ccd8") + _th("Revenue",  color="#b8ccd8") + _th("RevPAR", color="#b8ccd8")
        + _th("Occ %",   color="#b8ccd8", border_left=True) + _th("Rooms",   color="#b8ccd8")
        + _th("ADR",     color="#b8ccd8") + _th("Revenue",  color="#b8ccd8") + _th("RevPAR", color="#b8ccd8")
        + _th("Occ %",   color="#b8ccd8", border_left=True) + _th("Rooms",   color="#b8ccd8")
        + _th("ADR",     color="#b8ccd8") + _th("Revenue",  color="#b8ccd8") + _th("RevPAR", color="#b8ccd8")
        + '</tr></thead>'
    )

    def _build_row(rd, mo_lbl, is_cur, is_past, is_fut, bg_row=""):
        """Render one data row. rd = dict with ty_*/ly_*/bu_* keys."""
        # Comparison series based on toggle
        if compare_vs == "vs_ly":
            c_o  = rd.get("ly_o");  c_r  = rd.get("ly_r")
            c_a  = rd.get("ly_a");  c_v  = rd.get("ly_v");  c_rp = rd.get("ly_rp")
        else:
            c_o  = rd.get("bu_o");  c_r  = rd.get("bu_r")
            c_a  = rd.get("bu_a");  c_v  = rd.get("bu_v");  c_rp = rd.get("bu_rp")

        ty_c  = T if (is_cur or is_past) else DI
        mo_c  = T if is_cur else (WH if is_past else DI)
        mo_fw = "font-weight:800;" if is_cur else ""
        mo_fs = "15px" if is_cur else "13px"

        # TY row highlight: teal bg tint for current month
        if is_cur:
            row_bg = _mi_cur
        elif is_past:
            row_bg = bg_row
        else:
            row_bg = bg_row

        bg_s = f"background:{row_bg};" if row_bg else (f"background:{_mi_std};" if _mi_lt else "")
        cur_border = f"border-left:4px solid {_mi_cur_bar};" if is_cur else "border-left:4px solid transparent;"

        html = (
            f'<tr style="{bg_s}">'
            + f'<td style="padding:10px 14px;text-align:left;color:{mo_c};'
              f'font-family:Syne,sans-serif;font-size:{mo_fs};{mo_fw}'
              f'white-space:nowrap;border-right:1px solid {_mi_bdr};'
              f'border-bottom:1px solid {_mi_bdr3};{cur_border}">{mo_lbl}</td>'
            # TY cols
            + _td(_s(rd.get("ty_o"),"pct"),  ty_c,  delta=_dh(rd.get("ty_o"),  c_o,  "pct"),    bg=row_bg)
            + _td(_s(rd.get("ty_r"),"num"),  ty_c,  delta=_dh(rd.get("ty_r"),  c_r,  "num"),    bg=row_bg)
            + _td(_s(rd.get("ty_a")),        G, bold=True, delta=_dh(rd.get("ty_a"), c_a, "dollar"), bg=row_bg)
            + _td(_s(rd.get("ty_v")),        WH,    delta=_dh(rd.get("ty_v"),  c_v,  "dollar"), bg=row_bg)
            + _td(_s(rd.get("ty_rp")),       G, bold=True, delta=_dh(rd.get("ty_rp"),c_rp,"dollar"), bg=row_bg)
            # Budget cols
            + _td(_s(rd.get("bu_o"),"pct"),  PU, bg=row_bg, border_left=True)
            + _td(_s(rd.get("bu_r"),"num"),  PU, bg=row_bg)
            + _td(_s(rd.get("bu_a")),        G,  bold=True, bg=row_bg)
            + _td(_s(rd.get("bu_v")),        ("#3a5260"), bg=row_bg)
            + _td(_s(rd.get("bu_rp")),       G,  bold=True, bg=row_bg)
            # LY cols
            + _td(_s(rd.get("ly_o"),"pct"),  O,  bg=row_bg, border_left=True)
            + _td(_s(rd.get("ly_r"),"num"),  O,  bg=row_bg)
            + _td(_s(rd.get("ly_a")),        G,  bold=True, bg=row_bg)
            + _td(_s(rd.get("ly_v")),        ("#3a5260"), bg=row_bg)
            + _td(_s(rd.get("ly_rp")),       G,  bold=True, bg=row_bg)
            + "</tr>"
        )
        return html

    tbody = ""
    prev_q = None
    alt = False

    for r in rows_data:
        if r is None: continue
        q = QUARTERS[r["mn"]]
        if q != prev_q:
            if prev_q is not None:
                tbody += (f'<tr style="height:4px;"><td colspan="16" style="'
                          f'background:{_mi_qgap};border-bottom:1px solid {_mi_bdr};"></td></tr>')
            tbody += (f'<tr style="background:{_mi_qhdr};"><td colspan="16" style="'
                      f'padding:5px 14px;font-family:DM Mono,monospace;font-size:9px;'
                      f'letter-spacing:0.25em;color:{_mi_qlbl};text-transform:uppercase;'
                      f'border-bottom:1px solid {_mi_bdr};">{q}</td></tr>')
            prev_q = q
            alt = False

        row_bg = _mi_alt if alt else (_mi_std if _mi_lt else "")
        alt = not alt
        tbody += _build_row(r, r["name"], r["is_current"], r["is_past"], r["is_future"], row_bg)

    # Full Year row
    fy_rd = {"ty_o":fy_ty_o,"ty_r":fy_ty_r,"ty_a":fy_ty_a,"ty_v":fy_ty_v,"ty_rp":fy_ty_rp,
             "ly_o":fy_ly_o,"ly_r":fy_ly_r,"ly_a":fy_ly_a,"ly_v":fy_ly_v,"ly_rp":fy_ly_rp,
             "bu_o":fy_bu_o,"bu_r":fy_bu_r,"bu_a":fy_bu_a,"bu_v":fy_bu_v,"bu_rp":fy_bu_rp}
    tbody += f'<tr><td colspan="16" style="height:3px;background:{_mi_qgap};border-bottom:2px solid {_mi_sep};"></td></tr>'
    tbody += _build_row(fy_rd, "Full Year", False, True, False, _mi_fy)

    # Header bar
    yr_lbl = str(today.year)
    hdr = (
        f'<div style="background:{_mi_hdr};padding:10px 16px;border-bottom:1px solid {_mi_bdr};'
        'display:flex;align-items:center;gap:16px;flex-wrap:wrap;">'
        f'<b style="font-family:Syne,sans-serif;font-size:15px;font-weight:700;'
        f'color:white;-webkit-text-fill-color:white;">'
        f'{yr_lbl} Monthly Summary</b>'
        f'<b style="font-family:DM Mono,monospace;font-size:10px;'
        f'color:white;-webkit-text-fill-color:white;">'
        f'▌ = current month</b>'
        f'<b style="font-family:DM Mono,monospace;font-size:10px;font-weight:400;'
        f'color:white;-webkit-text-fill-color:white;margin-left:auto;">'
        f'Δ arrows on TY = {delta_lbl}</b>'
        '</div>'
    )

    table_html = (
        f'<div style="background:{_mi_card};border:1px solid {_mi_bdr};border-radius:10px;overflow:hidden;">'
        + hdr
        + '<div style="overflow-x:auto;">'
        + f'<table style="width:100%;border-collapse:collapse;background:{_mi_card};table-layout:fixed;">'
        + '<colgroup>'
        + '<col style="width:7%">'
        + '<col style="width:5.5%">'
        + '<col style="width:5%">'
        + '<col style="width:5%">'
        + '<col style="width:8%">'
        + '<col style="width:4.5%">'
        + '<col style="width:5.5%">'
        + '<col style="width:5%">'
        + '<col style="width:5%">'
        + '<col style="width:7%">'
        + '<col style="width:4.5%">'
        + '<col style="width:5.5%">'
        + '<col style="width:5%">'
        + '<col style="width:5%">'
        + '<col style="width:7%">'
        + '<col style="width:4.5%">'
        + '</colgroup>'
        + thead
        + f'<tbody>{tbody}</tbody>'
        + '</table></div></div>'
    )
    st.markdown(table_html, unsafe_allow_html=True)




def _render_ihg_corp_segments(data, hotel):
    """
    IHG SRP Intelligence tab — Corporate Account & Segment Production Report.
    Displays segment summary (top) and company detail (bottom).
    Source: Corp_Segments.xlsx via load_ihg_corp_segments() in data_loader.py.
    """
    cs = data.get("corp_segments")
    st.markdown('<div class="section-head">SRP Intelligence</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Corporate account & segment production · past 2 weeks · IHG</div>',
                unsafe_allow_html=True)

    if cs is None:
        err = data.get("corp_segments_error", "")
        if "errno 13" in err.lower() or "permission denied" in err.lower():
            st.warning("⚠️ Corp_Segments.xlsx is open in Excel — please close it, then click ⟳ Refresh Data.")
        else:
            st.warning("Corp Segments file not loaded. Place Corp_Segments.xlsx in the hotel folder.")
        return

    seg_df = cs.get("segments", pd.DataFrame())
    co_df  = cs.get("companies", pd.DataFrame())

    if seg_df.empty and co_df.empty:
        st.info("No production data found in Corp Segments file.")
        return

    _lt = False
    CARD   = "#ffffff"
    BORDER = "#c4cfd4"
    HDR_BG = "#dce5e8"
    HDR_FG = "#4e6878"
    ACCt   = "#2E618D"
    GOLD_C = "#6A924D"
    DIM    = "#5a7080"
    TXT    = "#1e2d35"
    RED_C  = "#a03020"
    SEG_COLORS = [ACCt, GOLD_C, "#556848", "#b44820", "#e84393",
                  "#44bbff", "#aaffcc", "#ff9966", "#99ccff", "#cc99ff",
                  "#ffcc66", "#66ffcc"]

    td = "padding:7px 10px;border-bottom:1px solid {bdr};text-align:{al};font-family:DM Mono,monospace;font-size:11px;"
    th_s = (f"padding:6px 8px;text-align:center;font-weight:600;"
            f"border-bottom:2px solid {BORDER};white-space:nowrap;"
            f"position:sticky;top:0;z-index:2;background:{HDR_BG};"
            f"font-family:DM Mono,monospace;font-size:10px;letter-spacing:0.06em;"
            f"text-transform:uppercase;color:{HDR_FG};")

    def _fmt_rms(v):  return f"{int(round(v)):,}" if v and v == v else "—"
    def _fmt_rev(v):  return f"${v:,.0f}"         if v and v == v else "—"
    def _fmt_adr(v):  return f"${v:,.2f}"          if v and v == v else "—"
    def _delta(v, fmt="rooms"):
        if v is None or (isinstance(v, float) and v != v): return ""
        d = float(v)
        if abs(d) < 0.01: return ""
        color = ACCt if d > 0 else RED_C
        arrow = "▲" if d > 0 else "▼"
        if fmt == "rooms": s = f"{int(abs(round(d))):,}"
        elif fmt == "rev": s = f"${abs(d):,.0f}"
        else: s = f"{abs(d):.0f}"
        return f'<span style="color:{color};font-size:10px;margin-left:3px;">{arrow}{s}</span>'

    # ── Segment Summary ──────────────────────────────────────────────────────
    if not seg_df.empty:
        st.markdown(f'<div class="section-head" style="margin-top:4px;">Segment Summary</div>',
                    unsafe_allow_html=True)

        seg_show = seg_df[(seg_df["CY_Rooms"] > 0) | (seg_df["LY_Rooms"] > 0)].copy()

        th_row = (f'<th style="{th_s}text-align:left;">Segment</th>'
                  f'<th style="{th_s}color:{ACCt};">CY Rooms</th>'
                  f'<th style="{th_s}color:{ACCt};">CY Rev</th>'
                  f'<th style="{th_s}color:{ACCt};">CY ADR</th>'
                  f'<th style="{th_s}color:{GOLD_C};">LY Rooms</th>'
                  f'<th style="{th_s}color:{GOLD_C};">LY Rev</th>'
                  f'<th style="{th_s}color:{GOLD_C};">LY ADR</th>'
                  f'<th style="{th_s}">Var Rooms</th>'
                  f'<th style="{th_s}">Var Rev</th>')

        tbody_s = ""
        for ri, (_, r) in enumerate(seg_show.iterrows()):
            rb = (f"background:{'#e8eef1' if not _lt else '#e8f0f8'};" if ri % 2 == 0
                  else f"background:{'#ffffff' if not _lt else ''};" if not _lt else "")
            bdr = BORDER
            seg_color = SEG_COLORS[ri % len(SEG_COLORS)]
            tbody_s += (
                f'<tr>'
                f'<td style="{td.format(bdr=bdr, al="left")}{rb}font-weight:600;color:{seg_color};">'
                f'{r["Segment"]}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{ACCt};font-weight:600;">'
                f'{_fmt_rms(r["CY_Rooms"])}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{ACCt};">'
                f'{_fmt_rev(r["CY_Rev"])}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{ACCt};">'
                f'{_fmt_adr(r.get("CY_ADR"))}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{GOLD_C};">'
                f'{_fmt_rms(r["LY_Rooms"])}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{GOLD_C};">'
                f'{_fmt_rev(r["LY_Rev"])}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{GOLD_C};">'
                f'{_fmt_adr(r.get("LY_ADR"))}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{TXT};">'
                f'{_fmt_rms(r["Var_Rooms"])}{_delta(r["Var_Rooms"])}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{TXT};">'
                f'{_fmt_rev(r["Var_Rev"])}{_delta(r["Var_Rev"], "rev")}</td>'
                f'</tr>'
            )

        # Totals row
        tot_cy_r = seg_show["CY_Rooms"].sum()
        tot_ly_r = seg_show["LY_Rooms"].sum()
        tot_cy_v = seg_show["CY_Rev"].sum()
        tot_ly_v = seg_show["LY_Rev"].sum()
        tot_cy_a = tot_cy_v / tot_cy_r if tot_cy_r > 0 else None
        tot_ly_a = tot_ly_v / tot_ly_r if tot_ly_r > 0 else None
        tot_vr   = tot_cy_r - tot_ly_r
        tot_vv   = tot_cy_v - tot_ly_v
        tot_bg   = f"background:{'#dce5e8' if not _lt else '#d0e0f0'};"
        tbody_s += (
            f'<tr>'
            f'<td style="{td.format(bdr=BORDER, al="left")}{tot_bg}font-weight:700;color:{TXT};">TOTAL</td>'
            f'<td style="{td.format(bdr=BORDER, al="center")}{tot_bg}font-weight:700;color:{ACCt};">{_fmt_rms(tot_cy_r)}</td>'
            f'<td style="{td.format(bdr=BORDER, al="center")}{tot_bg}font-weight:700;color:{ACCt};">{_fmt_rev(tot_cy_v)}</td>'
            f'<td style="{td.format(bdr=BORDER, al="center")}{tot_bg}font-weight:700;color:{ACCt};">{_fmt_adr(tot_cy_a)}</td>'
            f'<td style="{td.format(bdr=BORDER, al="center")}{tot_bg}font-weight:700;color:{GOLD_C};">{_fmt_rms(tot_ly_r)}</td>'
            f'<td style="{td.format(bdr=BORDER, al="center")}{tot_bg}font-weight:700;color:{GOLD_C};">{_fmt_rev(tot_ly_v)}</td>'
            f'<td style="{td.format(bdr=BORDER, al="center")}{tot_bg}font-weight:700;color:{GOLD_C};">{_fmt_adr(tot_ly_a)}</td>'
            f'<td style="{td.format(bdr=BORDER, al="center")}{tot_bg}font-weight:700;color:{TXT};">'
            f'{_fmt_rms(tot_vr)}{_delta(tot_vr)}</td>'
            f'<td style="{td.format(bdr=BORDER, al="center")}{tot_bg}font-weight:700;color:{TXT};">'
            f'{_fmt_rev(tot_vv)}{_delta(tot_vv, "rev")}</td>'
            f'</tr>'
        )

        seg_html = (
            f'<div style="overflow-x:auto;border:1px solid {BORDER};border-radius:6px;margin-bottom:24px;">'
            f'<table style="width:100%;border-collapse:collapse;background:{CARD};">'
            f'<thead><tr>{th_row}</tr></thead>'
            f'<tbody>{tbody_s}</tbody>'
            f'</table></div>'
        )
        st.markdown(seg_html, unsafe_allow_html=True)

    # ── Company Detail ───────────────────────────────────────────────────────
    if not co_df.empty:
        st.markdown(f'<div class="section-head">Corporate Account Detail</div>',
                    unsafe_allow_html=True)
        st.markdown(f'<div class="section-sub">One row per company · rooms summed across all rate codes · sorted by CY rooms</div>',
                    unsafe_allow_html=True)

        co_show = co_df[(co_df["CY_Rooms"] > 0) | (co_df["LY_Rooms"] > 0)].copy()

        th_co = (f'<th style="{th_s}text-align:left;">Company</th>'
                 f'<th style="{th_s}text-align:left;">Segment</th>'
                 f'<th style="{th_s}color:{ACCt};">CY Rooms</th>'
                 f'<th style="{th_s}color:{ACCt};">CY Rev</th>'
                 f'<th style="{th_s}color:{ACCt};">CY ADR</th>'
                 f'<th style="{th_s}color:{GOLD_C};">LY Rooms</th>'
                 f'<th style="{th_s}color:{GOLD_C};">LY ADR</th>'
                 f'<th style="{th_s}">Var Rooms</th>')

        tbody_c = ""
        for ri, (_, r) in enumerate(co_show.iterrows()):
            rb = (f"background:{'#e8eef1' if not _lt else '#e8f0f8'};" if ri % 2 == 0
                  else f"background:{'#ffffff' if not _lt else ''};" if not _lt else "")
            bdr = BORDER
            # Color-code by dominant segment
            seg_list = list(seg_df["Segment"].tolist()) if not seg_df.empty else []
            seg_idx  = seg_list.index(r.get("Dom_Segment","")) if r.get("Dom_Segment","") in seg_list else 0
            seg_col  = SEG_COLORS[seg_idx % len(SEG_COLORS)]
            tbody_c += (
                f'<tr>'
                f'<td style="{td.format(bdr=bdr, al="left")}{rb}font-weight:600;color:{TXT};">'
                f'{r["CompanyName"]}</td>'
                f'<td style="{td.format(bdr=bdr, al="left")}{rb}font-size:10px;color:{seg_col};">'
                f'{r.get("Dom_Segment","")}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{ACCt};font-weight:600;">'
                f'{_fmt_rms(r["CY_Rooms"])}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{ACCt};">'
                f'{_fmt_rev(r["CY_Rev"])}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{ACCt};">'
                f'{_fmt_adr(r.get("CY_ADR"))}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{GOLD_C};">'
                f'{_fmt_rms(r["LY_Rooms"])}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{GOLD_C};">'
                f'{_fmt_adr(r.get("LY_ADR"))}</td>'
                f'<td style="{td.format(bdr=bdr, al="center")}{rb}color:{TXT};">'
                f'{_fmt_rms(r["Var_Rooms"])}{_delta(r["Var_Rooms"])}</td>'
                f'</tr>'
            )

        co_html = (
            f'<div style="overflow-x:auto;border:1px solid {BORDER};border-radius:6px;margin-bottom:16px;">'
            f'<table style="width:100%;border-collapse:collapse;background:{CARD};">'
            f'<thead><tr>{th_co}</tr></thead>'
            f'<tbody>{tbody_c}</tbody>'
            f'</table></div>'
        )
        st.markdown(co_html, unsafe_allow_html=True)


def render_srp_tab(data, hotel):
    # ── IHG branch: show corporate segment production report ─────────────────
    if data.get("_ihg_hotel"):
        _render_ihg_corp_segments(data, hotel)
        return

    srp = data.get("booking")
    st.markdown('<div class="section-head">SRP Intelligence</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Individual reservation data · Jan 2024 – Sep 2026</div>',
                unsafe_allow_html=True)

    if srp is None:
        st.warning("SRP Activity file not loaded.")
        return

    # ── Date range filter — narrow column ──
    filter_col, _ = st.columns([1, 3])
    with filter_col:
        min_d = srp["Arrival_Date"].min().date()
        max_d = srp["Arrival_Date"].max().date()
        _default_end   = min(datetime.now().date(), max_d)
        _default_start = max((datetime.now() - timedelta(days=14)).date(), min_d)
        _default_start = min(_default_start, _default_end)
        date_range = st.date_input("Arrival Date Range",
                                   value=(_default_start, _default_end),
                                   min_value=min_d, max_value=max_d)

    # Apply date filter only — all segments always included
    filtered = srp.copy()
    if len(date_range) == 2:
        filtered = filtered[
            (filtered["Arrival_Date"] >= pd.Timestamp(date_range[0])) &
            (filtered["Arrival_Date"] <= pd.Timestamp(date_range[1]))
        ]

    if filtered.empty:
        st.info("No data for selected date range.")
        return

    # Segment color palette — consistent across both charts
    all_segments = sorted(filtered["MCAT"].dropna().unique().tolist())
    seg_colors = [TEAL, PURPLE, ORANGE, GOLD, "#e84393", "#44bbff", "#aaffcc",
                  "#ff9966", "#99ccff", "#cc99ff", "#ffcc66", "#66ffcc"]
    seg_color_map = {seg: seg_colors[i % len(seg_colors)]
                     for i, seg in enumerate(all_segments)}

    # ── 1. Daily Revenue by Segment — grouped bars, full width ──
    daily_seg = (
        filtered.groupby(["Arrival_Date", "MCAT"])["Room_Revenue"]
        .sum().reset_index()
        .rename(columns={"MCAT": "Segment", "Room_Revenue": "Revenue"})
    )

    fig_bar = go.Figure()
    for seg in all_segments:
        seg_data = daily_seg[daily_seg["Segment"] == seg]
        fig_bar.add_trace(go.Bar(
            x=seg_data["Arrival_Date"],
            y=seg_data["Revenue"],
            name=seg,
            marker_color=seg_color_map[seg],
            hovertemplate=f"<b>{seg}</b><br>%{{x|%b %d}}<br>${{y:,.0f}}<extra></extra>",
        ))

    _bar_layout = {k: v for k, v in chart_layout().items()
                   if k not in ("height", "legend", "margin", "xaxis", "yaxis")}
    _srp_ax_col = "#3a5260"
    fig_bar.update_layout(
        **_bar_layout,
        barmode="group",
        height=400,
        title=dict(text="Daily Revenue by Segment", font=dict(size=13, color="#1e2d35")),
        yaxis=dict(tickprefix="$", gridcolor=BORDER, tickfont=dict(color=_srp_ax_col), color=_srp_ax_col),
        xaxis=dict(tickfont=dict(color=_srp_ax_col), color=_srp_ax_col, gridcolor="#dce5e8"),
        legend=dict(orientation="h", yanchor="top", y=-0.18,
                    xanchor="center", x=0.5, font=dict(size=11, color="#3a5260")),
        margin=dict(t=40, b=140, l=50, r=20),
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    # ── 2. Revenue by Segment — pie full width, labels outside ──
    seg_sum = (
        filtered.groupby("MCAT")["Room_Revenue"]
        .sum().reset_index()
        .rename(columns={"MCAT": "Segment", "Room_Revenue": "Revenue"})
        .sort_values("Revenue", ascending=False)
    )
    total_rev = seg_sum["Revenue"].sum()
    seg_sum["Pct"] = (seg_sum["Revenue"] / total_rev * 100).round(1)
    seg_sum["Label"] = seg_sum.apply(
        lambda r: f"{r['Segment']}  {r['Pct']:.1f}%", axis=1
    )

    _pie_layout = {k: v for k, v in chart_layout().items()
                   if k not in ("xaxis", "yaxis", "height", "legend", "margin")}
    fig_pie = go.Figure(go.Pie(
        labels=seg_sum["Segment"],
        values=seg_sum["Revenue"],
        text=seg_sum["Label"],
        textinfo="text",
        textposition="outside",
        hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>",
        marker=dict(colors=[seg_color_map[s] for s in seg_sum["Segment"]],
                    line=dict(color=BG, width=2)),
        hole=0.08,
    ))
    fig_pie.update_layout(
        **_pie_layout,
        height=420,
        title=dict(text="Revenue by Segment", font=dict(size=13, color="#1e2d35")),
        showlegend=False,
        margin=dict(t=40, b=20, l=120, r=120),
    )
    st.plotly_chart(fig_pie, use_container_width=True)

    # ── 3. Top SRP codes ──
    st.markdown('<div class="section-head">Top SRP Codes</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-sub" style="margin-bottom:10px;">Click a rate name to view individual reservations</div>',
        unsafe_allow_html=True,
    )

    srp_sum = filtered.groupby(["SRP_Code", "SRP_Name", "MCAT"]).agg(
        Nights=("Room_Nights", "sum"),
        Revenue=("Room_Revenue", "sum"),
    ).reset_index()
    srp_sum = srp_sum.rename(columns={"MCAT": "Segment"})
    srp_sum["ADR"] = (srp_sum["Revenue"] / srp_sum["Nights"]).round(2)
    srp_sum = srp_sum.sort_values("Revenue", ascending=False).head(25).reset_index(drop=True)

    # ── Build clickable HTML table (SRP_Code hidden, SRP_Name is the drill link) ──
    _srp_lt = False
    _srp_tbl_bg   = "#dce5e8"
    _srp_hdr_col  = "#4e6878"
    _srp_bdr      = "#c4cfd4"
    _srp_row_col  = "#1e2d35"
    _srp_div_bdr  = "#c4cfd4"

    active_code = st.session_state.srp_drill

    table_rows = ""
    for _ri, (_, row) in enumerate(srp_sum.iterrows()):
        code     = row["SRP_Code"]
        name     = row["SRP_Name"]
        segment  = row["Segment"]
        nights   = int(row["Nights"])
        revenue  = f"${row['Revenue']:,.0f}"
        adr      = f"${row['ADR']:,.2f}"
        is_active = (code == active_code)

        _srp_r_lt = False
        if is_active:
            row_bg = "background:#c8dff0;" if _srp_r_lt else "background:#314B63;"
        else:
            _srp_alt_bg = ("#e8f0f8" if _srp_r_lt else "#dce5e8") if _ri % 2 == 0 else _srp_tbl_bg
            row_bg = f"background:{_srp_alt_bg};"
        link_col = "#ffffff" if is_active else "#2E618D"
        icon     = " ▾" if is_active else ""

        _name_cell_bg = ("#c8dff0" if _srp_r_lt else "#314B63") if is_active else _srp_tbl_bg
        # SRP_Name cell is the drill trigger (uses Streamlit button trick via form-less button)
        name_cell = (
            f'<td style="padding:8px 12px;border-bottom:1px solid {_srp_bdr};background:{_name_cell_bg};">'
            f'<span style="color:{link_col};cursor:pointer;text-decoration:underline '
            f'dotted;font-weight:{"600" if is_active else "400"};"'
            f' title="Click to view reservations">{name}{icon}</span></td>'
        )

        table_rows += (
            f'<tr style="{row_bg}" data-code="{code}">'
            f'{name_cell}'
            f'<td style="padding:8px 14px;border-bottom:1px solid {_srp_bdr};color:{_srp_row_col};text-align:center;{row_bg}">{segment}</td>'
            f'<td style="padding:8px 14px;border-bottom:1px solid {_srp_bdr};color:{_srp_row_col};text-align:center;{row_bg}">{nights}</td>'
            f'<td style="padding:8px 14px;border-bottom:1px solid {_srp_bdr};color:{_srp_row_col};text-align:center;{row_bg}">{revenue}</td>'
            f'<td style="padding:8px 14px;border-bottom:1px solid {_srp_bdr};color:{_srp_row_col};text-align:center;{row_bg}">{adr}</td>'
            f'</tr>'
        )

    srp_table_html = f"""
    <div style="overflow-x:auto;border:1px solid {_srp_div_bdr};border-radius:6px;margin-bottom:8px;">
      <table style="width:100%;border-collapse:collapse;font-size:13px;font-family:'DM Sans',sans-serif;background:{_srp_tbl_bg};">
        <thead>
          <tr style="background:{_srp_tbl_bg};">
            <th style="padding:9px 12px;text-align:left;color:{_srp_hdr_col};font-weight:700;
                       border-bottom:2px solid {_srp_bdr};">Rate Name</th>
            <th style="padding:9px 12px;text-align:center;color:{_srp_hdr_col};font-weight:700;
                       border-bottom:2px solid {_srp_bdr};">Segment</th>
            <th style="padding:9px 12px;text-align:center;color:{_srp_hdr_col};font-weight:700;
                       border-bottom:2px solid {_srp_bdr};">Nights</th>
            <th style="padding:9px 12px;text-align:center;color:{_srp_hdr_col};font-weight:700;
                       border-bottom:2px solid {_srp_bdr};">Revenue</th>
            <th style="padding:9px 12px;text-align:center;color:{_srp_hdr_col};font-weight:700;
                       border-bottom:2px solid {_srp_bdr};">ADR</th>
          </tr>
        </thead>
        <tbody>{table_rows}</tbody>
      </table>
    </div>
    """
    st.markdown(srp_table_html, unsafe_allow_html=True)

    # ── Streamlit buttons to handle row clicks (one per row, hidden under the HTML table) ──
    # We render invisible buttons that mirror each row; clicking the rate name label
    # via a visible selectbox is the cleanest Streamlit-native approach.
    rate_options = ["— select a rate to drill down —"] + srp_sum["SRP_Name"].tolist()
    code_lookup  = dict(zip(srp_sum["SRP_Name"], srp_sum["SRP_Code"]))

    sel_col, clear_col = st.columns([3, 1])
    with sel_col:
        chosen_name = st.selectbox(
            "Drill into rate:",
            options=rate_options,
            index=(rate_options.index(
                next((r["SRP_Name"] for _, r in srp_sum.iterrows()
                      if r["SRP_Code"] == active_code), rate_options[0])
            ) if active_code else 0),
            label_visibility="collapsed",
        )
    with clear_col:
        if st.button("✕ Clear", key="srp_clear", use_container_width=True):
            st.session_state.srp_drill = None
            st.rerun()

    # Update drill state when selection changes
    if chosen_name != rate_options[0]:
        new_code = code_lookup.get(chosen_name)
        if new_code != st.session_state.srp_drill:
            st.session_state.srp_drill = new_code
            st.rerun()

    # ── Inline drill-down panel ──
    if st.session_state.srp_drill:
        drill_code = st.session_state.srp_drill
        drill_name = next(
            (r["SRP_Name"] for _, r in srp_sum.iterrows() if r["SRP_Code"] == drill_code),
            drill_code,
        )
        drill_segment = next(
            (r["Segment"] for _, r in srp_sum.iterrows() if r["SRP_Code"] == drill_code),
            "",
        )

        reservations = filtered[filtered["SRP_Code"] == drill_code].copy()
        reservations = reservations.sort_values("Arrival_Date")

        # Build reservation rows HTML
        res_rows = ""
        for _, res in reservations.iterrows():
            conf      = str(res.get("Confirmation", "—"))
            booked    = res["Booked_Date"].strftime("%b %d, %Y") if pd.notna(res.get("Booked_Date")) else "—"
            arrival   = res["Arrival_Date"].strftime("%b %d, %Y") if pd.notna(res["Arrival_Date"]) else "—"
            departure = res["Departure_Date"].strftime("%b %d, %Y") if pd.notna(res.get("Departure_Date")) else "—"
            nights_r  = int(res["Room_Nights"]) if pd.notna(res["Room_Nights"]) else "—"
            rev_r     = f"${res['Room_Revenue']:,.0f}" if pd.notna(res.get("Room_Revenue")) else "—"
            adr_r     = f"${res['ADR']:,.2f}" if pd.notna(res.get("ADR")) else "—"

            res_rows += (
                f'<tr>'
                f'<td style="padding:7px 12px;border-bottom:1px solid #1a2e42;color:#1e2d35;text-align:center;">{conf}</td>'
                f'<td style="padding:7px 12px;border-bottom:1px solid #1a2e42;color:#1e2d35;text-align:center;">{booked}</td>'
                f'<td style="padding:7px 12px;border-bottom:1px solid #1a2e42;color:{TEAL};text-align:center;">{arrival}</td>'
                f'<td style="padding:7px 12px;border-bottom:1px solid #1a2e42;color:#1e2d35;text-align:center;">{departure}</td>'
                f'<td style="padding:7px 12px;border-bottom:1px solid #1a2e42;color:#1e2d35;text-align:center;">{nights_r}</td>'
                f'<td style="padding:7px 12px;border-bottom:1px solid #1a2e42;color:#1e2d35;text-align:center;">{rev_r}</td>'
                f'<td style="padding:7px 12px;border-bottom:1px solid #1a2e42;color:#1e2d35;text-align:center;">{adr_r}</td>'
                f'</tr>'
            )

        total_nights = int(reservations["Room_Nights"].sum())
        total_rev    = reservations["Room_Revenue"].sum()
        avg_adr      = total_rev / total_nights if total_nights > 0 else 0
        n_res        = len(reservations)

        drill_html = f"""
        <div style="margin-top:4px;border:1px solid {TEAL};border-radius:8px;
                    background:#ffffff;border:1px solid #c4cfd4;border-radius:8px;overflow:hidden;animation:fadeIn 0.2s ease;">
          <!-- Panel header -->
          <div style="background:#2E618D;padding:12px 16px;border-bottom:1px solid {TEAL};
                      display:flex;align-items:center;justify-content:space-between;">
            <div>
              <span style="color:{TEAL};font-weight:700;font-size:14px;">{drill_name}</span>
              <span style="color:#4e6878;font-size:12px;margin-left:12px;">{drill_segment}</span>
              <span style="color:#4e6878;font-size:12px;margin-left:12px;">·</span>
              <span style="color:#4e6878;font-size:12px;margin-left:8px;">{n_res} reservation{"s" if n_res != 1 else ""}</span>
            </div>
            <div style="display:flex;gap:24px;">
              <div style="text-align:right;">
                <div style="color:#4e6878;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">Total Nights</div>
                <div style="color:#1e2d35;font-weight:700;">{total_nights}</div>
              </div>
              <div style="text-align:right;">
                <div style="color:#4e6878;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">Total Revenue</div>
                <div style="color:{GOLD};font-weight:700;">${total_rev:,.0f}</div>
              </div>
              <div style="text-align:right;">
                <div style="color:#4e6878;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">Avg ADR</div>
                <div style="color:#1e2d35;font-weight:700;">${avg_adr:,.2f}</div>
              </div>
            </div>
          </div>
          <!-- Reservation table -->
          <div style="overflow-x:auto;max-height:380px;overflow-y:auto;">
            <table style="width:100%;border-collapse:collapse;font-size:12.5px;
                          font-family:'DM Sans',sans-serif;">
              <thead style="position:sticky;top:0;z-index:1;">
                <tr style="background:#314B63;">
                  <th style="padding:8px 12px;text-align:center;color:#4e6878;
                             font-weight:600;border-bottom:1px solid #1a2e42;">Confirmation</th>
                  <th style="padding:8px 12px;text-align:center;color:#4e6878;
                             font-weight:600;border-bottom:1px solid #1a2e42;">Booked</th>
                  <th style="padding:8px 12px;text-align:center;color:#4e6878;
                             font-weight:600;border-bottom:1px solid #1a2e42;">Arrival</th>
                  <th style="padding:8px 12px;text-align:center;color:#4e6878;
                             font-weight:600;border-bottom:1px solid #1a2e42;">Departure</th>
                  <th style="padding:8px 12px;text-align:center;color:#4e6878;
                             font-weight:600;border-bottom:1px solid #1a2e42;">Nights</th>
                  <th style="padding:8px 12px;text-align:center;color:#4e6878;
                             font-weight:600;border-bottom:1px solid #1a2e42;">Revenue</th>
                  <th style="padding:8px 12px;text-align:center;color:#4e6878;
                             font-weight:600;border-bottom:1px solid #1a2e42;">ADR</th>
                </tr>
              </thead>
              <tbody>{res_rows}</tbody>
            </table>
          </div>
        </div>
        <style>@keyframes fadeIn {{ from {{ opacity:0; transform:translateY(-6px); }} to {{ opacity:1; transform:translateY(0); }} }}</style>
        """
        st.markdown(drill_html, unsafe_allow_html=True)




def render_groups_tab(data, hotel):
    groups = data.get("groups")

    if groups is None:
        st.warning("Group Wash Report file not loaded.")
        return

    today = pd.Timestamp(datetime.now().date())
    active = groups[groups["Departure"] >= today] if "Departure" in groups.columns else groups

    # ── Alerts ──
    if "Pickup_Pct" in active.columns and "Days_To_Arrival" in active.columns:
        at_risk = active[
            (active["Pickup_Pct"] < 50) &
            (active["Days_To_Arrival"].between(0, 21))
        ]
        if not at_risk.empty:
            names = ", ".join(at_risk["Group_Name"].dropna().unique()[:5].tolist())
            st.markdown(f'<div class="alert-box">⚠ <strong>Groups at risk:</strong> {names} — pickup below 50% with cutoff approaching</div>',
                        unsafe_allow_html=True)

    # ── Summary by group ──
    if "Group_Name" in active.columns:
        summary_cols = [c for c in ["Group_Name", "Arrival", "Departure", "Cutoff",
                                     "Sales_Manager", "Block", "Pickup", "Avail_Block",
                                     "Pickup_Pct", "Rate", "Days_To_Arrival"] if c in active.columns]
        grp_sum = active.groupby("Group_Name").agg({
            c: "first" if c in ["Arrival","Departure","Cutoff","Sales_Manager","Rate","Days_To_Arrival"]
               else "sum"
            for c in summary_cols if c != "Group_Name"
        }).reset_index()

        if "Pickup_Pct" in grp_sum.columns:
            grp_sum["Pickup_Pct"] = (grp_sum["Pickup"] / grp_sum["Block"] * 100).where(
                grp_sum["Block"] > 0).round(1)

        # Recalculate Avail_Block from aggregated Block/Pickup to ensure consistency
        if "Avail_Block" in grp_sum.columns and "Block" in grp_sum.columns and "Pickup" in grp_sum.columns:
            grp_sum["Avail_Block"] = (grp_sum["Block"] - grp_sum["Pickup"]).clip(lower=0).astype(int)

        # ── Group Totals by Month ──
        if "Block" in grp_sum.columns and "Pickup" in grp_sum.columns and "Arrival" in grp_sum.columns:
            st.markdown('<div class="section-head" style="margin-top:8px;">Group Totals</div>',
                        unsafe_allow_html=True)
            st.markdown('<div class="section-sub">Blocked · Picked up · Pickup % by month · next 6 months</div>',
                        unsafe_allow_html=True)

            # Build month buckets for next 6 months
            month_cards_html = ""
            for m_offset in range(6):
                month_start = (today + pd.DateOffset(months=m_offset)).replace(day=1)
                month_end   = (month_start + pd.DateOffset(months=1)) - pd.Timedelta(days=1)
                month_label = month_start.strftime("%B %Y")

                mask = (
                    grp_sum["Arrival"].notna() &
                    (grp_sum["Arrival"] >= month_start) &
                    (grp_sum["Arrival"] <= month_end)
                )
                m_df = grp_sum[mask]

                blocked  = int(m_df["Block"].sum())  if not m_df.empty else 0
                pickedup = int(m_df["Pickup"].sum()) if not m_df.empty else 0
                pct      = round(pickedup / blocked * 100, 1) if blocked > 0 else 0
                n_groups = len(m_df)

                # Color thresholds
                _gt_lt = False
                _c_teal   = "#006a55" if _gt_lt else TEAL
                _c_gold   = "#7a4e00" if _gt_lt else GOLD
                _c_orange = "#cc3300" if _gt_lt else ORANGE
                _c_dim    = "#6a8090"
                _c_bar0   = "#c4cfd4"
                if pct >= 75:
                    pct_color   = _c_teal
                    bar_color   = _c_teal
                elif pct >= 50:
                    pct_color   = _c_gold
                    bar_color   = _c_gold
                elif pct > 0:
                    pct_color   = _c_orange
                    bar_color   = _c_orange
                else:
                    pct_color   = _c_dim
                    bar_color   = _c_bar0

                avail = blocked - pickedup
                groups_txt = f"{n_groups} group{'s' if n_groups != 1 else ''}" if n_groups > 0 else "No groups"

                month_cards_html += f"""
<div class="grp-month-card">
  <div class="grp-month-label">{month_label}</div>
  <div class="grp-month-pct" style="color:{pct_color}">{pct:.0f}<span class="grp-month-pct-sym">%</span></div>
  <div class="grp-month-bar-track">
    <div class="grp-month-bar-fill" style="width:{min(pct,100):.0f}%;background:{bar_color};"></div>
  </div>
  <div class="grp-month-stats">
    <div class="grp-month-stat">
      <div class="grp-month-stat-val" style="color:{"#4a3aaa" if _gt_lt else PURPLE}">{blocked:,}</div>
      <div class="grp-month-stat-lbl">Blocked</div>
    </div>
    <div class="grp-month-stat">
      <div class="grp-month-stat-val" style="color:{_c_teal}">{pickedup:,}</div>
      <div class="grp-month-stat-lbl">Pickup</div>
    </div>
    <div class="grp-month-stat">
      <div class="grp-month-stat-val" style="color:{_c_dim}">{avail:,}</div>
      <div class="grp-month-stat-lbl">Remaining</div>
    </div>
  </div>
  <div class="grp-month-groups">{groups_txt}</div>
</div>"""

            _gt_card_bg  = "#e8f0f5" if _gt_lt else "#dce8f0"
            _gt_card_bdr = "#c8d8e4" if _gt_lt else "#bccedd"
            _gt_lbl_col  = "#3a5260"
            _gt_bar_trk  = "#c0d0de" if _gt_lt else BORDER
            _gt_lbl2_col = "#6a8090"
            st.markdown(f"""
<style>
.grp-month-row {{
    display: grid;
    grid-template-columns: repeat(6, 1fr);
    gap: 10px;
    margin: 10px 0 24px 0;
    width: 100%;
}}
.grp-month-card {{
    background: {_gt_card_bg} !important;
    border: 1px solid {_gt_card_bdr};
    border-radius: 10px;
    padding: 14px 16px 12px;
    display: flex;
    flex-direction: column;
    gap: 6px;
}}
.grp-month-label {{
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: {_gt_lbl_col};
    font-weight: 600;
}}
.grp-month-pct {{
    font-family: 'Syne', sans-serif;
    font-size: 36px;
    font-weight: 800;
    line-height: 1;
    margin: 2px 0 0;
}}
.grp-month-pct-sym {{
    font-size: 18px;
    font-weight: 600;
    vertical-align: super;
}}
.grp-month-bar-track {{
    width: 100%;
    height: 4px;
    background: {_gt_bar_trk};
    border-radius: 2px;
    overflow: hidden;
    margin: 2px 0;
}}
.grp-month-bar-fill {{
    height: 100%;
    border-radius: 2px;
    transition: width 0.3s ease;
}}
.grp-month-stats {{
    display: flex;
    gap: 0;
    justify-content: space-between;
    margin-top: 4px;
}}
.grp-month-stat {{
    text-align: center;
    flex: 1;
}}
.grp-month-stat-val {{
    font-family: 'DM Mono', monospace;
    font-size: 15px;
    font-weight: 500;
    line-height: 1.2;
}}
.grp-month-stat-lbl {{
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: {_gt_lbl2_col};
    margin-top: 1px;
}}
.grp-month-groups {{
    font-size: 10px;
    color: {_gt_lbl2_col};
    margin-top: 2px;
    font-style: italic;
}}
</style>
<div class="grp-month-row">{month_cards_html}</div>
""", unsafe_allow_html=True)

        # ── Block vs Pickup — full width, next 20 groups by arrival date ──
        if "Block" in grp_sum.columns and "Pickup" in grp_sum.columns:
            chart_grp = grp_sum.copy()
            if "Arrival" in chart_grp.columns:
                chart_grp = chart_grp.sort_values("Arrival").head(20)
            else:
                chart_grp = chart_grp.sort_values("Block", ascending=False).head(20)
            chart_grp = chart_grp.reset_index(drop=True)

            n_bars = len(chart_grp)
            # Each bar occupies roughly (chart_width / n_bars) px.
            # At ~11px per character, estimate chars that fit per bar.
            # Chart is ~full browser width ≈ 1400px, minus margins ≈ 1340px usable.
            chars_per_bar = max(6, int((1340 / max(n_bars, 1)) / 7))

            def wrap_name(name, max_chars):
                """Split name into lines no longer than max_chars, breaking on spaces."""
                words = name.split()
                lines, current = [], ""
                for word in words:
                    if current and len(current) + 1 + len(word) > max_chars:
                        lines.append(current)
                        current = word
                    else:
                        current = f"{current} {word}".strip() if current else word
                if current:
                    lines.append(current)
                return "<br>".join(lines)

            # Build tick labels: wrapped name + cutoff date line
            tick_vals = list(range(n_bars))
            tick_texts = []
            for _, row in chart_grp.iterrows():
                raw_name = str(row["Group_Name"])
                wrapped = wrap_name(raw_name, chars_per_bar)
                cutoff_line = ""
                if "Cutoff" in chart_grp.columns:
                    try:
                        cutoff_ts = pd.Timestamp(row["Cutoff"])
                        if pd.notna(cutoff_ts):
                            days_to_cutoff = (cutoff_ts - today).days
                            cutoff_str = cutoff_ts.strftime('%b %d')
                            if days_to_cutoff < 0:
                                cutoff_line = f"<span style='color:{GOLD}'>✂ {cutoff_str} ✓</span>"
                            elif days_to_cutoff <= 7:
                                cutoff_line = f"<span style='color:#b44820'>✂ {cutoff_str} ({days_to_cutoff}d)</span>"
                            else:
                                cutoff_line = f"<span style='color:{GOLD}'>✂ {cutoff_str}</span>"
                    except Exception:
                        pass
                label = f"{wrapped}<br>{cutoff_line}" if cutoff_line else wrapped
                tick_texts.append(label)

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=tick_vals, y=chart_grp["Block"],
                name="Block", marker_color=PURPLE,
                hovertemplate="<b>%{customdata}</b><br>Block: %{y}<extra></extra>",
                customdata=chart_grp["Group_Name"],
            ))
            fig.add_trace(go.Bar(
                x=tick_vals, y=chart_grp["Pickup"],
                name="Pickup", marker_color=TEAL,
                hovertemplate="<b>%{customdata}</b><br>Pickup: %{y}<extra></extra>",
                customdata=chart_grp["Group_Name"],
            ))

            # Month divider shapes + timeline annotations below the tick labels
            month_shapes = []
            month_annotations = []
            if "Arrival" in chart_grp.columns:
                prev_month = None
                for i, row in chart_grp.iterrows():
                    try:
                        arr_ts = pd.Timestamp(row["Arrival"])
                        if pd.notna(arr_ts):
                            month_key = (arr_ts.year, arr_ts.month)
                            if month_key != prev_month:
                                if prev_month is not None and i > 0:
                                    month_shapes.append(dict(
                                        type="line",
                                        x0=i - 0.5, x1=i - 0.5,
                                        y0=0, y1=1,
                                        xref="x", yref="paper",
                                        line=dict(color="#2a4060", width=1, dash="dot"),
                                    ))
                                # Month label sits at y=-0.52 (well below two-line tick labels)
                                month_annotations.append(dict(
                                    x=i,
                                    y=-0.52,
                                    xref="x", yref="paper",
                                    text=f"<b>{arr_ts.strftime('%B %Y')}</b>",
                                    showarrow=False,
                                    font=dict(size=11, color="#3a5260"),
                                    xanchor="left",
                                ))
                                prev_month = month_key
                    except Exception:
                        pass

            _layout = {k: v for k, v in chart_layout().items()
                       if k not in ("xaxis", "legend", "margin", "height")}
            fig.update_layout(
                **_layout,
                barmode="overlay",
                height=420,
                title=dict(text="Block vs Pickup by Group  ·  next 20 arrivals",
                           font=dict(size=13, color="#1e2d35")),
                margin=dict(t=40, b=160, l=40, r=20),
                annotations=month_annotations,
                shapes=month_shapes,
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="right", x=1,
                            font=dict(size=11, color="#1e2d35")),
            )
            _tick_col = "#3a5260"
            fig.update_xaxes(
                tickvals=tick_vals,
                ticktext=tick_texts,
                tickangle=0,
                tickfont=dict(size=12, color=_tick_col),
                gridcolor=BORDER,
                showgrid=False,
            )
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.markdown('<div class="section-head">Group Detail Table</div>', unsafe_allow_html=True)
        for col in ["Arrival", "Departure", "Cutoff"]:
            if col in grp_sum.columns:
                grp_sum[col] = grp_sum[col].dt.strftime("%b %d, %Y")
        if "Rate" in grp_sum.columns:
            grp_sum["Rate"] = grp_sum["Rate"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "")
        if "Pickup_Pct" in grp_sum.columns:
            grp_sum["Pickup_Pct"] = grp_sum["Pickup_Pct"].apply(
                lambda x: f"{x:.1f}%" if pd.notna(x) else "")

        # Drop Days_To_Arrival, keep everything else
        display_cols = [c for c in grp_sum.columns if c != "Days_To_Arrival"]
        display_df = grp_sum[display_cols].copy()

        # Build row data as JSON for JS sorting
        import json
        rows_json = []
        for _, row in display_df.iterrows():
            rows_json.append({c: (str(row[c]) if pd.notna(row[c]) else "—") for c in display_cols})

        _gd_lt = False
        _gd_txt     = "#1e2d35"
        _gd_bdr     = "#b0c0d0"  if _gd_lt else BORDER
        _gd_teal    = "#006a55"  if _gd_lt else TEAL
        _gd_gold    = "#7a4e00"  if _gd_lt else GOLD
        _gd_orange  = "#cc3300"  if _gd_lt else ORANGE
        _gd_dim     = "#6a8090"
        _gd_hdr_bg  = "#314B63"
        _gd_hdr_col = "#c8d8e4"
        _gd_card    = "#f0f4f8"  if _gd_lt else CARD
        _gd_alt     = "#f0f4f6"
        def cell_style(c, val):
            base = (f"text-align:{'left' if c == 'Group_Name' else 'center'};"
                    f"padding:7px 12px;font-size:13px;border-bottom:1px solid {_gd_bdr};color:{_gd_txt};")
            if c == "Pickup_Pct":
                try:
                    pct_val = float(str(val).replace("%",""))
                    col = _gd_teal if pct_val >= 75 else (_gd_gold if pct_val >= 50 else (_gd_orange if pct_val > 0 else _gd_dim))
                    base += f"color:{col};font-weight:600;"
                except Exception: pass
            elif c == "Cutoff":
                try:
                    cutoff_ts = pd.Timestamp(val)
                    days = (cutoff_ts - today).days
                    if 0 <= days <= 7: base += f"color:{_gd_orange};font-weight:600;"
                    elif days < 0: base += f"color:{_gd_dim};"
                except Exception: pass
            return base

        # Pre-compute per-cell styles (these are static — sorting only reorders rows)
        cell_styles = {}
        for row_dict in rows_json:
            for c in display_cols:
                cell_styles[f"{c}"] = cell_style(c, row_dict[c])

        col_labels = {c: c.replace("_", " ") for c in display_cols}

        tbl_id = "grp_sort_tbl"
        table_html = f"""
<style>
#{tbl_id} th {{ cursor:pointer; position:relative; }}
#{tbl_id} th .sort-arrow {{ opacity:0; margin-left:5px; font-size:10px; transition:opacity 0.15s; }}
#{tbl_id} th:hover .sort-arrow {{ opacity:1; }}
#{tbl_id} th.sorted-asc .sort-arrow::after {{ content:'▲'; opacity:1 !important; }}
#{tbl_id} th.sorted-desc .sort-arrow::after {{ content:'▼'; opacity:1 !important; }}
#{tbl_id} th.sorted-asc .sort-arrow, #{tbl_id} th.sorted-desc .sort-arrow {{ opacity:1; }}
#{tbl_id} tr {{ transition:background 0.15s; }}
#{tbl_id} tr:hover {{ background:{_gd_alt} !important; }}
</style>
<div style="overflow-x:auto;margin-top:8px;">
<table id="{tbl_id}" style="width:100%;border-collapse:collapse;background:{_gd_card};border-radius:8px;overflow:hidden;">
  <thead><tr id="{tbl_id}_head" style="background:{_gd_hdr_bg};">
    {"".join(
        f'<th onclick="sortGrpTable(\'{tbl_id}\',{i})" '
        f'style="text-align:{"left" if c == "Group_Name" else "center"};padding:8px 12px;'
        f'font-family:DM Mono,monospace;font-size:11px;letter-spacing:0.08em;'
        f'text-transform:uppercase;color:{_gd_hdr_col};border-bottom:1px solid {_gd_bdr};white-space:nowrap;">'
        f'{col_labels[c]}<span class="sort-arrow"></span></th>'
        for i, c in enumerate(display_cols)
    )}
  </tr></thead>
  <tbody id="{tbl_id}_body">
    {"".join(
        "<tr>" + "".join(
            f'<td style="{cell_style(c, row_dict[c])}">{row_dict[c]}</td>'
            for c in display_cols
        ) + "</tr>"
        for row_dict in rows_json
    )}
  </tbody>
</table>
</div>
<script>
(function() {{
  var sortState = {{}};
  window.sortGrpTable = function(tblId, colIdx) {{
    var tbody = document.getElementById(tblId + '_body');
    var head  = document.getElementById(tblId + '_head');
    var rows  = Array.from(tbody.querySelectorAll('tr'));
    var asc   = sortState[colIdx] !== true;
    sortState = {{}};
    sortState[colIdx] = asc;

    // Update header classes
    Array.from(head.querySelectorAll('th')).forEach(function(th, i) {{
      th.classList.remove('sorted-asc','sorted-desc');
      if (i === colIdx) th.classList.add(asc ? 'sorted-asc' : 'sorted-desc');
    }});

    rows.sort(function(a, b) {{
      var av = a.querySelectorAll('td')[colIdx].textContent.trim();
      var bv = b.querySelectorAll('td')[colIdx].textContent.trim();
      // Try numeric (strip $, %, commas)
      var an = parseFloat(av.replace(/[$,%]/g,'').replace(/,/g,''));
      var bn = parseFloat(bv.replace(/[$,%]/g,'').replace(/,/g,''));
      if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
      // Try date
      var ad = Date.parse(av), bd = Date.parse(bv);
      if (!isNaN(ad) && !isNaN(bd)) return asc ? ad - bd : bd - ad;
      // String fallback
      return asc ? av.localeCompare(bv) : bv.localeCompare(av);
    }});
    rows.forEach(function(r) {{ tbody.appendChild(r); }});
  }};
}})();
</script>"""
        # components.html runs in its own iframe where JS executes fully
        # Calculate height: header row ~40px + each data row ~38px + padding
        tbl_height = 60 + len(rows_json) * 38
        components.html(table_html, height=tbl_height, scrolling=False)


def render_rates_tab(data, hotel):
    rates_ov   = data.get("rates_overview")
    rates_comp = data.get("rates_comp")
    rates_vs7  = data.get("rates_vs7")
    year_df    = data.get("year")
    total_rooms = hotel.get("total_rooms", 112)

    st.markdown('<div class="section-head">Rate Surveillance</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Rate positioning vs comp set · OTB & forecast occupancy by window</div>',
                unsafe_allow_html=True)

    if rates_ov is None:
        st.warning("Rates file not loaded.")
        return

    today = pd.Timestamp(datetime.now().date())

    # ── Three-window rate + occupancy summary ──────────────────────────────────
    windows = [
        ("Next 7 Days",  7),
        ("Next 14 Days", 14),
        ("Next 30 Days", 30),
    ]

    def window_stats(days):
        end = today + pd.Timedelta(days=days - 1)
        mask_r = (rates_ov["Date"] >= today) & (rates_ov["Date"] <= end)
        r_slice = rates_ov[mask_r]

        our_rate     = r_slice["Our_Rate"].dropna().mean()
        median_comp  = r_slice["Median_Comp"].dropna().mean()

        otb_pct = fcst_pct = None
        if year_df is not None and "Date" in year_df.columns:
            mask_y = (year_df["Date"] >= today) & (year_df["Date"] <= end)
            y_slice = year_df[mask_y]
            if not y_slice.empty:
                if "OTB" in y_slice.columns:
                    total_otb  = y_slice["OTB"].sum()
                    total_cap  = len(y_slice) * total_rooms
                    otb_pct    = (total_otb / total_cap * 100) if total_cap > 0 else None
                if "Forecast_Rooms" in y_slice.columns:
                    total_fcst = y_slice["Forecast_Rooms"].dropna().sum()
                    total_cap  = len(y_slice) * total_rooms
                    fcst_pct   = (total_fcst / total_cap * 100) if total_cap > 0 else None

        return our_rate, median_comp, otb_pct, fcst_pct

    def rate_delta_color(our, comp):
        _rs_card_lt2 = False
        if our is None or comp is None:
            return "#4e6878"
        diff = our - comp
        if diff > 5:   return "#cc3300" if _rs_card_lt2 else ORANGE
        if diff < -5:  return "#006a55" if _rs_card_lt2 else TEAL
        return "#7a4e00" if _rs_card_lt2 else GOLD

    def occ_color(pct):
        _rs_card_lt3 = False
        if pct is None: return "#4e6878"
        if pct >= 80:   return "#006a55" if _rs_card_lt3 else TEAL
        if pct >= 60:   return "#7a4e00" if _rs_card_lt3 else GOLD
        return "#cc3300" if _rs_card_lt3 else ORANGE

    _rs_card_lt = False
    _rc_bg       = "#dce5e8"
    _rc_bdr      = "#c4cfd4"
    _rc_lbl      = "#4e6878"
    _rc_sub      = "#4e6878"
    _rc_comp     = "#1e2d35"
    _rc_occ_bg   = "#e8eef2"
    _rc_div      = "#c4cfd4"
    panels_html = '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:24px;">'

    for label, days in windows:
        our_rate, median_comp, otb_pct, fcst_pct = window_stats(days)

        our_disp   = f"${our_rate:.0f}"   if our_rate   is not None else "—"
        comp_disp  = f"${median_comp:.0f}" if median_comp is not None else "—"
        otb_disp   = f"{otb_pct:.1f}%"    if otb_pct    is not None else "—"
        fcst_disp  = f"{fcst_pct:.1f}%"   if fcst_pct   is not None else "—"

        rate_col   = rate_delta_color(our_rate, median_comp)
        otb_col    = occ_color(otb_pct)
        fcst_col   = occ_color(fcst_pct)

        # Delta badge
        if our_rate is not None and median_comp is not None:
            diff = our_rate - median_comp
            sign = "+" if diff >= 0 else ""
            delta_html = (
                f'<span style="font-size:11px;color:{rate_col};font-weight:600;'
                f'margin-left:8px;">{sign}${diff:.0f} vs comp</span>'
            )
        else:
            delta_html = ""

        panel = (
            f'<div style="background:{_rc_bg};border:1px solid {_rc_bdr};border-radius:10px;'
            'padding:20px 22px;position:relative;overflow:hidden;">'
            '<div style="position:absolute;top:0;left:0;right:0;height:3px;'
            f'background:linear-gradient(90deg,{TEAL},{PURPLE});border-radius:10px 10px 0 0;"></div>'
            f'<div style="font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;'
            f'color:{_rc_lbl};margin-bottom:16px;">{label}</div>'
            '<div style="margin-bottom:14px;">'
            f'<div style="font-size:10px;color:{_rc_sub};text-transform:uppercase;'
            'letter-spacing:.08em;margin-bottom:4px;">Our Avg Rate</div>'
            f'<div style="font-size:28px;font-weight:800;color:{rate_col};'
            f'font-family:DM Mono,monospace;line-height:1;">{our_disp}</div>'
            f'{delta_html}'
            '</div>'
            f'<div style="height:1px;background:{_rc_div};margin-bottom:14px;"></div>'
            '<div style="margin-bottom:14px;">'
            f'<div style="font-size:10px;color:{_rc_sub};text-transform:uppercase;'
            'letter-spacing:.08em;margin-bottom:4px;">Median Comp Rate</div>'
            f'<div style="font-size:20px;font-weight:700;color:{_rc_comp};'
            f'font-family:DM Mono,monospace;">{comp_disp}</div>'
            '</div>'
            f'<div style="display:flex;gap:12px;margin-top:4px;">'
            f'<div style="flex:1;background:{_rc_occ_bg};border-radius:6px;padding:10px 12px;">'
            f'<div style="font-size:9px;color:{_rc_sub};text-transform:uppercase;'
            'letter-spacing:.08em;margin-bottom:4px;">OTB Occ</div>'
            f'<div style="font-size:18px;font-weight:700;color:{otb_col};'
            f'font-family:DM Mono,monospace;">{otb_disp}</div>'
            '</div>'
            f'<div style="flex:1;background:{_rc_occ_bg};border-radius:6px;padding:10px 12px;">'
            f'<div style="font-size:9px;color:{_rc_sub};text-transform:uppercase;'
            'letter-spacing:.08em;margin-bottom:4px;">Fcst Occ</div>'
            f'<div style="font-size:18px;font-weight:700;color:{fcst_col};'
            f'font-family:DM Mono,monospace;">{fcst_disp}</div>'
            '</div>'
            '</div>'
            '</div>'
        )
        panels_html += panel

    panels_html += "</div>"
    st.markdown(panels_html, unsafe_allow_html=True)

    # ── 90-day forward rate chart ──
    # ── Day-window selector — shared by both charts ──
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    _w_col, _ = st.columns([2, 5])
    with _w_col:
        _window_labels = {"30 Days": 30, "60 Days": 60, "90 Days": 90}
        _rate_window = st.radio(
            "View window",
            options=list(_window_labels.keys()),
            index=0,
            horizontal=True,
            key="rate_window",
            label_visibility="collapsed",
        )
    _days = _window_labels[_rate_window]

    # ── Our Rate vs Median Comp chart ──
    fwd = rates_ov[rates_ov["Date"] >= today].head(_days)
    if not fwd.empty:
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=fwd["Date"], y=fwd["Our_Rate"],
            name="Our Rate", mode="lines+markers",
            line=dict(color=TEAL, width=2.5),
            marker=dict(size=4),
        ))
        fig.add_trace(go.Scatter(
            x=fwd["Date"], y=fwd["Median_Comp"],
            name="Median Comp", mode="lines",
            line=dict(color=PURPLE, width=1.5, dash="dot"),
        ))
        _layout_r = {k: v for k, v in chart_layout().items() if k not in ("yaxis2",)}
        fig.update_layout(
            **_layout_r,
            title=dict(text=f"Our Rate vs Median Comp · Next {_days} Days",
                       font=dict(size=13, color="#1e2d35")),
        )
        _rs_ax_col = "#3a5260"
        fig.update_xaxes(tickfont=dict(color=_rs_ax_col), color=_rs_ax_col)
        fig.update_yaxes(tickprefix="$", gridcolor=BORDER, tickfont=dict(color=_rs_ax_col))
        st.plotly_chart(fig, use_container_width=True)

    # ── Comp hotel rate grid ──
    if rates_comp is not None:
        st.markdown("---")
        st.markdown('<div class="section-head">Comp Set Rate Grid</div>', unsafe_allow_html=True)
        fwd_comp = rates_comp[
            (rates_comp["Date"] >= today) &
            (rates_comp["Date"] < today + pd.Timedelta(days=_days))
        ]

        # Pivot: dates as rows, hotels as columns.
        # Use Rate_Numeric (pure float, NaN for Sold Out) so Plotly always gets numeric y-values.
        _rate_val_col = "Rate_Numeric" if "Rate_Numeric" in fwd_comp.columns else "Rate"
        pivot = fwd_comp.pivot_table(index="Date", columns="Hotel", values=_rate_val_col, aggfunc="first")
        pivot_display = pivot.copy()
        pivot_display.index = pd.to_datetime(pivot_display.index).strftime("%b %d")
        pivot_display = pivot_display.reset_index().rename(columns={"Date": "Date"})

        # Line chart — all comp hotels
        fig2 = go.Figure()
        comp_colors = [TEAL, PURPLE, ORANGE, GOLD, "#e84393", "#44bbff", "#aaffcc"]
        _rate_cols = [c for c in pivot_display.columns if c != "Date"]
        for i, hotel_name in enumerate(_rate_cols):
            fig2.add_trace(go.Scatter(
                x=pivot_display["Date"], y=pivot_display[hotel_name],
                name=hotel_name, mode="lines",
                line=dict(color=comp_colors[i % len(comp_colors)], width=1.5),
            ))
        # Percentile-based y-axis range: clips sold-out outlier spikes so the
        # typical rate band fills the chart rather than being squashed to the bottom.
        _all_rates = pd.to_numeric(pivot_display[_rate_cols].stack(), errors="coerce").dropna()
        if not _all_rates.empty:
            _p05 = float(_all_rates.quantile(0.05))
            _p95 = float(_all_rates.quantile(0.95))
            _spread = max(_p95 - _p05, 20)
            _y_min = max(0, _p05 - _spread * 0.15)
            _y_max = _p95 + _spread * 0.15
        else:
            _y_min, _y_max = 0, 500
        fig2.update_layout(
            **chart_layout(),
            title=dict(text=f"Comp Set BAR Rates · Next {_days} Days",
                       font=dict(size=13, color="#1e2d35")),
        )
        _rs2_grid = BORDER
        fig2.update_xaxes(tickfont=dict(color=_rs_ax_col), color=_rs_ax_col, gridcolor=_rs2_grid)
        fig2.update_yaxes(
            tickprefix="$",
            gridcolor=_rs2_grid,
            tickfont=dict(color=_rs_ax_col),
            range=[_y_min, _y_max],
            autorange=False,
            tickmode="auto",
            nticks=8,
        )
        st.plotly_chart(fig2, use_container_width=True)
        # ── Comp rate table — currency formatted, all columns centered ──
        tbl_df = pivot_display.head(_days)

        # Pin our hotel (any column containing "(Us)") first, then remaining comp hotels alphabetically
        our_col   = next((c for c in tbl_df.columns if "(us)" in c.lower()), None)
        other_cols = [c for c in tbl_df.columns if c not in ("Date", our_col)]
        hotel_cols = ([our_col] if our_col in tbl_df.columns else []) + sorted(other_cols)

        # Build OTB / Forecast lookup from year_df keyed by formatted date string
        otb_lookup  = {}
        fcst_lookup = {}
        if year_df is not None and "Date" in year_df.columns:
            yr_window = year_df[
                (year_df["Date"] >= today) &
                (year_df["Date"] < today + pd.Timedelta(days=_days))
            ]
            for _, yr in yr_window.iterrows():
                key = yr["Date"].strftime("%b %d")
                if "OTB" in yr_window.columns:
                    otb_lookup[key]  = int(yr["OTB"])  if pd.notna(yr.get("OTB"))            else None
                if "Forecast_Rooms" in yr_window.columns:
                    fcst_lookup[key] = int(yr["Forecast_Rooms"]) if pd.notna(yr.get("Forecast_Rooms")) else None

        # ── Rate Changes persistence — same pattern as snapshot rate change ──
        import json as _json_r
        _rr_path  = cfg.get_hotel_folder(hotel) / "rate_change_rates.json"
        _rr_today = datetime.now().strftime("%Y-%m-%d")
        _rr_hotel = hotel.get("id", "hotel")

        def _rr_load():
            try:
                if _rr_path.exists():
                    d = _json_r.loads(_rr_path.read_text(encoding="utf-8"))
                    if d.get("date") == _rr_today and d.get("hotel") == _rr_hotel:
                        return d.get("values", {})
            except Exception:
                pass
            return {}

        def _rr_save(vals):
            try:
                _rr_path.write_text(_json_r.dumps(
                    {"date": _rr_today, "hotel": _rr_hotel, "values": vals},
                    ensure_ascii=False), encoding="utf-8")
            except Exception:
                pass

        _rr_vals = _rr_load()

        # ── Header row ──
        _th_pad = 'padding:6px 8px'  # slightly tighter padding to make room for Changes col
        _rs_tbl_lt = False
        _rs_hdr_bg   = "#314B63"
        _rs_hdr_bdr  = "#c4cfd4"
        _rs_hdr_acc  = "#7ec8e3"
        _rs_hdr_std  = "#a8c4d4"
        _rs_hdr_chg  = "#ff8c69"
        _rs_tbl_bg   = "transparent"
        _rs_hdr_our  = "#90e090"
        def _th(label, accent=False, changes=False, our=False):
            color = (_rs_hdr_our if our else
                     _rs_hdr_chg if changes else
                     _rs_hdr_acc if accent else _rs_hdr_std)
            min_w = "min-width:80px;" if changes else "min-width:70px;"
            fw = "font-weight:700;" if our else "font-weight:600;"
            return (f'<th style="{_th_pad};text-align:center;{fw}'
                    f'border-bottom:2px solid {_rs_hdr_bdr};white-space:nowrap;'
                    f'position:sticky;top:0;z-index:2;background:{_rs_hdr_bg};'
                    f'{min_w}font-family:DM Mono,monospace;font-size:10px;'
                    f'letter-spacing:0.06em;text-transform:uppercase;'
                    f'color:{color} !important;-webkit-text-fill-color:{color};">{label}</th>')
        th  = _th("Date")
        th += _th("DOW")
        th += _th("OTB", accent=True)
        th += _th("Fcst", accent=True)
        if our_col and our_col in tbl_df.columns:
            th += _th("Our Hotel", our=True)
        th += _th("Changes", changes=True)
        for h in sorted(other_cols):
            th += _th(h)

        # ── Data rows ──
        # Build a vs7 lookup keyed by formatted date string for inline change deltas
        vs7_lookup = {}  # {date_str: {hotel_short: change_val}}
        if rates_vs7 is not None and not rates_vs7.empty:
            for _, vr in rates_vs7.iterrows():
                ds = pd.Timestamp(vr["Date"]).strftime("%b %d")
                vs7_lookup[ds] = {}
                for col in rates_vs7.columns:
                    if col.endswith("__change"):
                        h_name = col.replace("__change", "")
                        vs7_lookup[ds][h_name] = vr[col]

        tbody = ""
        _rs_lt = False
        _rs_date_col = "#1e2d35"
        _rs_cell_bdr = "#d0dde2"
        _rs_otb_col  = "#006a55" if _rs_lt else TEAL
        _rs_fcst_col = "#7a4e00" if _rs_lt else GOLD
        _rs_our_col  = "#6A924D"
        _rs_comp_col = "#3a5260"
        _rs_teal     = "#2E618D"
        _rs_red      = "#a03020"
        _rs_our_bg   = "#e4eaed"
        _td_pad = "padding:7px 10px"

        def _rs_rate_disp(val):
            """Format a rate value safely."""
            if val is None or (isinstance(val, float) and math.isnan(val)) or val == "":
                return "—", False
            if isinstance(val, str):
                return ("SOLD" if "sold" in val.lower() else val[:6]), False
            try:
                return f"${int(round(float(val))):,}", True
            except (ValueError, TypeError):
                return "—", False

        def _rs_chg_html(chg_val):
            if chg_val is None or (isinstance(chg_val, float) and math.isnan(chg_val)):
                return ""
            try:
                chg = int(float(str(chg_val).split("[")[0].strip()))
                if chg > 0:
                    return f'<span style="color:{_rs_teal};font-size:11px;font-weight:600;margin-left:3px;">+{chg}</span>'
                elif chg < 0:
                    return f'<span style="color:{_rs_red};font-size:11px;font-weight:600;margin-left:3px;">{chg}</span>'
            except (ValueError, TypeError):
                pass
            return ""

        for ri, (_, row) in enumerate(tbl_df.iterrows()):
            date_str = row["Date"]
            otb_val  = otb_lookup.get(date_str)
            fcst_val = fcst_lookup.get(date_str)
            otb_disp  = str(otb_val)  if otb_val  is not None else "—"
            fcst_disp = str(fcst_val) if fcst_val is not None else "—"
            row_bg = (_rs_our_bg if ri == 0 else ("#e8eef1")) if ri % 2 == 0 else ("" if _rs_lt else "#ffffff")
            row_bg_style = f"background:{row_bg};" if row_bg else ""

            td  = (f'<td style="{_td_pad};{row_bg_style}border-bottom:1px solid {_rs_cell_bdr};'
                   f'color:{_rs_date_col};font-family:DM Mono,monospace;font-size:13px;'
                   f'font-weight:600;text-align:center;white-space:nowrap;">{date_str}</td>')
            # Day of week
            try:
                _dow_yr   = today.year if datetime.strptime(date_str + f" {today.year}", "%b %d %Y").month >= today.month else today.year + 1
                _dow_dt   = datetime.strptime(date_str + f" {_dow_yr}", "%b %d %Y")
                _dow_str  = _dow_dt.strftime("%a")
                _dow_wknd = _dow_str in ("Fri", "Sat")
                _dow_col  = ("#6A924D") if _dow_wknd else ("#4e6878")
            except Exception:
                _dow_str  = "—"
                _dow_col  = _rs_date_col
            td += (f'<td style="{_td_pad};{row_bg_style}border-bottom:1px solid {_rs_cell_bdr};'
                   f'color:{_dow_col};font-family:DM Mono,monospace;font-size:13px;'
                   f'font-weight:600;text-align:center;white-space:nowrap;">{_dow_str}</td>')
            # OTB Rooms
            td += (f'<td style="{_td_pad};{row_bg_style}border-bottom:1px solid {_rs_cell_bdr};'
                   f'text-align:center;font-family:DM Mono,monospace;font-size:13px;'
                   f'color:{_rs_otb_col};font-weight:600;">{otb_disp}</td>')
            # Forecast Rooms
            td += (f'<td style="{_td_pad};{row_bg_style}border-bottom:1px solid {_rs_cell_bdr};'
                   f'text-align:center;font-family:DM Mono,monospace;font-size:13px;'
                   f'color:{_rs_fcst_col};font-weight:600;">{fcst_disp}</td>')

            # Rate columns — Our Hotel first, then comps
            _changes_inserted = False
            for h in hotel_cols:
                val = row.get(h)
                rate_str, is_num = _rs_rate_disp(val)
                chg_val  = vs7_lookup.get(date_str, {}).get(h)
                chg_html = _rs_chg_html(chg_val) if is_num else ""
                is_our_h = (our_col is not None and h == our_col)
                color  = _rs_our_col if is_our_h else _rs_comp_col
                weight = "700" if is_our_h else "400"
                font   = "14px" if is_our_h else "13px"
                bdr_l  = f"border-left:2px solid {_rs_cell_bdr};" if is_our_h else ""
                td += (f'<td style="{_td_pad};{row_bg_style}{bdr_l}border-bottom:1px solid {_rs_cell_bdr};'
                       f'text-align:center;font-family:DM Mono,monospace;font-size:{font};'
                       f'color:{color};font-weight:{weight};white-space:nowrap;">'
                       f'{rate_str}{chg_html}</td>')
                # Changes column always appears right after Our Hotel (or after Fcst if no our_col)
                if is_our_h:
                    rc_val  = _rr_vals.get(date_str, "")
                    rc_disp = rc_val if rc_val else "—"
                    rc_col  = ("#b44820") if rc_val else ("#4a6070")
                    td += (f'<td style="{_td_pad};{row_bg_style}border-bottom:1px solid {_rs_cell_bdr};'
                           f'text-align:center;font-family:DM Mono,monospace;font-size:13px;'
                           f'font-weight:700;color:{rc_col};min-width:80px;'
                           f'border-left:1px solid {_rs_cell_bdr};">{rc_disp}</td>')
                    _changes_inserted = True
            # If no our_col was found, still insert Changes col so header aligns
            if not _changes_inserted:
                rc_val  = _rr_vals.get(date_str, "")
                rc_disp = rc_val if rc_val else "—"
                rc_col  = ("#b44820") if rc_val else ("#4a6070")
                td += (f'<td style="{_td_pad};{row_bg_style}border-bottom:1px solid {_rs_cell_bdr};'
                       f'text-align:center;font-family:DM Mono,monospace;font-size:13px;'
                       f'font-weight:700;color:{rc_col};min-width:80px;'
                       f'border-left:1px solid {_rs_cell_bdr};">{rc_disp}</td>')
            tbody += f"<tr>{td}</tr>"

        comp_tbl_html = (
            f'<div style="overflow-x:auto;overflow-y:auto;max-height:480px;border:1px solid {_rs_hdr_bdr};border-radius:6px;margin-top:4px;">'
            f'<table style="width:100%;border-collapse:collapse;font-size:13px;font-family:DM Sans,sans-serif;background:{_rs_tbl_bg};">'
            f'<thead style="position:sticky;top:0;z-index:2;"><tr style="background:{_rs_hdr_bg};">{th}</tr></thead>'
            f'<tbody>{tbody}</tbody>'
            '</table></div>'
        )
        st.markdown(comp_tbl_html, unsafe_allow_html=True)

        # ── Changes column editor — same persistence pattern as snapshot rate change ──
        st.markdown(
            '<div style="background:#ffffff;border:1px solid #c4cfd4;border-radius:8px;overflow:hidden;margin-top:10px;">'
            '<div style="background:#2E618D;border-bottom:1px solid #1e4a6e;padding:8px 14px;display:flex;align-items:center;gap:12px;">'
            '<span style="font-family:DM Mono,monospace;font-size:10px;letter-spacing:0.18em;text-transform:uppercase;color:#b44820;font-weight:700;">📝 Rate Changes</span>'
            '<span style="font-family:DM Mono,monospace;font-size:10px;color:#5a7080;">Enter planned rate changes by date · saved until midnight · auto-reset each day</span>'
            '</div></div>',
            unsafe_allow_html=True)

        _hid_r = hotel.get("id", "hotel")
        date_keys = [row["Date"] for _, row in tbl_df.iterrows()]
        n_cols = min(len(date_keys), 15)   # show up to 15 per row to match table width
        rc_changed = False
        # Chunk into rows of 15
        for chunk_start in range(0, len(date_keys), 15):
            chunk = date_keys[chunk_start:chunk_start + 15]
            ecols = st.columns(len(chunk))
            for ecol, dk in zip(ecols, chunk):
                cur = _rr_vals.get(dk, "")
                nv  = ecol.text_input(dk, value=cur, placeholder="—",
                                      key=f"rrc_{_hid_r}_{dk}",
                                      max_chars=10)
                if nv != cur:
                    _rr_vals[dk] = nv
                    rc_changed = True
        if rc_changed:
            _rr_save(_rr_vals)
            st.rerun()


def render_str_tab(data, hotel):
    str_data  = data.get("str")
    str_ranks = data.get("str_ranks")
    files     = cfg.detect_files(hotel) if isinstance(hotel, dict) else {}

    st.markdown('<div class="section-head">Tactical Outcome Analysis</div>', unsafe_allow_html=True)

    if str_data is None:
        st.warning("STR file not loaded.")
        return

    period   = str_data.get("period", "")
    weekly   = str_data.get("weekly")
    day28    = str_data.get("28day")

    st.markdown(f'<div class="section-sub">Week of: {period}</div>', unsafe_allow_html=True)

    if weekly is not None:
        total_w  = weekly[weekly["Day"] == "Total"]
        total_28 = day28[day28["Day"] == "Total"] if day28 is not None else pd.DataFrame()

        if not total_w.empty:
            tw = total_w.iloc[0]
            t28 = total_28.iloc[0] if not total_28.empty else None

            TOTAL_ROOMS = hotel.get("total_rooms", 112)

            def safe(row, col, default=None):
                if row is None: return default
                v = row.get(col)
                return v if pd.notna(v) else default

            # ── Calculate metrics for Current Week ──
            occ_mine_w  = safe(tw, "Occ_Mine")
            occ_comp_w  = safe(tw, "Occ_Comp")
            adr_mine_w  = safe(tw, "ADR_Mine")
            adr_comp_w  = safe(tw, "ADR_Comp")
            rev_mine_w  = safe(tw, "RevPAR_Mine")
            rev_comp_w  = safe(tw, "RevPAR_Comp")
            mpi_w       = safe(tw, "MPI")
            ari_w       = safe(tw, "ARI")
            rgi_w       = safe(tw, "RGI")

            rooms_mine_w  = round(occ_mine_w / 100 * TOTAL_ROOMS * 7) if occ_mine_w else None
            rooms_comp_w  = round(occ_comp_w / 100 * TOTAL_ROOMS * 7) if occ_comp_w else None
            rev_mine_w_r  = round(rooms_mine_w * adr_mine_w) if rooms_mine_w and adr_mine_w else None
            rev_comp_w_r  = round(rooms_comp_w * adr_comp_w) if rooms_comp_w and adr_comp_w else None

            # ── Calculate metrics for Running 28 Days ──
            occ_mine_28 = safe(t28, "Occ_Mine")
            occ_comp_28 = safe(t28, "Occ_Comp")
            adr_mine_28 = safe(t28, "ADR_Mine")
            adr_comp_28 = safe(t28, "ADR_Comp")
            mpi_28      = safe(t28, "MPI")
            ari_28      = safe(t28, "ARI")
            rgi_28      = safe(t28, "RGI")

            rooms_mine_28 = round(occ_mine_28 / 100 * TOTAL_ROOMS * 28) if occ_mine_28 else None
            rooms_comp_28 = round(occ_comp_28 / 100 * TOTAL_ROOMS * 28) if occ_comp_28 else None
            rev_mine_28_r = round(rooms_mine_28 * adr_mine_28) if rooms_mine_28 and adr_mine_28 else None
            rev_comp_28_r = round(rooms_comp_28 * adr_comp_28) if rooms_comp_28 and adr_comp_28 else None

            # ── Helpers ──
            def fmt_int(v):   return f"{int(v):,}" if v is not None else "—"
            def fmt_dol(v):   return f"${v:,.2f}"  if v is not None else "—"
            def fmt_dol0(v):  return f"${int(v):,}" if v is not None else "—"
            def fmt_idx(v):   return f"{v:.1f}"    if v is not None else "—"
            def fmt_pct(v):   return f"{v:.1f}%"   if v is not None else "—"

            def delta_int(mine, comp):
                if mine is None or comp is None: return "—", "#4e6878"
                d = mine - comp
                col = TEAL if d >= 0 else ORANGE
                return (f"+{int(d):,}" if d >= 0 else f"{int(d):,}"), col

            def delta_dol(mine, comp):
                if mine is None or comp is None: return "—", "#4e6878"
                d = mine - comp
                col = TEAL if d >= 0 else ORANGE
                sign = "+" if d >= 0 else ""
                return f"{sign}${d:,.2f}", col

            def delta_dol0(mine, comp):
                if mine is None or comp is None: return "—", "#4e6878"
                d = mine - comp
                col = TEAL if d >= 0 else ORANGE
                sign = "+" if d >= 0 else ""
                return f"{sign}${int(d):,}", col

            def idx_color(v):
                if v is None: return "#4e6878"
                return TEAL if v >= 100 else ORANGE

            # ── Deltas ──
            d_rooms_w,   dc_rooms_w   = delta_int(rooms_mine_w,  rooms_comp_w)
            d_adr_w,     dc_adr_w     = delta_dol(adr_mine_w,    adr_comp_w)
            d_rev_w,     dc_rev_w     = delta_dol0(rev_mine_w_r, rev_comp_w_r)
            d_rooms_28,  dc_rooms_28  = delta_int(rooms_mine_28, rooms_comp_28)
            d_adr_28,    dc_adr_28    = delta_dol(adr_mine_28,   adr_comp_28)
            d_rev_28,    dc_rev_28    = delta_dol0(rev_mine_28_r, rev_comp_28_r)

            # ── Pull ranks ──
            rk = str_ranks or {}
            occ_rk   = rk.get("occ",    {"week": "—", "run28": "—", "mtd": "—"})
            adr_rk   = rk.get("adr",    {"week": "—", "run28": "—", "mtd": "—"})
            rev_rk   = rk.get("revpar", {"week": "—", "run28": "—", "mtd": "—"})

            # ── Theme-aware style tokens ──
            _lt = False
            _panel_bg = '#e8f0f8' if _lt else '#dce5e8'
            _hdr_bg   = '#314B63'
            _hdr_col  = '#ffffff'
            _hdr_bdr  = '#1a3a5c' if _lt else '#c4cfd4'
            _lbl_col  = '#4e6878'
            _cell_bdr = '#c0d0e0' if _lt else '#d0dde2'
            _val_col  = '#1e2d35'
            _rank_col = '#1a3a5c' if _lt else '#1e2d35'
            _foot_bg  = '#314B63'
            _foot_col = '#e8eef2'
            _panel_bdr = '#1a3a5c' if _lt else '#c4cfd4'

            HDR_S  = (f'background:{_hdr_bg};color:{_hdr_col};font-weight:600;font-size:11px;'
                      f'padding:7px 10px;text-align:center;border-bottom:1px solid {_hdr_bdr};'
                       'letter-spacing:.05em;text-transform:uppercase;white-space:nowrap;')
            LBL_S  = (f'padding:7px 10px;color:{_lbl_col};font-size:12px;font-weight:600;'
                      f'white-space:nowrap;border-bottom:1px solid {_cell_bdr};')
            FOOT_S = (f'background:{_foot_bg};padding:5px 12px 7px;font-size:10px;color:{_foot_col};'
                       'letter-spacing:.06em;text-transform:uppercase;')

            def th(label):
                return f'<th style="{HDR_S}">{label}</th>'

            def th_blank():
                return f'<th style="{HDR_S}text-align:left;"></th>'

            def vc(val, bold=False, color=None):
                c = color if color else _val_col
                w = "font-weight:700;" if bold else ""
                return (f'<td style="padding:7px 10px;text-align:center;font-family:DM Mono,monospace;'
                        f'font-size:13px;{w}color:{c};border-bottom:1px solid {_cell_bdr};white-space:nowrap;">{val}</td>')

            def lbl(text):
                return f'<td style="{LBL_S}">{text}</td>'

            def delta_cell(val, color):
                return (f'<td style="padding:7px 10px;text-align:center;font-family:DM Mono,monospace;'
                        f'font-size:12px;font-weight:700;color:{color};'
                        f'border-bottom:1px solid {_cell_bdr};">{val}</td>')

            # Fixed height for all three panels so they match exactly
            TBL_H = "height:134px;"

            # ── Occ% deltas ──
            def delta_pct(mine, comp):
                if mine is None or comp is None: return "—", "#4e6878"
                d = mine - comp
                col = TEAL if d >= 0 else ORANGE
                return (f"+{d:.1f}%" if d >= 0 else f"{d:.1f}%"), col

            d_occ_w,  dc_occ_w  = delta_pct(occ_mine_w,  occ_comp_w)
            d_occ_28, dc_occ_28 = delta_pct(occ_mine_28, occ_comp_28)

            # ── Build the three-panel HTML ──
            panel_html = (
                '<div style="display:flex;gap:10px;margin-bottom:20px;align-items:stretch;">'

                # ── Panel 1: Current Week ──
                f'<div style="flex:1;min-width:0;background:{_panel_bg};border:1px solid {_panel_bdr};border-radius:8px;overflow:hidden;display:flex;flex-direction:column;">'
                f'<table style="width:100%;border-collapse:collapse;font-size:12px;{TBL_H}">'
                f'<thead><tr>{th_blank()}{th("Rooms Sold")}{th("Occ %")}{th("ADR")}{th("Revenue")}</tr></thead>'
                '<tbody>'
                f'<tr>{lbl("Our Hotel")}{vc(fmt_int(rooms_mine_w),bold=True)}{vc(fmt_pct(occ_mine_w),bold=True)}{vc(fmt_dol(adr_mine_w),bold=True)}{vc(fmt_dol0(rev_mine_w_r),bold=True)}</tr>'
                f'<tr>{lbl("Avg Comp Hotel")}{vc(fmt_int(rooms_comp_w))}{vc(fmt_pct(occ_comp_w))}{vc(fmt_dol(adr_comp_w))}{vc(fmt_dol0(rev_comp_w_r))}</tr>'
                f'<tr style="border-top:2px solid {_panel_bdr};">'
                f'<td style="{LBL_S}color:{TEAL};">Results +/-</td>'
                f'{delta_cell(d_rooms_w,dc_rooms_w)}{delta_cell(d_occ_w,dc_occ_w)}{delta_cell(d_adr_w,dc_adr_w)}{delta_cell(d_rev_w,dc_rev_w)}</tr>'
                '</tbody></table>'
                f'<div style="{FOOT_S}margin-top:auto;">Current Week</div>'
                '</div>'

                # ── Panel 2: Index + Rank ──
                f'<div style="flex:0 0 380px;background:{_panel_bg};border:1px solid {_panel_bdr};border-radius:8px;overflow:hidden;display:flex;flex-direction:column;">'
                f'<table style="width:100%;border-collapse:collapse;font-size:12px;{TBL_H}">'
                f'<thead><tr>{th_blank()}{th("This Week")}{th("Rank")}{th("28 Days")}{th("Rank")}</tr></thead>'
                '<tbody>'
                f'<tr>{lbl("MPI (Occ)")}'
                f'{vc(fmt_idx(mpi_w),   bold=True, color=idx_color(mpi_w))}'
                f'{vc(occ_rk["week"],   color=_rank_col)}'
                f'{vc(fmt_idx(mpi_28),  bold=True, color=idx_color(mpi_28))}'
                f'{vc(occ_rk["run28"],  color=_rank_col)}'
                f'</tr>'
                f'<tr>{lbl("ARI (ADR)")}'
                f'{vc(fmt_idx(ari_w),   bold=True, color=idx_color(ari_w))}'
                f'{vc(adr_rk["week"],   color=_rank_col)}'
                f'{vc(fmt_idx(ari_28),  bold=True, color=idx_color(ari_28))}'
                f'{vc(adr_rk["run28"],  color=_rank_col)}'
                f'</tr>'
                f'<tr>{lbl("RGI (RevPAR)")}'
                f'{vc(fmt_idx(rgi_w),   bold=True, color=idx_color(rgi_w))}'
                f'{vc(rev_rk["week"],   color=_rank_col)}'
                f'{vc(fmt_idx(rgi_28),  bold=True, color=idx_color(rgi_28))}'
                f'{vc(rev_rk["run28"],  color=_rank_col)}'
                f'</tr>'
                '</tbody></table>'
                f'<div style="{FOOT_S}margin-top:auto;">Index Scores · &gt;100 = above comp set</div>'
                '</div>'

                # ── Panel 3: Running 28 Days ──
                f'<div style="flex:1;min-width:0;background:{_panel_bg};border:1px solid {_panel_bdr};border-radius:8px;overflow:hidden;display:flex;flex-direction:column;">'
                f'<table style="width:100%;border-collapse:collapse;font-size:12px;{TBL_H}">'
                f'<thead><tr>{th_blank()}{th("Rooms Sold")}{th("Occ %")}{th("ADR")}{th("Revenue")}</tr></thead>'
                '<tbody>'
                f'<tr>{lbl("Our Hotel")}{vc(fmt_int(rooms_mine_28),bold=True)}{vc(fmt_pct(occ_mine_28),bold=True)}{vc(fmt_dol(adr_mine_28),bold=True)}{vc(fmt_dol0(rev_mine_28_r),bold=True)}</tr>'
                f'<tr>{lbl("Avg Comp Hotel")}{vc(fmt_int(rooms_comp_28))}{vc(fmt_pct(occ_comp_28))}{vc(fmt_dol(adr_comp_28))}{vc(fmt_dol0(rev_comp_28_r))}</tr>'
                f'<tr style="border-top:2px solid {_panel_bdr};">'
                f'<td style="{LBL_S}color:{TEAL};">Results +/-</td>'
                f'{delta_cell(d_rooms_28,dc_rooms_28)}{delta_cell(d_occ_28,dc_occ_28)}{delta_cell(d_adr_28,dc_adr_28)}{delta_cell(d_rev_28,dc_rev_28)}</tr>'
                '</tbody></table>'
                f'<div style="{FOOT_S}margin-top:auto;">Running 28 Days</div>'
                '</div>'

                '</div>'  # end flex container
            )
            st.markdown(panel_html, unsafe_allow_html=True)

        view = st.radio("Period", ["This Week", "Running 28 Days"], horizontal=True)
        df   = weekly if view == "This Week" else (day28 if day28 is not None else weekly)
        day_df = df[df["Day"] != "Total"]

        cols = st.columns(2)
        with cols[0]:
            fig = fig_bar_index(day_df, "Day", "Occ_Mine", "Occ_Comp", "MPI",
                                "Occupancy % by Day of Week", y_fmt="pct")
            st.plotly_chart(fig, use_container_width=True)
        with cols[1]:
            fig2 = fig_bar_index(day_df, "Day", "ADR_Mine", "ADR_Comp", "ARI",
                                 "ADR by Day of Week", y_fmt="dollar")
            st.plotly_chart(fig2, use_container_width=True)

        fig3 = fig_bar_index(day_df, "Day", "RevPAR_Mine", "RevPAR_Comp", "RGI",
                              "RevPAR by Day of Week", y_fmt="dollar")
        st.plotly_chart(fig3, use_container_width=True)

        # ── Weekly STR Report — formatted like official STR report ──────────
        with st.expander("📋 STR Report"):
            days_order = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Total"]
            df_disp = df.copy()
            # Ensure rows in day order
            df_disp["_sort"] = df_disp["Day"].apply(lambda d: days_order.index(d) if d in days_order else 99)
            df_disp = df_disp.sort_values("_sort").drop(columns=["_sort"])

            cols_days = [d for d in days_order if d in df_disp["Day"].values]

            # ── Theme-aware style tokens — Snapshot palette ──────────────
            _str_lt = False
            _s_h0_bg  = "#dce5e8"   # Snapshot bg_header
            _s_h0_col = "#4e6878"   # Snapshot txt_header
            _s_bdr    = "#c4cfd4"   # Snapshot bdr
            _s_sec_bg = "#314B63"   # Snapshot bg_lbl_cell
            _s_sec_col= "#ffffff"   # white on navy section bar
            _s_r1_bg  = "#f5f8f9"   # Snapshot bg_row_std
            _s_r1_col = "#1e2d35"   # Snapshot txt_pri
            _s_r1_bdr = "#d0dde2"   # Snapshot bdr2
            _s_r2_bg  = "#eaeff1"   # Snapshot bg_row_alt
            _s_r2_col = "#3a5260"   # Snapshot txt_sec
            _s_lbl1_col = "#1e2d35"  # Snapshot txt_pri
            _s_lbl2_col = "#3a5260"  # Snapshot txt_sec
            _s_lbli_col = "#4e6878"  # Snapshot txt_dim
            # Light mode index colors matching Snapshot accent palette
            _s_idx_pos_col = "#1a4a7a" if _str_lt else TEAL     # Snapshot txt_teal
            _s_idx_neg_col = "#cc3300" if _str_lt else ORANGE   # Snapshot txt_orange

            H0 = f"background:{_s_h0_bg};color:{_s_h0_col};font-weight:700;font-size:11px;padding:8px 10px;text-align:center;border:1px solid {_s_bdr};letter-spacing:.05em;text-transform:uppercase;white-space:nowrap;"
            HL = f"background:{_s_h0_bg};color:{_s_h0_col};font-weight:700;font-size:11px;padding:8px 10px;text-align:left;border:1px solid {_s_bdr};letter-spacing:.05em;text-transform:uppercase;white-space:nowrap;"
            SEC = f"background:{_s_sec_bg};color:{_s_sec_col};font-weight:700;font-size:11px;padding:6px 10px;text-align:left;border:1px solid {_s_bdr};letter-spacing:.12em;text-transform:uppercase;"
            R1  = f"background:{_s_r1_bg};color:{_s_r1_col};font-size:13px;font-family:DM Mono,monospace;padding:7px 10px;text-align:center;border:1px solid {_s_r1_bdr};"
            R2  = f"background:{_s_r2_bg};color:{_s_r2_col};font-size:13px;font-family:DM Mono,monospace;padding:7px 10px;text-align:center;border:1px solid {_s_r1_bdr};"
            IDX_POS = f"background:{_s_r2_bg};color:{_s_idx_pos_col};font-weight:700;font-size:13px;font-family:DM Mono,monospace;padding:7px 10px;text-align:center;border:1px solid {_s_r1_bdr};"
            IDX_NEG = f"background:{_s_r2_bg};color:{_s_idx_neg_col};font-weight:700;font-size:13px;font-family:DM Mono,monospace;padding:7px 10px;text-align:center;border:1px solid {_s_r1_bdr};"
            R1_CHG  = f"background:{_s_r1_bg};color:{_s_r1_col};font-size:11px;font-style:italic;font-family:DM Mono,monospace;padding:7px 6px;text-align:center;border:1px solid {_s_r1_bdr};"
            R2_CHG  = f"background:{_s_r2_bg};color:{_s_r2_col};font-size:11px;font-style:italic;font-family:DM Mono,monospace;padding:7px 6px;text-align:center;border:1px solid {_s_r1_bdr};"
            IDX_CHG = f"background:{_s_r2_bg};font-size:11px;font-style:italic;font-family:DM Mono,monospace;padding:7px 6px;text-align:center;border:1px solid {_s_r1_bdr};"
            LBL1 = f"background:{_s_r1_bg};color:{_s_lbl1_col};font-size:12px;padding:7px 12px;border:1px solid {_s_r1_bdr};white-space:nowrap;"
            LBL2 = f"background:{_s_r2_bg};color:{_s_lbl2_col};font-size:12px;padding:7px 12px;border:1px solid {_s_r1_bdr};white-space:nowrap;"
            LBLI = f"background:{_s_r2_bg};color:{_s_lbli_col};font-size:12px;font-style:italic;padding:7px 12px;border:1px solid {_s_r1_bdr};white-space:nowrap;"

            def get_val(day, col):
                row = df_disp[df_disp["Day"] == day]
                if row.empty: return None
                v = row.iloc[0].get(col)
                return v if pd.notna(v) else None

            def td(val, style):
                return f'<td style="{style}">{val}</td>'

            # Build header row
            header = f'<tr><th style="{HL}"></th>'
            for d in cols_days:
                short = d[:3] if d != "Total" else "Total"
                header += f'<th style="{H0}">{short}</th><th style="{H0}">% Chg</th>'
            header += "</tr>"

            def fmt_chg(v):
                """Format % change value: show as e.g. 16.0 or -2.6 (no % sign, matching STR style)."""
                if v is None: return "—"
                return f"{v:+.1f}" if v != 0 else "0.0"

            def chg_style(v, row_type):
                """Return italic % chg style with snapshot-themed color based on value and row type."""
                base = IDX_CHG if row_type == "idx" else (R1_CHG if row_type == "mine" else R2_CHG)
                color = _s_idx_pos_col if (v is not None and v >= 0) else _s_idx_neg_col
                return base + f"color:{color};"

            def make_section(section_label, mine_col, comp_col, idx_col,
                             mine_chg_col, comp_chg_col, idx_chg_col,
                             fmt_mine, fmt_comp, fmt_idx, idx_label):
                """Build 3 rows: My Property, Comp Set, Index — each with value + % Chg columns."""
                ncols = 1 + len(cols_days) * 2
                sec_row = f'<tr><td colspan="{ncols}" style="{SEC}">{section_label}</td></tr>'

                r_mine = f'<tr><td style="{LBL1}">My Property</td>'
                r_comp = f'<tr><td style="{LBL2}">Comp Set</td>'
                r_idx  = f'<tr><td style="{LBLI}">{idx_label}</td>'

                for d in cols_days:
                    mv  = get_val(d, mine_col);    mc  = get_val(d, mine_chg_col)
                    cv  = get_val(d, comp_col);    cc  = get_val(d, comp_chg_col)
                    iv  = get_val(d, idx_col);     ic  = get_val(d, idx_chg_col)

                    r_mine += td(fmt_mine(mv) if mv is not None else "—", R1)
                    r_mine += td(fmt_chg(mc), chg_style(mc, "mine"))
                    r_comp += td(fmt_comp(cv) if cv is not None else "—", R2)
                    r_comp += td(fmt_chg(cc), chg_style(cc, "comp"))
                    idx_s = IDX_POS if (iv and iv >= 100) else IDX_NEG
                    r_idx  += td(fmt_idx(iv) if iv is not None else "—", idx_s)
                    r_idx  += td(fmt_chg(ic), chg_style(ic, "idx"))

                r_mine += "</tr>"; r_comp += "</tr>"; r_idx += "</tr>"
                # Blank spacer row between sections
                return sec_row + r_mine + r_comp + r_idx

            occ_section = make_section(
                "Occupancy",
                "Occ_Mine", "Occ_Comp", "MPI",
                "Occ_Mine_Chg", "Occ_Comp_Chg", "MPI_Chg",
                lambda v: f"{v:.1f}%",
                lambda v: f"{v:.1f}%",
                lambda v: f"{v:.1f}",
                "Index (MPI)"
            )
            adr_section = make_section(
                "ADR",
                "ADR_Mine", "ADR_Comp", "ARI",
                "ADR_Mine_Chg", "ADR_Comp_Chg", "ARI_Chg",
                lambda v: f"${v:.2f}",
                lambda v: f"${v:.2f}",
                lambda v: f"{v:.1f}",
                "Index (ARI)"
            )
            rev_section = make_section(
                "RevPAR",
                "RevPAR_Mine", "RevPAR_Comp", "RGI",
                "RevPAR_Mine_Chg", "RevPAR_Comp_Chg", "RGI_Chg",
                lambda v: f"${v:.2f}",
                lambda v: f"${v:.2f}",
                lambda v: f"{v:.1f}",
                "Index (RGI)"
            )

            raw_html = (
                '<div style="overflow-x:auto;margin-top:8px;">'
                '<table style="border-collapse:collapse;width:100%;font-family:DM Sans,sans-serif;">'
                f'<thead>{header}</thead>'
                f'<tbody>{occ_section}{adr_section}{rev_section}</tbody>'
                '</table></div>'
            )
            st.markdown(raw_html, unsafe_allow_html=True)

        # ── STR Segmentation Report ──────────────────────────────────────────
        with st.expander("📊 STR Segmentation Report"):
            str_seg = data.get("str_seg")
            if str_seg is None:
                st.info("Segmentation Glance sheet not found in STR file.")
            else:
                # Uses the same 'view' toggle as the charts above (This Week / Running 28 Days)
                seg_df = str_seg.get("weekly") if view == "This Week" else str_seg.get("28day")
                seg_period = str_seg.get("period", "")

                if seg_df is None or seg_df.empty:
                    st.info("No segmentation data available for this period.")
                else:
                    # ── Reuse theme tokens from STR Report above ──────────────
                    SEGS_ORDER = ["Transient", "Group", "Contract", "Total"]
                    seg_df = seg_df.copy()
                    seg_df["_sort"] = seg_df["Segment"].apply(
                        lambda s: SEGS_ORDER.index(s) if s in SEGS_ORDER else 99)
                    seg_df = seg_df.sort_values("_sort").drop(columns=["_sort"])

                    def get_seg_val(seg, col):
                        row = seg_df[seg_df["Segment"] == seg]
                        if row.empty: return None
                        v = row.iloc[0].get(col)
                        return v if pd.notna(v) else None

                    segs_present = [s for s in SEGS_ORDER if s in seg_df["Segment"].values]

                    # ── Date header row ───────────────────────────────────────
                    n_seg_cols = len(segs_present) * 2 + 1
                    date_hdr = (
                        f'<tr><td colspan="{n_seg_cols}" style="'
                        f'background:{_s_h0_bg};color:{_s_h0_col};font-weight:700;'
                        f'font-size:12px;padding:10px 14px;text-align:center;'
                        f'border:1px solid {_s_bdr};letter-spacing:.04em;">'
                        f'{seg_period}</td></tr>'
                    )

                    # ── Column header rows ────────────────────────────────────
                    seg_hdr1 = f'<tr><th style="{HL}"></th>'
                    for seg in segs_present:
                        seg_hdr1 += (
                            f'<th colspan="2" style="{H0}">{seg}</th>'
                        )
                    seg_hdr1 += "</tr>"

                    seg_hdr2 = f'<tr><th style="{HL}"></th>'
                    for _ in segs_present:
                        seg_hdr2 += (
                            f'<th style="{H0}">Value</th>'
                            f'<th style="{H0}">% Chg</th>'
                        )
                    seg_hdr2 += "</tr>"

                    seg_thead = date_hdr + seg_hdr1 + seg_hdr2

                    def make_seg_section(section_label,
                                         mine_col, mine_chg_col,
                                         comp_col, comp_chg_col,
                                         idx_col,  idx_chg_col,
                                         fmt_mine, fmt_comp, fmt_idx,
                                         idx_label):
                        ncols = 1 + len(segs_present) * 2
                        sec_row = f'<tr><td colspan="{ncols}" style="{SEC}">{section_label}</td></tr>'

                        r_mine = f'<tr><td style="{LBL1}">My Property</td>'
                        r_comp = f'<tr><td style="{LBL2}">Comp set</td>'
                        r_idx  = f'<tr><td style="{LBLI}">{idx_label}</td>'

                        for seg in segs_present:
                            mv = get_seg_val(seg, mine_col);  mc = get_seg_val(seg, mine_chg_col)
                            cv = get_seg_val(seg, comp_col);  cc = get_seg_val(seg, comp_chg_col)
                            iv = get_seg_val(seg, idx_col);   ic = get_seg_val(seg, idx_chg_col)

                            r_mine += td(fmt_mine(mv) if mv is not None else "—", R1)
                            r_mine += td(fmt_chg(mc), chg_style(mc, "mine"))
                            r_comp += td(fmt_comp(cv) if cv is not None else "—", R2)
                            r_comp += td(fmt_chg(cc), chg_style(cc, "comp"))
                            idx_s = IDX_POS if (iv and iv >= 100) else IDX_NEG
                            r_idx  += td(fmt_idx(iv) if iv is not None else "—", idx_s)
                            r_idx  += td(fmt_chg(ic), chg_style(ic, "idx"))

                        r_mine += "</tr>"; r_comp += "</tr>"; r_idx += "</tr>"
                        return sec_row + r_mine + r_comp + r_idx

                    seg_occ = make_seg_section(
                        "Occupancy",
                        "Occ_Mine", "Occ_Mine_Chg",
                        "Occ_Comp", "Occ_Comp_Chg",
                        "MPI",      "MPI_Chg",
                        lambda v: f"{v:.1f}%",
                        lambda v: f"{v:.1f}%",
                        lambda v: f"{v:.1f}",
                        "Index (MPI)"
                    )
                    seg_adr = make_seg_section(
                        "ADR",
                        "ADR_Mine", "ADR_Mine_Chg",
                        "ADR_Comp", "ADR_Comp_Chg",
                        "ARI",      "ARI_Chg",
                        lambda v: f"${v:.2f}",
                        lambda v: f"${v:.2f}",
                        lambda v: f"{v:.1f}",
                        "Index (ARI)"
                    )
                    seg_rev = make_seg_section(
                        "RevPAR",
                        "RevPAR_Mine", "RevPAR_Mine_Chg",
                        "RevPAR_Comp", "RevPAR_Comp_Chg",
                        "RGI",         "RGI_Chg",
                        lambda v: f"${v:.2f}",
                        lambda v: f"${v:.2f}",
                        lambda v: f"{v:.1f}",
                        "Index (RGI)"
                    )

                    seg_html = (
                        '<div style="overflow-x:auto;margin-top:8px;">'
                        '<table style="border-collapse:collapse;width:100%;font-family:DM Sans,sans-serif;">'
                        f'<thead>{seg_thead}</thead>'
                        f'<tbody>{seg_occ}{seg_adr}{seg_rev}</tbody>'
                        '</table></div>'
                    )
                    st.markdown(seg_html, unsafe_allow_html=True)




def render_srp_pace_tab(data, hotel):
    """SRP Pace -- segment-level booking pace vs same time last year."""
    import datetime as _dt

    pace_df = data.get("srp_pace")

    st.markdown('''<div class="section-head">SRP Pace</div>
<div class="section-sub">Booking pace by rate segment vs same time last year</div>''',
        unsafe_allow_html=True)

    if pace_df is None or pace_df.empty:
        st.info("No SRP Pace file found. Add an SRP Pace.xlsx to the hotel folder.")
        return

    # ── IHG hotels: monthly segment data — render a different layout ──────────
    if data.get("_ihg_hotel"):
        _render_ihg_segment_tab(pace_df, hotel, data)
        return

    _lt = False

    # ── Date range filter: default = last 14 days of actual data in file ───────
    today = _dt.date.today()
    min_d = pace_df["Date"].min().date()
    max_d = pace_df["Date"].max().date()
    # Find last date that actually has non-zero values
    _nonzero = pace_df.groupby("Date")[["OTB","STLY_OTB","Revenue","STLY_Revenue"]].sum()
    _nonzero = _nonzero[_nonzero.sum(axis=1) > 0]
    if not _nonzero.empty:
        last_data_d = _nonzero.index.max().date()
    else:
        last_data_d = max_d
    _today_pace   = _dt.date.today()
    default_end   = min(last_data_d, _today_pace)
    default_start = max(min_d, default_end - _dt.timedelta(days=13))

    _lbl_color = "#3a5260"
    st.markdown(
        f'<style>'
        f'div[data-testid="stDateInput"] label {{ color: {_lbl_color} !important; }}'
        f'div[data-testid="stDateInput"] p {{ color: {_lbl_color} !important; }}'
        f'</style>',
        unsafe_allow_html=True
    )
    fc1, fc2, fc3 = st.columns([2, 2, 6])
    with fc1:
        sel_start = st.date_input("From", value=default_start,
                                  min_value=min_d, max_value=max_d,
                                  key="srp_pace_start")
    with fc2:
        sel_end = st.date_input("To", value=default_end,
                                min_value=min_d, max_value=max_d,
                                key="srp_pace_end")

    # ── Filter to date range ─────────────────────────────────────────────────
    mask = (pace_df["Date"].dt.date >= sel_start) & (pace_df["Date"].dt.date <= sel_end)
    df = pace_df[mask].copy()

    if df.empty:
        st.warning("No data in selected date range.")
        return

    # ── Aggregate by segment (weighted ADR = Revenue / OTB) ─────────────────
    SEGMENTS = ["BAR","CMP","CMTG","CNR","CONS","CONV","DISC","GOV","GT","IT","LNR","MKT","SMRF"]

    seg_agg = []
    for seg in SEGMENTS:
        grp = df[df["Segment"] == seg]
        if grp.empty:
            continue
        tot_otb      = grp["OTB"].sum()
        tot_stly_otb = grp["STLY_OTB"].sum()
        tot_rev      = grp["Revenue"].sum()
        tot_stly_rev = grp["STLY_Revenue"].sum()
        adr      = tot_rev      / tot_otb      if tot_otb      > 0 else 0.0
        stly_adr = tot_stly_rev / tot_stly_otb if tot_stly_otb > 0 else 0.0
        seg_agg.append({
            "Segment":      seg,
            "OTB":          tot_otb,
            "STLY_OTB":     tot_stly_otb,
            "Var_OTB":      tot_otb - tot_stly_otb,
            "Revenue":      tot_rev,
            "STLY_Revenue": tot_stly_rev,
            "Var_Revenue":  tot_rev - tot_stly_rev,
            "ADR":          adr,
            "STLY_ADR":     stly_adr,
            "Var_ADR":      adr - stly_adr,
        })

    # Hotel total row
    tot_otb      = df["OTB"].sum()
    tot_stly_otb = df["STLY_OTB"].sum()
    tot_rev      = df["Revenue"].sum()
    tot_stly_rev = df["STLY_Revenue"].sum()
    tot_adr      = tot_rev      / tot_otb      if tot_otb      > 0 else 0.0
    tot_stly_adr = tot_stly_rev / tot_stly_otb if tot_stly_otb > 0 else 0.0

    total_row = {
        "Segment":      "_TOTAL",
        "OTB":          tot_otb,
        "STLY_OTB":     tot_stly_otb,
        "Var_OTB":      tot_otb - tot_stly_otb,
        "Revenue":      tot_rev,
        "STLY_Revenue": tot_stly_rev,
        "Var_Revenue":  tot_rev - tot_stly_rev,
        "ADR":          tot_adr,
        "STLY_ADR":     tot_stly_adr,
        "Var_ADR":      tot_adr - tot_stly_adr,
    }

    # ── Theme tokens ─────────────────────────────────────────────────────────
    if _lt:
        hdr_bg   = "#314B63"; hdr_col  = "#ffffff"
        hdr_bdr  = "#1a3a5c"; cell_bdr = "#c0d0e0"
        tbl_bg   = "#e8f0f8"; row_alt  = "#dce8f4"
        lbl_col  = "#0d1f2d"; val_col  = "#0d1f2d"
        tot_bg   = "#2E618D"; tot_col  = "#ffffff"
        stly_col = "#007a5c"; var_pos  = "#7c3fbf"
        var_neg  = "#cc2200"; zero_col = "#8899aa"
    else:
        hdr_bg   = "#dce5e8"; hdr_col  = "#4e6878"
        hdr_bdr  = "#c4cfd4"; cell_bdr = "#d0dde2"
        tbl_bg   = "#ffffff"; row_alt  = "#dce5e8"
        lbl_col  = "#4e6878"; val_col  = "#1e2d35"
        tot_bg   = "#2E618D"; tot_col  = "#ffffff"
        stly_col = "#2E618D"; var_pos  = "#556848"
        var_neg  = "#a03020"; zero_col = "#6a8090"

    def _fmt_int(v):
        return "0" if v == 0 else f"{int(round(v)):,}"

    def _fmt_dol(v):
        return "$0.00" if v == 0 else f"${v:,.2f}"   # ADR — $ with 2 decimals

    def _fmt_rev(v):
        return "$0" if v == 0 else f"${v:,.0f}"       # Revenue — $ no decimals

    def _fmt_var_int(v):
        if v == 0: return "0"
        return ("+%s" if v > 0 else "%s") % f"{int(round(v)):,}"

    def _fmt_var_dol(v):
        if v == 0: return "$0.00"
        return ("+$%s" if v > 0 else "-$%s") % f"{abs(v):,.2f}"  # ADR variance

    def _fmt_var_rev(v):
        if v == 0: return "$0"
        return ("+$%s" if v > 0 else "-$%s") % f"{abs(v):,.0f}"  # Revenue variance

    def _var_col(v):
        return zero_col if v == 0 else (var_pos if v > 0 else var_neg)

    # ── Cell builders — ALL data cells centered ───────────────────────────────
    BASE = f"padding:7px 10px;font-family:DM Mono,monospace;font-size:13px;text-align:center;border-bottom:1px solid {cell_bdr};white-space:nowrap;"

    def td_lbl(text, bg, bold=False, color=None):
        fw = "font-weight:700;" if bold else ""
        col = color or lbl_col
        return f'<td style="padding:7px 14px;text-align:center;font-size:13px;{fw}color:{col};background:{bg};border-bottom:1px solid {cell_bdr};white-space:nowrap;">{text}</td>'

    def td_val(text, bg, color=None):
        return f'<td style="{BASE}color:{color or val_col};background:{bg};">{text}</td>'

    def td_stly(text, bg):
        return f'<td style="{BASE}color:{stly_col};background:{bg};">{text}</td>'

    def td_var(text, v, bg):
        return f'<td style="{BASE}color:{_var_col(v)};background:{bg};">{text}</td>'

    def build_row(row, bg, is_total=False):
        lbl_text = hotel.get("display_name", "Hotel Total") if is_total else row["Segment"]
        lbl = td_lbl(lbl_text, bg, bold=is_total, color=tot_col if is_total else None)
        if is_total:
            # All cells white — hardcode color directly, no filter tricks
            def _td(text):
                return (f'<td data-tot="1" style="padding:7px 10px;font-family:DM Mono,monospace;font-size:13px;' +
                        f'text-align:center;border-bottom:1px solid {cell_bdr};' +
                        f'white-space:nowrap;background:{bg};color:#ffffff !important;">' +
                        f'<b style="color:#ffffff !important;font-weight:700;">{text}</b></td>')
            return (
                f"<tr>{lbl}"
                + _td(_fmt_int(row["OTB"]))
                + _td(_fmt_int(row["STLY_OTB"]))
                + _td(_fmt_var_int(row["Var_OTB"]))
                + _td(_fmt_dol(row["ADR"]))
                + _td(_fmt_dol(row["STLY_ADR"]))
                + _td(_fmt_var_dol(row["Var_ADR"]))
                + _td(_fmt_rev(row["Revenue"]))
                + _td(_fmt_rev(row["STLY_Revenue"]))
                + _td(_fmt_var_rev(row["Var_Revenue"]))
                + "</tr>"
            )
        return (
            f"<tr>{lbl}"
            + td_val( _fmt_int(row["OTB"]),              bg)
            + td_stly(_fmt_int(row["STLY_OTB"]),         bg)
            + td_var( _fmt_var_int(row["Var_OTB"]),      row["Var_OTB"],      bg)
            + td_val( _fmt_dol(row["ADR"]),               bg)
            + td_stly(_fmt_dol(row["STLY_ADR"]),         bg)
            + td_var( _fmt_var_dol(row["Var_ADR"]),      row["Var_ADR"],      bg)
            + td_val( _fmt_rev(row["Revenue"]),           bg)
            + td_stly(_fmt_rev(row["STLY_Revenue"]),     bg)
            + td_var( _fmt_var_rev(row["Var_Revenue"]),  row["Var_Revenue"],  bg)
            + "</tr>"
        )

    # ── Header (2-row: group spans + sub-labels) ──────────────────────────────
    TH  = f"padding:8px 10px;text-align:center;font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{hdr_col};background:{hdr_bg};border-bottom:2px solid {hdr_bdr};white-space:nowrap;"
    THL = f"padding:8px 14px;text-align:center;font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:{hdr_col};background:{hdr_bg};border-bottom:2px solid {hdr_bdr};white-space:nowrap;"

    header = (
        f'<tr>'
        f'<th style="{THL}" rowspan="2">Segment</th>'
        f'<th style="{TH}" colspan="3">Occupancy On Books</th>'
        f'<th style="{TH}" colspan="3">ADR On Books</th>'
        f'<th style="{TH}" colspan="3">Revenue On Books</th>'
        f'</tr><tr>'
        + "".join(f'<th style="{TH}">{lbl}</th>' for lbl in
                  ["Current","STLY","Variance","Current","STLY","Variance","Current","STLY","Variance"])
        + "</tr>"
    )

    # ── Build rows ────────────────────────────────────────────────────────────
    total_html = build_row(total_row, tot_bg, is_total=True)
    rows_html  = "".join(
        build_row(row, row_alt if i % 2 == 0 else tbl_bg)
        for i, row in enumerate(seg_agg)
    )

    date_label = f"{sel_start.strftime('%b %d, %Y')} &nbsp;—&nbsp; {sel_end.strftime('%b %d, %Y')}"
    days_count = (sel_end - sel_start).days + 1
    meta_col   = "#5a7a9a"

    html = f"""
<div style="margin-top:16px;">
  <div style="font-size:12px;color:{meta_col};margin-bottom:12px;font-family:DM Sans,sans-serif;">
    {date_label} &nbsp;&middot;&nbsp; {days_count} days
  </div>
  <div style="overflow-x:auto;">
    <table style="width:100%;border-collapse:collapse;font-size:13px;background:{tbl_bg};border:1px solid {hdr_bdr};border-radius:8px;overflow:hidden;">
      <thead>{header}</thead>
      <tbody>{total_html}{rows_html}</tbody>
    </table>
  </div>
</div>
"""
    st.markdown(html, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# IHG MONTHLY SEGMENT TAB
# Renders Property_Segment_Data monthly rollup for IHG hotels.
# Called from render_srp_pace_tab() when data["_ihg_hotel"] is True.
# ─────────────────────────────────────────────────────────────────────────────

def _render_ihg_segment_tab(pace_df, hotel, data):
    """
    IHG SRP Pace tab — Corp Segments view.
    Two sortable tables from Corp_Segments.xlsx:
      1. By Segment  — IHG segment names, sorted by CY Rooms desc by default
      2. By Company  — one row per company (rooms summed), sorted by CY Rooms desc by default
    Features: centered name column, equal-width CY/LY cols, narrow diff cols, sort carets.
    """
    cs = data.get("corp_segments")

    if cs is None:
        err = data.get("corp_segments_error", "")
        if "errno 13" in err.lower() or "permission denied" in err.lower():
            st.warning("⚠️ Corp_Segments.xlsx is open in Excel — please close it, then click ⟳ Refresh Data.")
        else:
            st.warning("Corp Segments file not loaded. Place Corp_Segments.xlsx in the hotel folder.")
        return

    seg_df = cs.get("segments", pd.DataFrame())
    co_df  = cs.get("companies", pd.DataFrame())

    if seg_df.empty and co_df.empty:
        st.info("No production data found in Corp Segments file.")
        return

    _lt = False
    hotel_id = hotel.get("id", "ihg")

    CARD = "#ffffff"
    BDR  = "#c4cfd4"
    HDR  = "#dce5e8"
    TXT  = "#1e2d35"
    DIM  = "#6a8090"
    CY_C = "#2E618D"
    LY_C = "#6A924D"
    POS  = "#2E618D"
    NEG  = "#a03020"
    HDR_TXT = "#4e6878"

    # ── Sort state ────────────────────────────────────────────────────────────
    # Keys: ihg_seg_sort_col, ihg_seg_sort_asc, ihg_co_sort_col, ihg_co_sort_asc
    for sk, sv in [
        (f"ihg_seg_sort_col_{hotel_id}", "CY_Rooms"),
        (f"ihg_seg_sort_asc_{hotel_id}", False),
        (f"ihg_co_sort_col_{hotel_id}",  "CY_Rooms"),
        (f"ihg_co_sort_asc_{hotel_id}",  False),
    ]:
        if sk not in st.session_state:
            st.session_state[sk] = sv

    SORT_COLS = {
        "CY Rooms Qty": "CY_Rooms", "LY Rooms Qty": "LY_Rooms",
        "CY Revenue": "CY_Rev", "LY Revenue": "LY_Rev",
        "CY ADR": "CY_ADR", "LY ADR": "LY_ADR",
        "Diff Rooms": "Var_Rooms", "Diff Rev": "Var_Rev", "Diff ADR": "Var_ADR",
    }

    def _fmt_r(v):
        try: return f"{int(round(float(v))):,}" if pd.notna(v) else "0"
        except: return "0"
    def _fmt_v(v):
        try: return f"${float(v):,.0f}" if pd.notna(v) else "$0"
        except: return "$0"
    def _fmt_a(v):
        try: return f"${float(v):,.0f}" if pd.notna(v) else "$0"
        except: return "$0"

    def _delta_r(d):
        try:
            d = int(round(float(d)))
            c = POS if d>0 else (NEG if d<0 else DIM)
            return f'<span style="color:{c};font-weight:600;">{"+" if d>0 else ""}{d:,}</span>'
        except: return "0"
    def _delta_v(d):
        try:
            d = float(d)
            c = POS if d>0 else (NEG if d<0 else DIM)
            s = f"+${abs(d):,.0f}" if d>=0 else f"-${abs(d):,.0f}"
            return f'<span style="color:{c};font-weight:600;">{s}</span>'
        except: return "$0"
    def _delta_a(d):
        try:
            d = float(d)
            c = POS if d>0 else (NEG if d<0 else DIM)
            s = f"+${abs(d):.0f}" if d>=0 else f"-${abs(d):.0f}"
            return f'<span style="color:{c};">{s}</span>'
        except: return "$0"

    # ── Shared CSS injected once ──────────────────────────────────────────────
    css_id = f"ihg_srp_pace_{hotel_id}"
    st.markdown(f"""
    <style>
    .ihg-tbl-{css_id} {{ width:100%; border-collapse:collapse; table-layout:fixed; }}
    .ihg-tbl-{css_id} th {{
        padding:6px 6px; background:{HDR}; border-bottom:2px solid {BDR};
        font-family:DM Mono,monospace; font-size:10px; letter-spacing:0.06em;
        text-transform:uppercase; white-space:nowrap; cursor:pointer;
        position:relative; user-select:none; text-align:center;
    }}
    .ihg-tbl-{css_id} th .sort-caret {{
        opacity:0; margin-left:4px; font-size:9px; transition:opacity 0.15s;
    }}
    .ihg-tbl-{css_id} th:hover .sort-caret {{ opacity:1; }}
    .ihg-tbl-{css_id} th.sorted .sort-caret {{ opacity:1; }}
    .ihg-tbl-{css_id} td {{
        padding:7px 6px; font-family:DM Mono,monospace; font-size:11px;
        border-bottom:1px solid {BDR}; text-align:center;
    }}
    /* Column widths: name=17%, CY/LY cols=9% each, diff cols=7% each */
    .ihg-tbl-{css_id} .col-name  {{ width:17%; text-align:center; font-weight:600; color:{TXT}; }}
    .ihg-tbl-{css_id} .col-cy    {{ width:9%;  color:{CY_C}; font-weight:600; }}
    .ihg-tbl-{css_id} .col-ly    {{ width:9%;  color:{LY_C}; }}
    .ihg-tbl-{css_id} .col-diff  {{ width:7%; }}
    .ihg-tbl-{css_id} tr:hover td {{ background:#e8f0f5 !important; }}
    .ihg-section-title {{
        font-family:DM Sans,sans-serif; font-size:14px; font-weight:700;
        color:{TXT}; margin:16px 0 4px;
    }}
    .ihg-section-sub {{
        font-family:DM Sans,sans-serif; font-size:11px; color:{DIM};
        margin-bottom:8px;
    }}
    </style>
    """, unsafe_allow_html=True)

    def _render_table(df, name_col, title, subtitle,
                      sort_col_key, sort_asc_key, table_id, show_zero=False):

        sort_col = st.session_state[sort_col_key]
        sort_asc = st.session_state[sort_asc_key]

        # Ensure variance columns exist
        if "Var_ADR" not in df.columns:
            df = df.copy()
            df["Var_ADR"] = (df.get("CY_ADR", pd.Series(0.0, index=df.index)).fillna(0) -
                              df.get("LY_ADR", pd.Series(0.0, index=df.index)).fillna(0)).round(2)
        if "Var_Rooms" not in df.columns:
            df = df.copy()
            df["Var_Rooms"] = df["CY_Rooms"] - df["LY_Rooms"]
            df["Var_Rev"]   = df["CY_Rev"]   - df["LY_Rev"]

        show = df if show_zero else df[(df["CY_Rooms"]>0)|(df["LY_Rooms"]>0)]
        # Initial sort
        if sort_col in show.columns:
            show = show.sort_values(sort_col, ascending=sort_asc).reset_index(drop=True)

        st.markdown(f'<div class="ihg-section-title">{title}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="ihg-section-sub">{subtitle}</div>', unsafe_allow_html=True)

        # Build rows JSON for JS sorting (same pattern as Group Detail table)
        import json as _json
        display_cols = [name_col, "CY_Rooms", "LY_Rooms", "Var_Rooms",
                        "CY_Rev", "LY_Rev", "Var_Rev",
                        "CY_ADR", "LY_ADR", "Var_ADR"]
        col_labels = {
            name_col:    "",
            "CY_Rooms":  "CY Rooms Qty",
            "LY_Rooms":  "LY Rooms Qty",
            "Var_Rooms": "Difference",
            "CY_Rev":    "CY Revenue",
            "LY_Rev":    "LY Revenue",
            "Var_Rev":   "Difference",
            "CY_ADR":    "CY ADR",
            "LY_ADR":    "LY ADR",
            "Var_ADR":   "Difference",
        }
        # Column index → sort session key mapping (used by JS callback)
        sort_keys_js = _json.dumps({
            str(i): c for i, c in enumerate(display_cols)
        })

        rows_json = []
        for _, r in show.iterrows():
            cy_r = int(r.get("CY_Rooms", 0) or 0)
            ly_r = int(r.get("LY_Rooms", 0) or 0)
            cy_v = float(r.get("CY_Rev",   0) or 0)
            ly_v = float(r.get("LY_Rev",   0) or 0)
            cy_a = float(r.get("CY_ADR",   0) or 0) if pd.notna(r.get("CY_ADR")) else 0
            ly_a = float(r.get("LY_ADR",   0) or 0) if pd.notna(r.get("LY_ADR")) else 0
            d_r  = cy_r - ly_r
            d_v  = cy_v - ly_v
            d_a  = cy_a - ly_a
            rows_json.append({
                name_col:    str(r[name_col]) if name_col in r.index else str(r.iloc[0]),
                "CY_Rooms":  cy_r, "LY_Rooms": ly_r, "Var_Rooms": d_r,
                "CY_Rev":    cy_v, "LY_Rev":   ly_v, "Var_Rev":   d_v,
                "CY_ADR":    cy_a, "LY_ADR":   ly_a, "Var_ADR":   d_a,
            })

        # Total row
        tot_cy_r = int(show["CY_Rooms"].sum())
        tot_ly_r = int(show["LY_Rooms"].sum())
        tot_cy_v = float(show["CY_Rev"].sum())
        tot_ly_v = float(show["LY_Rev"].sum())
        tot_cy_a = tot_cy_v / tot_cy_r if tot_cy_r > 0 else 0
        tot_ly_a = tot_ly_v / tot_ly_r if tot_ly_r > 0 else 0

        _tbl_lt   = _lt
        _tbl_card = "#ffffff" if not _tbl_lt else "#ffffff"
        _tbl_hdr  = "#dce5e8" if not _tbl_lt else "#1a3a5c"
        _tbl_bdr  = "#c4cfd4" if not _tbl_lt else "#b0c0d0"
        _tbl_txt  = "#1e2d35"
        _tbl_alt  = "#e8eef1" if not _tbl_lt else "#edf4fc"
        _tbl_tot  = "#dce5e8" if not _tbl_lt else "#d0dff0"
        _cy_c     = "#2E618D" if not _tbl_lt else "#006a55"
        _ly_c     = "#6A924D" if not _tbl_lt else "#7a4e00"
        _pos_c    = "#2E618D" if not _tbl_lt else "#006a55"
        _neg_c    = "#a03020" if not _tbl_lt else "#cc0000"
        _dim_c    = "#6a8090" if not _tbl_lt else "#8aaabf"
        _hdr_txt  = "#4e6878" if not _tbl_lt else "#ffffff"

        tbl_id = f"ihg_srp_{table_id}"

        table_html = f"""
<style>
#{tbl_id} {{ width:100%; border-collapse:collapse; table-layout:auto; background:{_tbl_card}; border-radius:8px; overflow:hidden; }}
#{tbl_id} th {{
    cursor:pointer; padding:8px 10px;
    background:{_tbl_hdr}; border-bottom:2px solid {_tbl_bdr};
    font-family:DM Mono,monospace; font-size:10px; letter-spacing:0.07em;
    text-transform:uppercase; white-space:nowrap; text-align:center;
    position:relative; user-select:none;
}}
#{tbl_id} th:first-child {{ text-align:center; }}
#{tbl_id} th .sort-arrow {{ opacity:0; margin-left:4px; font-size:9px; }}
#{tbl_id} th:hover .sort-arrow {{ opacity:0.6; }}
#{tbl_id} th.sorted-asc  .sort-arrow::after {{ content:'▲'; opacity:1 !important; }}
#{tbl_id} th.sorted-desc .sort-arrow::after {{ content:'▼'; opacity:1 !important; }}
#{tbl_id} th.sorted-asc .sort-arrow,
#{tbl_id} th.sorted-desc .sort-arrow {{ opacity:1; }}
#{tbl_id} td {{
    padding:7px 10px; font-family:DM Mono,monospace; font-size:11px;
    border-bottom:1px solid {_tbl_bdr}; text-align:center;
}}
#{tbl_id} td:first-child {{ text-align:center; font-weight:600; color:{_tbl_txt}; }}
#{tbl_id} tr:hover td {{ background:{_tbl_alt} !important; }}
</style>
<div style="overflow-x:auto; border:1px solid {_tbl_bdr}; border-radius:8px; margin-bottom:8px;">
<table id="{tbl_id}">
  <thead><tr id="{tbl_id}_head">
    <th onclick="sortIhgTbl('{tbl_id}',0)" style="color:{_hdr_txt};">&nbsp;<span class="sort-arrow"></span></th>
    <th onclick="sortIhgTbl('{tbl_id}',1)" style="color:{_cy_c};">CY Rooms Qty<span class="sort-arrow"></span></th>
    <th onclick="sortIhgTbl('{tbl_id}',2)" style="color:{_ly_c};">LY Rooms Qty<span class="sort-arrow"></span></th>
    <th onclick="sortIhgTbl('{tbl_id}',3)" style="color:{_hdr_txt};">Difference<span class="sort-arrow"></span></th>
    <th onclick="sortIhgTbl('{tbl_id}',4)" style="color:{_cy_c};">CY Revenue<span class="sort-arrow"></span></th>
    <th onclick="sortIhgTbl('{tbl_id}',5)" style="color:{_ly_c};">LY Revenue<span class="sort-arrow"></span></th>
    <th onclick="sortIhgTbl('{tbl_id}',6)" style="color:{_hdr_txt};">Difference<span class="sort-arrow"></span></th>
    <th onclick="sortIhgTbl('{tbl_id}',7)" style="color:{_cy_c};">CY ADR<span class="sort-arrow"></span></th>
    <th onclick="sortIhgTbl('{tbl_id}',8)" style="color:{_ly_c};">LY ADR<span class="sort-arrow"></span></th>
    <th onclick="sortIhgTbl('{tbl_id}',9)" style="color:{_hdr_txt};">Difference<span class="sort-arrow"></span></th>
  </tr></thead>
  <tbody id="{tbl_id}_body">
"""
        def _dr(v):
            try:
                v = int(round(float(v)))
                c = _pos_c if v > 0 else (_neg_c if v < 0 else _dim_c)
                s = f"+{v:,}" if v > 0 else f"{v:,}"
                return f'<span style="color:{c};font-weight:600;">{s}</span>'
            except: return "0"
        def _dv(v):
            try:
                v = float(v)
                c = _pos_c if v > 0 else (_neg_c if v < 0 else _dim_c)
                s = f"+${abs(v):,.0f}" if v >= 0 else f"-${abs(v):,.0f}"
                return f'<span style="color:{c};font-weight:600;">{s}</span>'
            except: return "$0"
        def _da(v):
            try:
                v = float(v)
                c = _pos_c if v > 0 else (_neg_c if v < 0 else _dim_c)
                s = f"+${abs(v):.0f}" if v >= 0 else f"-${abs(v):.0f}"
                return f'<span style="color:{c};">{s}</span>'
            except: return "$0"

        for ri, r in enumerate(rows_json):
            bg = f"background:{_tbl_alt};" if ri % 2 == 0 else ""
            table_html += (
                f'<tr style="{bg}">'
                f'<td>{r[name_col]}</td>'
                f'<td style="color:{_cy_c};font-weight:600;">{r["CY_Rooms"]:,}</td>'
                f'<td style="color:{_ly_c};">{r["LY_Rooms"]:,}</td>'
                f'<td>{_dr(r["Var_Rooms"])}</td>'
                f'<td style="color:{_cy_c};font-weight:600;">${r["CY_Rev"]:,.0f}</td>'
                f'<td style="color:{_ly_c};">${r["LY_Rev"]:,.0f}</td>'
                f'<td>{_dv(r["Var_Rev"])}</td>'
                f'<td style="color:{_cy_c};font-weight:600;">${r["CY_ADR"]:,.0f}</td>'
                f'<td style="color:{_ly_c};">${r["LY_ADR"]:,.0f}</td>'
                f'<td>{_da(r["Var_ADR"])}</td>'
                f'</tr>'
            )

        # Total row
        table_html += (
            f'<tr style="background:{_tbl_tot};">'
            f'<td style="font-weight:700;">TOTAL</td>'
            f'<td style="color:{_cy_c};font-weight:700;">{tot_cy_r:,}</td>'
            f'<td style="color:{_ly_c};font-weight:700;">{tot_ly_r:,}</td>'
            f'<td>{_dr(tot_cy_r - tot_ly_r)}</td>'
            f'<td style="color:{_cy_c};font-weight:700;">${tot_cy_v:,.0f}</td>'
            f'<td style="color:{_ly_c};font-weight:700;">${tot_ly_v:,.0f}</td>'
            f'<td>{_dv(tot_cy_v - tot_ly_v)}</td>'
            f'<td style="color:{_cy_c};font-weight:700;">${tot_cy_a:,.0f}</td>'
            f'<td style="color:{_ly_c};font-weight:700;">${tot_ly_a:,.0f}</td>'
            f'<td>{_da(tot_cy_a - tot_ly_a)}</td>'
            f'</tr>'
        )

        table_html += f"""
  </tbody>
</table>
</div>
<script>
(function() {{
  var sortState = {{}};
  window.sortIhgTbl = function(tblId, colIdx) {{
    var tbody = document.getElementById(tblId + '_body');
    var head  = document.getElementById(tblId + '_head');
    if (!tbody || !head) return;
    var rows = Array.from(tbody.querySelectorAll('tr'));
    var asc  = sortState[colIdx] !== true;
    sortState = {{}};
    sortState[colIdx] = asc;
    Array.from(head.querySelectorAll('th')).forEach(function(th, i) {{
      th.classList.remove('sorted-asc','sorted-desc');
      if (i === colIdx) th.classList.add(asc ? 'sorted-asc' : 'sorted-desc');
    }});
    rows.sort(function(a, b) {{
      var av = a.querySelectorAll('td')[colIdx].textContent.trim();
      var bv = b.querySelectorAll('td')[colIdx].textContent.trim();
      if (av === 'TOTAL' || bv === 'TOTAL') return 0;
      var an = parseFloat(av.replace(/[$,+%]/g,''));
      var bn = parseFloat(bv.replace(/[$,+%]/g,''));
      if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
      return asc ? av.localeCompare(bv) : bv.localeCompare(av);
    }});
    rows.forEach(function(r) {{ tbody.appendChild(r); }});
  }};
  // Set initial sort indicator
  var head = document.getElementById('{tbl_id}_head');
  if (head) {{
    var ths = head.querySelectorAll('th');
    if (ths[1]) ths[1].classList.add('sorted-desc');
  }}
}})();
</script>"""

        tbl_height = 60 + (len(rows_json) + 1) * 38
        components.html(table_html, height=tbl_height, scrolling=False)
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    _render_table(seg_df, "Segment", "By Segment", "IHG segment rollup · CY vs LY",
                  f"ihg_seg_sort_col_{hotel_id}", f"ihg_seg_sort_asc_{hotel_id}",
                  f"seg_{hotel_id}", show_zero=True)

    _render_table(co_df, "CompanyName", "By Company",
                  "One row per company · rooms summed · sorted by CY rooms",
                  f"ihg_co_sort_col_{hotel_id}", f"ihg_co_sort_asc_{hotel_id}",
                  f"co_{hotel_id}")


def render_demand_tab(data, hotel):
    import json as _json
    from pathlib import Path as _Path

    year_df   = data.get("year")
    budget_df = data.get("budget")

    st.markdown('<div class="section-head">Demand Nights</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Need Nights vs High Demand Nights based on forecasted occupancy</div>',
                unsafe_allow_html=True)

    if year_df is None:
        st.warning("Year data file not loaded.")
        return

    _lt = False
    today = pd.Timestamp(datetime.now().date())
    total_rooms = hotel.get("total_rooms", 112)

    # ── Persistent thresholds (saved per hotel) ───────────────────────────────
    _thresh_path = _Path(str(cfg.get_hotel_folder(hotel) / "demand_thresholds.json"))
    _defaults    = {"need_pct": 30, "high_pct": 90}
    _saved       = _defaults.copy()
    if _thresh_path.exists():
        try:
            _saved.update(_json.loads(_thresh_path.read_text(encoding="utf-8")))
        except Exception:
            pass

    tc1, tc2, tc3 = st.columns([1, 1, 6])
    with tc1:
        need_pct = st.number_input(
            "Need Nights — Max Occ %",
            min_value=1, max_value=99,
            value=int(_saved.get("need_pct", 30)),
            step=1, key="demand_need_pct",
            help="Dates forecast below this occupancy % are flagged as Need Nights"
        )
    with tc2:
        high_pct = st.number_input(
            "High Demand — Min Occ %",
            min_value=1, max_value=100,
            value=int(_saved.get("high_pct", 90)),
            step=1, key="demand_high_pct",
            help="Dates forecast at or above this occupancy % are flagged as High Demand"
        )

    # Save thresholds if changed
    if need_pct != _saved.get("need_pct") or high_pct != _saved.get("high_pct"):
        try:
            _thresh_path.write_text(
                _json.dumps({"need_pct": need_pct, "high_pct": high_pct}),
                encoding="utf-8"
            )
        except Exception:
            pass

    # ── Build forward-looking data ────────────────────────────────────────────
    fwd = year_df[year_df["Date"] >= today].copy()
    if fwd.empty:
        st.info("No forward-looking data available.")
        return

    # Compute forecast occupancy %
    cap_col = "Capacity" if "Capacity" in fwd.columns else None
    if "Forecast_Rooms" in fwd.columns and cap_col:
        fwd["Fcst_Occ_Pct"] = (fwd["Forecast_Rooms"] / fwd[cap_col].replace(0, total_rooms) * 100).round(1)
    elif "Forecast_Rooms" in fwd.columns:
        fwd["Fcst_Occ_Pct"] = (fwd["Forecast_Rooms"] / total_rooms * 100).round(1)
    elif "OTB" in fwd.columns:
        fwd["Fcst_Occ_Pct"] = (fwd["OTB"] / total_rooms * 100).round(1)
    else:
        st.warning("No forecast or OTB rooms column found.")
        return

    # Merge budget for suggested group rate context
    if budget_df is not None and "Occ_Pct" in budget_df.columns:
        fwd = fwd.merge(budget_df[["Date","Occ_Rooms","Occ_Pct","ADR","RevPAR"]],
                        on="Date", how="left", suffixes=("","_Budget"))

    # ── Segment into need / high demand ──────────────────────────────────────
    need_df = fwd[fwd["Fcst_Occ_Pct"] <= need_pct].copy()
    high_df = fwd[fwd["Fcst_Occ_Pct"] >= high_pct].copy()

    # Suggested group rate: budget ADR only, rounded to nearest dollar
    # Need Nights = 90% of budget ADR | High Demand = 120% of budget ADR
    def _suggested_rate(row, multiplier):
        # Budget ADR only (from budget_df merge — column "ADR" with _Budget suffix or plain "ADR")
        adr = None
        for col in ("ADR_Budget", "ADR"):
            v = row.get(col) if col in row.index else None
            if v is not None and pd.notna(v) and float(v) > 0:
                adr = float(v)
                break
        if adr:
            return f"${round(adr * multiplier):,}"
        return "—"

    # ── Theme tokens ─────────────────────────────────────────────────────────
    if _lt:
        hdr_need  = "#1a4a2a"; hdr_need_t  = "#ffffff"
        hdr_high  = "#4a1a1a"; hdr_high_t  = "#ffffff"
        tbl_bg    = "#e8f0f8"; row_alt      = "#dce8f4"
        cell_bdr  = "#c0d0e0"; hdr_bdr      = "#1a3a5c"
        lbl_col   = "#0d1f2d"; val_col      = "#0d1f2d"
        evt_col   = "#2E618D"; occ_col      = "#1e2d35"
    else:
        hdr_need  = "#d8f0e4"; hdr_need_t  = "#2E618D"
        hdr_high  = "#f5dcd8"; hdr_high_t  = "#b44820"
        tbl_bg    = "#ffffff"; row_alt      = "#dce5e8"
        cell_bdr  = "#d0dde2"; hdr_bdr      = "#c4cfd4"
        lbl_col   = "#4e6878"; val_col      = "#1e2d35"
        evt_col   = "#556848"; occ_col      = "#1e2d35"

    def _build_table(df, title, title_color, header_bg, header_txt, suggested_mult, empty_msg):
        TH = (f"padding:8px 12px;text-align:center;font-size:11px;font-weight:700;"
              f"letter-spacing:.05em;text-transform:uppercase;color:{header_txt};"
              f"background:{header_bg};border-bottom:2px solid {hdr_bdr};white-space:nowrap;")
        THL = TH.replace("text-align:center", "text-align:left")

        header = (
            f'<tr><th style="{TH}">Date</th><th style="{TH}">Day</th>'
            f'<th style="{TH}">Event</th><th style="{TH}">Fcst Occ %</th>'
            f'<th style="{TH}">Fcst Rooms</th><th style="{TH}">Suggested<br>Group Rate</th></tr>'
        )

        if df.empty:
            body = f'<tr><td colspan="6" style="padding:20px;text-align:center;color:{lbl_col};font-style:italic;">{empty_msg}</td></tr>'
        else:
            rows = ""
            for i, (_, row) in enumerate(df.iterrows()):
                bg      = row_alt if i % 2 == 0 else tbl_bg
                date_s  = pd.Timestamp(row["Date"]).strftime("%m/%d/%Y")
                _dow_raw = row.get("DayOfWeek", None)
                dow = (str(_dow_raw)[:3] if (_dow_raw and pd.notna(_dow_raw) and str(_dow_raw).strip() not in ("", "—", "nan"))
                       else pd.Timestamp(row["Date"]).strftime("%a"))
                event   = str(row.get("Event","")) if pd.notna(row.get("Event","")) else ""
                occ_pct = f'{row["Fcst_Occ_Pct"]:.0f}%' if pd.notna(row.get("Fcst_Occ_Pct")) else "—"
                fcst_r  = f'{int(row["Forecast_Rooms"])}' if "Forecast_Rooms" in row.index and pd.notna(row.get("Forecast_Rooms")) else (
                           f'{int(row["OTB"])}' if "OTB" in row.index and pd.notna(row.get("OTB")) else "—")
                sug_rate = _suggested_rate(row, suggested_mult)

                BASE = f"padding:7px 12px;border-bottom:1px solid {cell_bdr};background:{bg};font-size:13px;"
                def td_c(v, color=val_col):
                    return f'<td style="{BASE}text-align:center;font-family:DM Mono,monospace;color:{color};">{v}</td>'
                def td_l(v, color=val_col):
                    return f'<td style="{BASE}text-align:left;color:{color};">{v}</td>'

                rows += (
                    f"<tr>"
                    + td_c(date_s)
                    + td_c(dow)
                    + td_c(event, color=evt_col if event else val_col)
                    + td_c(occ_pct, color=occ_col)
                    + td_c(fcst_r)
                    + td_c(sug_rate, color=("#7a4e00" if _lt else GOLD))
                    + "</tr>"
                )
            body = rows

        return f"""
<div style="margin-bottom:32px;">
  <div style="font-size:15px;font-weight:700;color:{title_color};margin-bottom:10px;
              font-family:DM Sans,sans-serif;letter-spacing:.02em;">{title}</div>
  <div style="overflow-x:auto;">
    <table style="width:100%;border-collapse:collapse;background:{tbl_bg};
                  border:1px solid {hdr_bdr};border-radius:8px;overflow:hidden;">
      <thead>{header}</thead>
      <tbody>{body}</tbody>
    </table>
  </div>
</div>"""

    # ── Render side by side ───────────────────────────────────────────────────
    st.markdown("<div style='margin-top:20px;'></div>", unsafe_allow_html=True)
    col_left, col_right = st.columns(2)

    need_title = f"Need Nights  ({need_pct}%- Occ Forecast)"
    high_title = f"High Demand Nights  ({high_pct}%+ Occ Forecast)"
    need_color = "#2E618D"
    high_color = "#b44820"

    with col_left:
        st.markdown(
            _build_table(need_df, need_title, need_color,
                         hdr_need, hdr_need_t, 0.90,
                         f"No dates forecast below {need_pct}% occupancy"),
            unsafe_allow_html=True
        )
        st.caption(f"{len(need_df)} need night{'s' if len(need_df)!=1 else ''} found")

    with col_right:
        st.markdown(
            _build_table(high_df, high_title, high_color,
                         hdr_high, hdr_high_t, 1.20,
                         f"No dates forecast at or above {high_pct}% occupancy"),
            unsafe_allow_html=True
        )
        st.caption(f"{len(high_df)} high demand night{'s' if len(high_df)!=1 else ''} found")


def render_events_tab(data, hotel):
    year_df          = data.get("year")
    rates_ov         = data.get("rates_overview")
    lighthouse_events = data.get("lighthouse_events")
    is_ihg           = bool(data.get("_ihg_hotel")) or hotel.get("brand", "") == "IHG"

    st.markdown('<div class="section-head">Events & Demand Calendar</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Special events flagged by date</div>', unsafe_allow_html=True)

    today = pd.Timestamp(datetime.now().date())

    # ── Theme tokens (shared) ─────────────────────────────────────────────────
    _ev_lt = False
    _ev_hdr_bg   = "#dce5e8"
    _ev_hdr_col  = "#4e6878"
    _ev_bdr      = "#c4cfd4"
    _ev_txt      = "#1e2d35"
    _ev_tbl_bg   = "transparent"
    _ev_date_col = "#1e2d35"
    _ev_otb_col  = "#1e2d35"
    _ev_fcst_col = "#1e2d35"
    _ev_ly_col   = "#4e6878"
    _ev_event_col = TEAL
    _ev_cat_col  = "#5ad0a0"
    _row_bg      = "transparent"

    _th = lambda label: (
        f'<th style="padding:8px 12px;text-align:center;color:{_ev_hdr_col};font-weight:700;'
        f'font-size:11px;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap;'
        f'border-bottom:2px solid {_ev_bdr};background:{_ev_hdr_bg};'
        f'position:sticky;top:0;z-index:2;">{label}</th>'
    )

    def td(v, color=None, mono=False, align="center"):
        c  = color or _ev_txt
        ff = "font-family:DM Mono,monospace;" if mono else ""
        return (f'<td style="padding:7px 12px;border-bottom:1px solid {_ev_bdr};'
                f'color:{c};font-size:13px;text-align:{align};background:{_row_bg};{ff}">{v}</td>')

    # Build date → Our_Rate lookup from rates
    rate_lookup = {}
    if rates_ov is not None and "Date" in rates_ov.columns:
        for _, rr in rates_ov.iterrows():
            if pd.notna(rr["Date"]):
                rate_lookup[pd.Timestamp(rr["Date"]).normalize()] = rr.get("Our_Rate")

    # ── IHG path: Lighthouse Events ──────────────────────────────────────────
    if is_ihg:
        if lighthouse_events is None or lighthouse_events.empty:
            err = data.get("lighthouse_events_error")
            if err:
                st.warning(f"Lighthouse Events file found but failed to load: {err}")
            else:
                st.info("No Lighthouse Events file found. Place 'Lighthouse Events.xlsx' in the hotel folder.")
            return

        upcoming = lighthouse_events[lighthouse_events["Date"] >= today].copy()
        if upcoming.empty:
            st.info("No upcoming events found in the Lighthouse Events file.")
            return

        # Build date → OTB / Forecast_Rooms lookup from year_df
        yr_lookup = {}
        if year_df is not None and not year_df.empty:
            for _, yr in year_df.iterrows():
                d = pd.Timestamp(yr["Date"]).normalize()
                yr_lookup[d] = {
                    "OTB":            yr.get("OTB"),
                    "Forecast_Rooms": yr.get("Forecast_Rooms"),
                }

        # Deduplicate: for multi-day events show only first day row with duration badge
        # Group by Start_Date + Event name so we get one row per event
        seen = set()
        deduped = []
        for _, row in upcoming.sort_values(["Start_Date", "Date"]).iterrows():
            key = (row["Start_Date"], row["Event"])
            if key not in seen:
                seen.add(key)
                deduped.append(row)
        upcoming_deduped = pd.DataFrame(deduped)

        thead = (
            _th("Start Date") +
            _th("End Date") +
            _th("Event") +
            _th("Category") +
            _th("Location") +
            _th("Current Selling Rate") +
            _th("OTB") +
            _th("Forecast Rooms")
        )

        tbody = ""
        for _, row in upcoming_deduped.iterrows():
            start_ts  = pd.Timestamp(row["Start_Date"]).normalize()
            end_ts    = pd.Timestamp(row["End_Date"]).normalize()
            duration  = int(row.get("Duration_Days", 1))

            start_str = start_ts.strftime("%a, %b %d %Y")
            end_str   = end_ts.strftime("%a, %b %d %Y") if duration > 1 else "—"
            event     = str(row.get("Event", "—"))
            category  = str(row.get("Category", ""))
            location  = str(row.get("Location", "")) or "—"

            # OTB / Forecast from year_df for start date
            yr_day   = yr_lookup.get(start_ts, {})
            otb      = yr_day.get("OTB")
            fcst     = yr_day.get("Forecast_Rooms")

            our_rate = rate_lookup.get(start_ts)
            try:
                our_rate = float(our_rate) if our_rate is not None else None
            except (ValueError, TypeError):
                our_rate = None

            rate_disp = f"${our_rate:,.0f}" if our_rate is not None else "—"
            rate_color = (GOLD if _ev_lt else GOLD) if our_rate else "#4e6878"
            otb_disp  = f"{int(otb):,}"  if otb  is not None and pd.notna(otb)  else "—"
            fcst_disp = f"{fcst:.0f}"    if fcst is not None and pd.notna(fcst) else "—"

            # Category badge color
            cat_lower = category.lower()
            if "holiday" in cat_lower:
                cat_col = "#6A924D"
            elif "sport" in cat_lower:
                cat_col = "#00aaff"
            elif "concert" in cat_lower or "festival" in cat_lower or "performing" in cat_lower:
                cat_col = "#cc66ff"
            else:
                cat_col = _ev_cat_col  # Conference / default — teal

            tbody += (
                f"<tr>"
                f"{td(start_str, _ev_date_col)}"
                f"{td(end_str, _ev_ly_col)}"
                f"{td(event, _ev_event_col, align='left')}"
                f"{td(category, cat_col)}"
                f"{td(location, _ev_txt, align='left')}"
                f"{td(rate_disp, rate_color, mono=True)}"
                f"{td(otb_disp, _ev_otb_col, mono=True)}"
                f"{td(fcst_disp, _ev_fcst_col, mono=True)}"
                f"</tr>"
            )

        ev_html = (
            f'<div style="overflow-x:auto;overflow-y:auto;max-height:600px;'
            f'border:1px solid {_ev_bdr};border-radius:6px;background:{_ev_tbl_bg};">'
            f'<table style="width:100%;border-collapse:collapse;font-size:13px;'
            f'font-family:DM Sans,sans-serif;background:{_ev_tbl_bg};">'
            f'<thead><tr>{thead}</tr></thead>'
            f'<tbody>{tbody}</tbody>'
            f'</table></div>'
        )
        st.markdown(ev_html, unsafe_allow_html=True)
        st.markdown(
            f'<div style="margin-top:8px;font-family:DM Sans,sans-serif;font-size:11px;'
            f'color:{_ev_ly_col};">Source: Lighthouse Events · {len(upcoming_deduped)} upcoming events</div>',
            unsafe_allow_html=True)
        return

    # ── Hilton / standard path: year_df Event column ─────────────────────────
    if year_df is None:
        st.warning("Year data file not loaded.")
        return

    if "Event" not in year_df.columns:
        st.info("No event data found in the year file.")
        return

    events   = year_df[year_df["Event"].str.strip().str.len() > 0].copy()
    upcoming = events[events["Date"] >= today].sort_values("Date")

    if upcoming.empty:
        st.info("No upcoming events found in the data.")
        return

    thead = (
        _th("Date") +
        _th("Day") +
        _th("Event") +
        _th("Current Selling Rate") +
        _th("OTB") +
        _th("Forecast Rooms") +
        _th("Rooms Sold LY")
    )

    tbody = ""
    for _, row in upcoming.iterrows():
        date_ts  = pd.Timestamp(row["Date"]).normalize()
        date_str = date_ts.strftime("%a, %b %d %Y")
        dow_raw  = row.get("DayOfWeek", None)
        dow      = (str(dow_raw)[:3] if (dow_raw and pd.notna(dow_raw) and str(dow_raw).strip() not in ("", "—", "nan"))
                    else date_ts.strftime("%a"))
        event    = str(row.get("Event", "—"))
        otb      = row.get("OTB")
        otb_ly   = row.get("OTB_LY")
        fcst     = row.get("Forecast_Rooms")
        our_rate = rate_lookup.get(date_ts)
        try:
            our_rate = float(our_rate) if our_rate is not None else None
        except (ValueError, TypeError):
            our_rate = None

        otb_disp    = f"{int(otb):,}"     if pd.notna(otb)   and otb   is not None else "—"
        otb_ly_disp = f"{int(otb_ly):,}"  if pd.notna(otb_ly) and otb_ly is not None else "—"
        fcst_disp   = f"{fcst:.0f}"       if pd.notna(fcst)  and fcst  is not None else "—"
        rate_disp   = f"${our_rate:,.0f}" if our_rate is not None and pd.notna(our_rate) else "—"
        rate_color  = GOLD if our_rate else "#4e6878"
        _ev_rate_col = ("#7a4e00" if our_rate else "#3a5a7a") if _ev_lt else rate_color

        tbody += (
            f"<tr>"
            f"{td(date_str, _ev_date_col)}"
            f"{td(dow, _ev_txt)}"
            f"{td(event, _ev_event_col)}"
            f"{td(rate_disp, _ev_rate_col, mono=True)}"
            f"{td(otb_disp, _ev_otb_col, mono=True)}"
            f"{td(fcst_disp, _ev_fcst_col, mono=True)}"
            f"{td(otb_ly_disp, _ev_ly_col, mono=True)}"
            f"</tr>"
        )

    ev_html = (
        f'<div style="overflow-x:auto;overflow-y:auto;max-height:600px;'
        f'border:1px solid {_ev_bdr};border-radius:6px;background:{_ev_tbl_bg};">'
        f'<table style="width:100%;border-collapse:collapse;font-size:13px;'
        f'font-family:DM Sans,sans-serif;background:{_ev_tbl_bg};">'
        f'<thead><tr>{thead}</tr></thead>'
        f'<tbody>{tbody}</tbody>'
        f'</table></div>'
    )
    st.markdown(ev_html, unsafe_allow_html=True)



def _compute_seg_note(data) -> str:
    """Compute segment revenue string for Call Recap (used by both PDF and tab).

    Uses the same date-window logic as the SRP Pace tab: find the last date
    with actual non-zero OTB/Revenue in the file, cap at today, then go back
    14 days from there.  This matches exactly what the user sees on screen.

    For IHG hotels, uses corp_segments data (CY_Rev by Segment).
    """
    import pandas as pd
    from datetime import datetime

    # ── IHG: use corp_segments CY_Rev breakdown ──────────────────────────────
    cs = data.get("corp_segments")
    if cs is not None:
        seg_df = cs.get("segments", pd.DataFrame())
        if not seg_df.empty and "Segment" in seg_df.columns and "CY_Rev" in seg_df.columns:
            try:
                seg_rev = seg_df[seg_df["CY_Rev"] > 0].set_index("Segment")["CY_Rev"].sort_values(ascending=False)
                total = seg_rev.sum()
                if total > 0:
                    parts = [f"{s} ({seg_rev[s]/total*100:.0f}%)" for s in seg_rev.index]
                    return "Segments by Revenue (past 2 wks): " + ", ".join(parts)
            except Exception:
                pass

    sp = data.get("srp_pace")
    if sp is not None and not sp.empty and "Date" in sp.columns:
        seg_col = "Segment" if "Segment" in sp.columns else ("MCAT" if "MCAT" in sp.columns else None)
        rev_col = "Revenue" if "Revenue" in sp.columns else ("Room_Revenue" if "Room_Revenue" in sp.columns else None)
        otb_col = "OTB"     if "OTB"     in sp.columns else None
        if seg_col and (rev_col or otb_col):
            try:
                today = pd.Timestamp(datetime.now().date())
                # Find last date with any non-zero data (mirrors SRP Pace tab logic)
                check_cols = [c for c in [rev_col, otb_col, "STLY_OTB", "STLY_Revenue"] if c and c in sp.columns]
                _nonzero = sp.groupby("Date")[check_cols].sum()
                _nonzero = _nonzero[_nonzero.sum(axis=1) > 0]
                if _nonzero.empty:
                    return ""
                last_data_d = _nonzero.index.max()
                window_end   = min(last_data_d, today)
                window_start = window_end - pd.Timedelta(days=13)

                pace_2w = sp[(sp["Date"] >= window_start) & (sp["Date"] <= window_end)]
                if pace_2w.empty:
                    return ""

                use_col = rev_col if rev_col else otb_col
                seg_rev = pace_2w.groupby(seg_col)[use_col].sum().sort_values(ascending=False)
                seg_rev = seg_rev[seg_rev > 0]
                total   = seg_rev.sum()
                if total > 0:
                    parts = [f"{s} ({seg_rev[s]/total*100:.0f}%)" for s in seg_rev.index]
                    label = "Revenue" if use_col == rev_col else "OTB Rooms"
                    return f"Segments by {label} (past 2 wks): " + ", ".join(parts)
            except Exception:
                pass

    # Fallback: booking (SRP Activity)
    bk = data.get("booking")
    if bk is not None and not bk.empty and "MCAT" in bk.columns:
        rev_col  = "Room_Revenue" if "Room_Revenue" in bk.columns else ("Revenue" if "Revenue" in bk.columns else None)
        date_col = "Arrival_Date" if "Arrival_Date" in bk.columns else ("Date" if "Date" in bk.columns else None)
        if rev_col and date_col:
            try:
                today  = pd.Timestamp(datetime.now().date())
                _2wk   = today - pd.Timedelta(days=14)
                bk_2w  = bk[pd.to_datetime(bk[date_col]) >= _2wk]
                seg_rev = bk_2w.groupby("MCAT")[rev_col].sum().sort_values(ascending=False)
                seg_rev = seg_rev[seg_rev > 0]
                total   = seg_rev.sum()
                if total > 0:
                    parts = [f"{s} ({seg_rev[s]/total*100:.0f}%)" for s in seg_rev.index]
                    return "Segments (past 2 wks): " + ", ".join(parts)
            except Exception:
                pass

    return ""


def _compute_auto_rows(data, hotel) -> list:
    """Compute all 4 auto-populated meeting note rows from live data.
    Used by both the PDF and the Call Recap tab so they always match.
    Returns list of 4 strings: [str_report, pace, segments, top_companies]
    """
    import pandas as pd
    from datetime import datetime

    today      = pd.Timestamp(datetime.now().date())
    rows       = ["", "", "", ""]
    hotel_name = hotel.get("name", "Hotel") if isinstance(hotel, dict) else "Hotel"
    short_name = hotel.get("short_name", hotel_name) if isinstance(hotel, dict) else hotel_name

    # ── Row 0: STR Report ────────────────────────────────────────────────────
    try:
        str_data  = data.get("str")
        str_ranks = data.get("str_ranks") or {}
        if str_data:
            weekly = str_data.get("weekly")
            if weekly is not None:
                total_w = weekly[weekly["Day"] == "Total"]
                if not total_w.empty:
                    tw = total_w.iloc[0]
                    def _sv(col):
                        v = tw.get(col); return float(v) if (v is not None and pd.notna(v)) else None
                    occ   = _sv("Occ_Mine");    occ_c   = _sv("Occ_Comp")
                    adr   = _sv("ADR_Mine");    adr_c   = _sv("ADR_Comp")
                    revp  = _sv("RevPAR_Mine")
                    occ_chg  = _sv("Occ_Mine_Chg")
                    adr_chg  = _sv("ADR_Mine_Chg")
                    revp_chg = _sv("RevPAR_Mine_Chg")
                    mpi = _sv("MPI"); ari = _sv("ARI"); rgi = _sv("RGI")
                    tr  = hotel.get("total_rooms", 112) if isinstance(hotel, dict) else 112
                    rms_mine = round(occ  / 100 * tr * 7) if occ  else None
                    rms_comp = round(occ_c / 100 * tr * 7) if occ_c else None
                    rev_mine = round(rms_mine * adr)   if (rms_mine and adr)   else None
                    rev_comp = round(rms_comp * adr_c) if (rms_comp and adr_c) else None
                    rms_diff = (rms_mine - rms_comp) if (rms_mine is not None and rms_comp is not None) else None
                    rev_diff = (rev_mine - rev_comp) if (rev_mine is not None and rev_comp is not None) else None
                    rk_occ  = str_ranks.get("occ",    {}).get("week", "—")
                    rk_adr  = str_ranks.get("adr",    {}).get("week", "—")
                    rk_revp = str_ranks.get("revpar", {}).get("week", "—")
                    parts = []
                    occ_str = f"OCC {occ:.1f}% (Rank #{rk_occ})" if (occ and rk_occ != "—") else (f"OCC {occ:.1f}%" if occ else "")
                    if occ_chg is not None and abs(occ_chg) >= 10:
                        occ_str += f" ▲{occ_chg:.1f}% YOY" if occ_chg > 0 else f" ▼{abs(occ_chg):.1f}% YOY"
                    if occ_str: parts.append(occ_str)
                    adr_str = f"ADR ${adr:.2f} (Rank #{rk_adr})" if (adr and rk_adr != "—") else (f"ADR ${adr:.2f}" if adr else "")
                    if adr_chg is not None and abs(adr_chg) >= 10:
                        adr_str += f" ▲{adr_chg:.1f}% YOY" if adr_chg > 0 else f" ▼{abs(adr_chg):.1f}% YOY"
                    if adr_str: parts.append(adr_str)
                    revp_str = f"RevPAR ${revp:.2f} (Rank #{rk_revp})" if (revp and rk_revp != "—") else (f"RevPAR ${revp:.2f}" if revp else "")
                    if revp_chg is not None and abs(revp_chg) >= 10:
                        revp_str += f" ▲{revp_chg:.1f}% YOY" if revp_chg > 0 else f" ▼{abs(revp_chg):.1f}% YOY"
                    if revp_str: parts.append(revp_str)
                    if mpi: parts.append(f"MPI {mpi:.1f} / ARI {ari:.1f} / RGI {rgi:.1f}" if (ari and rgi) else f"MPI {mpi:.1f}")
                    if rms_diff is not None:
                        parts.append(f"{abs(rms_diff):,} {'more' if rms_diff > 0 else 'fewer'} rooms than avg comp")
                    if rev_diff is not None:
                        parts.append(f"${abs(rev_diff):,} {'more' if rev_diff > 0 else 'less'} revenue than avg comp")
                    rows[0] = " · ".join(p for p in parts if p)
    except Exception:
        pass

    # ── Row 1: Pace — rooms + revenue vs LY ─────────────────────────────────
    try:
        yr = data.get("year")
        if yr is not None:
            mo_start = today.replace(day=1)
            mo_end   = (mo_start + pd.DateOffset(months=1)) - pd.Timedelta(days=1)
            mo_df    = yr[(yr["Date"] >= mo_start) & (yr["Date"] <= mo_end)].copy()
            mo_lbl   = today.strftime("%B")

            if not mo_df.empty:
                is_ihg = hotel.get("brand", "") == "IHG"

                if is_ihg:
                    # IHG: past days use OTB, future days use Forecast_Rooms
                    past  = mo_df[mo_df["Date"] <  today]
                    future= mo_df[mo_df["Date"] >= today]
                    ty_rooms_past   = int(past["OTB"].sum())           if "OTB"            in past.columns  and not past.empty  else 0
                    ty_rooms_future = int(future["Forecast_Rooms"].fillna(0).sum()) if "Forecast_Rooms" in future.columns and not future.empty else 0
                    ty_rooms = ty_rooms_past + ty_rooms_future

                    # Revenue: past=Revenue_OTB, future=Forecast_Rev
                    rev_past   = float(past["Revenue_OTB"].fillna(0).sum())   if "Revenue_OTB" in past.columns   and not past.empty   else 0.0
                    rev_future = float(future["Forecast_Rev"].fillna(0).sum()) if "Forecast_Rev" in future.columns and not future.empty else 0.0
                    ty_rev = rev_past + rev_future if (rev_past + rev_future) > 0 else None
                else:
                    rooms_col = ("Forecast_Rooms" if "Forecast_Rooms" in mo_df.columns
                                 else ("OTB" if "OTB" in mo_df.columns else None))
                    ty_rooms  = int(mo_df[rooms_col].sum()) if rooms_col else 0
                    rev_col   = ("Revenue_OTB" if "Revenue_OTB" in mo_df.columns
                                  else ("Revenue_Forecast" if "Revenue_Forecast" in mo_df.columns else None))
                    ty_rev    = float(mo_df[rev_col].sum()) if (rev_col and mo_df[rev_col].notna().any()) else None

                ly_rooms = int(mo_df["OTB_LY"].sum())      if "OTB_LY"     in mo_df.columns and mo_df["OTB_LY"].notna().any()     else None
                ly_rev   = float(mo_df["Revenue_LY"].sum()) if "Revenue_LY" in mo_df.columns and mo_df["Revenue_LY"].notna().any() else None

                pace_parts = []
                if ly_rooms is not None:
                    diff = ty_rooms - ly_rooms
                    if diff > 0:   pace_parts.append(f"pacing {diff:,} rooms ahead for {mo_lbl} vs STLY")
                    elif diff < 0: pace_parts.append(f"pacing {abs(diff):,} rooms behind for {mo_lbl} vs STLY")
                    else:          pace_parts.append(f"pacing even on rooms for {mo_lbl} vs STLY")
                if ty_rev is not None and ly_rev is not None:
                    diff = ty_rev - ly_rev
                    if diff > 0:   pace_parts.append(f"${abs(diff):,.0f} ahead in revenue")
                    elif diff < 0: pace_parts.append(f"${abs(diff):,.0f} behind in revenue")
                    else:          pace_parts.append("even on revenue")
                if pace_parts:
                    rows[1] = f"{short_name} is " + " and ".join(pace_parts)
    except Exception:
        pass

    # ── Row 2: Revenue by Segment ────────────────────────────────────────────
    rows[2] = _compute_seg_note(data)

    # ── Row 3: Top Companies — Local Negotiated, past 14 days ───────────────
    try:
        # IHG: use corp_segments By Company table (CY_Rooms desc)
        cs = data.get("corp_segments")
        if cs is not None:
            co_df = cs.get("companies", pd.DataFrame())
            if not co_df.empty and "CompanyName" in co_df.columns and "CY_Rooms" in co_df.columns:
                top = co_df[co_df["CY_Rooms"] > 0].sort_values("CY_Rooms", ascending=False).head(7)
                if not top.empty:
                    parts = [f"{r['CompanyName']} ({int(r['CY_Rooms']):,})" for _, r in top.iterrows()]
                    rows[3] = "Top Companies (past 2 wks): " + ", ".join(parts)
        else:
            booking = data.get("booking")
            if booking is not None and "MCAT" in booking.columns and "SRP_Name" in booking.columns:
                _2wk = today - pd.Timedelta(days=14)
                date_col = ("Arrival_Date" if "Arrival_Date" in booking.columns
                            else ("Date" if "Date" in booking.columns else None))
                if date_col:
                    local_neg = booking[
                        (pd.to_datetime(booking[date_col]) >= _2wk) &
                        (booking["MCAT"].str.strip().str.lower() == "local negotiated")
                    ]
                    if not local_neg.empty and "Room_Nights" in local_neg.columns:
                        co = local_neg.groupby("SRP_Name")["Room_Nights"].sum().sort_values(ascending=False).head(7)
                        if not co.empty:
                            parts = [f"{n} ({int(v):,})" for n, v in co.items()]
                            rows[3] = "Local Negotiated (past 2 wks): " + ", ".join(parts)
    except Exception:
        pass

    return rows


def _generate_call_recap_pdf(data, hotel) -> bytes:
    """Build 2-page Call Recap PDF. Returns bytes."""

    import io
    from datetime import datetime as _dt
    import pandas as pd
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph,
        Spacer, PageBreak, HRFlowable
    )
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from reportlab.pdfbase.pdfmetrics import stringWidth

    # ── Constants ────────────────────────────────────────────────────────────────
    PAGE_W, PAGE_H = landscape(letter)   # 792 × 612
    MARGIN  = 0.40 * inch
    TW      = PAGE_W - 2 * MARGIN        # ~727 pts usable

    # ── Colour palette — matches RevPar MD app exactly ─────────────────────────
    # Brand: navy #314B63, brand-green #6DA14B, brand-blue #1B70A7
    C = dict(
        # Core UI
        navy    = colors.HexColor("#314B63"),   # primary dark header bg
        dkblue  = colors.HexColor("#1B70A7"),   # brand blue
        midblue = colors.HexColor("#dce5e8"),   # light panel bg / sub-header row bg
        row1    = colors.HexColor("#ffffff"),   # table row 1 — white
        row2    = colors.HexColor("#eef2f3"),   # table row 2 — app alt row
        rowtot  = colors.HexColor("#314B63"),   # total row bg — navy
        wkband  = colors.HexColor("#dce5e8"),   # week band bg — app mid-panel
        border  = colors.HexColor("#c4cfd4"),   # cell borders
        # Section group header backgrounds (dark, coloured)
        gro_h   = colors.HexColor("#556848"),   # System Forecast — dark green
        gro_s   = colors.HexColor("#6DA14B"),   # brand green
        ps_h    = colors.HexColor("#b44820"),   # PS Forecast — rust/orange
        bud_h   = colors.HexColor("#2E618D"),   # Budget — brand blue
        diff_h  = colors.HexColor("#7a1a10"),   # Difference — deep red
        grp_h   = colors.HexColor("#dce5e8"),
        risk_h  = colors.HexColor("#a03020"),
        meet_h  = colors.HexColor("#314B63"),
        meet_s  = colors.HexColor("#dce5e8"),
        # Text on light (white/near-white) row backgrounds
        white   = colors.white,
        offwhite= colors.HexColor("#ffffff"),   # white — for labels on dark row/total bg
        dim     = colors.HexColor("#6a8090"),   # muted — dashes, locked cells
        pritxt  = colors.HexColor("#1e2d35"),   # primary dark text on light rows
        sectxt  = colors.HexColor("#3a5260"),   # secondary text on light rows (budget values)
        # Accent data value colours — readable on white/light row backgrounds
        teal    = colors.HexColor("#1B70A7"),   # brand blue — positive variance
        gold    = colors.HexColor("#556848"),   # dark green — PS forecast values (readable on white)
        orange  = colors.HexColor("#b44820"),   # rust — negative variance
        green   = colors.HexColor("#556848"),   # dark green
        red_txt = colors.HexColor("#a03020"),   # red negative
        # Header text on coloured section header backgrounds (light, for contrast)
        gro_sub = colors.HexColor("#d4e8c8"),   # light green — header text on dark-green bg
        gro_val = colors.HexColor("#3d5c2e"),   # dark green — System Forecast values on white rows
        ps_sub  = colors.HexColor("#ffd4b0"),   # light orange on rust bg
        diff_sub= colors.HexColor("#ffccc8"),   # light pink on deep-red bg
        # Text on dark header bar (navy bg)
        hdr_dim = colors.HexColor("#8ab4cc"),   # muted light blue — metadata on navy
    )

    BDR = C["border"]

    def _style(size=7, color=None, bold=False, align=TA_CENTER, leading=None, italic=False):
        fn = ("Helvetica-BoldOblique" if bold and italic else
              "Helvetica-Bold" if bold else
              "Helvetica-Oblique" if italic else "Helvetica")
        return ParagraphStyle("_", fontName=fn, fontSize=size,
                              textColor=color or C["pritxt"],
                              leading=leading or max(size * 1.3, size + 3),
                              alignment=align, spaceAfter=0, spaceBefore=0)

    def _p(text, size=7, color=None, bold=False, align=TA_CENTER, italic=False):
        return Paragraph(str(text), _style(size, color, bold, align, italic=italic))

    def _hdr_bar(text, bg, width=None, size=8, left=False):
        """Full-width coloured header bar."""
        w = width or TW
        align = TA_LEFT if left else TA_CENTER
        data = [[_p(text, size=size, color=C["white"], bold=True, align=align)]]
        t = Table(data, colWidths=[w])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), bg),
            ("LEFTPADDING",  (0,0), (-1,-1), 10 if left else 6),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
            ("TOPPADDING",   (0,0), (-1,-1), 5),
            ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ]))
        return t

    def _build_pdf(data, hotel) -> bytes:
        buf = io.BytesIO()

        hotel_name   = hotel.get("display_name", "") if isinstance(hotel, dict) else str(hotel)
        hotel_sub    = hotel.get("subtitle", "")     if isinstance(hotel, dict) else ""
        hotel_id     = hotel.get("id", "hotel")      if isinstance(hotel, dict) else "hotel"
        total_rooms  = hotel.get("total_rooms", 112) if isinstance(hotel, dict) else 112

        today        = pd.Timestamp(_dt.now().date())
        cur_month    = today.replace(day=1)
        months       = [cur_month + pd.DateOffset(months=i) for i in range(3)]
        month_lbls   = [m.strftime("%B") for m in months]
        month_range  = f"{month_lbls[0]} – {month_lbls[2]} {months[0].year}"
        gen_ts       = _dt.now().strftime("%B %d, %Y  ·  %I:%M %p")

        year_df   = data.get("year")
        budget_df = data.get("budget")
        groups_df = data.get("groups")
        str_ranks = data.get("str_ranks") or {}

        # ── Data helpers ─────────────────────────────────────────────────────────
        def _parse(s):
            try: return float(str(s).replace(",","").replace("$","").strip())
            except: return None

        def _fr(v): return f"${int(round(v)):,}" if v is not None else "—"
        def _fn(v): return f"{int(round(v)):,}"  if v is not None else "—"
        def _fa(v): return f"${v:,.2f}"           if v is not None else "—"

        def _gro_month(m):
            if year_df is None: return None, None, None
            mask = (year_df["Date"].dt.year==m.year)&(year_df["Date"].dt.month==m.month)
            sub  = year_df[mask]
            if sub.empty: return None,None,None
            nts = sub["Forecast_Rooms"].sum() if "Forecast_Rooms" in sub else None
            # Hilton: Revenue_Forecast covers all days fully
            if "Revenue_Forecast" in sub.columns and sub["Revenue_Forecast"].notna().any():
                rev = sub["Revenue_Forecast"].sum()
            # IHG: Forecast_Rev is future-days only. Blend past Revenue_OTB + future Forecast_Rev
            # so the current month includes both what's already on the books AND remaining forecast.
            elif "Forecast_Rev" in sub.columns and sub["Forecast_Rev"].notna().any():
                fut_rev  = float(sub["Forecast_Rev"].fillna(0).sum())
                past_rev = float(sub.loc[sub["Forecast_Rev"].isna(), "Revenue_OTB"].fillna(0).sum()) \
                           if "Revenue_OTB" in sub.columns else 0.0
                rev = fut_rev + past_rev
            else:
                rev = None
            adr = (rev/nts) if (nts and nts>0 and rev) else None
            return (round(rev,0) if rev else None, round(nts,0) if nts else None,
                    round(adr,2) if adr else None)

        def _bud_month(m):
            if budget_df is None: return None,None,None
            mask = (budget_df["Date"].dt.year==m.year)&(budget_df["Date"].dt.month==m.month)
            sub  = budget_df[mask]
            if sub.empty: return None,None,None
            nts = sub["Occ_Rooms"].sum(); rev = sub["Revenue"].sum()
            adr = (rev/nts) if nts and nts>0 else None
            return round(rev,0), round(nts,0), (round(adr,2) if adr else None)

        # Load PS values
        import json as _j; from pathlib import Path as _P
        _cr_folder = _P(hotel["folder_name"]) if isinstance(hotel,dict) and "folder_name" in hotel else (_P(hotel.get("folder",".")) if isinstance(hotel,dict) else _P("."))
        # Resolve to full OneDrive path if not absolute
        if not _cr_folder.is_absolute():
            try:
                import hotel_config as _hcfg
                _cr_folder = _hcfg.get_hotel_folder(hotel)
            except Exception:
                pass
        _cr_path   = _cr_folder / "call_recap_ps.json"
        _cr_mkey   = cur_month.strftime("%Y-%m")
        _ps_vals   = {}
        _w1_sub_date = ""
        try:
            if _cr_path.exists():
                _d = _j.loads(_cr_path.read_text(encoding="utf-8"))
                if _d.get("month")==_cr_mkey and _d.get("hotel")==hotel_id:
                    _ps_vals     = _d.get("values", {})
                    _w1_sub_date = _d.get("w1_submitted_date","")
        except: pass

        _w2_unlocked = False
        if _w1_sub_date:
            try: _w2_unlocked = today.date() > pd.Timestamp(_w1_sub_date).date()
            except: pass

        # ── Column widths — full-width table matching header bar ────────────────
        # LBL | GRO(REV,NTS,ADR) | PS(REV,NTS,ADR) | BUD(REV,NTS,ADR) | DIFF(REV,NTS,ADR) | NOTES
        FC_LBL  = 50
        FC_REV  = 62   # "$942,693" — 8 chars at 7pt bold
        FC_NTS  = 42   # "NIGHTS" — 6 chars at 6pt bold needs ~35pt; 42 gives comfortable padding
        FC_ADR  = 39   # "$158.14" — 7 chars at 7pt bold
        FC_NOTE = int(TW - FC_LBL - (FC_REV + FC_NTS + FC_ADR) * 4)  # ~112pt remaining
        FC_COLS = [FC_LBL] + [FC_REV, FC_NTS, FC_ADR] * 4 + [FC_NOTE]  # 14 cols, sums to TW

        # ── Story builder ─────────────────────────────────────────────────────────
        story = []

        def _footer(canvas, doc):
            canvas.saveState()
            canvas.setFont("Helvetica", 6)
            canvas.setFillColor(C["dim"])
            y = MARGIN * 0.45
            canvas.drawString(MARGIN, y, f"RevPar MD  ·  {hotel_name}  ·  Confidential")
            canvas.drawRightString(PAGE_W-MARGIN, y, f"Page {doc.page} of 2")
            canvas.restoreState()

        # ═══════════════ PAGE 1 ══════════════════════════════════════════════════

        # ── Top header bar ────────────────────────────────────────────────────────
        hdr_rows = [[
            _p("REVENUE MANAGEMENT CALL RECAP", size=11, color=C["white"], bold=True, align=TA_LEFT),
            _p(f"{hotel_name}  ·  {hotel_sub}", size=8, color=C["offwhite"], bold=True, align=TA_LEFT),
            _p(f"Forecast Review + Pace  ·  {month_range}", size=7, color=C["hdr_dim"], align=TA_LEFT),
            _p(gen_ts, size=7, color=C["hdr_dim"], align=TA_RIGHT),
        ]]
        hdr_t = Table(hdr_rows, colWidths=[TW*0.30, TW*0.25, TW*0.28, TW*0.17])
        hdr_t.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(-1,-1), C["navy"]),
            ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1), 10),
            ("RIGHTPADDING",(0,0),(-1,-1),8),
            ("TOPPADDING", (0,0),(-1,-1), 10),
            ("BOTTOMPADDING",(0,0),(-1,-1),10),
            ("LINEBELOW",(0,0),(-1,-1), 2.5, C["teal"]),
        ]))
        story.append(hdr_t)
        story.append(Spacer(1, 5))

        # ── Build PS forecast table rows ──────────────────────────────────────────
        def _cell(val, color=None, bold=False, align=TA_CENTER, size=7):
            return _p(val, size=size, color=color or C["pritxt"], bold=bold, align=align)

        def _diff_cell(v, is_adr=False, is_nts=False):
            if v is None: return _cell("—", color=C["dim"])
            col = C["teal"] if v >= 0 else C["orange"]
            pfx = "+" if v >= 0 else "-"
            if is_adr:   s = f"{pfx}${abs(v):,.2f}"
            elif is_nts: s = f"{pfx}{int(abs(v)):,}"
            else:        s = f"{pfx}${int(abs(v)):,}"
            return _p(s, color=col, bold=True)

        def _ps_week_data(wk_num, locked):
            wk = f"w{wk_num}"
            rows_out = []
            tot = {k:0.0 for k in ["gr","gn","br","bn","pr","pn"]}
            tv  = {k:False for k in tot}

            for m, lbl in zip(months, month_lbls):
                gr, gn, ga = _gro_month(m)
                br, bn, ba = _bud_month(m)
                m_str = m.strftime("%Y-%m")
                pr  = _parse(_ps_vals.get(f"{wk}_{m_str}_rev","")) if not locked else None
                pn  = _parse(_ps_vals.get(f"{wk}_{m_str}_nts","")) if not locked else None
                pa  = (pr/pn) if (pr and pn and pn>0) else None

                diff_r = (pr-br) if (pr is not None and br is not None) else None
                diff_n = (pn-bn) if (pn is not None and bn is not None) else None
                diff_a = (pa-ba) if (pa is not None and ba is not None) else None

                for k,v in [("gr",gr),("gn",gn),("br",br),("bn",bn),("pr",pr),("pn",pn)]:
                    if v is not None: tot[k]+=v; tv[k]=True

                dim = C["dim"]
                gro_r = _cell(_fr(gr), color=C["gro_val"]) if gr else _cell("—",color=dim)
                gro_n = _cell(_fn(gn), color=C["gro_val"]) if gn else _cell("—",color=dim)
                gro_a = _cell(_fa(ga), color=C["gro_val"]) if ga else _cell("—",color=dim)

                if locked:
                    ps_r = ps_n = ps_a = _cell("—", color=dim)
                else:
                    ps_r = _cell(_fr(pr), color=C["gold"]) if pr else _cell("—",color=dim)
                    ps_n = _cell(_fn(pn), color=C["gold"]) if pn else _cell("—",color=dim)
                    ps_a = _cell(_fa(pa), color=C["gold"]) if pa else _cell("—",color=dim)

                bud_r = _cell(_fr(br), color=C["sectxt"]) if br else _cell("—",color=dim)
                bud_n = _cell(_fn(bn), color=C["sectxt"]) if bn else _cell("—",color=dim)
                bud_a = _cell(_fa(ba), color=C["sectxt"]) if ba else _cell("—",color=dim)

                if locked:
                    df_r = df_n = df_a = _cell("—", color=dim)
                else:
                    df_r = _diff_cell(diff_r)
                    df_n = _diff_cell(diff_n, is_nts=True)
                    df_a = _diff_cell(diff_a, is_adr=True)

                rows_out.append([
                    _cell(lbl, align=TA_LEFT, color=C["pritxt"], bold=True),
                    gro_r, gro_n, gro_a,
                    ps_r,  ps_n,  ps_a,
                    bud_r, bud_n, bud_a,
                    df_r,  df_n,  df_a,
                    _cell("", color=C["dim"]),  # notes — filled via rowspan below
                ])

            # Totals row
            tga = tot["gr"]/tot["gn"] if tv["gn"] and tot["gn"]>0 else None
            tba = tot["br"]/tot["bn"] if tv["bn"] and tot["bn"]>0 else None
            tpa = tot["pr"]/tot["pn"] if tv["pn"] and tot["pn"]>0 else None
            dr  = (tot["pr"]-tot["br"]) if (tv["pr"] and tv["br"]) else None
            dn  = (tot["pn"]-tot["bn"]) if (tv["pn"] and tv["bn"]) else None
            da  = (tpa-tba) if (tpa and tba) else None
            dim = C["dim"]

            def _tc(v, fmt, color): return _cell(fmt(v), color=color, bold=True) if v else _cell("—",color=dim,bold=True)

            if locked:
                ps_tot  = [_cell("—",color=dim,bold=True)]*3
                dif_tot = [_cell("—",color=dim,bold=True)]*3
            else:
                ps_tot  = [_tc(tot["pr"] if tv["pr"] else None, _fr, C["offwhite"]),
                           _tc(tot["pn"] if tv["pn"] else None, _fn, C["offwhite"]),
                           _tc(tpa, _fa, C["offwhite"])]
                # Total row is navy bg — use white for all difference cells
                def _diff_tot(v, is_adr=False, is_nts=False):
                    if v is None: return _cell("—", color=C["offwhite"], bold=True)
                    pfx = "+" if v >= 0 else "-"
                    if is_adr:   s = f"{pfx}${abs(v):,.2f}"
                    elif is_nts: s = f"{pfx}{int(abs(v)):,}"
                    else:        s = f"{pfx}${int(abs(v)):,}"
                    return _p(s, color=C["offwhite"], bold=True)
                dif_tot = [_diff_tot(dr), _diff_tot(dn,is_nts=True), _diff_tot(da,is_adr=True)]

            rows_out.append([
                _cell("TOTAL", align=TA_LEFT, color=C["white"], bold=True),
                _tc(tot["gr"] if tv["gr"] else None, _fr, C["offwhite"]),
                _tc(tot["gn"] if tv["gn"] else None, _fn, C["offwhite"]),
                _tc(tga, _fa, C["offwhite"]),
                *ps_tot,
                _tc(tot["br"] if tv["br"] else None, _fr, C["offwhite"]),
                _tc(tot["bn"] if tv["bn"] else None, _fn, C["offwhite"]),
                _tc(tba, _fa, C["offwhite"]),
                *dif_tot,
            ])
            return rows_out

        # ── Build complete forecast Table object ──────────────────────────────────
        def _build_forecast_tbl():
            # Header row 1: group spans
            def _gh(txt, bg, fc, span=1):
                p = ParagraphStyle("gh", fontName="Helvetica-Bold", fontSize=6,
                                   textColor=fc, alignment=TA_CENTER, leading=8)
                return Paragraph(txt, p)

            sub_fc = {
                "gro":  C["gro_sub"],    # light green on dark-green
                "ps":   C["ps_sub"],     # light orange on rust
                "bud":  C["offwhite"],   # white on brand-blue
                "diff": C["diff_sub"],   # light pink on deep-red
            }

            r1 = [
                _gh("GUEST ROOMS", C["navy"], C["offwhite"]),
                _gh("SYSTEM FORECAST",C["gro_h"],C["gro_sub"]), "","",
                _gh("PS FORECAST", C["ps_h"], C["ps_sub"]),  "","",
                _gh("BUDGET",      C["bud_h"],C["sectxt"]),   "","",
                _gh("DIFF TO PS FORECAST",  C["diff_h"],C["diff_sub"]),"","",
                _gh("FORECAST NOTES", C["navy"], C["offwhite"]),
            ]
            r2 = [
                "",
                _gh("REVENUE",C["gro_h"],sub_fc["gro"]), _gh("NIGHTS",C["gro_h"],sub_fc["gro"]), _gh("ADR",C["gro_h"],sub_fc["gro"]),
                _gh("REVENUE",C["ps_h"], sub_fc["ps"]),  _gh("NIGHTS",C["ps_h"], sub_fc["ps"]),  _gh("ADR",C["ps_h"], sub_fc["ps"]),
                _gh("REVENUE",C["bud_h"],sub_fc["bud"]), _gh("NIGHTS",C["bud_h"],sub_fc["bud"]), _gh("ADR",C["bud_h"],sub_fc["bud"]),
                _gh("REVENUE",C["diff_h"],sub_fc["diff"]),_gh("NIGHTS",C["diff_h"],sub_fc["diff"]),_gh("ADR",C["diff_h"],sub_fc["diff"]),
                "",
            ]

            all_rows = [r1, r2]
            style_cmds = [
                # Header backgrounds
                ("BACKGROUND",(0,0),(0,1),   C["navy"]),
                ("BACKGROUND",(1,0),(3,0),   C["gro_h"]),
                ("BACKGROUND",(4,0),(6,0),   C["ps_h"]),
                ("BACKGROUND",(7,0),(9,0),   C["bud_h"]),
                ("BACKGROUND",(10,0),(12,0), C["diff_h"]),
                ("BACKGROUND",(13,0),(13,1), C["dkblue"]),   # FORECAST NOTES col header
                ("BACKGROUND",(1,1),(3,1),   C["gro_h"]),
                ("BACKGROUND",(4,1),(6,1),   C["ps_h"]),
                ("BACKGROUND",(7,1),(9,1),   C["bud_h"]),
                ("BACKGROUND",(10,1),(12,1), C["diff_h"]),
                # Spans
                ("SPAN",(0,0),(0,1)),
                ("SPAN",(1,0),(3,0)), ("SPAN",(4,0),(6,0)),
                ("SPAN",(7,0),(9,0)), ("SPAN",(10,0),(12,0)),
                # Alignment
                ("ALIGN",    (0,0),(-1,-1), "CENTER"),
                ("VALIGN",   (0,0),(-1,-1), "MIDDLE"),
                ("FONTNAME", (0,0),(-1,-1), "Helvetica-Bold"),
                ("FONTSIZE", (0,0),(-1,-1), 6),
                ("TOPPADDING",(0,0),(-1,-1), 4),
                ("BOTTOMPADDING",(0,0),(-1,-1), 4),
                ("GRID",(0,0),(-1,-1), 0.4, BDR),
            ]

            for wk_num in [1, 2]:
                locked = (wk_num == 2 and not _w2_unlocked)
                note   = _ps_vals.get(f"w{wk_num}_note","") if not locked else ""

                # Week band
                wk_label = f"  WEEK {wk_num}" + ("   ·   Locked – unlocks the day after Week 1 is complete" if locked else "")
                band_fc  = C["dim"] if locked else C["pritxt"]
                band_bg  = C["midblue"] if not locked else C["wkband"]
                ri = len(all_rows)
                band_row = [_p(wk_label, size=6.5, color=band_fc, bold=True, align=TA_LEFT)] + [""]*13
                all_rows.append(band_row)
                style_cmds += [
                    ("SPAN",(0,ri),(-1,ri)),
                    ("BACKGROUND",(0,ri),(-1,ri), band_bg),
                    ("LEFTPADDING",(0,ri),(-1,ri), 8),
                    ("TOPPADDING",(0,ri),(-1,ri), 5),
                    ("BOTTOMPADDING",(0,ri),(-1,ri), 5),
                    ("LINEABOVE",(0,ri),(-1,ri), 1.0, C["dkblue"] if not locked else BDR),
                ]

                # Data rows
                data_rows = _ps_week_data(wk_num, locked)
                for ii, dr in enumerate(data_rows):
                    ri2 = len(all_rows)
                    is_total = (ii == len(data_rows)-1)
                    bg = C["rowtot"] if is_total else (C["row1"] if ii%2==0 else C["row2"])
                    all_rows.append(dr)
                    style_cmds += [
                        ("BACKGROUND",(0,ri2),(-1,ri2), bg),
                        ("TOPPADDING",(0,ri2),(-1,ri2), 3),
                        ("BOTTOMPADDING",(0,ri2),(-1,ri2), 3),
                        ("ALIGN",(0,ri2),(0,ri2), "LEFT"),
                        ("LEFTPADDING",(0,ri2),(0,ri2), 7),
                    ]
                    if is_total:
                        style_cmds += [
                            ("LINEABOVE",(0,ri2),(-1,ri2), 1.0, C["dkblue"]),
                            ("LINEBELOW",(0,ri2),(-1,ri2), 0.5, BDR),
                        ]

                # Notes: fill col 13 of first data row, rowspan across 4 rows (3 months + total)
                note_start_ri = len(all_rows) - len(data_rows)
                note_para = _p(note if note else "", size=6.5, color=C["pritxt"],
                               bold=False, align=TA_LEFT, italic=True)
                all_rows[note_start_ri][13] = note_para
                note_bg = C["row1"]  # always white — note text is dark, readable on white
                style_cmds += [
                    ("SPAN",        (13,note_start_ri),(13,note_start_ri+3)),
                    ("BACKGROUND",  (13,note_start_ri),(13,note_start_ri+3), note_bg),
                    ("LEFTPADDING", (13,note_start_ri),(13,note_start_ri+3), 6),
                    ("TOPPADDING",  (13,note_start_ri),(13,note_start_ri+3), 5),
                    ("VALIGN",      (13,note_start_ri),(13,note_start_ri+3), "TOP"),
                ]

            # Global data rows style
            style_cmds += [
                ("ALIGN",(0,2),(-1,-1), "CENTER"),
                ("VALIGN",(0,2),(-1,-1), "MIDDLE"),
                ("FONTSIZE",(0,2),(-1,-1), 7),
                ("GRID",(0,0),(-1,-1), 0.4, BDR),
            ]

            t = Table(all_rows, colWidths=FC_COLS, repeatRows=2)
            t.setStyle(TableStyle(style_cmds))
            return t

        # ── Build 14-day daily recap table ───────────────────────────────────────
        def _build_daily():
            if year_df is None:
                return Table([[_p("No data.", color=C["dim"])]], colWidths=[DR_W])

            past14 = year_df[
                (year_df["Date"] >= today - pd.Timedelta(days=14)) &
                (year_df["Date"] < today)
            ].copy().sort_values("Date", ascending=False)

            if past14.empty:
                return Table([[_p("No recent data.", color=C["dim"])]], colWidths=[DR_W])

            def _dh(txt, bg):
                return _p(txt, size=6, color=C["white"], bold=True)

            hdr = [[_dh("DATE",C["navy"]), _dh("RMS\nTY",C["gro_h"]), _dh("REV TY",C["gro_h"]),
                    _dh("RMS\nLY",C["bud_h"]), _dh("REV LY",C["bud_h"])]]
            style = [
                ("BACKGROUND",(0,0),(0,0), C["navy"]),
                ("BACKGROUND",(1,0),(2,0), C["gro_h"]),
                ("BACKGROUND",(3,0),(4,0), C["bud_h"]),
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("TOPPADDING",(0,0),(-1,-1),3),
                ("BOTTOMPADDING",(0,0),(-1,-1),3),
                ("FONTSIZE",(0,0),(-1,-1),6),
                ("GRID",(0,0),(-1,-1),0.4,BDR),
            ]

            rows = []
            for i, (_idx, r) in enumerate(past14.iterrows()):
                rb = C["row1"] if i%2==0 else C["row2"]
                ts = pd.Timestamp(r["Date"])
                dt_str = f"{ts.strftime('%a')} {ts.month}/{ts.day}"

                ty_r = next((int(r[c]) for c in ["OTB","Forecast_Rooms"]
                             if c in r.index and pd.notna(r.get(c))), None)
                ty_v = next((float(r[c]) for c in ["Revenue_OTB","Revenue_Forecast"]
                             if c in r.index and pd.notna(r.get(c))), None)
                ly_r = int(float(r["OTB_LY"])) if ("OTB_LY" in r.index and pd.notna(r.get("OTB_LY"))) else None
                ly_v = float(r["Revenue_LY"])   if ("Revenue_LY" in r.index and pd.notna(r.get("Revenue_LY"))) else None

                rows.append([
                    _p(dt_str, size=6.5, color=C["offwhite"], align=TA_LEFT),
                    _p(f"{ty_r:,}" if ty_r else "—", size=6.5, color=C["gro_val"]),
                    _p(f"${int(ty_v):,}" if ty_v else "—", size=6.5, color=C["gro_val"]),
                    _p(f"{ly_r:,}" if ly_r else "—", size=6.5, color=C["sectxt"]),
                    _p(f"${int(ly_v):,}" if ly_v else "—", size=6.5, color=C["sectxt"]),
                ])
                style.append(("BACKGROUND",(0,i+1),(-1,i+1), rb))
                style.append(("LEFTPADDING",(0,i+1),(0,i+1), 4))

            t = Table(hdr+rows, colWidths=DR_COLS, repeatRows=1)
            t.setStyle(TableStyle(style))
            return t

        # ── Page 1: full-width forecast table ────────────────────────────────────
        fc_tbl = _build_forecast_tbl()
        story.append(fc_tbl)


        # ═══════════════ SRP PACE BY SEGMENT (Page 1, below forecast) ═══════════════
        def _build_srp_pace_tbl():
            """Build the SRP Pace by Segment table.
            IHG: uses corp_segments (CY/LY Rooms, Revenue, ADR by Segment).
            Hilton: uses srp_pace with short MCAT codes."""

            is_ihg_hotel = hotel.get("brand", "") == "IHG"

            # ── IHG path: corp_segments ──────────────────────────────────────
            if is_ihg_hotel:
                cs = data.get("corp_segments")
                if cs is None:
                    return None, None
                seg_df = cs.get("segments", pd.DataFrame())
                if seg_df.empty or "Segment" not in seg_df.columns:
                    return None, None

                # Filter to segments with any activity
                active_segs = seg_df[(seg_df["CY_Rooms"] > 0) | (seg_df["LY_Rooms"] > 0)].copy()
                if active_segs.empty:
                    return None, None

                SEG_W  = 160
                DATA_W = round((TW - SEG_W) / 9, 1)
                SW = [SEG_W] + [DATA_W] * 9

                def _sv(v, fmt="int"):
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        return _p("—", size=6.5, color=C["dim"])
                    if fmt == "int":   s = f"{int(round(v)):,}"
                    elif fmt == "rev": s = f"${int(round(v)):,}"
                    elif fmt == "adr": s = f"${v:,.2f}"
                    else: s = str(v)
                    return _p(s, size=6.5, color=C["pritxt"])

                def _var(v, fmt="int"):
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        return _p("—", size=6.5, color=C["dim"])
                    col = C["teal"] if v >= 0 else C["orange"]
                    pfx = "+" if v > 0 else ("-" if v < 0 else "")
                    if fmt == "int":   s = f"{pfx}{int(round(abs(v))):,}"
                    elif fmt == "rev": s = f"{pfx}${int(round(abs(v))):,}"
                    elif fmt == "adr": s = f"{pfx}${abs(v):,.2f}"
                    else: s = str(v)
                    return _p(s, size=6.5, color=col, bold=True)

                def _seg_label(seg):
                    st = ParagraphStyle("sl", fontName="Helvetica-Bold", fontSize=6,
                                        textColor=C["offwhite"], leading=8.5,
                                        alignment=TA_LEFT, spaceAfter=0, spaceBefore=0)
                    return Paragraph(str(seg).upper(), st)

                grp_hdr = [[
                    _p("SEGMENT",           size=6, color=C["white"], bold=True),
                    _p("OCCUPANCY ON BOOKS", size=6, color=C["white"], bold=True),
                    _p(""), _p(""),
                    _p("ADR ON BOOKS",      size=6, color=C["white"], bold=True),
                    _p(""), _p(""),
                    _p("REVENUE ON BOOKS",  size=6, color=C["white"], bold=True),
                    _p(""), _p(""),
                ]]
                sub_hdr = [[
                    _p("",        size=6, color=C["sectxt"]),
                    _p("CY",      size=6, color=C["sectxt"], bold=True),
                    _p("LY",      size=6, color=C["sectxt"], bold=True),
                    _p("VARIANCE",size=6, color=C["sectxt"], bold=True),
                    _p("CY",      size=6, color=C["sectxt"], bold=True),
                    _p("LY",      size=6, color=C["sectxt"], bold=True),
                    _p("VARIANCE",size=6, color=C["sectxt"], bold=True),
                    _p("CY",      size=6, color=C["sectxt"], bold=True),
                    _p("LY",      size=6, color=C["sectxt"], bold=True),
                    _p("VARIANCE",size=6, color=C["sectxt"], bold=True),
                ]]

                all_rows  = grp_hdr + sub_hdr
                data_start = len(all_rows)
                style_cmds = [
                    ("SPAN",(1,0),(3,0)), ("SPAN",(4,0),(6,0)), ("SPAN",(7,0),(9,0)),
                    ("BACKGROUND",(0,0),(0,0),  C["dkblue"]),
                    ("BACKGROUND",(1,0),(3,0),  colors.HexColor("#d8eed8")),
                    ("BACKGROUND",(4,0),(6,0),  colors.HexColor("#f5f0d8")),
                    ("BACKGROUND",(7,0),(9,0),  colors.HexColor("#dce8f5")),
                    ("BACKGROUND",(0,1),(-1,1), C["meet_s"]),
                    ("ALIGN",(0,0),(-1,-1),"CENTER"),
                    ("ALIGN",(0,2),(0,-1),"LEFT"),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                    ("FONTSIZE",(0,0),(-1,-1),6.5),
                    ("TOPPADDING",(0,0),(-1,-1),2),
                    ("BOTTOMPADDING",(0,0),(-1,-1),2),
                    ("LEFTPADDING",(0,0),(-1,-1),3),
                    ("RIGHTPADDING",(0,0),(-1,-1),3),
                    ("LEFTPADDING",(0,2),(0,-1),5),
                    ("GRID",(0,0),(-1,-1),0.4,BDR),
                ]

                for i, (_, row) in enumerate(active_segs.iterrows()):
                    rb = C["row1"] if i % 2 == 0 else C["row2"]
                    all_rows.append([
                        _seg_label(row["Segment"]),
                        _sv(row.get("CY_Rooms")),       _sv(row.get("LY_Rooms")),  _var(row.get("Var_Rooms")),
                        _sv(row.get("CY_ADR"), "adr"),  _sv(row.get("LY_ADR"), "adr"), _var(row.get("Var_ADR"), "adr"),
                        _sv(row.get("CY_Rev"),  "rev"),  _sv(row.get("LY_Rev"),  "rev"), _var(row.get("Var_Rev"),  "rev"),
                    ])
                    ri = data_start + i
                    style_cmds.append(("BACKGROUND",(0,ri),(-1,ri), rb))

                # Totals row
                tot_ri = len(all_rows)
                t_cy_r = active_segs["CY_Rooms"].sum()
                t_ly_r = active_segs["LY_Rooms"].sum()
                t_cy_v = active_segs["CY_Rev"].sum()
                t_ly_v = active_segs["LY_Rev"].sum()
                t_cy_a = t_cy_v / t_cy_r if t_cy_r > 0 else None
                t_ly_a = t_ly_v / t_ly_r if t_ly_r > 0 else None
                all_rows.append([
                    _p("TOTAL", size=6.5, color=C["offwhite"], bold=True),
                    _sv(t_cy_r),            _sv(t_ly_r),             _var(t_cy_r - t_ly_r),
                    _sv(t_cy_a, "adr"),     _sv(t_ly_a, "adr"),      _var((t_cy_a - t_ly_a) if (t_cy_a and t_ly_a) else None, "adr"),
                    _sv(t_cy_v, "rev"),     _sv(t_ly_v, "rev"),      _var(t_cy_v - t_ly_v, "rev"),
                ])
                style_cmds += [
                    ("BACKGROUND",(0,tot_ri),(-1,tot_ri), C["rowtot"]),
                    ("LINEABOVE",(0,tot_ri),(-1,tot_ri), 1.0, C["dkblue"]),
                ]

                tbl = Table(all_rows, colWidths=SW, repeatRows=2)
                tbl.setStyle(TableStyle(style_cmds))
                hdr_bar = _hdr_bar(
                    "REVENUE BY SEGMENT  ·  CURRENT YEAR vs LAST YEAR",
                    C["navy"], left=True
                )
                return hdr_bar, tbl

            # ── Hilton path: srp_pace with short MCAT codes ──────────────────
            sp = data.get("srp_pace")
            if sp is None or sp.empty or "Date" not in sp.columns:
                return None, None

            seg_col = "Segment" if "Segment" in sp.columns else ("MCAT" if "MCAT" in sp.columns else None)
            if not seg_col:
                return None, None

            # Determine window — same logic as _compute_seg_note
            try:
                check_cols = [c for c in ["OTB","Revenue","STLY_OTB","STLY_Revenue"] if c in sp.columns]
                _nonzero = sp.groupby("Date")[check_cols].sum()
                _nonzero = _nonzero[_nonzero.sum(axis=1) > 0]
                if _nonzero.empty:
                    return None, None
                last_data_d  = _nonzero.index.max()
                window_end   = min(last_data_d, today)
                window_start = window_end - pd.Timedelta(days=13)
            except Exception:
                return None, None

            pace_2w = sp[(sp["Date"] >= window_start) & (sp["Date"] <= window_end)]
            if pace_2w.empty:
                return None, None

            start_lbl  = window_start.strftime("%b %d")
            end_lbl    = window_end.strftime("%b %d, %Y")
            date_range = f"{start_lbl} – {end_lbl}"

            SEG_COLORS = {
                "BAR":  "#4fc3f7", "DISC": "#81c784", "CNR":  "#ffb74d",
                "MKT":  "#ce93d8", "LNR":  "#80cbc4", "CONS": "#fff176",
                "SMRF": "#ff8a65", "GOV":  "#a5d6a7", "IT":   "#90caf9",
                "CMP":  "#b0bec5", "CMTG": "#f48fb1", "GT":   "#bcaaa4",
                "CONV": "#ffe082",
            }
            SEGS_ORDER = ["BAR","CMP","CMTG","CNR","CONS","CONV","DISC","GOV","GT","IT","LNR","MKT","SMRF"]

            # Aggregate
            seg_agg = {}
            for seg in SEGS_ORDER:
                grp = pace_2w[pace_2w[seg_col] == seg]
                otb      = float(grp["OTB"].sum())      if "OTB"          in grp.columns else 0.0
                stly_otb = float(grp["STLY_OTB"].sum()) if "STLY_OTB"    in grp.columns else 0.0
                rev      = float(grp["Revenue"].sum())      if "Revenue"  in grp.columns else 0.0
                stly_rev = float(grp["STLY_Revenue"].sum()) if "STLY_Revenue" in grp.columns else 0.0
                adr      = rev / otb           if otb > 0      else None
                stly_adr = stly_rev / stly_otb if stly_otb > 0 else None
                seg_agg[seg] = dict(
                    otb=otb, stly_otb=stly_otb, var_otb=otb-stly_otb,
                    rev=rev, stly_rev=stly_rev, var_rev=rev-stly_rev,
                    adr=adr, stly_adr=stly_adr,
                    var_adr=(adr-stly_adr) if (adr is not None and stly_adr is not None) else None)

            tot_otb      = sum(v["otb"]      for v in seg_agg.values())
            tot_stly_otb = sum(v["stly_otb"] for v in seg_agg.values())
            tot_rev      = sum(v["rev"]      for v in seg_agg.values())
            tot_stly_rev = sum(v["stly_rev"] for v in seg_agg.values())
            tot_adr      = tot_rev / tot_otb           if tot_otb > 0      else None
            tot_stly_adr = tot_stly_rev / tot_stly_otb if tot_stly_otb > 0 else None

            # ── Equal column widths ──────────────────────────────────────────────
            # 10 cols: 1 seg label + 9 equal data cols
            SEG_W  = 56                           # segment label col
            DATA_W = round((TW - SEG_W) / 9, 1)  # 9 equal data columns
            SW = [SEG_W] + [DATA_W] * 9

            # ── Cell helpers — all data centered ────────────────────────────────
            def _dot(seg):
                c = colors.HexColor(SEG_COLORS.get(seg, "#888888"))
                st = ParagraphStyle("dot", fontName="Helvetica-Bold", fontSize=6.5,
                                    textColor=c, leading=9, alignment=TA_LEFT,
                                    spaceAfter=0, spaceBefore=0)
                return Paragraph(f"● {seg}", st)

            def _sv(v, fmt="int"):
                if v is None: return _p("—", size=6.5, color=C["dim"])
                if fmt == "int": s = f"{int(round(v)):,}"
                elif fmt == "rev": s = f"${int(round(v)):,}"
                elif fmt == "adr": s = f"${v:,.2f}"
                else: s = str(v)
                return _p(s, size=6.5, color=C["pritxt"])

            def _var(v, fmt="int"):
                if v is None: return _p("—", size=6.5, color=C["dim"])
                col = C["teal"] if v >= 0 else C["orange"]
                pfx = "+" if v > 0 else ("-" if v < 0 else "")
                if fmt == "int":  s = f"{pfx}{int(round(abs(v))):,}"
                elif fmt == "rev": s = f"{pfx}${int(round(abs(v))):,}"
                elif fmt == "adr": s = f"{pfx}${abs(v):,.2f}"
                else: s = str(v)
                return _p(s, size=6.5, color=col, bold=True)

            # ── Build rows ───────────────────────────────────────────────────────
            grp_hdr = [[
                _p("SEGMENT",          size=6, color=C["white"], bold=True),
                _p("OCCUPANCY ON BOOKS", size=6, color=C["white"], bold=True),
                _p(""), _p(""),
                _p("ADR ON BOOKS",     size=6, color=C["white"], bold=True),
                _p(""), _p(""),
                _p("REVENUE ON BOOKS", size=6, color=C["white"], bold=True),
                _p(""), _p(""),
            ]]
            _sh = colors.HexColor("#4e6878")  # dark muted blue — readable on light panel bg
            sub_hdr = [[
                _p("",         size=6, color=_sh),
                _p("CURRENT",  size=6, color=_sh, bold=True),
                _p("STLY",     size=6, color=_sh, bold=True),
                _p("VARIANCE", size=6, color=_sh, bold=True),
                _p("CURRENT",  size=6, color=_sh, bold=True),
                _p("STLY",     size=6, color=_sh, bold=True),
                _p("VARIANCE", size=6, color=_sh, bold=True),
                _p("CURRENT",  size=6, color=_sh, bold=True),
                _p("STLY",     size=6, color=_sh, bold=True),
                _p("VARIANCE", size=6, color=_sh, bold=True),
            ]]

            all_rows  = grp_hdr + sub_hdr
            data_start = len(all_rows)

            style_cmds = [
                # Group header spans + backgrounds — match SRP Pace tab colors
                ("SPAN",(1,0),(3,0)), ("SPAN",(4,0),(6,0)), ("SPAN",(7,0),(9,0)),
                ("BACKGROUND",(0,0),(0,0),  C["navy"]),        # SEGMENT label — navy
                ("BACKGROUND",(1,0),(3,0),  C["gro_h"]),       # OCCUPANCY — dark green
                ("BACKGROUND",(4,0),(6,0),  C["ps_h"]),        # ADR — rust/orange
                ("BACKGROUND",(7,0),(9,0),  C["dkblue"]),      # REVENUE — brand blue
                ("BACKGROUND",(0,1),(-1,1), C["midblue"]),     # sub-header row — light panel
                # All cells centered; segment col left-aligned
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("ALIGN",(0,2),(0,-1),"LEFT"),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("FONTSIZE",(0,0),(-1,-1),6.5),
                # Tight padding to keep table compact
                ("TOPPADDING",(0,0),(-1,-1),2),
                ("BOTTOMPADDING",(0,0),(-1,-1),2),
                ("LEFTPADDING",(0,0),(-1,-1),3),
                ("RIGHTPADDING",(0,0),(-1,-1),3),
                ("LEFTPADDING",(0,2),(0,-1),5),
                ("GRID",(0,0),(-1,-1),0.4,BDR),
            ]

            for i, seg in enumerate(SEGS_ORDER):
                v  = seg_agg[seg]
                rb = C["row1"] if i%2==0 else C["row2"]
                all_rows.append([
                    _dot(seg),
                    _sv(v["otb"]),       _sv(v["stly_otb"]),  _var(v["var_otb"]),
                    _sv(v["adr"],"adr"), _sv(v["stly_adr"],"adr"), _var(v["var_adr"],"adr"),
                    _sv(v["rev"],"rev"), _sv(v["stly_rev"],"rev"), _var(v["var_rev"],"rev"),
                ])
                ri = data_start + i
                style_cmds.append(("BACKGROUND",(0,ri),(-1,ri), rb))

            tot_ri = len(all_rows)
            all_rows.append([
                _p("TOTAL", size=6.5, color=C["offwhite"], bold=True),
                _sv(tot_otb),       _sv(tot_stly_otb),  _var(tot_otb-tot_stly_otb),
                _sv(tot_adr,"adr"), _sv(tot_stly_adr,"adr"),
                _var((tot_adr-tot_stly_adr) if (tot_adr and tot_stly_adr) else None,"adr"),
                _sv(tot_rev,"rev"), _sv(tot_stly_rev,"rev"), _var(tot_rev-tot_stly_rev,"rev"),
            ])
            style_cmds += [
                ("BACKGROUND",(0,tot_ri),(-1,tot_ri), C["rowtot"]),
                ("LINEABOVE",(0,tot_ri),(-1,tot_ri), 1.0, C["teal"]),
            ]

            tbl = Table(all_rows, colWidths=SW, repeatRows=2)
            tbl.setStyle(TableStyle(style_cmds))

            hdr_bar = _hdr_bar(
                f"SRP PACE BY SEGMENT  ·  PAST 14 DAYS  ·  {date_range}",
                colors.HexColor("#0a2030"), left=True
            )
            return hdr_bar, tbl

        _srp_hdr, _srp_tbl = _build_srp_pace_tbl()
        if _srp_hdr and _srp_tbl:
            story.append(Spacer(1, 8))
            story.append(_srp_hdr)
            story.append(_srp_tbl)

        # ═══════════════ PAGE 2 ══════════════════════════════════════════════════
        story.append(PageBreak())

        # Page 2 header
        p2hdr = Table([[
            _p("CALL RECAP  ·  GROUPS & MEETING NOTES", size=9, color=C["white"], bold=True, align=TA_LEFT),
            _p(f"{hotel_name}  ·  {gen_ts}", size=7, color=C["dim"], align=TA_RIGHT),
        ]], colWidths=[TW*0.55, TW*0.45])
        p2hdr.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), C["navy"]),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),10),
            ("RIGHTPADDING",(0,0),(-1,-1),8),
            ("TOPPADDING",(0,0),(-1,-1),7),
            ("BOTTOMPADDING",(0,0),(-1,-1),7),
            ("LINEBELOW",(0,0),(-1,-1),2,C["teal"]),
        ]))
        story.append(p2hdr)
        story.append(Spacer(1, 6))

        # ── Groups table builder ──────────────────────────────────────────────────
        GCOLS = ["Group_Name","Arrival","Departure","Cutoff","Sales_Manager",
                 "Block","Pickup","Avail_Block","Pickup_Pct","Rate"]
        GHDRS = ["GROUP NAME","ARRIVAL","DEPARTURE","CUTOFF","SALES MGR",
                 "BLOCK","PICKUP","AVAIL","PCT","RATE"]
        # Widths measured to fit all text
        GW    = [313, 46, 54, 42, 103, 38, 40, 32, 36, 30]  # sums to TW; all headers clear
        # GW already sums to TW — no adjustment needed

        def _build_group_tbl(df_in, hdr_bg, risk_mode=False):
            if df_in is None or df_in.empty: return None
            disp = [c for c in GCOLS if c in df_in.columns]
            col_w = [GW[GCOLS.index(c)] for c in disp]
            hdr_labels = [GHDRS[GCOLS.index(c)] for c in disp]

            # header
            hdr_row = [_p(h, size=6, color=C["white"], bold=True) for h in hdr_labels]
            rows = [hdr_row]
            style = [
                ("BACKGROUND",(0,0),(-1,0), hdr_bg),
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("ALIGN",(0,1),(0,-1),"LEFT"),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("FONTSIZE",(0,0),(-1,-1),6.5),
                ("TOPPADDING",(0,0),(-1,-1),3),
                ("BOTTOMPADDING",(0,0),(-1,-1),3),
                ("GRID",(0,0),(-1,-1),0.4,BDR),
                ("LEFTPADDING",(0,1),(0,-1),5),
            ]

            for i, (_, row) in enumerate(df_in.iterrows()):
                if risk_mode:
                    rb = colors.HexColor("#fdeee8") if i%2==0 else colors.HexColor("#fff5f5")
                else:
                    rb = C["row1"] if i%2==0 else C["row2"]
                cells = []
                for c in disp:
                    val = row.get(c)
                    fc = C["pritxt"]
                    bold = False
                    # format
                    if c in ["Arrival","Departure","Cutoff"] and pd.notna(val):
                        v = pd.Timestamp(val).strftime("%b %d")
                    elif c == "Rate" and pd.notna(val):
                        try: v = f"${float(val):,.0f}"
                        except: v = str(val)
                    elif c == "Pickup_Pct" and pd.notna(val):
                        try:
                            pv = float(val); v = f"{pv:.1f}%"
                            fc = C["teal"] if pv>=75 else (C["gold"] if pv>=50 else (C["orange"] if pv>0 else C["dim"]))
                            bold = True
                        except: v = str(val)
                    elif c in ["Block","Pickup","Avail_Block"] and pd.notna(val):
                        try: v = f"{int(float(val)):,}"
                        except: v = str(val)
                        if c == "Avail_Block" and risk_mode:
                            fc = C["red_txt"]; bold = True
                    elif pd.notna(val):
                        v = str(val)
                    else:
                        v = "—"
                    cells.append(_p(v, size=6.5, color=fc, bold=bold,
                                   align=TA_LEFT if c=="Group_Name" else TA_CENTER))
                rows.append(cells)
                style.append(("BACKGROUND",(0,i+1),(-1,i+1), rb))

            t = Table(rows, colWidths=col_w, repeatRows=1)
            t.setStyle(TableStyle(style))
            return t

        def _prep_groups(at_risk=False):
            if groups_df is None or "Arrival" not in groups_df.columns: return None
            active = groups_df[groups_df["Departure"]>=today] if "Departure" in groups_df.columns else groups_df
            sum_c  = [c for c in GCOLS if c in active.columns]
            grp_s  = active.groupby("Group_Name").agg({
                c: "first" if c in ["Arrival","Departure","Cutoff","Sales_Manager","Rate"] else "sum"
                for c in sum_c if c!="Group_Name"
            }).reset_index()
            if "Block" in grp_s.columns and "Pickup" in grp_s.columns:
                grp_s["Pickup_Pct"] = (grp_s["Pickup"]/grp_s["Block"]*100).where(grp_s["Block"]>0).round(1)
                # Recalculate Avail_Block from aggregated totals so it's always consistent
                # with Block/Pickup (raw per-night AU values from IHG sum incorrectly across nights).
                grp_s["Avail_Block"] = (grp_s["Block"] - grp_s["Pickup"]).clip(lower=0).astype(int)
            for dc in ["Arrival","Departure","Cutoff"]:
                if dc in grp_s.columns:
                    grp_s[dc] = pd.to_datetime(grp_s[dc], errors="coerce")

            if at_risk:
                if "Avail_Block" not in grp_s.columns: return None
                if "Cutoff" not in grp_s.columns: return None
                cutoff_dl = pd.Timestamp(today) + pd.Timedelta(days=7)
                mask = (
                    (grp_s["Avail_Block"].fillna(0) > 0) &
                    (grp_s["Arrival"].notna()) &
                    (grp_s["Arrival"] >= pd.Timestamp(today)) &
                    (grp_s["Cutoff"].notna()) &
                    (grp_s["Cutoff"] <= cutoff_dl)
                )
                df = grp_s[mask].sort_values("Arrival")
            else:
                df = grp_s[
                    grp_s["Arrival"].notna() &
                    (grp_s["Arrival"]>=today) &
                    (grp_s["Arrival"]<=today+pd.Timedelta(days=30))
                ].sort_values("Arrival")
            return df if not df.empty else None

        # Upcoming groups
        story.append(_hdr_bar("UPCOMING GROUPS  ·  NEXT 30 DAYS", C["grp_h"], left=True))
        ug_df = _prep_groups(at_risk=False)
        ug_t  = _build_group_tbl(ug_df, C["grp_h"])
        if ug_t:
            story.append(ug_t)
        else:
            story.append(_p("No groups arriving in the next 30 days.", color=C["dim"]))

        story.append(Spacer(1, 6))

        # At-risk groups
        story.append(_hdr_bar("⚠  GROUPS HOLDING ROOMS  ·  LESS THAN 7 DAYS TO CUTOFF",
                               C["risk_h"], left=True))
        ar_df = _prep_groups(at_risk=True)
        ar_t  = _build_group_tbl(ar_df, C["risk_h"], risk_mode=True)
        if ar_t:
            story.append(ar_t)
        else:
            story.append(_p("No groups with available block within 7 days of arrival.", color=C["dim"]))

        story.append(Spacer(1, 6))

        # ── Meeting Notes ─────────────────────────────────────────────────────────
        story.append(_hdr_bar("MEETING NOTES & RECAP", C["meet_h"], left=True))

        import json as _mnj
        _mn_path = _cr_folder / "meeting_notes.json"
        _mn_data = {"month":"","calls":{}}
        try:
            if _mn_path.exists():
                _mnd = _mnj.loads(_mn_path.read_text(encoding="utf-8"))
                if _mnd.get("month") == cur_month.strftime("%Y-%m"):
                    _mn_data = _mnd
        except: pass

        today_str  = today.strftime("%Y-%m-%d")
        call_keys  = [today_str]
        call_labels= ["TODAY'S CALL"]
        date_disps = [today.strftime("%B %d, %Y")]
        TOPICS = ["STR Report","Pace — Current Month",
                  "Revenue by Segment — Past 2 Weeks","Top Companies — Past 2 Weeks"]

        def _coerce(raw):
            if not isinstance(raw,list): raw=[]
            out=[]
            for item in raw:
                if isinstance(item,dict): out.append({"topic":item.get("topic",""),"note":item.get("note","")})
                else: out.append({"topic":"","note":str(item)if item else""})
            while len(out)<4: out.append({"topic":"","note":""})
            return out

        MN_W = [82, 110, TW-82-110]

        # Header row
        mn_hdr = [[
            _p("MEETING DATE", size=6, color=C["white"], bold=True),
            _p("TOPIC",        size=6, color=C["white"], bold=True),
            _p("NOTES",        size=6, color=C["white"], bold=True),
        ]]
        mn_rows  = list(mn_hdr)
        mn_style = [
            ("BACKGROUND",(0,0),(-1,0), C["meet_s"]),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("ALIGN",(2,1),(2,-1),"LEFT"),
            ("ALIGN",(0,1),(0,-1),"CENTER"),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("FONTSIZE",(0,0),(-1,-1),7),
            ("TOPPADDING",(0,0),(-1,-1),3),
            ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("LEFTPADDING",(2,1),(2,-1),6),
            ("GRID",(0,0),(-1,-1),0.4,BDR),
        ]

        # Compute all 4 auto rows live — never rely on frozen JSON
        _live_auto = _compute_auto_rows(data, hotel)

        for ci, (ck, cl, dd) in enumerate(zip(call_keys, call_labels, date_disps)):
            entry     = _mn_data["calls"].get(ck, {})
            auto_rows = list(_live_auto)   # always use live values
            # Merge in any manually saved overrides from JSON (non-empty frozen values win)
            frozen = list(entry.get("auto_rows", [""]*4))
            while len(frozen) < 4: frozen.append("")
            for _i in range(4):
                if frozen[_i] and _i != 2:   # keep live segment always; others: prefer frozen if set
                    auto_rows[_i] = frozen[_i]
            user_notes = _coerce(entry.get("user_notes",[]))

            total_topic_rows = 8
            for ri in range(total_topic_rows):
                if ri < 4:
                    # Auto rows 0-3: direct from auto_rows (row 2 always has _live_seg)
                    note  = auto_rows[ri]
                    topic = TOPICS[ri]
                    note_color = C["green"]
                else:
                    # User rows 4-7
                    un    = user_notes[ri-4] if ri-4 < len(user_notes) else {}
                    note  = un.get("note","")
                    topic = un.get("topic","")
                    note_color = C["pritxt"]

                note_style = _style(7, note_color, bold=False, align=TA_LEFT)
                note_para  = Paragraph(note or "", note_style)
                topic_para = _p(topic, size=6.5, color=C["sectxt"], italic=(ri<4))

                bg = C["row1"] if (len(mn_rows))%2==0 else C["row2"]
                row_i = len(mn_rows)

                if ri==0:
                    date_para = _p(f"{cl}\n{dd}", size=7, color=C["offwhite"], bold=True)
                    mn_rows.append([date_para, topic_para, note_para])
                    mn_style += [
                        ("SPAN",(0,row_i),(0,row_i+total_topic_rows-1)),
                        ("BACKGROUND",(0,row_i),(0,row_i+total_topic_rows-1), C["dkblue"]),
                        ("LINEABOVE",(0,row_i),(-1,row_i), 1.0, C["teal"]),
                    ]
                else:
                    mn_rows.append(["", topic_para, note_para])

                mn_style.append(("BACKGROUND",(1,row_i),(2,row_i), bg))

        mn_t = Table(mn_rows, colWidths=MN_W, repeatRows=1)
        mn_t.setStyle(TableStyle(mn_style))
        story.append(mn_t)

        # ── Build PDF ─────────────────────────────────────────────────────────────
        doc = SimpleDocTemplate(
            buf, pagesize=landscape(letter),
            leftMargin=MARGIN, rightMargin=MARGIN,
            topMargin=0.35*inch, bottomMargin=0.30*inch,
        )
        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        buf.seek(0)
        return buf.read()

    return _build_pdf(data, hotel)


def render_call_recap_tab(data, hotel):
    """Revenue Management Call Recap — 3-Month Forecast Review + Pace."""
    import json as _json_cr
    from pathlib import Path as _Path

    st.markdown('<div class="section-head">Revenue Management Call Recap</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Forecast Review + Pace · 3-Month Rolling Window</div>', unsafe_allow_html=True)

    # ── Save Call Recap PDF — single click, direct download ──────────────────
    _pdf_col, _spacer = st.columns([2, 8])
    with _pdf_col:
        _pdf_hotel_id = hotel.get("id","hotel") if isinstance(hotel,dict) else "hotel"
        _pdf_fname    = f"CallRecap_{_pdf_hotel_id}_{pd.Timestamp('today').strftime('%Y%m%d')}.pdf"
        try:
            _pdf_bytes = _generate_call_recap_pdf(data, hotel)
            st.download_button(
                label="📄 Save Call Recap",
                data=_pdf_bytes,
                file_name=_pdf_fname,
                mime="application/pdf",
                key="save_call_recap_pdf",
                use_container_width=True,
            )
        except Exception as _pdf_err:
            st.error(f"PDF generation failed: {_pdf_err}")

    year_df   = data.get("year")
    budget_df = data.get("budget")
    today     = pd.Timestamp(datetime.now().date())
    _lt = False

    # ── Rolling 3 months ─────────────────────────────────────────────────────
    cur_month  = today.replace(day=1)
    months     = [cur_month + pd.DateOffset(months=i) for i in range(3)]
    month_lbls = [m.strftime("%B") for m in months]

    # ── PS Forecast persistence ───────────────────────────────────────────────
    _cr_folder    = cfg.get_hotel_folder(hotel) if isinstance(hotel, dict) else _Path(".")
    _cr_path      = _cr_folder / "call_recap_ps.json"
    _cr_month_key = cur_month.strftime("%Y-%m")
    _cr_hotel_id  = hotel.get("id", "hotel") if isinstance(hotel, dict) else "hotel"

    def _cr_load():
        try:
            if _cr_path.exists():
                d = _json_cr.loads(_cr_path.read_text(encoding="utf-8"))
                if d.get("month") == _cr_month_key and d.get("hotel") == _cr_hotel_id:
                    return d.get("values", {}), d.get("w1_submitted_date", "")
        except Exception:
            pass
        return {}, ""

    def _cr_save(vals, w1_date=""):
        try:
            _cr_path.write_text(_json_cr.dumps(
                {"month": _cr_month_key, "hotel": _cr_hotel_id,
                 "values": vals, "w1_submitted_date": w1_date},
                ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass

    _ps_vals, _w1_submitted_date = _cr_load()

    # ── Parse helper (must be defined before _w1_complete) ───────────────────
    def _parse(s):
        try: return float(str(s).replace(",","").replace("$","").strip())
        except: return None

    # ── Week 2 unlock logic ───────────────────────────────────────────────────
    # Week 2 is editable only after the day AFTER Week 1 was fully submitted.
    # "Fully submitted" = all 3 months have both Revenue and Nights entered for w1.
    def _w1_complete():
        for m in months:
            m_str = m.strftime("%Y-%m")
            if not _parse(_ps_vals.get(f"w1_{m_str}_rev", "")): return False
            if not _parse(_ps_vals.get(f"w1_{m_str}_nts", "")): return False
        return True

    _w1_done     = _w1_complete()
    _today_str   = today.strftime("%Y-%m-%d")

    # Record submission date when w1 first becomes complete
    if _w1_done and not _w1_submitted_date:
        _w1_submitted_date = _today_str
        _cr_save(_ps_vals, _w1_submitted_date)

    # Week 2 unlocks the day AFTER w1 was submitted
    _w2_unlocked = False
    if _w1_submitted_date:
        try:
            sub_dt = pd.Timestamp(_w1_submitted_date).date()
            _w2_unlocked = (today.date() > sub_dt)
        except Exception:
            pass

    # ── Data helpers ─────────────────────────────────────────────────────────
    def _gro_for_month(m):
        if year_df is None: return None, None, None
        mask = (year_df["Date"].dt.year == m.year) & (year_df["Date"].dt.month == m.month)
        sub  = year_df[mask]
        if sub.empty: return None, None, None
        nts = sub["Forecast_Rooms"].sum() if "Forecast_Rooms" in sub.columns else None
        # Hilton: Revenue_Forecast covers all days fully
        if "Revenue_Forecast" in sub.columns and sub["Revenue_Forecast"].notna().any():
            rev = sub["Revenue_Forecast"].sum()
        # IHG: Forecast_Rev is future-days only. Blend past Revenue_OTB + future Forecast_Rev
        # so the current month includes both what's already on the books AND remaining forecast.
        elif "Forecast_Rev" in sub.columns and sub["Forecast_Rev"].notna().any():
            fut_rev  = float(sub["Forecast_Rev"].fillna(0).sum())
            past_rev = float(sub.loc[sub["Forecast_Rev"].isna(), "Revenue_OTB"].fillna(0).sum()) \
                       if "Revenue_OTB" in sub.columns else 0.0
            rev = fut_rev + past_rev
        else:
            rev = None
        adr = (rev / nts) if (nts and nts > 0 and rev) else None
        return (round(rev,0) if rev is not None else None,
                round(nts,0) if nts is not None else None,
                round(adr,2) if adr is not None else None)

    def _bud_for_month(m):
        if budget_df is None: return None, None, None
        mask = (budget_df["Date"].dt.year == m.year) & (budget_df["Date"].dt.month == m.month)
        sub  = budget_df[mask]
        if sub.empty: return None, None, None
        nts = sub["Occ_Rooms"].sum()
        rev = sub["Revenue"].sum()
        adr = (rev / nts) if nts and nts > 0 else None
        return round(rev,0), round(nts,0), (round(adr,2) if adr else None)

    def _fmt_rev(v): return f"${int(round(v)):,}" if v is not None else "—"
    def _fmt_nts(v): return f"{int(round(v)):,}"  if v is not None else "—"
    def _fmt_adr(v): return f"${v:,.2f}"           if v is not None else "—"

    def _diff_fmt(v, is_adr=False):
        if v is None: return "—", _r1_col
        col = _pos_col if v >= 0 else _neg_col
        s   = (f"+${abs(v):,.2f}" if v >= 0 else f"-${abs(v):,.2f}") if is_adr \
              else (f"+${int(abs(v)):,}" if v >= 0 else f"-${int(abs(v)):,}")
        return s, col

    def _nts_diff_fmt(v):
        if v is None: return "—", _r1_col
        return ((f"+{int(abs(v)):,}" if v >= 0 else f"-{int(abs(v)):,}"),
                (_pos_col if v >= 0 else _neg_col))

    # ── Theme tokens ─────────────────────────────────────────────────────────
    _hdr_bg   = "#314B63"          # app main navy — all section title bars
    _hdr_col  = "#ffffff"          # white text on navy
    _week_bg  = "#dde8f5"
    _week_col = "#a0c8e8"
    _r1_bg    = "#f5f8f9"
    _r1_col   = "#1e2d35"
    _r2_bg    = "#eaeff2"
    _r2_col   = "#3a5260"
    _tot_bg   = "#314B63"
    _tot_col  = "#e0eaf8"
    _bdr      = "#c4cfd4"
    _pos_col  = "#2E618D"
    _neg_col  = "#b44820"
    _gro_hdr  = "#2E618D"          # slate blue — System Forecast (was olive #556848)
    _gro_sub  = "#d0e4f5"          # light blue sub-header text
    _ps_hdr   = "#3a6b3a"          # clean forest green — PS Forecast (was dark brown #3a2800)
    _ps_sub   = "#c8eac8"          # light green sub-header text
    _bud_hdr  = "#4e6878"          # medium slate — Budget (was pale gray #dce5e8)
    _diff_hdr = "#8B3A2A"          # controlled terracotta — Difference (was heavy maroon #a03020)
    _diff_sub = "#f5c8be"          # soft salmon sub-header text
    _diff_bg  = _r1_bg
    _inp_bg   = _r1_bg
    _inp_col  = "#3a6b3a"          # matches new PS green
    _inp_bdr  = "#3a6b3a"          # matches new PS green (was brown #5a3a00)
    _note_bg  = "#f5f8f9"
    _note_col = "#1e2d35"
    _lock_bg  = _r1_bg
    _lock_col = "#4a6070"

    # ── CSS: scoped to call-recap inputs only ────────────────────────────────
    st.markdown(f"""<style>
    /* ── PS input styling — scoped via .cr-inp class on wrapper ── */
    .cr-inp div[data-testid="stTextInput"] input {{
        text-align: center !important;
        font-family: DM Mono, monospace !important;
        font-size: 13px !important;
        background: {_inp_bg} !important;
        color: {_inp_col} !important;
        border: 1px solid {_inp_bdr} !important;
        border-radius: 2px !important;
        padding: 2px 4px !important;
        height: 34px !important;
    }}
    /* Shrink label height to zero inside PS input wrappers */
    .cr-inp div[data-testid="stTextInput"] label {{
        display: none !important;
    }}
    .cr-inp div[data-testid="stTextInput"] {{
        margin-bottom: 0 !important;
        padding-bottom: 0 !important;
    }}
    /* Row gaps — only inside the call-recap left panel */
    .cr-rows div[data-testid="stHorizontalBlock"] {{
        gap: 2px !important;
        margin-bottom: 0px !important;
        padding-bottom: 0px !important;
        margin-top: 0px !important;
    }}
    .cr-rows div[data-testid="element-container"] {{
        margin-bottom: 0px !important;
        padding-bottom: 0px !important;
        margin-top: 0px !important;
        padding-top: 0px !important;
    }}
    /* Notes textarea */
    .cr-notes div[data-testid="stTextArea"] textarea {{
        background: {_note_bg} !important;
        color: {_note_col} !important;
        border: 1px solid {_bdr} !important;
        border-radius: 3px !important;
        font-family: DM Sans, sans-serif !important;
        font-size: 12px !important;
        resize: none !important;
    }}
    .cr-notes div[data-testid="stTextArea"] label {{
        display: none !important;
    }}
    </style>""", unsafe_allow_html=True)

    # ── Column width ratios ───────────────────────────────────────────────────
    # [label | gro×3 | ps×3 | bud×3 | diff×3 | notes]
    CW = [1.2, 1.0,1.0,1.0,  1.1,1.1,0.95,  1.0,1.0,1.0,  1.0,1.0,1.0,  1.9]

    # ── HTML cell helpers ────────────────────────────────────────────────────
    def _div(val, bg, col, fw="400", fs="13px", ff="DM Mono,monospace",
             h="36px", align="center", pad="6px 6px", extra=""):
        return (f'<div style="background:{bg};color:{col};font-weight:{fw};font-size:{fs};'
                f'font-family:{ff};padding:{pad};text-align:{align};border:1px solid {_bdr};'
                f'height:{h};display:flex;align-items:center;justify-content:'
                f'{"center" if align=="center" else "flex-start"};{extra}">{val}</div>')

    def _cell(col_obj, val, bg, fc, bold=False, fs="13px", ff="DM Mono,monospace",
              align="center", h="36px", pad="6px 6px", extra=""):
        col_obj.markdown(_div(val, bg, fc, "700" if bold else "400", fs, ff, h, align,
                              pad=pad, extra=extra), unsafe_allow_html=True)

    def _blank(col_obj, bg=None, h="36px"):
        col_obj.markdown(_div("", bg or _lock_bg, _lock_col, h=h),
                         unsafe_allow_html=True)

    # ── Merged header (HTML table for real colspan) ───────────────────────────
    # ── Shared column widths (px) — used by header AND all week tables ──────
    # One colgroup definition governs everything so header always aligns with data.
    W = [100, 82,78,72,  88,80,72,  82,78,72,  82,78,72,  160]
    # cols: label | gro×3 | ps×3 | bud×3 | diff×3 | notes

    COLGROUP   = "".join(f'<col style="width:{w}px">' for w in W)
    _TBL_TOTAL = sum(W)  # exact pixel total — no browser guessing
    TBL_STYLE  = (f'border-collapse:collapse;width:100%;min-width:{_TBL_TOTAL}px;table-layout:fixed;'
                  f'font-family:DM Sans,sans-serif;margin:0;')

    # ── Cell builders ────────────────────────────────────────────────────────
    def _th(txt, bg, fc, colspan=1, rowspan=1, fs="10px", pad="7px 4px"):
        cs = f' colspan="{colspan}"' if colspan > 1 else ""
        rs = f' rowspan="{rowspan}"' if rowspan > 1 else ""
        return (f'<th{cs}{rs} style="background:{bg};color:{fc};font-weight:700;'
                f'font-size:{fs};letter-spacing:.06em;text-transform:uppercase;'
                f'text-align:center;padding:{pad};border:1px solid {_bdr};'
                f'white-space:nowrap;overflow:hidden;">{txt}</th>')

    def _td(val, bg, fc, fw="400", align="center", pad="7px 6px"):
        return (f'<td style="background:{bg};color:{fc};font-weight:{fw};'
                f'font-size:13px;font-family:DM Mono,monospace;padding:{pad};'
                f'text-align:{align};border:1px solid {_bdr};'
                f'white-space:nowrap;overflow:hidden;">{val}</td>')

    def _tdl(val, bg, fc, fw="600"):
        return (f'<td style="background:{bg};color:{fc};font-weight:{fw};'
                f'font-size:12px;font-family:DM Sans,sans-serif;padding:7px 10px;'
                f'text-align:left;border:1px solid {_bdr};'
                f'white-space:nowrap;overflow:hidden;">{val}</td>')

    # ── Header rows (2 rows, shared table) ───────────────────────────────────
    hdr_row1 = (
        _th("GUEST ROOMS",    _hdr_bg,   _hdr_col,  rowspan=2, pad="8px 4px") +
        _th("SYSTEM FORECAST", _gro_hdr,  "#ffffff",  colspan=3) +
        _th("PS FORECAST",    _ps_hdr,   _ps_sub,   colspan=3) +
        _th("BUDGET",         _bud_hdr,  _hdr_col,  colspan=3) +
        _th("DIFFERENCE TO PS FORECAST",  _diff_hdr, _diff_sub, colspan=3) +
        _th("FORECAST NOTES", _hdr_bg,   _hdr_col,  rowspan=2, pad="8px 4px")
    )
    def _th2(txt, bg, fc):
        return _th(txt, bg, fc, fs="9px", pad="5px 4px")

    hdr_row2 = (
        _th2("REVENUE", _gro_hdr, _gro_sub) + _th2("NIGHTS", _gro_hdr, _gro_sub) + _th2("ADR", _gro_hdr, _gro_sub) +
        _th2("REVENUE", _ps_hdr,  _ps_sub)  + _th2("NIGHTS", _ps_hdr,  _ps_sub)  + _th2("ADR", _ps_hdr,  _ps_sub)  +
        _th2("REVENUE", _bud_hdr, _hdr_col) + _th2("NIGHTS", _bud_hdr, _hdr_col) + _th2("ADR", _bud_hdr, _hdr_col) +
        _th2("REVENUE", _diff_hdr,_diff_sub)+ _th2("NIGHTS", _diff_hdr,_diff_sub)+ _th2("ADR", _diff_hdr,_diff_sub)
    )

    # ── Week data builder — returns <tbody> rows as string ───────────────────
    def _week_rows(week_num, locked):
        wk_key   = f"w{week_num}"
        note_key = f"{wk_key}_note"

        rows_data = []
        tot = {k: 0.0 for k in ["gro_rev","gro_nts","bud_rev","bud_nts","ps_rev","ps_nts"]}
        tv  = {k: False for k in tot}

        for i, (m, lbl) in enumerate(zip(months, month_lbls)):
            gro_rev, gro_nts, gro_adr = _gro_for_month(m)
            bud_rev, bud_nts, bud_adr = _bud_for_month(m)
            m_str      = m.strftime("%Y-%m")
            cur_rev    = _ps_vals.get(f"{wk_key}_{m_str}_rev", "")
            cur_nts    = _ps_vals.get(f"{wk_key}_{m_str}_nts", "")
            ps_rev     = _parse(cur_rev)
            ps_nts     = _parse(cur_nts)
            ps_adr     = (ps_rev / ps_nts) if (ps_rev and ps_nts and ps_nts > 0) else None
            def _d(a, b): return round(a-b, 2) if (a is not None and b is not None) else None
            diff_rev = _d(ps_rev, bud_rev)
            diff_nts = _d(ps_nts, bud_nts)
            diff_adr = _d(ps_adr, bud_adr)
            for k, v in [("gro_rev",gro_rev),("gro_nts",gro_nts),
                          ("bud_rev",bud_rev),("bud_nts",bud_nts),
                          ("ps_rev",ps_rev),  ("ps_nts",ps_nts)]:
                if v is not None: tot[k] += v; tv[k] = True
            rows_data.append(dict(
                lbl=lbl, i=i,
                gro_rev=gro_rev, gro_nts=gro_nts, gro_adr=gro_adr,
                bud_rev=bud_rev, bud_nts=bud_nts, bud_adr=bud_adr,
                ps_rev=ps_rev, ps_nts=ps_nts, ps_adr=ps_adr,
                diff_rev=diff_rev, diff_nts=diff_nts, diff_adr=diff_adr,
                cur_rev=cur_rev, cur_nts=cur_nts,
                m_str=m_str,
                ps_rev_key=f"{wk_key}_{m_str}_rev",
                ps_nts_key=f"{wk_key}_{m_str}_nts",
            ))

        tga = (tot["gro_rev"]/tot["gro_nts"]) if tv["gro_nts"] and tot["gro_nts"]>0 else None
        tba = (tot["bud_rev"]/tot["bud_nts"]) if tv["bud_nts"] and tot["bud_nts"]>0 else None
        tpa = (tot["ps_rev"] /tot["ps_nts"])  if tv["ps_nts"]  and tot["ps_nts"] >0 else None
        tdr,tdrc = _diff_fmt((tot["ps_rev"]-tot["bud_rev"]) if (tv["ps_rev"] and tv["bud_rev"]) else None)
        tdn,tdnc = _nts_diff_fmt((tot["ps_nts"]-tot["bud_nts"]) if (tv["ps_nts"] and tv["bud_nts"]) else None)
        tda,tdac = _diff_fmt((tpa-tba) if (tpa and tba) else None, is_adr=True)

        wk_label = f"WEEK {week_num}" + ("  ·  🔒 Unlocks the day after Week 1 is complete" if locked else "")
        wk_bg    = _lock_bg if locked else _week_bg
        wk_fc    = _lock_col if locked else _week_col

        # Week label row — spans all 14 cols
        out = (f'<tr><td colspan="14" style="background:{wk_bg};color:{wk_fc};'
               f'font-weight:700;font-size:11px;letter-spacing:.1em;'
               f'text-transform:uppercase;padding:8px 12px;'
               f'border:1px solid {_bdr};">{wk_label}</td></tr>')

        # Notes cell: rowspan = 3 month rows + 1 total = 4
        cur_note     = _ps_vals.get(note_key, "")
        note_display = cur_note.replace("\n","<br>") if cur_note else ("🔒" if locked else "")
        note_style   = (f"background:{_lock_bg};color:{_lock_col};"
                        if locked else f"background:{_note_bg};color:{_note_col};")
        note_td = (f'<td rowspan="4" style="{note_style}font-size:12px;'
                   f'font-family:DM Sans,sans-serif;padding:8px;'
                   f'border:1px solid {_bdr};vertical-align:top;">'
                   f'{note_display}</td>')

        # Month rows
        for ri, r in enumerate(rows_data):
            rb  = _r1_bg  if r["i"] % 2 == 0 else _r2_bg
            rc  = _r1_col if r["i"] % 2 == 0 else _r2_col
            drv_s,drv_c = _diff_fmt(r["diff_rev"])
            dns_s,dns_c = _nts_diff_fmt(r["diff_nts"])
            dad_s,dad_c = _diff_fmt(r["diff_adr"], is_adr=True)
            ps_r = f"${int(round(r['ps_rev'])):,}" if r["ps_rev"] is not None else ""
            ps_n = f"{int(round(r['ps_nts'])):,}"  if r["ps_nts"] is not None else ""
            ps_a = _fmt_adr(r["ps_adr"])            if r["ps_adr"] is not None else "—"
            if locked:
                ps_cells = _td("",rb,rc) + _td("",rb,rc) + _td("—",rb,rc)
                df_cells = _td("—",rb,rc) + _td("—",rb,rc) + _td("—",rb,rc)
            else:
                ps_cells = _td(ps_r,rb,_inp_col) + _td(ps_n,rb,_inp_col) + _td(ps_a,rb,_inp_col)
                df_cells = (_td(drv_s,rb,drv_c,fw="700") +
                            _td(dns_s,rb,dns_c,fw="700") +
                            _td(dad_s,rb,dad_c,fw="700"))
            row_html = (
                "<tr>"
                + _tdl(r["lbl"], rb, rc)
                + _td(_fmt_rev(r["gro_rev"]),rb,rc)
                + _td(_fmt_nts(r["gro_nts"]),rb,rc)
                + _td(_fmt_adr(r["gro_adr"]),rb,rc)
                + ps_cells
                + _td(_fmt_rev(r["bud_rev"]),rb,rc)
                + _td(_fmt_nts(r["bud_nts"]),rb,rc)
                + _td(_fmt_adr(r["bud_adr"]),rb,rc)
                + df_cells
                + (note_td if ri == 0 else "")  # notes only on first row, rowspan=4
                + "</tr>"
            )
            out += row_html

        # Total row
        if locked:
            ps_tot  = _td("—",_tot_bg,_tot_col)*3
            dif_tot = _td("—",_tot_bg,_tot_col)*3
        else:
            ps_tot  = (_td(_fmt_rev(tot["ps_rev"]  if tv["ps_rev"]  else None),_tot_bg,_tot_col,fw="700") +
                       _td(_fmt_nts(tot["ps_nts"]  if tv["ps_nts"]  else None),_tot_bg,_tot_col,fw="700") +
                       _td(_fmt_adr(tpa),_tot_bg,_tot_col,fw="700"))
            dif_tot = (_td(tdr,_tot_bg,"#ffffff",fw="700") +
                       _td(tdn,_tot_bg,"#ffffff",fw="700") +
                       _td(tda,_tot_bg,"#ffffff",fw="700"))
        out += (
            "<tr>"
            + _tdl("Total",_tot_bg,_tot_col,fw="700")
            + _td(_fmt_rev(tot["gro_rev"] if tv["gro_rev"] else None),_tot_bg,_tot_col,fw="700")
            + _td(_fmt_nts(tot["gro_nts"] if tv["gro_nts"] else None),_tot_bg,_tot_col,fw="700")
            + _td(_fmt_adr(tga),_tot_bg,_tot_col,fw="700")
            + ps_tot
            + _td(_fmt_rev(tot["bud_rev"] if tv["bud_rev"] else None),_tot_bg,_tot_col,fw="700")
            + _td(_fmt_nts(tot["bud_nts"] if tv["bud_nts"] else None),_tot_bg,_tot_col,fw="700")
            + _td(_fmt_adr(tba),_tot_bg,_tot_col,fw="700")
            + dif_tot
            + "</tr>"
        )
        return out, rows_data, wk_key, note_key, cur_note

    # ── PS input expander (below the shared table, per-week) ─────────────────
    def _ps_expander(week_num, rows_data, wk_key, note_key, cur_note):
        _changed = [False]
        with st.expander(f"✏ Edit Week {week_num} PS Forecast", expanded=False):
            input_cols = st.columns(6)
            labels     = [f"{r['lbl'][:3]} Rev" if ri%2==0 else f"{rows_data[ri//2]['lbl'][:3]} Nts"
                          for ri in range(6) for r in [rows_data[ri//2]]]
            labels     = []
            keys_order = []
            for r in rows_data:
                labels.append(f"{r['lbl'][:3]} Rev")
                labels.append(f"{r['lbl'][:3]} Nts")
                keys_order.append((r["ps_rev_key"], r["cur_rev"], "rev"))
                keys_order.append((r["ps_nts_key"], r["cur_nts"], "nts"))

            for col, lbl, (key, cur, typ) in zip(input_cols, labels, keys_order):
                raw  = _parse(cur)
                disp = (f"${int(round(raw)):,}" if typ=="rev" else f"{int(round(raw)):,}") if raw is not None else ""
                new_val   = col.text_input(lbl, value=disp, key=f"cr_form_{key}")
                clean     = new_val.replace("$","").replace(",","").strip()
                old_clean = cur.replace("$","").replace(",","").strip()
                if clean != old_clean:
                    _ps_vals[key] = clean; _changed[0] = True

            new_note = st.text_area("Forecast Notes", value=cur_note,
                placeholder="Forecast commentary...",
                key=f"cr_form_note_{wk_key}", height=80)
            if new_note != cur_note:
                _ps_vals[note_key] = new_note; _changed[0] = True

            if st.button(f"💾 Save Week {week_num}", key=f"cr_save_{wk_key}"):
                w1_done_now = all(
                    _parse(_ps_vals.get(f"w1_{m.strftime('%Y-%m')}_rev","")) and
                    _parse(_ps_vals.get(f"w1_{m.strftime('%Y-%m')}_nts",""))
                    for m in months
                )
                save_date = _w1_submitted_date
                if w1_done_now and not save_date:
                    save_date = _today_str
                _cr_save(_ps_vals, save_date)
                st.rerun()

    # ── 14-day daily recap builder ───────────────────────────────────────────
    def _daily_recap_html():
        if year_df is None:
            return f'<div style="color:{_hdr_col};padding:16px;font-size:12px;">No data available.</div>'

        past14 = year_df[
            (year_df["Date"] >= today - pd.Timedelta(days=14)) &
            (year_df["Date"] < today)
        ].copy().sort_values("Date", ascending=False)

        if past14.empty:
            return f'<div style="color:{_hdr_col};padding:16px;font-size:12px;">No historical data in range.</div>'

        total_rooms = hotel.get("total_rooms", 0) if isinstance(hotel, dict) else 0

        # ── Column header row
        DCOLS = [90, 52, 70, 52, 70]   # DATE, OCC TY, REV TY, OCC LY, REV LY
        _DAILY_W = sum(DCOLS)
        DCOLGROUP = "".join(f'<col style="width:{w}px">' for w in DCOLS)

        def dth(txt, bg, fc, fs="9px"):
            return (f'<th style="background:{bg};color:{fc};font-weight:700;font-size:{fs};'
                    f'letter-spacing:.05em;text-transform:uppercase;text-align:center;'
                    f'padding:5px 6px;border:1px solid {_bdr};white-space:nowrap;">{txt}</th>')

        def dtd(val, bg, fc, fw="400", align="center"):
            return (f'<td style="background:{bg};color:{fc};font-weight:{fw};'
                    f'font-size:12px;font-family:DM Mono,monospace;'
                    f'padding:5px 6px;text-align:{align};border:1px solid {_bdr};'
                    f'white-space:nowrap;">{val}</td>')

        _dr_ty_fc   = "#ffffff"
        _dr_ly_fc   = "#ffffff"
        _dr_dat_fc  = "#ffffff"
        _dr_cell_ty = "#1a6a45" if _lt else "#2E618D"   # blue for TY values
        _dr_cell_ly = "#1a3a5c" if _lt else "#4e6878"   # slate for LY values
        _dr_cell_dt = "#0d1f2d" if _lt else _r1_col

        hdr = (
            dth("DATE",     _hdr_bg,   _dr_dat_fc) +
            dth("RMS TY",   _gro_hdr,  _dr_ty_fc) +
            dth("REV TY",   _gro_hdr,  _dr_ty_fc) +
            dth("RMS LY",   _bud_hdr,  _dr_ly_fc) +
            dth("REV LY",   _bud_hdr,  _dr_ly_fc)
        )

        rows_html = ""
        for i, (_, r) in enumerate(past14.iterrows()):
            rb  = _r1_bg  if i % 2 == 0 else _r2_bg
            rc  = _r1_col if i % 2 == 0 else _r2_col

            _ts    = pd.Timestamp(r["Date"])
            dt_str = f"{_ts.strftime('%a')} · {_ts.month}/{_ts.day}"

            # TY rooms sold
            ty_rooms = None
            for col in ["OTB", "Forecast_Rooms"]:
                if col in r and pd.notna(r.get(col)):
                    ty_rooms = int(r[col]); break
            ty_occ = f"{ty_rooms:,}" if ty_rooms is not None else "—"

            # TY revenue
            ty_rev = None
            for col in ["Revenue_OTB", "Revenue_Forecast"]:
                if col in r and pd.notna(r.get(col)):
                    ty_rev = float(r[col]); break
            ty_rev_s = f"${int(round(ty_rev)):,}" if ty_rev is not None else "—"

            # LY rooms sold
            ly_rooms = r.get("OTB_LY") if "OTB_LY" in r.index and pd.notna(r.get("OTB_LY")) else None
            ly_occ   = f"{int(float(ly_rooms)):,}" if ly_rooms is not None else "—"

            # LY revenue
            ly_rev   = r.get("Revenue_LY") if "Revenue_LY" in r.index and pd.notna(r.get("Revenue_LY")) else None
            ly_rev_s = f"${int(round(float(ly_rev))):,}" if ly_rev is not None else "—"

            rows_html += (
                "<tr>"
                + dtd(dt_str, rb, _dr_cell_dt, fw="600", align="left")
                + dtd(ty_occ, rb, _dr_cell_ty if ty_rooms else rc)
                + dtd(ty_rev_s, rb, _dr_cell_ty if ty_rev else rc)
                + dtd(ly_occ, rb, _dr_cell_ly)
                + dtd(ly_rev_s, rb, _dr_cell_ly)
                + "</tr>"
            )

        return (
            f'<table style="border-collapse:collapse;width:100%;min-width:{_DAILY_W}px;table-layout:fixed;'
            f'font-family:DM Sans,sans-serif;">'
            f'<colgroup>{DCOLGROUP}</colgroup>'
            f'<thead><tr>{hdr}</tr></thead>'
            f'<tbody>{rows_html}</tbody>'
            f'</table>'
        )

    # ── Two-column outer layout: forecast table left, daily recap right ───────
    col_left, col_right = st.columns([_TBL_TOTAL, 334], gap="medium")

    with col_left:
        # Title bar
        month_range = f"{month_lbls[0]} – {month_lbls[2]} {months[0].year}"
        st.markdown(
            f'<div style="overflow-x:auto;width:100%;min-width:{_TBL_TOTAL}px;">'
            f'<div style="background:{_hdr_bg};color:#ffffff;font-weight:700;font-size:13px;'
            f'letter-spacing:.08em;text-transform:uppercase;text-align:center;'
            f'padding:12px 16px;min-height:42px;line-height:1.4;width:100%;'
            f'box-sizing:border-box;border-radius:6px 6px 0 0;margin-bottom:0;display:flex;'
            f'align-items:center;justify-content:center;">'
            f'FORECAST REVIEW + PACE &nbsp;·&nbsp; {month_range}</div></div>',
            unsafe_allow_html=True)

        # Unified table
        w1_rows, w1_data, w1_key, w1_note_key, w1_note = _week_rows(1, locked=False)
        w2_rows, w2_data, w2_key, w2_note_key, w2_note = _week_rows(2, locked=not _w2_unlocked)

        full_table = (
            f'<div style="overflow-x:auto;width:100%;">'
            f'<table style="{TBL_STYLE}">'
            f'<colgroup>{COLGROUP}</colgroup>'
            f'<thead><tr>{hdr_row1}</tr><tr>{hdr_row2}</tr></thead>'
            f'<tbody>{w1_rows}{w2_rows}</tbody>'
            f'</table></div>'
        )
        st.markdown(full_table, unsafe_allow_html=True)

        # PS input expanders
        _ps_expander(1, w1_data, w1_key, w1_note_key, w1_note)
        if _w2_unlocked:
            _ps_expander(2, w2_data, w2_key, w2_note_key, w2_note)

    with col_right:
        st.markdown(
            f'<div style="background:{_hdr_bg};color:#ffffff;font-weight:700;font-size:13px;'
            f'letter-spacing:.08em;text-transform:uppercase;text-align:center;'
            f'padding:12px 8px;min-height:42px;line-height:1.4;'
            f'border-radius:6px 6px 0 0;display:flex;align-items:center;justify-content:center;">'
            f'14-DAY DAILY RECAP</div>',
            unsafe_allow_html=True)
        st.markdown(_daily_recap_html(), unsafe_allow_html=True)

    # ── Full-width Upcoming Groups (next 30 days) ─────────────────────────────
    groups_df = data.get("groups")
    if groups_df is not None and "Arrival" in groups_df.columns and "Group_Name" in groups_df.columns:
        _ug_lt      = _lt
        _ug_hdr_bg  = _hdr_bg
        _ug_txt     = "#0d1f2d" if _ug_lt else "#1e2d35"
        _ug_bdr     = "#b0c0d0" if _ug_lt else "#c4cfd4"
        _ug_hdr_col = "#ffffff"
        _ug_sub_bg  = "#314B63"
        _ug_card    = "#f5f8f9"
        _ug_alt     = "#eaeff2"
        _ug_teal    = "#006a55" if _ug_lt else "#00c49a"
        _ug_gold    = "#7a4e00" if _ug_lt else "#6A924D"
        _ug_orange  = "#cc3300" if _ug_lt else "#b44820"
        _ug_dim     = "#6a8aaa" if _ug_lt else "#6a8090"

        # Summarise to one row per group (same logic as Groups tab)
        active_g = groups_df[groups_df["Departure"] >= today] if "Departure" in groups_df.columns else groups_df
        sum_cols = [c for c in ["Group_Name","Arrival","Departure","Cutoff",
                                "Sales_Manager","Block","Pickup","Avail_Block",
                                "Pickup_Pct","Rate"] if c in active_g.columns]
        grp_s = active_g.groupby("Group_Name").agg({
            c: "first" if c in ["Arrival","Departure","Cutoff","Sales_Manager","Rate"] else "sum"
            for c in sum_cols if c != "Group_Name"
        }).reset_index()

        if "Block" in grp_s.columns and "Pickup" in grp_s.columns:
            grp_s["Pickup_Pct"] = (grp_s["Pickup"] / grp_s["Block"] * 100).where(
                grp_s["Block"] > 0).round(1)

        # Filter to arrivals in next 30 days
        if "Arrival" in grp_s.columns:
            grp_s["Arrival"] = pd.to_datetime(grp_s["Arrival"], errors="coerce")
            upcoming = grp_s[
                grp_s["Arrival"].notna() &
                (grp_s["Arrival"] >= today) &
                (grp_s["Arrival"] <= today + pd.Timedelta(days=30))
            ].sort_values("Arrival")
        else:
            upcoming = pd.DataFrame()

        st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)

        # Title bar
        st.markdown(
            f'<div style="background:{_ug_hdr_bg};color:#ffffff;font-weight:700;font-size:13px;'
            f'letter-spacing:.08em;text-transform:uppercase;text-align:center;'
            f'padding:12px 16px;min-height:42px;line-height:1.4;width:100%;'
            f'box-sizing:border-box;border-radius:6px 6px 0 0;display:flex;'
            f'align-items:center;justify-content:center;">'
            f'UPCOMING GROUPS · NEXT 30 DAYS</div>',
            unsafe_allow_html=True)

        if upcoming.empty:
            st.markdown(
                f'<div style="background:{_ug_card};color:{_ug_dim};font-size:13px;'
                f'padding:16px;text-align:center;border:1px solid {_ug_bdr};'
                f'border-top:none;border-radius:0 0 6px 6px;">'
                f'No groups arriving in the next 30 days.</div>',
                unsafe_allow_html=True)
        else:
            disp_cols = [c for c in ["Group_Name","Arrival","Departure","Cutoff",
                                     "Sales_Manager","Block","Pickup","Avail_Block",
                                     "Pickup_Pct","Rate"] if c in upcoming.columns]
            # Format dates for display
            fmt_df = upcoming[disp_cols].copy()
            for dc in ["Arrival","Departure","Cutoff"]:
                if dc in fmt_df.columns:
                    fmt_df[dc] = pd.to_datetime(fmt_df[dc], errors="coerce").dt.strftime("%b %d, %Y")
            if "Rate" in fmt_df.columns:
                fmt_df["Rate"] = fmt_df["Rate"].apply(
                    lambda x: f"${float(x):,.0f}" if pd.notna(x) and str(x) not in ["","—"] else "—")
            if "Pickup_Pct" in fmt_df.columns:
                fmt_df["Pickup_Pct"] = fmt_df["Pickup_Pct"].apply(
                    lambda x: f"{float(x):.1f}%" if pd.notna(x) and str(x) not in ["","—"] else "—")

            col_labels = {c: c.replace("_"," ").upper() for c in disp_cols}

            # Header row
            hdr_cells = "".join(
                f'<th style="background:{_ug_sub_bg};color:{_ug_hdr_col};font-weight:700;'
                f'font-size:10px;letter-spacing:.1em;text-transform:uppercase;'
                f'padding:7px 10px;text-align:{"left" if c == "Group_Name" else "center"};'
                f'border-bottom:2px solid {_ug_bdr};white-space:nowrap;">'
                f'{col_labels[c]}</th>'
                for c in disp_cols
            )

            # Data rows — bold if Pickup == 0 (rooms not picked up)
            rows_html = ""
            for i, (_, row) in enumerate(fmt_df.iterrows()):
                rb = _ug_card if i % 2 == 0 else _ug_alt
                # Determine if unpicked (Pickup == 0)
                raw_pickup = upcoming.iloc[i].get("Pickup", 1) if i < len(upcoming) else 1
                is_unpicked = (pd.notna(raw_pickup) and float(raw_pickup) == 0)
                fw = "700" if is_unpicked else "400"

                cells = ""
                for c in disp_cols:
                    val = row[c] if pd.notna(row[c]) else "—"
                    align = "left" if c == "Group_Name" else "center"
                    color = _ug_txt

                    if c == "Pickup_Pct":
                        try:
                            pv = float(str(val).replace("%",""))
                            color = _ug_teal if pv >= 75 else (_ug_gold if pv >= 50 else (_ug_orange if pv > 0 else _ug_dim))
                            fw_cell = "700"
                        except: fw_cell = fw
                    else:
                        fw_cell = fw

                    cells += (f'<td style="background:{rb};color:{color};font-weight:{fw_cell};'
                              f'font-size:12px;padding:7px 10px;text-align:{align};'
                              f'border-bottom:1px solid {_ug_bdr};white-space:nowrap;">{val}</td>')
                rows_html += f"<tr>{cells}</tr>"

            tbl = (
                f'<div style="overflow-x:auto;width:100%;">'
                f'<table style="border-collapse:collapse;width:100%;table-layout:auto;'
                f'font-family:DM Sans,sans-serif;">'
                f'<thead><tr>{hdr_cells}</tr></thead>'
                f'<tbody>{rows_html}</tbody>'
                f'</table></div>'
            )
            st.markdown(tbl, unsafe_allow_html=True)

    # ── Groups Holding Rooms < 7 Days to Arrival ──────────────────────────────
    if groups_df is not None and "Arrival" in groups_df.columns and "Group_Name" in groups_df.columns:
        st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)
        st.markdown(
            f'<div style="background:#a03020;color:#ffffff;font-weight:700;font-size:13px;'
            f'letter-spacing:.08em;text-transform:uppercase;text-align:center;'
            f'padding:12px 16px;min-height:42px;line-height:1.4;width:100%;'
            f'box-sizing:border-box;border-radius:6px 6px 0 0;display:flex;'
            f'align-items:center;justify-content:center;">'
            f'⚠ GROUPS HOLDING ROOMS · LESS THAN 7 DAYS TO CUTOFF</div>',
            unsafe_allow_html=True)

        # Reuse the already-summarised grp_s from upcoming groups section;
        # re-build if not in scope (groups_df always is)
        active_g2 = groups_df[groups_df["Departure"] >= today] if "Departure" in groups_df.columns else groups_df
        sum_cols2 = [c for c in ["Group_Name","Arrival","Departure","Cutoff",
                                  "Sales_Manager","Block","Pickup","Avail_Block",
                                  "Pickup_Pct","Rate"] if c in active_g2.columns]
        grp_s2 = active_g2.groupby("Group_Name").agg({
            c: "first" if c in ["Arrival","Departure","Cutoff","Sales_Manager","Rate"] else "sum"
            for c in sum_cols2 if c != "Group_Name"
        }).reset_index()
        if "Block" in grp_s2.columns and "Pickup" in grp_s2.columns:
            grp_s2["Pickup_Pct"] = (grp_s2["Pickup"] / grp_s2["Block"] * 100).where(
                grp_s2["Block"] > 0).round(1)
            # Recalculate Avail_Block from aggregated totals — raw per-night AU values
            # from IHG sum incorrectly across nights and produce false availability.
            grp_s2["Avail_Block"] = (grp_s2["Block"] - grp_s2["Pickup"]).clip(lower=0).astype(int)

        # Convert dates
        for dc in ["Arrival","Departure","Cutoff"]:
            if dc in grp_s2.columns:
                grp_s2[dc] = pd.to_datetime(grp_s2[dc], errors="coerce")

        # Filter: Avail_Block > 0, Arrival within 7 days, Cutoff within 7 days of Arrival
        if "Avail_Block" in grp_s2.columns and "Cutoff" in grp_s2.columns:
            at_risk = grp_s2[
                grp_s2["Avail_Block"].fillna(0) > 0
            ].copy()
            if "Arrival" in at_risk.columns:
                at_risk = at_risk[at_risk["Arrival"].notna()]
                # Arrival must be in the future, cutoff within 7 days of TODAY
                cutoff_deadline = pd.Timestamp(today) + pd.Timedelta(days=7)
                at_risk = at_risk[
                    (at_risk["Arrival"] >= pd.Timestamp(today)) &
                    (at_risk["Cutoff"].notna()) &
                    (at_risk["Cutoff"] <= cutoff_deadline)
                ]
                at_risk = at_risk.sort_values("Arrival")
        else:
            at_risk = pd.DataFrame()

        _ar_card  = "#fdf0f0"
        _ar_alt   = "#fae4e4"
        _ar_bdr   = "#d09090"
        _ar_hdr_bg = "#a03020"
        _ar_warn  = "#cc2200" if _ug_lt else "#ff6b6b"

        if at_risk.empty:
            st.markdown(
                f'<div style="background:{_ar_card};color:{_ug_dim};font-size:13px;'
                f'padding:16px;text-align:center;border:1px solid {_ar_bdr};'
                f'border-top:none;border-radius:0 0 6px 6px;">'
                f'No groups with available block within 7 days of arrival.</div>',
                unsafe_allow_html=True)
        else:
            disp_cols2 = [c for c in ["Group_Name","Arrival","Departure","Cutoff",
                                       "Sales_Manager","Block","Pickup","Avail_Block",
                                       "Pickup_Pct","Rate"] if c in at_risk.columns]
            fmt_ar = at_risk[disp_cols2].copy()
            for dc in ["Arrival","Departure","Cutoff"]:
                if dc in fmt_ar.columns:
                    fmt_ar[dc] = pd.to_datetime(fmt_ar[dc], errors="coerce").dt.strftime("%b %d, %Y")
            if "Rate" in fmt_ar.columns:
                fmt_ar["Rate"] = fmt_ar["Rate"].apply(
                    lambda x: f"${float(x):,.0f}" if pd.notna(x) and str(x) not in ["","—"] else "—")
            if "Pickup_Pct" in fmt_ar.columns:
                fmt_ar["Pickup_Pct"] = fmt_ar["Pickup_Pct"].apply(
                    lambda x: f"{float(x):.1f}%" if pd.notna(x) and str(x) not in ["","—"] else "—")

            col_labels2 = {c: c.replace("_"," ").upper() for c in disp_cols2}
            hdr_cells2 = "".join(
                f'<th style="background:{_ar_hdr_bg};color:#ffffff;font-weight:700;'
                f'font-size:10px;letter-spacing:.1em;text-transform:uppercase;'
                f'padding:7px 10px;text-align:{"left" if c == "Group_Name" else "center"};'
                f'border-bottom:2px solid {_ar_bdr};white-space:nowrap;">'
                f'{col_labels2[c]}</th>'
                for c in disp_cols2
            )
            rows_ar = ""
            for i, (_, row) in enumerate(fmt_ar.iterrows()):
                rb = _ar_card if i % 2 == 0 else _ar_alt
                raw_avail = at_risk.iloc[i].get("Avail_Block", 0) if i < len(at_risk) else 0
                cells = ""
                for c in disp_cols2:
                    val   = row[c] if pd.notna(row[c]) else "—"
                    align = "left" if c == "Group_Name" else "center"
                    color = _ug_txt
                    fw    = "400"
                    if c == "Avail_Block":
                        color = _ar_warn; fw = "700"
                    elif c == "Pickup_Pct":
                        try:
                            pv = float(str(val).replace("%",""))
                            color = _ug_teal if pv >= 75 else (_ug_gold if pv >= 50 else (_ug_orange if pv > 0 else _ug_dim))
                            fw = "700"
                        except: pass
                    cells += (f'<td style="background:{rb};color:{color};font-weight:{fw};'
                              f'font-size:12px;padding:7px 10px;text-align:{align};'
                              f'border-bottom:1px solid {_ar_bdr};white-space:nowrap;">{val}</td>')
                rows_ar += f"<tr>{cells}</tr>"

            tbl_ar = (
                f'<div style="overflow-x:auto;width:100%;">'
                f'<table style="border-collapse:collapse;width:100%;table-layout:auto;'
                f'font-family:DM Sans,sans-serif;">'
                f'<thead><tr>{hdr_cells2}</tr></thead>'
                f'<tbody>{rows_ar}</tbody>'
                f'</table></div>'
            )
            st.markdown(tbl_ar, unsafe_allow_html=True)

    # ── Meeting Notes & Recap ─────────────────────────────────────────────────
    import json as _mn_json

    _mn_path = _cr_folder / "meeting_notes.json"
    _today_str = today.strftime("%Y-%m-%d")
    _this_month = today.strftime("%Y-%m")

    def _mn_load():
        if _mn_path.exists():
            try:
                d = _mn_json.loads(_mn_path.read_text(encoding="utf-8"))
                if d.get("month") != _this_month:
                    return {"month": _this_month, "calls": {}}
                return d
            except Exception:
                pass
        return {"month": _this_month, "calls": {}}

    def _mn_save(d):
        try:
            _mn_path.write_text(_mn_json.dumps(d, indent=2), encoding="utf-8")
        except Exception:
            pass

    _mn_data = _mn_load()

    # ── Build auto-populated row data ─────────────────────────────────────────
    hotel_name = hotel.get("name", "Hotel") if isinstance(hotel, dict) else "Hotel"
    short_name = hotel.get("short_name", hotel_name) if isinstance(hotel, dict) else hotel_name

    def _auto_rows(data):
        """Delegate to module-level helper so PDF and tab always match."""
        return _compute_auto_rows(data, hotel)

    # ── Determine call dates — Today's Call only ─────────────────────────────
    call_keys     = [today.strftime("%Y-%m-%d")]
    call_labels   = ["TODAY'S CALL"]
    date_displays = [today.strftime("%B %d, %Y")]

    TOPICS = [
        "STR Report",
        "Pace - Current Month",
        "Revenue By Segment - Past 2 Weeks",
        "Top Companies - Past 2 Weeks",
    ]

    # ── Freeze / persist today's auto rows ────────────────────────────────────
    c0_entry  = _mn_data["calls"].get(call_keys[0], {})
    frozen_dt = c0_entry.get("frozen_date", "")
    c0_frozen = frozen_dt and frozen_dt < _today_str
    if c0_frozen:
        auto_vals_c0 = c0_entry.get("auto_rows", ["", "", "", ""])
    else:
        auto_vals_c0 = _auto_rows(data)
        c0_entry["auto_rows"]   = auto_vals_c0
        c0_entry["frozen_date"] = _today_str
        _mn_data["calls"][call_keys[0]] = c0_entry
        _mn_save(_mn_data)

    user_notes_c0 = c0_entry.get("user_notes", [])   # list of {"topic":…,"note":…}
    blank_note = {"topic": "", "note": ""}
    def _coerce_notes(raw):
        """Migrate old plain-string notes or non-list to list of dicts."""
        if not isinstance(raw, list): raw = []
        out = []
        for item in raw:
            if isinstance(item, dict):
                out.append({"topic": item.get("topic","") or "", "note": item.get("note","") or ""})
            else:
                out.append({"topic": "", "note": str(item) if item else ""})
        while len(out) < 4: out.append(dict(blank_note))
        return out
    user_notes_c0 = _coerce_notes(user_notes_c0)

    # ── Theme tokens ──────────────────────────────────────────────────────────
    _mn_lt       = _lt
    _mn_hdr_bg   = "#314B63"           # app navy — was green #6A924D
    _mn_sub_bg   = "#3a5a70"           # slightly lighter navy for column headers
    _mn_txt      = "#0d1f2d" if _mn_lt else "#1e2d35"
    _mn_bdr      = "#b0c0d0" if _mn_lt else "#c4cfd4"
    _mn_card     = "#f5f8f9"
    _mn_alt      = "#eaeff2"
    _mn_auto_fc  = "#1a4a35" if _mn_lt else "#1e2d35"   # dark readable text — was bright teal #7addb8
    _mn_user_fc  = "#0d1f2d" if _mn_lt else "#1e2d35"
    _mn_date_bg  = "#dce8f4" if _mn_lt else "#e4eaed"
    _mn_date_fc  = "#0d1f2d" if _mn_lt else "#3a5260"
    _mn_topic_fc = "#1a3a5c" if _mn_lt else "#2E618D"   # blue — unchanged, readable
    _mn_dim      = "#6a8aaa" if _mn_lt else "#5a7080"

    st.markdown("<div style='margin-top:24px'></div>", unsafe_allow_html=True)

    st.markdown(
        f'<div style="background:{_mn_hdr_bg};color:#ffffff;font-weight:700;font-size:13px;'
        f'letter-spacing:.08em;text-transform:uppercase;'
        f'padding:10px 16px;border-radius:6px 6px 0 0;">'
        f'MEETING NOTES &amp; RECAP</div>',
        unsafe_allow_html=True)

    col_hdr = (
        f'<tr>'
        f'<th style="background:{_mn_sub_bg};color:#ffffff;font-weight:700;font-size:10px;'
        f'letter-spacing:.1em;text-transform:uppercase;padding:7px 12px;text-align:center;'
        f'border:1px solid {_mn_bdr};width:130px;">MEETING DATE</th>'
        f'<th style="background:{_mn_sub_bg};color:#ffffff;font-weight:700;font-size:10px;'
        f'letter-spacing:.1em;text-transform:uppercase;padding:7px 12px;text-align:center;'
        f'border:1px solid {_mn_bdr};width:220px;">TOPIC</th>'
        f'<th style="background:{_mn_sub_bg};color:#ffffff;font-weight:700;font-size:10px;'
        f'letter-spacing:.1em;text-transform:uppercase;padding:7px 12px;text-align:center;'
        f'border:1px solid {_mn_bdr};">NOTES</th>'
        f'</tr>'
    )

    def _build_call_rows(ci, call_key, call_label, date_disp, auto_vals, user_notes_list):
        """Render one full call block (4 auto + 4 user rows)."""
        rows_html = ""
        total_rows = 8  # rowspan for date cell

        for ri in range(4):
            bg    = _mn_card if ri % 2 == 0 else _mn_alt
            note  = auto_vals[ri] if (auto_vals and ri < len(auto_vals)) else ""
            topic = TOPICS[ri]
            if ri == 0:
                date_cell = (f'<td rowspan="{total_rows}" style="background:{_mn_date_bg};'
                             f'color:{_mn_date_fc};font-weight:700;font-size:12px;'
                             f'padding:10px 8px;text-align:center;vertical-align:top;'
                             f'border:1px solid {_mn_bdr};width:130px;">'
                             f'<div style="font-size:10px;text-transform:uppercase;'
                             f'letter-spacing:.06em;margin-bottom:4px;">{call_label}</div>'
                             f'<div style="font-size:11px;">{date_disp}</div></td>')
            else:
                date_cell = ""
            rows_html += (
                f'<tr>{date_cell}'
                f'<td style="background:{bg};color:{_mn_topic_fc};font-size:12px;'
                f'font-style:italic;padding:7px 12px;text-align:center;'
                f'border:1px solid {_mn_bdr};">{topic}</td>'
                f'<td style="background:{bg};color:{_mn_auto_fc};font-size:12px;'
                f'padding:7px 14px;border:1px solid {_mn_bdr};">{note if note else "&nbsp;"}</td>'
                f'</tr>'
            )

        for bi in range(4):
            bg     = _mn_card if (4 + bi) % 2 == 0 else _mn_alt
            un     = user_notes_list[bi] if bi < len(user_notes_list) else blank_note
            utopic = un.get("topic", "") or ""
            unote  = un.get("note",  "") or ""
            rows_html += (
                f'<tr>'
                f'<td style="background:{bg};color:{_mn_user_fc};font-size:12px;'
                f'padding:7px 12px;text-align:center;'
                f'border:1px solid {_mn_bdr};">{utopic if utopic else "&nbsp;"}</td>'
                f'<td style="background:{bg};color:{_mn_user_fc};font-size:12px;'
                f'padding:7px 14px;border:1px solid {_mn_bdr};">{unote if unote else "&nbsp;"}</td>'
                f'</tr>'
            )
        return rows_html

    # ── Build full table — Today's Call only ─────────────────────────────────
    all_rows = _build_call_rows(0, call_keys[0], call_labels[0], date_displays[0],
                                auto_vals_c0, user_notes_c0)

    static_tbl = (
        f'<div style="overflow-x:auto;width:100%;">'
        f'<table style="border-collapse:collapse;width:100%;table-layout:fixed;'
        f'font-family:DM Sans,sans-serif;">'
        f'<colgroup>'
        f'<col style="width:130px"><col style="width:220px"><col style="width:auto">'
        f'</colgroup>'
        f'<thead>{col_hdr}</thead>'
        f'<tbody>{all_rows}</tbody>'
        f'</table></div>'
    )
    st.markdown(static_tbl, unsafe_allow_html=True)

    # ── Edit expander — Today's Call only ────────────────────────────────────
    call_key        = call_keys[0]
    call_label      = call_labels[0]
    date_disp       = date_displays[0]
    user_notes_list = user_notes_c0
    if True:
        call_entry = _mn_data["calls"].get(call_key, {})
        with st.expander(f"✏  Edit Notes — {call_label} ({date_disp})", expanded=False):
            new_notes = []
            for ni in range(4):
                un = user_notes_list[ni] if ni < len(user_notes_list) else blank_note
                c1, c2 = st.columns([1, 2])
                with c1:
                    t_val = st.text_input(f"Topic {ni+1}", value=un.get("topic",""),
                                          key=f"mn_{call_key}_topic_{ni}",
                                          placeholder="Topic...")
                with c2:
                    n_val = st.text_input(f"Note {ni+1}", value=un.get("note",""),
                                          key=f"mn_{call_key}_note_{ni}",
                                          placeholder="Notes...")
                new_notes.append({"topic": t_val, "note": n_val})

            if st.button(f"💾  Save Notes", key=f"mn_save_{call_key}"):
                call_entry["user_notes"] = new_notes
                if not call_entry.get("frozen_date"):
                    call_entry["frozen_date"] = _today_str
                    call_entry["auto_rows"]   = auto_vals_c0
                _mn_data["calls"][call_key] = call_entry
                _mn_save(_mn_data)
                st.success("Notes saved.")
                st.rerun()



# ══════════════════════════════════════════════════════════════════════════════
# HOTEL DASHBOARD SHELL
# ══════════════════════════════════════════════════════════════════════════════

def render_hotel_dashboard():
    hotel_id = st.session_state.hotel_id
    hotel    = next((h for h in cfg.HOTELS if h["id"] == hotel_id), None)

    if not hotel:
        st.error("Hotel not found.")
        st.session_state.hotel_id = None
        st.rerun()
        return

    # Force cache clear for IHG hotels to ensure DSS forecast data is always fresh.
    # IHG data pipeline (_build_ihg_daily_forecast) adds columns that may not exist
    # in previously cached results from older data_loader versions.
    if hotel_id in getattr(dl, "IHG_HOTELS", set()):
        _ver_key = f"_ihg_cache_ver_{hotel_id}"
        _cur_ver = getattr(dl, "_DL_VERSION", "")
        if st.session_state.get(_ver_key) != _cur_ver:
            load_hotel_data.clear()
            st.session_state[_ver_key] = _cur_ver

    # ── Restore last-used tab for this hotel ──
    if st.session_state.hotel_tab_memory.get(hotel_id):
        saved_tab = st.session_state.hotel_tab_memory[hotel_id]
        if saved_tab in TABS and st.session_state.active_tab == "Biweekly Snapshot":
            st.session_state.active_tab = saved_tab

    # ── Header ──
    st.markdown(f"""
    <div class="dash-header">
        <div>
            <div class="dash-hotel-name">{hotel['display_name']} · {hotel['subtitle']}</div>
            <div class="dash-hotel-sub">{hotel['brand']} · {hotel.get('total_rooms','?')} rooms</div>
        </div>
        <div class="dash-last-refresh">Last loaded: {datetime.now().strftime('%I:%M %p')}</div>
    </div>
    """, unsafe_allow_html=True)

    theme_toggle_btn("theme_dash")
    # ── Back + Refresh in same row, then tab radio ──
    back_col, ref_col, info_col, _ = st.columns([1.5, 1.5, 2.5, 4.2])
    with back_col:
        if st.button("← Portfolio", key="back_btn", use_container_width=True):
            # Save current tab for this hotel before leaving
            st.session_state.hotel_tab_memory[hotel_id] = st.session_state.active_tab
            st.session_state.hotel_id = None
            st.session_state.hotel_data = {}
            st.session_state.active_tab = "Biweekly Snapshot"
            st.rerun()
    with ref_col:
        if st.button("⟳ Refresh Data", key="refresh_btn", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with info_col:
        pass  # time is shown in the platform header live clock

    # Use Streamlit native tabs (works without JavaScript)
    selected_tab = st.radio(
        "Navigation",
        TABS,
        index=TABS.index(st.session_state.active_tab),
        horizontal=True,
        label_visibility="collapsed",
        key="tab_radio",
    )
    st.session_state.active_tab = selected_tab
    st.session_state.hotel_tab_memory[hotel_id] = selected_tab

    # Always load fresh — cache key is file mtimes so stale data never persists
    with st.spinner(f"Loading {hotel['display_name']} data..."):
        data = _get_fresh_hotel_data(hotel_id)

    # Show any load errors
    errors = [f"**{k.replace('_error','')}**: {v}"
              for k, v in data.items() if k.endswith("_error")]
    if errors:
        with st.expander("⚠ Some files had loading issues"):
            for e in errors:
                st.warning(e)

    # ── Route to tab ──
    st.markdown('<div class="content-area">', unsafe_allow_html=True)

    if selected_tab == "Biweekly Snapshot":
        render_snapshot_tab(data, hotel)
    elif selected_tab == "Calendar View":
        render_dashboard_tab(data, hotel)
    elif selected_tab == "Monthly Performance":
        render_biweekly_tab(data, hotel)
    elif selected_tab == "Segment Analysis":
        render_srp_tab(data, hotel)
    elif selected_tab == "Booking Pace by SRP":
        render_srp_pace_tab(data, hotel)
    elif selected_tab == "Groups":
        render_groups_tab(data, hotel)
    elif selected_tab == "Rate Insights":
        render_rates_tab(data, hotel)
    elif selected_tab == "STR":
        render_str_tab(data, hotel)
    elif selected_tab == "Demand Nights":
        render_demand_tab(data, hotel)
    elif selected_tab == "Market Events":
        render_events_tab(data, hotel)
    elif selected_tab == "Call Recap":
        render_call_recap_tab(data, hotel)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── File Sources & Sync Status — bottom of every page ──
    st.markdown("<div style='margin-top:40px;'></div>", unsafe_allow_html=True)
    _hotel_cfg   = next(h for h in cfg.HOTELS if h["id"] == hotel_id)
    _hotel_files = cfg.detect_files(_hotel_cfg)
    with st.expander("📁 File Sources & Sync Status", expanded=False):
        diag_rows = []
        for role, path in sorted(_hotel_files.items()):
            if path and os.path.exists(str(path)):
                mtime   = pd.Timestamp(os.path.getmtime(str(path)), unit="s").strftime("%Y-%m-%d %H:%M:%S")
                size_kb = os.path.getsize(str(path)) // 1024
                diag_rows.append({"File Role": role, "Path": str(path), "Last Modified": mtime, "Size (KB)": size_kb})
            else:
                diag_rows.append({"File Role": role, "Path": str(path) if path else "NOT FOUND", "Last Modified": "—", "Size (KB)": "—"})
        if diag_rows:
            st.dataframe(pd.DataFrame(diag_rows), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# ROUTER — Landing vs Dashboard
# ══════════════════════════════════════════════════════════════════════════════

# ── Handle calendar date click (?cal_date=YYYY-MM-DD&hotel=hotel_id) ──
# Calendar cells use JS pushState (no page reload) — session state is preserved.
# This block is a safety fallback for any direct URL access with cal_date param.
# render_dashboard_tab reads ?cal_date= and stores it in session state.
_cal_qp = st.query_params.get("cal_date", None)
_hot_qp = st.query_params.get("hotel", None)
if _cal_qp and _hot_qp and st.session_state.hotel_id is None:
    _matched2 = next((h for h in cfg.HOTELS if h["id"] == _hot_qp and h.get("active", True)), None)
    if _matched2:
        st.session_state.hotel_id   = _hot_qp
        st.session_state.active_tab = "Calendar View"
        st.session_state.hotel_data = {}
        # Leave ?cal_date= + ?hotel= in place — render_dashboard_tab consumes them

if st.session_state.hotel_id is None:
    render_landing()
else:
    render_hotel_dashboard()
