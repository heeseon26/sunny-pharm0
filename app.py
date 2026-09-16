from __future__ import annotations

from pathlib import Path
from datetime import datetime, date
from zoneinfo import ZoneInfo
import base64
import html
import re

import pandas as pd
import streamlit as st

from services.excel_source import (
    DEFAULT_EXCEL,
    available_sheets,
    load_sheet,
    clean_empty,
    parse_incoming,
    parse_order_queue,
    parse_owm,
    parse_coupang,
    parse_direct_orders,
    parse_vendor_products,
    parse_substitution,
    parse_influencers,
    parse_sets,
    parse_hub_safe,
    parse_burn_down,
    save_preview,
    load_preview,
)


# -----------------------------------------------------------------------------
# App config
# -----------------------------------------------------------------------------
BASE = Path(__file__).resolve().parent
PROFILE = BASE / "assets" / "sunny_profile_web.png"
FACE = BASE / "assets" / "sunny_face.png"
HERO = BASE / "assets" / "sunny_hero.png"

st.set_page_config(
    page_title="SUNNY PHARM 365",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# -----------------------------------------------------------------------------
# Small helpers
# -----------------------------------------------------------------------------
def image_data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


PROFILE_URI = image_data_uri(PROFILE)
FACE_URI = image_data_uri(FACE) or PROFILE_URI
HERO_URI = image_data_uri(HERO) or PROFILE_URI


def esc(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return html.escape(str(value))


def nonempty(value) -> bool:
    if value is None:
        return False
    if isinstance(value, float) and pd.isna(value):
        return False
    s = str(value).strip()
    return s not in {"", "nan", "None", "False"}


def money(value) -> str:
    try:
        if pd.isna(value):
            return "-"
        return f"{float(value):,.0f}원"
    except Exception:
        return str(value)


def truthy(value) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {
        "true", "1", "yes", "y", "완료", "입고완료"
    }


def search_df(df: pd.DataFrame, query: str) -> pd.DataFrame:
    if df.empty or not query:
        return df
    q = str(query).strip()
    mask = df.astype(str).apply(
        lambda col: col.str.contains(q, case=False, na=False, regex=False)
    ).any(axis=1)
    return df[mask]


def loose_date(value) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    s = str(value).strip()
    if not s:
        return None

    # 2026.08.19 / 2026-08-19 / 2026년 8월 19일
    m = re.search(r"(20\d{2})\D{0,3}(\d{1,2})\D{1,3}(\d{1,2})", s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass

    # 8.24 / 8/12 style - assume current year
    m = re.search(r"(?<!\d)(\d{1,2})\D{1,3}(\d{1,2})(?!\d)", s)
    if m:
        try:
            return date(datetime.now(ZoneInfo("Asia/Seoul")).year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            pass
    return None


def compact_text(value, limit: int = 46) -> str:
    s = str(value).replace("\n", " ").strip() if nonempty(value) else ""
    if len(s) > limit:
        return s[: limit - 1] + "…"
    return s


# -----------------------------------------------------------------------------
# Password (optional on Streamlit Cloud)
# -----------------------------------------------------------------------------
try:
    _app_password = st.secrets.get("APP_PASSWORD", "")
except Exception:
    _app_password = ""

if _app_password and not st.session_state.get("_sunny_auth_ok", False):
    st.markdown(
        """
        <div style="max-width:480px;margin:12vh auto 20px;text-align:center">
          <div style="font-size:3rem">☀️</div>
          <h2 style="margin:.2rem 0">SUNNY PHARM 365</h2>
          <p style="color:#7f817d">약사님 전용 워크스페이스</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _pw = st.text_input("프리뷰 비밀번호", type="password", placeholder="비밀번호를 입력해주세요")
    if st.button("SUNNY 열기", use_container_width=True):
        if _pw == _app_password:
            st.session_state["_sunny_auth_ok"] = True
            st.rerun()
        else:
            st.error("비밀번호가 맞지 않아요.")
    st.stop()


# -----------------------------------------------------------------------------
# Design system
# -----------------------------------------------------------------------------
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700;800&family=Nunito:wght@600;700;800;900&display=swap');

:root {
  --sun: #F7C948;
  --sun-deep: #E9A91C;
  --sun-soft: #FFF1B7;
  --cream: #FFFCF5;
  --paper: #FFFFFF;
  --sage: #83A27F;
  --sage-deep: #55755A;
  --sage-soft: #EEF5EC;
  --mint-soft: #EDF8F1;
  --blue-soft: #EEF5FF;
  --rose-soft: #FFF0F2;
  --orange-soft: #FFF6E4;
  --ink: #2B302C;
  --muted: #7C827E;
  --line: #ECEAE4;
  --shadow: 0 12px 34px rgba(67, 70, 60, .07);
}

html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
.stApp {
  background: radial-gradient(circle at 80% 0%, #FFF6D6 0%, transparent 30%),
              linear-gradient(145deg, #FFFDF8 0%, #F9FBF8 58%, #FFF9E9 100%);
  color: var(--ink);
}
#MainMenu, footer { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }
.block-container {
  max-width: 1500px;
  padding-top: 1.15rem;
  padding-bottom: 3rem;
}
h1, h2, h3, h4 { color: var(--ink); letter-spacing: -.035em; }
p { line-height: 1.6; }

/* Sidebar */
section[data-testid="stSidebar"] {
  width: 268px !important;
  background: linear-gradient(180deg, #FFFDF7 0%, #FFFAE9 100%);
  border-right: 1px solid #EFE8D7;
}
section[data-testid="stSidebar"] > div { width: 268px !important; }
section[data-testid="stSidebar"] [data-testid="stSidebarContent"] { padding: 1.2rem 1rem 1rem; }
.sidebar-brand { text-align:center; padding: 5px 8px 13px; }
.sidebar-logo {
  font-family:'Nunito',sans-serif; font-weight:900; font-size:2.0rem;
  color:#E5A91A; letter-spacing:.02em; line-height:1.05;
}
.sidebar-sublogo { font-size:.68rem; color:#75906F; letter-spacing:.26em; font-weight:800; margin-top:6px; }
.sidebar-tagline { margin-top:14px; color:#9C8265; font-size:.73rem; line-height:1.65; }
.sidebar-profile {
  width:54px; height:54px; border-radius:50%; object-fit:cover;
  border:3px solid #fff; box-shadow:0 5px 16px rgba(80,71,47,.12);
}
.sidebar-user {
  display:flex; align-items:center; gap:10px; padding:10px;
  background:rgba(255,255,255,.72); border:1px solid #EFE8D7;
  border-radius:17px; margin:8px 0 12px;
}
.sidebar-user-name { font-weight:800; font-size:.84rem; color:#383D39; }
.sidebar-user-role { font-size:.67rem; color:#96978F; margin-top:2px; }

section[data-testid="stSidebar"] div[role="radiogroup"] { gap:4px; }
section[data-testid="stSidebar"] div[role="radiogroup"] > label {
  border-radius:14px;
  padding:.66rem .75rem;
  transition:all .18s ease;
  color:#515852;
  font-weight:650;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
  background:#FFF4C9;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
  background:linear-gradient(90deg,#FFF0AE 0%,#FFF7D9 100%);
  box-shadow:inset 0 0 0 1px #F5DE86;
  color:#3C403C;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child { display:none; }
section[data-testid="stSidebar"] [data-testid="stFileUploader"] { border-radius:14px; }

/* Top bar */
.top-greeting .mini { color:#999C97; font-size:.78rem; font-weight:600; }
.top-greeting .big {
  font-family:'Nunito','Noto Sans KR',sans-serif;
  font-size:1.73rem; font-weight:900; color:#3A342E; line-height:1.2;
}
.top-profile {
  display:flex; justify-content:flex-end; align-items:center; gap:10px; padding-top:2px;
}
.top-profile img { width:45px; height:45px; border-radius:50%; object-fit:cover; border:3px solid white; box-shadow:0 4px 12px rgba(71,64,45,.12); }
.top-profile .name { font-weight:800; font-size:.84rem; }
.top-profile .role { color:#9A9C97; font-size:.65rem; margin-top:2px; }

/* Streamlit inputs */
.stTextInput input, .stSelectbox [data-baseweb="select"] > div {
  border-radius:14px !important;
  border-color:#E8E7E1 !important;
  background:rgba(255,255,255,.94) !important;
  min-height:46px;
  box-shadow:0 4px 12px rgba(60,65,58,.025);
}
.stTextInput input:focus { border-color:#E9C44C !important; box-shadow:0 0 0 1px #E9C44C !important; }
.stButton > button {
  border-radius:13px; border:1px solid #ECD582; background:#FFF7D2;
  color:#624F0F; font-weight:800; min-height:40px;
}
.stButton > button:hover { border-color:#DDBB42; color:#4E3F0E; background:#FFF1B4; }

/* Home hero */
.home-hero {
  position:relative; min-height:245px; overflow:hidden;
  border:1px solid #F0E1A9; border-radius:26px;
  background:
    radial-gradient(circle at 14% 28%, rgba(255,221,92,.40) 0 50px, transparent 51px),
    linear-gradient(120deg,#FFF1B5 0%,#FFF8D8 52%,#F2F6E8 100%);
  box-shadow:var(--shadow); margin:12px 0 16px;
}
.home-hero:after {
  content:'☀︎   ♡     ✦        ♡';
  position:absolute; right:34px; top:20px; color:rgba(190,146,48,.38);
  font-size:1.4rem; letter-spacing:1.1rem; transform:rotate(-4deg);
}
.hero-photo-wrap { position:absolute; left:20px; bottom:-12px; width:31%; max-width:330px; min-width:245px; }
.hero-photo-wrap img { width:100%; filter:drop-shadow(0 14px 18px rgba(91,74,33,.13)); }
.hero-copy-wrap { margin-left:32%; padding:38px 36px 30px; position:relative; z-index:2; }
.hero-eyebrow { color:#A98416; font-family:'Nunito',sans-serif; letter-spacing:.15em; font-weight:900; font-size:.73rem; }
.hero-title { margin-top:8px; font-size:1.55rem; line-height:1.6; font-weight:800; color:#3D3B34; letter-spacing:-.04em; }
.hero-title b { color:#83630A; }
.hero-caption { margin-top:13px; color:#847B68; font-size:.82rem; }
.hero-logo { position:absolute; right:36px; bottom:27px; text-align:center; opacity:.95; }
.hero-logo .sunny { font-family:'Nunito',sans-serif; font-weight:900; color:#E5A91A; font-size:2rem; }
.hero-logo .pharm { font-size:.62rem; letter-spacing:.26em; color:#799072; font-weight:800; }

/* KPI cards */
.kpi-card {
  position:relative; overflow:hidden; border-radius:20px; padding:17px 18px 15px;
  min-height:116px; border:1px solid rgba(225,225,218,.86); box-shadow:0 7px 22px rgba(70,72,65,.05);
}
.kpi-card.rose { background:linear-gradient(135deg,#FFF7F8,#FFF0F2); }
.kpi-card.green { background:linear-gradient(135deg,#F7FCF8,#EAF7EF); }
.kpi-card.blue { background:linear-gradient(135deg,#F7FAFF,#EDF5FF); }
.kpi-card.yellow { background:linear-gradient(135deg,#FFFDF7,#FFF5D8); }
.kpi-row { display:flex; align-items:center; justify-content:space-between; }
.kpi-icon {
  width:42px; height:42px; border-radius:14px; display:flex; align-items:center; justify-content:center;
  background:rgba(255,255,255,.72); font-size:1.18rem; box-shadow:inset 0 0 0 1px rgba(255,255,255,.85);
}
.kpi-label { font-size:.78rem; color:#6E756F; font-weight:750; }
.kpi-value { font-family:'Nunito',sans-serif; font-size:1.8rem; font-weight:900; color:#303631; line-height:1; margin-top:10px; }
.kpi-hint { color:#A0A19B; font-size:.67rem; margin-top:8px; }

/* Generic cards / panels */
.panel {
  background:rgba(255,255,255,.93); border:1px solid #ECEAE4; border-radius:20px;
  padding:16px 17px; box-shadow:0 7px 22px rgba(62,66,59,.045); margin-bottom:12px;
}
.panel-title { display:flex; align-items:center; gap:8px; font-weight:850; color:#343A35; font-size:1rem; margin-bottom:12px; }
.panel-title .accent { color:#E0A312; }
.section-title { font-size:1.08rem; font-weight:850; color:#343A35; margin:18px 0 10px; }
.page-head { display:flex; align-items:flex-end; justify-content:space-between; margin:5px 0 15px; }
.page-kicker { color:#AD8B2D; font-family:'Nunito',sans-serif; font-size:.7rem; font-weight:900; letter-spacing:.16em; }
.page-title { font-size:1.85rem; font-weight:850; line-height:1.15; margin-top:4px; }
.page-desc { color:#8A8F8A; font-size:.82rem; margin-top:6px; }
.page-chip { display:inline-flex; align-items:center; gap:6px; background:#FFF2B8; color:#836511; border-radius:999px; padding:7px 11px; font-size:.7rem; font-weight:800; }

/* Task list */
.task-table { width:100%; border-collapse:collapse; }
.task-table th { color:#A0A19B; font-size:.67rem; text-align:left; font-weight:700; padding:0 8px 9px; }
.task-table td { border-top:1px solid #F0EFEA; padding:10px 8px; font-size:.76rem; vertical-align:middle; }
.task-title { font-weight:700; color:#3C443E; }
.pill { display:inline-flex; align-items:center; border-radius:999px; padding:4px 8px; font-size:.64rem; font-weight:800; white-space:nowrap; }
.pill.red { background:#FFE9ED; color:#C25D6B; }
.pill.orange { background:#FFF1D0; color:#AF7D18; }
.pill.green { background:#EAF7EE; color:#4B8C5A; }
.pill.blue { background:#EAF3FF; color:#4C7FC6; }
.pill.gray { background:#F0F1EF; color:#777D78; }

/* Visit cards */
.visit-row { display:grid; grid-template-columns:76px 1fr auto; gap:11px; align-items:center; padding:11px 4px; border-top:1px solid #F0EFEA; }
.visit-row:first-of-type { border-top:0; }
.visit-date { font-family:'Nunito',sans-serif; font-weight:900; font-size:.82rem; color:#515951; }
.visit-name { font-weight:800; font-size:.78rem; color:#353B36; }
.visit-meta { color:#959993; font-size:.66rem; margin-top:3px; }

/* Product finder */
.finder-result { background:#fff; border:1px solid #ECEAE4; border-radius:16px; padding:13px 14px; margin-bottom:8px; }
.finder-vendor { color:#98700C; font-size:.65rem; font-weight:850; }
.finder-product { color:#343A35; font-weight:800; margin-top:4px; font-size:.82rem; }
.finder-meta { color:#989A95; font-size:.65rem; margin-top:4px; }

/* Note and set cards */
.soft-card { background:#fff; border:1px solid #ECEAE4; border-radius:18px; padding:15px 16px; margin-bottom:10px; box-shadow:0 5px 18px rgba(61,65,58,.035); }
.badge { display:inline-flex; border-radius:999px; padding:4px 8px; font-size:.63rem; font-weight:850; background:#EEF5EF; color:#5C7862; }
.badge-yellow { background:#FFF1BA; color:#8A690D; }
.badge-red { background:#FDEBEE; color:#B05F69; }
.product-name { font-weight:800; color:#354038; font-size:.88rem; margin:6px 0 3px; }
.subtext { color:#898E89; font-size:.7rem; line-height:1.55; }
.notice { background:#F3F7F2; border:1px solid #DFEADF; border-radius:15px; padding:12px 14px; color:#647169; font-size:.73rem; }

/* Dataframe / tabs */
[data-testid="stVerticalBlockBorderWrapper"] { border-radius:20px; border-color:#ECEAE4 !important; background:rgba(255,255,255,.88); box-shadow:0 7px 22px rgba(62,66,59,.035); }
[data-testid="stDataFrame"], [data-testid="stDataEditor"] {
  border-radius:17px; overflow:hidden; border:1px solid #ECEAE4;
  box-shadow:0 6px 18px rgba(62,66,59,.035);
}
.stTabs [data-baseweb="tab-list"] {
  gap:5px; background:#F4F3EE; border-radius:13px; padding:4px; margin-bottom:8px;
}
.stTabs [data-baseweb="tab"] { border-radius:10px; padding:8px 13px; font-weight:700; }
.stTabs [aria-selected="true"] { background:white; box-shadow:0 2px 8px rgba(58,62,57,.06); }
div[data-testid="stMetric"] { background:white; border:1px solid #ECEAE4; padding:12px; border-radius:15px; }
hr { border-color:#EEECE5; }

@media (max-width: 900px) {
  .hero-photo-wrap { opacity:.28; left:-20px; width:50%; }
  .hero-copy-wrap { margin-left:0; padding:30px 24px; }
  .hero-logo { display:none; }
  .home-hero { min-height:225px; }
  .top-profile { display:none; }
  .visit-row { grid-template-columns:68px 1fr; }
  .visit-row > :last-child { display:none; }
}
</style>
""",
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Data source / upload
# -----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def materialize_upload(uploaded):
    out = BASE / "data" / "_uploaded_preview.xlsx"
    out.write_bytes(uploaded.getvalue())
    return out


if "excel_path" not in st.session_state:
    st.session_state.excel_path = str(DEFAULT_EXCEL) if DEFAULT_EXCEL.exists() else None


# -----------------------------------------------------------------------------
# Sidebar navigation
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-brand">
          <div class="sidebar-logo">☘ SUNNY</div>
          <div class="sidebar-sublogo">PHARM APP</div>
          <div class="sidebar-tagline">Pharmacist for a brighter<br>tomorrow ♡</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if PROFILE_URI:
        st.markdown(
            f"""
            <div class="sidebar-user">
              <img class="sidebar-profile" src="{FACE_URI}">
              <div>
                <div class="sidebar-user-name">SUNNY</div>
                <div class="sidebar-user-role">Pharmacist workspace</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    pages = [
        "🏠  오늘의 SUNNY",
        "📦  입고관리",
        "🛒  주문센터",
        "✨  인플루언서",
        "🔎  상품 찾기",
        "💊  대체조제",
        "🎁  세트상품",
        "📚  SUNNY NOTE",
        "🗂️  원본 시트",
    ]
    page = st.radio("메뉴", pages, label_visibility="collapsed")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    with st.expander("⚙️ 데이터 · 설정", expanded=False):
        uploaded = st.file_uploader(
            "다른 SUNNY 엑셀로 바꾸기",
            type=["xlsx"],
            help="프리뷰용으로 다른 엑셀을 임시 연결할 수 있어요.",
        )
        if uploaded is not None:
            st.session_state.excel_path = str(materialize_upload(uploaded))
            st.cache_data.clear()
            st.success("새 엑셀을 연결했어요.")
        st.caption("현재 데이터")
        st.code("SUNNY.xlsx" if st.session_state.excel_path else "연결된 파일 없음", language=None)
        st.caption("비밀번호 시트는 앱에서 제외됩니다.")

    st.markdown(
        """
        <div style="text-align:center;margin-top:18px;color:#9E9279;font-size:.68rem;line-height:1.7">
          🌱 작은 약이<br>더 건강한 내일을 만듭니다.<br><br>
          <span style="letter-spacing:.14em;font-size:.61rem">SUNNY PHARM APP · v2</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


if not st.session_state.excel_path or not Path(st.session_state.excel_path).exists():
    st.warning("SUNNY 엑셀 파일을 왼쪽에서 업로드해 주세요.")
    st.stop()

XLSX = Path(st.session_state.excel_path)


# -----------------------------------------------------------------------------
# Cached parsers
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def incoming(p):
    return parse_incoming(p)


@st.cache_data(show_spinner=False)
def orders(p):
    return parse_order_queue(p)


@st.cache_data(show_spinner=False)
def owm(p):
    return parse_owm(p)


@st.cache_data(show_spinner=False)
def coupang(p):
    return parse_coupang(p)


@st.cache_data(show_spinner=False)
def direct_orders(p):
    return parse_direct_orders(p)


@st.cache_data(show_spinner=False)
def vendors(p):
    return parse_vendor_products(p)


@st.cache_data(show_spinner=False)
def substitutes(p):
    return parse_substitution(p)


@st.cache_data(show_spinner=False)
def influencers(p):
    return parse_influencers(p)


@st.cache_data(show_spinner=False)
def sets(p):
    return parse_sets(p)


@st.cache_data(show_spinner=False)
def hub(p):
    return parse_hub_safe(p)


@st.cache_data(show_spinner=False)
def burn(p):
    return parse_burn_down(p)


# -----------------------------------------------------------------------------
# UI components
# -----------------------------------------------------------------------------
def page_header(title: str, desc: str, kicker: str, chip: str | None = None):
    chip_html = f'<span class="page-chip">{esc(chip)}</span>' if chip else ""
    st.markdown(
        f"""
        <div class="page-head">
          <div>
            <div class="page-kicker">{esc(kicker)}</div>
            <div class="page-title">{esc(title)}</div>
            <div class="page-desc">{esc(desc)}</div>
          </div>
          <div>{chip_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_card(icon: str, label: str, value, hint: str, variant: str):
    st.markdown(
        f"""
        <div class="kpi-card {variant}">
          <div class="kpi-row">
            <div class="kpi-label">{esc(label)}</div>
            <div class="kpi-icon">{icon}</div>
          </div>
          <div class="kpi-value">{esc(value)}</div>
          <div class="kpi-hint">{esc(hint)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_topbar(global_query_key: str = "global_search") -> str:
    now = datetime.now(ZoneInfo("Asia/Seoul"))
    if now.hour < 12:
        english = "Good morning, SUNNY"
    elif now.hour < 18:
        english = "Good afternoon, SUNNY"
    else:
        english = "Good evening, SUNNY"

    c1, c2, c3 = st.columns([1.25, 1.65, .72], vertical_alignment="center")
    with c1:
        st.markdown(
            f"""
            <div class="top-greeting">
              <div class="mini">오늘도 좋은 하루예요! · {now.strftime('%m월 %d일')}</div>
              <div class="big">{english} ☀️</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        query = st.text_input(
            "전체 검색",
            placeholder="🔍  상품명, 업체명, 주문 상품 등을 검색해보세요…",
            label_visibility="collapsed",
            key=global_query_key,
        )
    with c3:
        if PROFILE_URI:
            st.markdown(
                f"""
                <div class="top-profile">
                  <img src="{FACE_URI}">
                  <div>
                    <div class="name">SUNNY</div>
                    <div class="role">약사님, 오늘도 화이팅!</div>
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
    return query


def render_home_hero():
    photo = f'<div class="hero-photo-wrap"><img src="{HERO_URI}"></div>' if HERO_URI else ""
    st.markdown(
        f"""
        <div class="home-hero">
          {photo}
          <div class="hero-copy-wrap">
            <div class="hero-eyebrow">SUNNY PHARM 365</div>
            <div class="hero-title">약사님의 하루가<br>더 빛날 수 있도록,<br><b>SUNNY</b>가 함께할게요. 🌱</div>
            <div class="hero-caption">작은 변화가, 더 건강한 내일을 만듭니다. ♡</div>
          </div>
          <div class="hero-logo">
            <div class="sunny">SUNNY♥</div>
            <div class="pharm">PHARM APP</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def task_rows(inc_df: pd.DataFrame, order_df: pd.DataFrame, max_rows: int = 6) -> list[dict]:
    rows: list[dict] = []
    if not inc_df.empty and "입고상태" in inc_df.columns:
        target = inc_df[
            inc_df["입고상태"].fillna("").astype(str).str.contains("미입고|부분입고", regex=True, na=False)
        ].head(4)
        for _, r in target.iterrows():
            stat = str(r.get("입고상태", "확인필요"))
            rows.append(
                {
                    "title": f"{compact_text(r.get('상품명', '입고 상품'), 36)} 입고 확인",
                    "kind": "입고",
                    "status": stat,
                    "status_class": "red" if "미입고" in stat else "orange",
                    "source": compact_text(r.get("주문처", "입고관리"), 16),
                }
            )

    if not order_df.empty:
        completed = order_df.get("주문완료", pd.Series(False, index=order_df.index)).apply(truthy)
        target = order_df[~completed].head(4)
        for _, r in target.iterrows():
            rows.append(
                {
                    "title": f"{compact_text(r.get('제품명', '주문 상품'), 36)} 주문",
                    "kind": "주문",
                    "status": "주문 대기",
                    "status_class": "blue",
                    "source": compact_text(r.get("주문채널", "주문센터"), 16),
                }
            )

    return rows[:max_rows]


def render_task_panel(rows: list[dict]):
    if not rows:
        body = '<tr><td colspan="4" style="color:#9B9E99;padding:18px 8px">지금 급하게 확인할 업무가 없어요 ☀️</td></tr>'
    else:
        body = "".join(
            f"""
            <tr>
              <td><span class="task-title">{esc(r['title'])}</span></td>
              <td><span class="pill {'orange' if r['kind']=='입고' else 'blue'}">{esc(r['kind'])}</span></td>
              <td><span class="pill {esc(r['status_class'])}">{esc(r['status'])}</span></td>
              <td style="color:#989C97">{esc(r['source'])}</td>
            </tr>
            """
            for r in rows
        )
    st.markdown(
        f"""
        <div class="panel">
          <div class="panel-title"><span class="accent">☑</span> 오늘 먼저 확인할 일</div>
          <table class="task-table">
            <thead><tr><th>업무 내용</th><th>구분</th><th>상태</th><th>채널</th></tr></thead>
            <tbody>{body}</tbody>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


def influencer_display_rows(df: pd.DataFrame, count: int = 4) -> list[dict]:
    today = datetime.now(ZoneInfo("Asia/Seoul")).date()
    parsed = []
    for idx, r in df.iterrows():
        d = loose_date(r.get("날짜"))
        if d:
            parsed.append((idx, d, r))

    future = [x for x in parsed if x[1] >= today]
    if future:
        chosen = sorted(future, key=lambda x: x[1])[:count]
    else:
        chosen = sorted(parsed, key=lambda x: x[1], reverse=True)[:count]

    result = []
    for _, d, r in chosen:
        name_candidates = [r.get("닉네임/이름"), r.get("팔로워/언어"), r.get("분류")]
        name = next((compact_text(v, 28) for v in name_candidates if nonempty(v)), "방문 일정")
        platform_candidates = [r.get("플랫폼/특이사항"), r.get("제품/가이드"), r.get("분류")]
        meta = next((compact_text(v, 34) for v in platform_candidates if nonempty(v)), "일정 확인")

        time_val = ""
        for key in ["시간", "장소/시간2"]:
            v = r.get(key)
            if nonempty(v):
                time_val = compact_text(v, 12)
                break
        result.append({"date": d, "name": name, "meta": meta, "time": time_val})
    return result


def render_visit_panel(rows: list[dict]):
    if not rows:
        body = '<div style="color:#9B9E99;font-size:.75rem;padding:12px 4px">등록된 방문 일정이 없어요.</div>'
    else:
        body = "".join(
            f"""
            <div class="visit-row">
              <div class="visit-date">{r['date'].strftime('%m/%d')}<br><span style="font-size:.62rem;color:#A4A7A2">{esc(r['time'])}</span></div>
              <div><div class="visit-name">{esc(r['name'])}</div><div class="visit-meta">{esc(r['meta'])}</div></div>
              <div><span class="pill green">일정</span></div>
            </div>
            """
            for r in rows
        )
    st.markdown(
        f"""
        <div class="panel">
          <div class="panel-title"><span class="accent">📣</span> 최근 · 다가오는 인플루언서 일정</div>
          {body}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_product_results(df: pd.DataFrame, limit: int = 8):
    if df.empty:
        st.info("검색 결과가 없어요.")
        return
    rows = df.head(limit).to_dict("records")
    cols = st.columns(2)
    for i, r in enumerate(rows):
        with cols[i % 2]:
            st.markdown(
                f"""
                <div class="finder-result">
                  <div class="finder-vendor">{esc(r.get('업체','거래처'))}</div>
                  <div class="finder-product">{esc(r.get('상품명',''))}</div>
                  <div class="finder-meta">{esc(r.get('구분/방명',''))}{' · ' if nonempty(r.get('구분/방명')) and nonempty(r.get('공급처/메모')) else ''}{esc(r.get('공급처/메모',''))}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


# -----------------------------------------------------------------------------
# HOME
# -----------------------------------------------------------------------------
if page == "🏠  오늘의 SUNNY":
    global_q = render_topbar("home_global_search")

    if global_q:
        vendor_hits = search_df(vendors(str(XLSX)), global_q).head(6)
        incoming_hits = search_df(incoming(str(XLSX)), global_q).head(4)
        with st.expander(f"🔎 ‘{global_q}’ 전체 검색 결과", expanded=True):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**상품 · 거래처**")
                if vendor_hits.empty:
                    st.caption("검색 결과 없음")
                else:
                    render_product_results(vendor_hits, 6)
            with c2:
                st.markdown("**입고관리**")
                if incoming_hits.empty:
                    st.caption("검색 결과 없음")
                else:
                    show = [c for c in ["상품명", "주문처", "입고상태", "주문수량"] if c in incoming_hits.columns]
                    st.dataframe(incoming_hits[show], use_container_width=True, hide_index=True)

    render_home_hero()

    inc = incoming(str(XLSX))
    oq = orders(str(XLSX))
    inf = influencers(str(XLSX))
    bd = burn(str(XLSX))

    pending_inc = partial = complete_inc = 0
    if not inc.empty and "입고상태" in inc.columns:
        s = inc["입고상태"].fillna("").astype(str)
        pending_inc = int(s.str.contains("미입고", na=False).sum())
        partial = int(s.str.contains("부분", na=False).sum())
        complete_inc = int(s.str.contains("입고완료", na=False).sum())

    pending_order = 0
    if not oq.empty and "주문완료" in oq.columns:
        pending_order = int((~oq["주문완료"].apply(truthy)).sum())

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("📦", "미입고", pending_inc, "확인이 필요한 입고 건이에요", "rose")
    with c2:
        kpi_card("🛒", "주문 대기", pending_order, "아직 주문완료 전이에요", "green")
    with c3:
        kpi_card("📅", "인플 일정", len(inf), "등록된 방문 일정이에요", "blue")
    with c4:
        kpi_card("⚠️", "소진 우선", len(bd), "먼저 안내하면 좋은 상품이에요", "yellow")

    left, right = st.columns([1.35, .9], gap="medium")
    with left:
        render_task_panel(task_rows(inc, oq, 6))
        st.markdown(
            f"""
            <div class="notice">🌱 현재 입고완료 <b>{complete_inc}</b>건 · 부분입고 <b>{partial}</b>건을 확인할 수 있어요. 프리뷰에서 수정한 값은 원본 Excel이 아니라 별도 저장본에 기록됩니다.</div>
            """,
            unsafe_allow_html=True,
        )
    with right:
        render_visit_panel(influencer_display_rows(inf, 4))

        with st.container(border=True):
            st.markdown('<div class="panel-title"><span class="accent">🔎</span> 빠른 상품 찾기</div>', unsafe_allow_html=True)
            quick_q = st.text_input(
                "빠른 상품 찾기",
                placeholder="상품명 또는 업체명을 입력하세요",
                label_visibility="collapsed",
                key="home_quick_product",
            )
            quick_df = search_df(vendors(str(XLSX)), quick_q) if quick_q else pd.DataFrame()
            if quick_q:
                render_product_results(quick_df, 4)
            else:
                st.caption("예: 센시아, 마데카솔, 리쥬란, 유산균…")

    st.markdown(
        """
        <div style="margin-top:4px;background:linear-gradient(90deg,#EEF5EC,#FFF8DE);border-radius:15px;padding:12px 16px;color:#6D776C;font-size:.74rem;text-align:center">
          🌿 “좋은 약이, 더 건강한 내일을 만듭니다.” <span style="margin-left:18px;color:#A4A59F">SUNNY와 함께하는 건강한 하루 ♡</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# INCOMING
# -----------------------------------------------------------------------------
elif page == "📦  입고관리":
    df = incoming(str(XLSX))
    saved = load_preview("입고관리")
    if saved is not None:
        df = saved

    status_series = df.get("입고상태", pd.Series(dtype=str)).fillna("").astype(str)
    page_header("입고관리", "미입고와 부분입고를 먼저 보고, 실제 입고 수량까지 한 번에 확인해요.", "DELIVERY DESK", f"총 {len(df)}개")

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        kpi_card("📋", "전체 품목", len(df), "현재 관리 중인 행", "blue")
    with m2:
        kpi_card("📦", "미입고", int(status_series.str.contains("미입고", na=False).sum()), "아직 도착 전", "rose")
    with m3:
        kpi_card("🌓", "부분입고", int(status_series.str.contains("부분", na=False).sum()), "추가 확인 필요", "yellow")
    with m4:
        kpi_card("✅", "입고완료", int(status_series.str.contains("입고완료", na=False).sum()), "처리 완료", "green")

    st.markdown('<div class="section-title">입고 품목 검색 · 필터</div>', unsafe_allow_html=True)
    a, b, c = st.columns([1.8, 1, 1])
    with a:
        q = st.text_input("입고 상품 검색", placeholder="상품명, 주문처, 메모 검색", label_visibility="collapsed")
    with b:
        opts = ["전체"] + sorted([x for x in status_series.unique() if x])
        status = st.selectbox("입고상태", opts, label_visibility="collapsed")
    with c:
        vendor_opts = ["전체"] + sorted(df.get("주문처", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        vendor = st.selectbox("주문처", vendor_opts, label_visibility="collapsed")

    view = search_df(df, q)
    if status != "전체" and "입고상태" in view:
        view = view[view["입고상태"].astype(str) == status]
    if vendor != "전체" and "주문처" in view:
        view = view[view["주문처"].astype(str) == vendor]

    st.caption(f"조건에 맞는 품목 {len(view):,}개")
    cols = [c for c in ["주문처", "주문날짜", "상품명", "주문수량", "입고수량", "입고상태", "메모", "유통기한/유기", "실제온수량", "비고"] if c in view.columns]
    edited = st.data_editor(
        view[cols],
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        height=585,
        column_config={
            "입고상태": st.column_config.SelectboxColumn("입고상태", options=["미입고", "부분입고", "입고완료", "확인필요"]),
            "상품명": st.column_config.TextColumn("상품명", width="large"),
            "메모": st.column_config.TextColumn("메모", width="medium"),
        },
    )
    if st.button("💾 프리뷰 편집내용 저장", key="save_incoming"):
        merged = df.copy()
        for idx in edited.index:
            if idx in merged.index:
                for col in edited.columns:
                    merged.loc[idx, col] = edited.loc[idx, col]
        save_preview("입고관리", merged)
        st.success("원본 Excel은 그대로 두고 프리뷰 저장본에 저장했어요.")


# -----------------------------------------------------------------------------
# ORDERS
# -----------------------------------------------------------------------------
elif page == "🛒  주문센터":
    oq = orders(str(XLSX))
    ow = owm(str(XLSX))
    cp = coupang(str(XLSX))
    dr = direct_orders(str(XLSX))
    pending = int((~oq.get("주문완료", pd.Series(False, index=oq.index)).apply(truthy)).sum()) if not oq.empty else 0

    page_header("주문센터", "주문해야 하는 약부터 OWM · 쿠팡 · 직거래 내역까지 한 곳에서 관리해요.", "ORDER CENTER", f"주문 대기 {pending}")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        kpi_card("📝", "주문 대기", pending, "주문완료 전", "rose")
    with m2:
        kpi_card("🌿", "OWM 내역", len(ow), "등록된 구매 품목", "green")
    with m3:
        kpi_card("📦", "쿠팡 내역", len(cp), "구매 기록", "blue")
    with m4:
        kpi_card("🤝", "직거래", len(dr), "직거래 주문 행", "yellow")

    t1, t2, t3, t4 = st.tabs(["주문해야 하는 약", "OWM 구매내역", "쿠팡", "직거래 주문"])

    with t1:
        df = oq.copy()
        saved = load_preview("주문해야하는약")
        if saved is not None:
            df = saved
        q = st.text_input("약 검색", placeholder="제품명 또는 주문채널 검색", key="orderq")
        view = search_df(df, q)
        edited = st.data_editor(
            view,
            use_container_width=True,
            hide_index=True,
            height=535,
            column_config={
                "장바구니 넣기": st.column_config.CheckboxColumn("장바구니 넣기"),
                "장바구니": st.column_config.CheckboxColumn("장바구니"),
                "주문완료": st.column_config.CheckboxColumn("주문완료"),
                "입고": st.column_config.CheckboxColumn("입고"),
                "제품명": st.column_config.TextColumn("제품명", width="large"),
            },
        )
        if st.button("💾 주문 프리뷰 저장"):
            merged = df.copy()
            for idx in edited.index:
                if idx in merged.index:
                    for col in edited.columns:
                        merged.loc[idx, col] = edited.loc[idx, col]
            save_preview("주문해야하는약", merged)
            st.success("프리뷰 저장 완료")

    with t2:
        q = st.text_input("OWM 구매내역 검색", placeholder="상품명 또는 주문번호", key="owmq")
        view = search_df(ow, q)
        if "총액" in view.columns:
            total = pd.to_numeric(view["총액"], errors="coerce").sum()
            st.markdown(f"<div class='page-chip'>표시 합계 · {total:,.0f}원</div>", unsafe_allow_html=True)
        st.dataframe(
            view,
            use_container_width=True,
            hide_index=True,
            height=540,
            column_config={
                "이미지주소": st.column_config.ImageColumn("상품 이미지"),
                "상품명": st.column_config.TextColumn("상품명", width="large"),
            },
        )

    with t3:
        q = st.text_input("쿠팡 구매내역 검색", placeholder="상품명 검색", key="cq")
        view = search_df(cp, q)
        st.dataframe(
            view,
            use_container_width=True,
            hide_index=True,
            height=540,
            column_config={
                "이미지주소": st.column_config.ImageColumn("이미지"),
                "입고확인": st.column_config.CheckboxColumn("입고확인"),
                "상품명": st.column_config.TextColumn("상품명", width="large"),
            },
        )

    with t4:
        q = st.text_input("직거래 주문 검색", placeholder="업체 또는 제품 검색", key="dq")
        view = search_df(dr, q)
        st.dataframe(view, use_container_width=True, hide_index=True, height=540)


# -----------------------------------------------------------------------------
# INFLUENCERS
# -----------------------------------------------------------------------------
elif page == "✨  인플루언서":
    df = influencers(str(XLSX))
    page_header("인플루언서 일정", "여러 업체에서 들어오는 방문 일정과 협업 정보를 카드처럼 빠르게 확인해요.", "VISIT CALENDAR", f"총 {len(df)}건")

    a, b = st.columns([2, 1])
    with a:
        q = st.text_input("인플루언서 검색", placeholder="닉네임 · 국가 · 플랫폼 · 제품 검색", label_visibility="collapsed")
    with b:
        sources = ["전체"] + sorted([str(x) for x in df.get("분류", pd.Series(dtype=str)).dropna().unique() if str(x).strip()])
        source = st.selectbox("업체/분류", sources, label_visibility="collapsed")

    view = search_df(df, q)
    if source != "전체" and "분류" in view.columns:
        view = view[view["분류"].astype(str) == source]

    st.caption(f"검색 결과 {len(view)}건")
    cards = view.head(12).to_dict("records")
    grid = st.columns(2)
    for i, r in enumerate(cards):
        d = loose_date(r.get("날짜"))
        date_text = d.strftime("%Y.%m.%d") if d else compact_text(r.get("날짜", "날짜 미정"), 20)
        vals = [r.get("닉네임/이름"), r.get("팔로워/언어"), r.get("분류")]
        name = next((compact_text(v, 28) for v in vals if nonempty(v)), "방문 일정")
        detail_parts = [r.get("분류"), r.get("플랫폼/특이사항"), r.get("제품/가이드")]
        detail = " · ".join(compact_text(v, 30) for v in detail_parts if nonempty(v))
        with grid[i % 2]:
            st.markdown(
                f"""
                <div class="soft-card">
                  <div style="display:flex;justify-content:space-between;align-items:center;gap:10px">
                    <span class="badge badge-yellow">{esc(date_text)}</span>
                    <span style="font-size:.66rem;color:#9A9D98">{esc(compact_text(r.get('시간',''),12))}</span>
                  </div>
                  <div class="product-name" style="font-size:.94rem">{esc(name)}</div>
                  <div class="subtext">{esc(detail)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    with st.expander("전체 일정 표로 보기"):
        st.dataframe(view, use_container_width=True, hide_index=True, height=500)


# -----------------------------------------------------------------------------
# PRODUCT FINDER
# -----------------------------------------------------------------------------
elif page == "🔎  상품 찾기":
    df = vendors(str(XLSX))
    page_header("어디서 주문하지?", "상품명만 입력하면 직거래 업체와 주문 힌트를 바로 찾아요.", "QUICK FINDER", f"등록 상품 {len(df)}개")

    with st.container(border=True):
        st.markdown('<div style="font-size:.78rem;color:#8A908B;margin-bottom:7px">상품 또는 업체 통합 검색</div>', unsafe_allow_html=True)
        q = st.text_input("상품명 또는 업체명", placeholder="🔎  예: 센시아, 마데카솔, 레스노베…", label_visibility="collapsed")

    if q:
        view = search_df(df, q)
        st.markdown(f'<div class="section-title">검색 결과 · {len(view)}건</div>', unsafe_allow_html=True)
        render_product_results(view, 24)
        if len(view) > 24:
            with st.expander("전체 검색 결과 표로 보기"):
                st.dataframe(view, use_container_width=True, hide_index=True)
    else:
        st.markdown('<div class="section-title">거래처 빠른 보기</div>', unsafe_allow_html=True)
        vendor_counts = df.groupby("업체", dropna=True).size().sort_values(ascending=False).head(12)
        cols = st.columns(4)
        for i, (name, count) in enumerate(vendor_counts.items()):
            with cols[i % 4]:
                st.markdown(
                    f"""
                    <div class="soft-card" style="min-height:92px">
                      <span class="badge">VENDOR</span>
                      <div class="product-name">{esc(name)}</div>
                      <div class="subtext">등록 상품 {count}개</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        st.caption("찾고 싶은 상품명을 위 검색창에 입력하면 실제 거래처가 바로 나와요.")


# -----------------------------------------------------------------------------
# SUBSTITUTION
# -----------------------------------------------------------------------------
elif page == "💊  대체조제":
    df = substitutes(str(XLSX))
    saved = load_preview("대체보고")
    if saved is not None:
        df = saved

    checked = int(df.get("확인", pd.Series(dtype=bool)).apply(truthy).sum()) if not df.empty else 0
    total = len(df)
    pct = round(checked / total * 100) if total else 0
    page_header("대체조제 확인", "날짜별 통보 확인 여부와 확인자를 빠르게 체크해요.", "SUBSTITUTION CHECK", f"확인 {pct}%")

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("✅", "확인 완료", checked, "체크된 날짜", "green")
    with c2:
        kpi_card("⏳", "미확인", max(total - checked, 0), "추가 확인 필요", "rose")
    with c3:
        kpi_card("📆", "관리 일수", total, "현재 표에 등록", "blue")

    edited = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        height=560,
        column_config={
            "확인": st.column_config.CheckboxColumn("확인"),
            "일": st.column_config.NumberColumn("일", format="%d"),
            "확인자": st.column_config.TextColumn("확인자", width="medium"),
        },
    )
    if st.button("💾 대체조제 프리뷰 저장"):
        save_preview("대체보고", edited)
        st.success("프리뷰 저장 완료")


# -----------------------------------------------------------------------------
# SETS
# -----------------------------------------------------------------------------
elif page == "🎁  세트상품":
    df = sets(str(XLSX))
    page_header("세트상품", "세트별 구성품과 가격을 묶어서 한눈에 확인해요.", "SET GUIDE")
    q = st.text_input("세트명 또는 구성품 검색", placeholder="예: 헛개, 남자불, 반하…", label_visibility="collapsed")
    view = search_df(df, q)

    if "세트상품명" in view.columns:
        groups = list(view.groupby("세트상품명", dropna=True))
        st.caption(f"표시 중인 세트 {len(groups)}개")
        cols = st.columns(2)
        for i, (name, g) in enumerate(groups[:30]):
            comp = []
            for _, r in g.iterrows():
                product = r.get("구성상품명")
                price = r.get("구성품가격")
                if nonempty(product):
                    comp.append(f"{compact_text(product, 34)} · {money(price)}")
            if not comp:
                continue
            sale = next((r.get("세트판매가격") for _, r in g.iterrows() if nonempty(r.get("세트판매가격"))), None)
            price_line = f"<div style='font-family:Nunito;font-weight:900;color:#8A6A13;margin-top:8px'>{esc(money(sale))}</div>" if nonempty(sale) else ""
            with cols[i % 2]:
                st.markdown(
                    f"""
                    <div class="soft-card">
                      <span class="badge badge-yellow">SET</span>
                      <div class="product-name" style="font-size:1rem">{esc(name)}</div>
                      <div class="subtext">{'<br>'.join(esc(x) for x in comp[:8])}</div>
                      {price_line}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        if len(groups) > 30:
            with st.expander("전체 세트 데이터 보기"):
                st.dataframe(view, use_container_width=True, hide_index=True)
    else:
        st.dataframe(view, use_container_width=True, hide_index=True)


# -----------------------------------------------------------------------------
# SUNNY NOTE
# -----------------------------------------------------------------------------
elif page == "📚  SUNNY NOTE":
    df = hub(str(XLSX))
    page_header("SUNNY NOTE", "HUB에 있던 업무 메모와 응대 문구를 검색하기 쉽게 정리했어요.", "KNOWLEDGE", "민감정보 제외")
    q = st.text_input("메모 검색", placeholder="예: Tax Free, 영수증, 결제, antihistamine…", label_visibility="collapsed")
    view = search_df(df, q)

    if view.empty:
        st.info("검색 결과가 없어요.")
    else:
        cols = st.columns(2)
        for i, text in enumerate(view["내용"].head(60)):
            with cols[i % 2]:
                st.markdown(
                    f"""
                    <div class="soft-card">
                      <span class="badge">NOTE</span>
                      <div class="subtext" style="font-size:.77rem;color:#58635B;margin-top:7px">{esc(text)}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


# -----------------------------------------------------------------------------
# RAW EXCEL EXPLORER
# -----------------------------------------------------------------------------
elif page == "🗂️  원본 시트":
    page_header("원본 시트 보기", "앱으로 가공하기 전 Excel 내용을 그대로 확인할 수 있어요.", "EXCEL EXPLORER", "비밀번호 시트 제외")
    names = available_sheets(str(XLSX))
    a, b = st.columns([1, 2])
    with a:
        sheet = st.selectbox("시트 선택", names)
    try:
        df = clean_empty(load_sheet(sheet, str(XLSX)))
        with b:
            q = st.text_input("현재 시트 검색", placeholder="현재 시트 안에서 검색", key="rawq")
        view = search_df(df, q)
        st.caption(f"{len(view):,}행 × {len(view.columns):,}열")
        st.dataframe(view, use_container_width=True, hide_index=True, height=640)
    except Exception as e:
        st.error(f"이 시트는 구조가 특수해서 표 형태로 읽지 못했어요: {e}")
