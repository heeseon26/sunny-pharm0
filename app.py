from __future__ import annotations

from pathlib import Path
from datetime import datetime
import math
import pandas as pd
import streamlit as st

from services.excel_source import (
    DEFAULT_EXCEL, available_sheets, load_sheet, clean_empty,
    parse_incoming, parse_order_queue, parse_owm, parse_coupang,
    parse_direct_orders, parse_vendor_products, parse_substitution,
    parse_influencers, parse_sets, parse_hub_safe, parse_burn_down,
    save_preview, load_preview,
)

BASE = Path(__file__).resolve().parent
PROFILE = BASE / 'assets' / 'sunny_profile.png'

st.set_page_config(
    page_title='SUNNY PHARM 365',
    page_icon='☀️',
    layout='wide',
    initial_sidebar_state='expanded',
)

# Optional preview password. Set APP_PASSWORD in Streamlit Cloud Secrets if desired.
try:
    _app_password = st.secrets.get('APP_PASSWORD', '')
except Exception:
    _app_password = ''
if _app_password:
    if not st.session_state.get('_sunny_auth_ok', False):
        st.markdown('## ☀️ SUNNY PHARM 365')
        _pw = st.text_input('프리뷰 비밀번호', type='password')
        if st.button('앱 열기'):
            if _pw == _app_password:
                st.session_state['_sunny_auth_ok'] = True
                st.rerun()
            else:
                st.error('비밀번호가 맞지 않아요.')
        st.stop()

# ---------- design ----------
st.markdown('''
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700;800&family=Nunito:wght@600;700;800&display=swap');
:root {
  --sun:#F7C948; --sun2:#FFE9A6; --cream:#FFFDF6; --paper:#FFFFFF;
  --sage:#759A80; --sage-soft:#EAF3EC; --ink:#26352B; --muted:#758078;
  --line:#EAE8DD; --rose:#F6DED8;
}
html, body, [class*="css"] {font-family:'Noto Sans KR',sans-serif;}
.stApp {background:linear-gradient(145deg,#FFFDF7 0%,#F8FBF7 58%,#FFF9E8 100%); color:var(--ink);}
[data-testid="stSidebar"] {background:#FFF9E8; border-right:1px solid #F1E7C8;}
[data-testid="stSidebar"] > div:first-child {padding-top:1.2rem;}
#MainMenu, footer {visibility:hidden;}
.block-container {padding-top:1.7rem; padding-bottom:3rem; max-width:1450px;}
h1,h2,h3 {color:var(--ink); letter-spacing:-.03em;}
.sunny-hero {background:linear-gradient(135deg,#FFF3B9 0%,#FFFDF7 46%,#EAF4EC 100%); border:1px solid #F2E7C5; border-radius:30px; padding:28px 30px; margin-bottom:22px; box-shadow:0 16px 45px rgba(61,72,59,.07);}
.eyebrow {font-family:'Nunito',sans-serif; font-size:.78rem; font-weight:800; letter-spacing:.12em; color:#9B7A14; margin-bottom:7px;}
.hero-title {font-size:2.35rem; font-weight:800; letter-spacing:-.045em; margin:0; color:#24352A;}
.hero-copy {margin-top:9px; color:#667268; font-size:.97rem; line-height:1.7;}
.kpi {background:rgba(255,255,255,.87); border:1px solid #ECEADF; border-radius:21px; padding:18px 19px; min-height:120px; box-shadow:0 8px 24px rgba(59,74,62,.045);}
.kpi .label {font-size:.78rem; color:#7C857E; font-weight:700; margin-bottom:10px;}
.kpi .num {font-family:'Nunito',sans-serif; font-size:2rem; font-weight:800; color:#2C4033; line-height:1;}
.kpi .hint {font-size:.74rem; color:#9A9D96; margin-top:9px;}
.section-head {display:flex; align-items:center; gap:10px; font-size:1.22rem; font-weight:800; margin:14px 0 10px;}
.soft-card {background:#fff; border:1px solid #ECEADF; border-radius:20px; padding:17px 18px; margin-bottom:10px; box-shadow:0 6px 22px rgba(50,68,55,.04);}
.badge {display:inline-flex; border-radius:999px; padding:5px 10px; font-size:.72rem; font-weight:800; background:#EEF5EF; color:#597663;}
.badge-yellow {background:#FFF0B7; color:#916D00;}
.badge-red {background:#FCE9E5; color:#A35549;}
.product-name {font-weight:800; color:#304237; font-size:.96rem; margin:7px 0 3px;}
.subtext {font-size:.78rem; color:#838B85; line-height:1.55;}
.notice {background:#F3F7F3; border:1px solid #DDEADF; border-radius:16px; padding:13px 15px; color:#637269; font-size:.82rem;}
.stButton>button {border-radius:13px; border:1px solid #E6D99B; background:#FFF8D9; color:#624F0F; font-weight:800;}
.stButton>button:hover {border-color:#DCC46B; color:#4E3F0E;}
div[data-testid="stMetric"] {background:#fff; border:1px solid #eceae0; padding:14px; border-radius:18px;}
[data-testid="stDataFrame"], [data-testid="stDataEditor"] {border-radius:18px; overflow:hidden; border:1px solid #ECEADF;}
.stTabs [data-baseweb="tab-list"] {gap:6px; background:#F4F1E7; border-radius:14px; padding:5px;}
.stTabs [data-baseweb="tab"] {border-radius:10px; padding:8px 14px;}
.stTabs [aria-selected="true"] {background:white;}
hr {border-color:#eeeade;}
</style>
''', unsafe_allow_html=True)

# ---------- source ----------
@st.cache_resource(show_spinner=False)
def materialize_upload(uploaded):
    out = BASE / 'data' / '_uploaded_preview.xlsx'
    out.write_bytes(uploaded.getvalue())
    return out

if 'excel_path' not in st.session_state:
    st.session_state.excel_path = str(DEFAULT_EXCEL) if DEFAULT_EXCEL.exists() else None

with st.sidebar:
    if PROFILE.exists():
        st.image(str(PROFILE), use_container_width=True)
    st.markdown("<div style='text-align:center;font-family:Nunito;font-size:1.12rem;font-weight:800;color:#35473B;margin-top:-6px'>SUNNY PHARM 365</div>", unsafe_allow_html=True)
    st.markdown("<div style='text-align:center;font-size:.72rem;color:#94886C;margin-bottom:14px'>Pharmacist workspace · GitHub preview</div>", unsafe_allow_html=True)

    uploaded = st.file_uploader('다른 SUNNY 엑셀로 바꾸기', type=['xlsx'], label_visibility='collapsed')
    if uploaded is not None:
        st.session_state.excel_path = str(materialize_upload(uploaded))
        st.cache_data.clear()

    pages = ['☀️ 오늘의 SUNNY','📦 입고관리','🛒 주문센터','🔎 상품 · 거래처','✨ 인플루언서','🎁 세트상품','✅ 대체조제','📚 SUNNY NOTE','🗂️ 원본 시트']
    page = st.radio('메뉴', pages, label_visibility='collapsed')
    st.markdown('---')
    st.caption('현재 데이터')
    st.write('`SUNNY.xlsx`' if st.session_state.excel_path else '연결된 파일 없음')
    st.caption('※ 비밀번호 시트는 앱에서 제외됩니다.')

if not st.session_state.excel_path or not Path(st.session_state.excel_path).exists():
    st.warning('SUNNY 엑셀 파일을 왼쪽에서 업로드해 주세요.')
    st.stop()

XLSX = Path(st.session_state.excel_path)

@st.cache_data(show_spinner=False)
def incoming(p): return parse_incoming(p)
@st.cache_data(show_spinner=False)
def orders(p): return parse_order_queue(p)
@st.cache_data(show_spinner=False)
def owm(p): return parse_owm(p)
@st.cache_data(show_spinner=False)
def coupang(p): return parse_coupang(p)
@st.cache_data(show_spinner=False)
def direct_orders(p): return parse_direct_orders(p)
@st.cache_data(show_spinner=False)
def vendors(p): return parse_vendor_products(p)
@st.cache_data(show_spinner=False)
def substitutes(p): return parse_substitution(p)
@st.cache_data(show_spinner=False)
def influencers(p): return parse_influencers(p)
@st.cache_data(show_spinner=False)
def sets(p): return parse_sets(p)
@st.cache_data(show_spinner=False)
def hub(p): return parse_hub_safe(p)
@st.cache_data(show_spinner=False)
def burn(p): return parse_burn_down(p)


def safe_int(x):
    try:
        return int(float(x)) if pd.notna(x) else 0
    except Exception:
        return 0


def money(x):
    try:
        if pd.isna(x): return '-'
        return f'{float(x):,.0f}원'
    except Exception:
        return str(x)


def truthy(v):
    if isinstance(v, bool): return v
    return str(v).strip().lower() in {'true','1','yes','y','완료','입고완료'}


def hero(title, desc, eyebrow='SUNNY PHARM 365'):
    st.markdown(f"""<div class='sunny-hero'><div class='eyebrow'>{eyebrow}</div><div class='hero-title'>{title}</div><div class='hero-copy'>{desc}</div></div>""", unsafe_allow_html=True)


def kpi(label, value, hint='', cls=''):
    st.markdown(f"<div class='kpi'><div class='label'>{label}</div><div class='num'>{value}</div><div class='hint'>{hint}</div></div>", unsafe_allow_html=True)


def search_df(df, q):
    if df.empty or not q: return df
    mask = df.astype(str).apply(lambda c: c.str.contains(q, case=False, na=False)).any(axis=1)
    return df[mask]

# ---------- HOME ----------
if page == '☀️ 오늘의 SUNNY':
    hero('Good day, SUNNY ☀️', '엑셀에 흩어진 입고 · 주문 · 방문 일정을 한 화면에서 빠르게 확인해요. 지금은 비밀번호 시트를 제거한 SUNNY 엑셀 스냅샷을 직접 읽는 GitHub 프리뷰예요.')
    inc, oq, inf, bd = incoming(str(XLSX)), orders(str(XLSX)), influencers(str(XLSX)), burn(str(XLSX))
    pending_inc = 0
    partial = 0
    if not inc.empty and '입고상태' in inc:
        s=inc['입고상태'].fillna('').astype(str)
        pending_inc=int((s.str.contains('미입고',na=False)).sum())
        partial=int((s.str.contains('부분',na=False)).sum())
    pending_order=0
    if not oq.empty and '주문완료' in oq:
        pending_order=int((~oq['주문완료'].apply(truthy)).sum())
    c1,c2,c3,c4=st.columns(4)
    with c1: kpi('미입고', pending_inc, '아직 도착하지 않은 품목')
    with c2: kpi('부분입고', partial, '추가 확인이 필요한 품목')
    with c3: kpi('주문 대기', pending_order, '주문완료 체크 전')
    with c4: kpi('소진 우선', len(bd), '먼저 안내하고 싶은 상품')

    st.markdown("<div class='section-head'>🌤️ 지금 먼저 볼 것</div>", unsafe_allow_html=True)
    left,right=st.columns([1.25,1])
    with left:
        if not inc.empty:
            priority=inc[inc['입고상태'].fillna('').astype(str).str.contains('미입고|부분입고', regex=True, na=False)].head(7)
            if priority.empty:
                st.info('현재 표시할 미입고/부분입고 품목이 없어요.')
            else:
                for _,r in priority.iterrows():
                    stat=str(r.get('입고상태',''))
                    badge='badge-red' if '미입고' in stat else 'badge-yellow'
                    st.markdown(f"<div class='soft-card'><span class='badge {badge}'>{stat or '확인필요'}</span><div class='product-name'>{r.get('상품명','')}</div><div class='subtext'>{r.get('주문처','')} · 주문수량 {r.get('주문수량','-')} · 실제온수량 {r.get('실제온수량','-')}</div></div>", unsafe_allow_html=True)
    with right:
        st.markdown("<div class='soft-card'><span class='badge badge-yellow'>SUNNY PICK</span><div class='product-name'>소진해야 하는 상품</div>", unsafe_allow_html=True)
        if bd.empty:
            st.caption('등록된 항목이 없어요.')
        else:
            for p in bd['상품명'].head(8): st.write(f'• {p}')
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("<div class='notice'>💡 이 프리뷰에서는 GitHub에 포함된 안전한 Excel 스냅샷을 읽습니다. 원본 엑셀은 건드리지 않습니다. 체크/수정 테스트는 별도 프리뷰 저장본에 저장돼요. Google Sheets 권한을 받으면 같은 UI에서 실제 시트로 바로 저장하도록 바꿀 수 있어요.</div>", unsafe_allow_html=True)

# ---------- INCOMING ----------
elif page == '📦 입고관리':
    hero('입고관리', '상품명으로 검색하고, 미입고·부분입고만 골라서 빠르게 확인해요.', 'DELIVERY DESK')
    df = incoming(str(XLSX))
    saved = load_preview('입고관리')
    if saved is not None: df=saved
    a,b=st.columns([2,1])
    with a: q=st.text_input('입고 상품 검색', placeholder='예: 도미나, 미녹시딜, 태극…')
    with b:
        opts=['전체']+sorted([x for x in df.get('입고상태',pd.Series()).dropna().astype(str).unique() if x])
        status=st.selectbox('입고상태',opts)
    view=search_df(df,q)
    if status!='전체' and '입고상태' in view: view=view[view['입고상태'].astype(str)==status]
    st.caption(f'{len(view):,}개 품목')
    cols=[c for c in ['주문처','주문날짜','상품명','주문수량','입고수량','입고상태','메모','유통기한/유기','실제온수량','비고'] if c in view.columns]
    edited=st.data_editor(view[cols], use_container_width=True, hide_index=True, num_rows='fixed', height=590,
        column_config={
            '입고상태': st.column_config.SelectboxColumn('입고상태', options=['미입고','부분입고','입고완료','확인필요']),
            '상품명': st.column_config.TextColumn('상품명', width='large'),
        })
    if st.button('💾 프리뷰 편집내용 저장', key='save_incoming'):
        merged = df.copy()
        for idx in edited.index:
            if idx in merged.index:
                for c in edited.columns:
                    merged.loc[idx, c] = edited.loc[idx, c]
        save_preview('입고관리', merged)
        st.success('원본 Excel은 그대로 두고 프리뷰 저장본에 저장했어요.')

# ---------- ORDERS ----------
elif page == '🛒 주문센터':
    hero('주문센터', '주문해야 할 약, OWM 구매내역, 쿠팡, 직거래 주문을 한 곳에서 봐요.', 'ORDER CENTER')
    t1,t2,t3,t4=st.tabs(['주문해야 하는 약','OWM','쿠팡','직거래 주문'])
    with t1:
        df=orders(str(XLSX)); saved=load_preview('주문해야하는약'); df=saved if saved is not None else df
        q=st.text_input('약 검색', placeholder='제품명 또는 주문채널', key='orderq')
        view=search_df(df,q)
        edited=st.data_editor(view, use_container_width=True, hide_index=True, height=560,
            column_config={
                '장바구니 넣기':st.column_config.CheckboxColumn('장바구니 넣기'),
                '장바구니':st.column_config.CheckboxColumn('장바구니'),
                '주문완료':st.column_config.CheckboxColumn('주문완료'),
                '입고':st.column_config.CheckboxColumn('입고'),
                '제품명':st.column_config.TextColumn('제품명',width='large'),
            })
        if st.button('💾 주문 프리뷰 저장'):
            merged = df.copy()
            for idx in edited.index:
                if idx in merged.index:
                    for c in edited.columns:
                        merged.loc[idx, c] = edited.loc[idx, c]
            save_preview('주문해야하는약', merged)
            st.success('프리뷰 저장 완료')
    with t2:
        df=owm(str(XLSX)); q=st.text_input('OWM 구매내역 검색', key='owmq'); view=search_df(df,q)
        if '총액' in view.columns:
            total=pd.to_numeric(view['총액'],errors='coerce').sum(); st.metric('표시된 구매 총액',f'{total:,.0f}원')
        st.dataframe(view, use_container_width=True, hide_index=True, height=560,
            column_config={'이미지주소':st.column_config.ImageColumn('상품 이미지'), '상품명':st.column_config.TextColumn('상품명',width='large')})
    with t3:
        df=coupang(str(XLSX)); q=st.text_input('쿠팡 구매내역 검색',key='cq'); view=search_df(df,q)
        st.dataframe(view,use_container_width=True,hide_index=True,height=560,
            column_config={'이미지주소':st.column_config.ImageColumn('이미지'),'입고확인':st.column_config.CheckboxColumn('입고확인'),'상품명':st.column_config.TextColumn('상품명',width='large')})
    with t4:
        df=direct_orders(str(XLSX)); q=st.text_input('직거래 주문 검색',key='dq'); view=search_df(df,q)
        st.dataframe(view,use_container_width=True,hide_index=True,height=560)

# ---------- PRODUCTS ----------
elif page == '🔎 상품 · 거래처':
    hero('어디서 주문하지?', '상품명 하나만 입력하면 직거래 업체와 주문 힌트를 찾아줘요.', 'QUICK FINDER')
    df=vendors(str(XLSX))
    q=st.text_input('상품명 또는 업체명', placeholder='예: 센시아, 마데카솔, 리쥬란…')
    view=search_df(df,q) if q else df.head(30)
    if q:
        st.caption(f'검색 결과 {len(view)}건')
    st.dataframe(view,use_container_width=True,hide_index=True,height=580,
        column_config={'업체':st.column_config.TextColumn('업체',width='medium'),'상품명':st.column_config.TextColumn('상품명',width='large')})

# ---------- INFLUENCERS ----------
elif page == '✨ 인플루언서':
    hero('인플루언서 일정', '여러 업체에서 들어온 방문 일정을 한 화면에서 검색해요.', 'VISIT CALENDAR')
    df=influencers(str(XLSX)); q=st.text_input('닉네임 · 국가 · 플랫폼 · 제품 검색',placeholder='예: 중국, 샤오홍슈, 에띠어리…')
    view=search_df(df,q)
    # compact cards for first rows
    for _,r in view.head(8).iterrows():
        vals=[str(v) for v in r.iloc[:10].tolist() if pd.notna(v) and str(v).strip() not in {'','False','nan'}]
        if not vals: continue
        title=vals[0] if len(vals)==1 else ' · '.join(vals[:2])
        desc=' · '.join(vals[2:6])
        st.markdown(f"<div class='soft-card'><span class='badge'>VISIT</span><div class='product-name'>{title}</div><div class='subtext'>{desc}</div></div>",unsafe_allow_html=True)
    with st.expander('전체 일정 표로 보기', expanded=len(view)>8):
        st.dataframe(view,use_container_width=True,hide_index=True,height=520)

# ---------- SETS ----------
elif page == '🎁 세트상품':
    hero('세트상품', '세트 구성품과 가격을 묶어서 보기 쉽게 확인해요.', 'SET GUIDE')
    df=sets(str(XLSX)); q=st.text_input('세트명 또는 구성품 검색',placeholder='예: 헛개, 반하…'); view=search_df(df,q)
    if '세트상품명' in view.columns:
        for name,g in view.groupby('세트상품명',dropna=True):
            comp=[]
            for _,r in g.iterrows():
                p=r.get('구성상품명'); price=r.get('구성품가격')
                if pd.notna(p): comp.append(f"{p} · {money(price)}")
            if comp:
                st.markdown(f"<div class='soft-card'><span class='badge badge-yellow'>SET</span><div class='product-name'>{name}</div><div class='subtext'>{'<br>'.join(comp[:8])}</div></div>",unsafe_allow_html=True)
    else: st.dataframe(view,use_container_width=True)

# ---------- SUBSTITUTION ----------
elif page == '✅ 대체조제':
    hero('대체조제 확인', '날짜별 통보 확인 여부와 확인자를 빠르게 체크해요.', 'SUBSTITUTION CHECK')
    df=substitutes(str(XLSX)); saved=load_preview('대체보고'); df=saved if saved is not None else df
    edited=st.data_editor(df,use_container_width=True,hide_index=True,height=600,
        column_config={'확인':st.column_config.CheckboxColumn('확인'),'일':st.column_config.NumberColumn('일',format='%d')})
    if st.button('💾 대체조제 프리뷰 저장'):
        save_preview('대체보고',edited); st.success('프리뷰 저장 완료')

# ---------- NOTE ----------
elif page == '📚 SUNNY NOTE':
    hero('SUNNY NOTE', 'HUB에 있던 약국 정보 · 응대 문구 · 업무 메모를 검색하기 쉽게 모았어요. 비밀번호/접속코드는 자동 제외돼요.', 'KNOWLEDGE')
    df=hub(str(XLSX)); q=st.text_input('메모 검색',placeholder='예: 운전, Tax Free, antihistamine…'); view=search_df(df,q)
    for text in view['내용'].head(60) if not view.empty else []:
        st.markdown(f"<div class='soft-card'><div class='subtext' style='font-size:.9rem;color:#536158'>{text}</div></div>",unsafe_allow_html=True)

# ---------- RAW ----------
elif page == '🗂️ 원본 시트':
    hero('원본 시트 보기', '엑셀의 다른 시트도 그대로 확인할 수 있어요. “홈페이지 아이디, 비번” 시트는 의도적으로 제외돼요.', 'EXCEL EXPLORER')
    names=available_sheets(str(XLSX))
    sheet=st.selectbox('시트 선택',names)
    try:
        df=clean_empty(load_sheet(sheet,str(XLSX)))
        q=st.text_input('현재 시트 검색',key='rawq')
        view=search_df(df,q)
        st.caption(f'{len(view):,}행 × {len(view.columns):,}열')
        st.dataframe(view,use_container_width=True,hide_index=True,height=650)
    except Exception as e:
        st.error(f'이 시트는 구조가 특수해서 표 형태로 읽지 못했어요: {e}')
