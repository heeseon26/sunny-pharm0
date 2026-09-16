from __future__ import annotations

from pathlib import Path
import re
import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_EXCEL = BASE_DIR / 'data' / 'SUNNY.xlsx'
EDIT_DIR = BASE_DIR / 'preview_edits'
EDIT_DIR.mkdir(exist_ok=True)

BLOCKED_SHEETS = {'홈페이지 아이디, 비번'}


def _dedupe_headers(values):
    seen = {}
    out = []
    for i, raw in enumerate(values):
        name = '' if pd.isna(raw) else str(raw).strip()
        if not name or name.lower().startswith('unnamed'):
            name = f'컬럼_{i+1}'
        if name in seen:
            seen[name] += 1
            name = f'{name}_{seen[name]}'
        else:
            seen[name] = 1
        out.append(name)
    return out


def load_sheet(sheet_name: str, path: Path | str = DEFAULT_EXCEL, header: int | None = 0) -> pd.DataFrame:
    path = Path(path)
    if sheet_name in BLOCKED_SHEETS:
        return pd.DataFrame()
    df = pd.read_excel(path, sheet_name=sheet_name, header=header, engine='openpyxl')
    if header is None:
        return df
    df.columns = _dedupe_headers(df.columns)
    return df


def available_sheets(path: Path | str = DEFAULT_EXCEL):
    path = Path(path)
    xls = pd.ExcelFile(path, engine='openpyxl')
    return [s for s in xls.sheet_names if s not in BLOCKED_SHEETS]


def load_all(path: Path | str = DEFAULT_EXCEL):
    return {s: load_sheet(s, path) for s in available_sheets(path)}


def clean_empty(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df = df.dropna(axis=0, how='all').dropna(axis=1, how='all')
    return df.reset_index(drop=True)


def parse_incoming(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name='입고관리', header=None, engine='openpyxl')
    if raw.empty:
        return pd.DataFrame()
    raw = raw.iloc[1:].copy()
    raw = raw.dropna(how='all')
    raw.columns = [f'c{i}' for i in range(raw.shape[1])]
    def col(i):
        return raw[f'c{i}'] if f'c{i}' in raw else pd.Series(index=raw.index, dtype=object)
    out = pd.DataFrame({
        '주문처': col(0),
        '주문날짜': col(1),
        '택배도착날짜': col(2),
        '이미지주소': col(3),
        '상품명': col(5),
        '주문수량': col(6),
        '단가': col(7),
        '총액': col(8),
        '입고기록': col(9),
        '입고수량': col(10),
        '입고상태': col(14),
        '메모': col(15),
        '명세서': col(16),
        '택배': col(17),
        '유통기한/유기': col(18),
        '실제온수량': col(19),
        '비고': col(20),
    })
    out = out[out['상품명'].notna() | out['주문처'].notna()].reset_index(drop=True)
    out['주문처'] = out['주문처'].ffill()
    out['주문날짜'] = out['주문날짜'].ffill()
    return out


def parse_order_queue(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    df = clean_empty(load_sheet('주문해야하는약', path))
    return df


def parse_owm(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    df = clean_empty(load_sheet('OWM', path))
    keep = [c for c in ['주문번호','주문일시','이미지주소','상품명','수량','단가','총액'] if c in df.columns]
    return df[keep] if keep else df


def parse_coupang(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name='쿠팡', header=None, engine='openpyxl')
    raw = raw.iloc[1:].copy().dropna(how='all')
    raw.columns = [f'c{i}' for i in range(raw.shape[1])]
    def col(i):
        return raw[f'c{i}'] if f'c{i}' in raw else pd.Series(index=raw.index, dtype=object)
    return pd.DataFrame({
        '구분': col(0), '주문날짜': col(1), '이미지주소': col(3), '상품명': col(4),
        '단가': col(5), '주문수량': col(6), '총액': col(7), '입고확인': col(8),
        '입고날짜': col(9), '입고수량': col(10), '비고': col(11),
    }).reset_index(drop=True)


def parse_direct_orders(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    df = clean_empty(load_sheet('직거래주문관리', path))
    if df.empty:
        return df
    for c in ['업체', '카톡방명', '주문날짜']:
        if c in df.columns:
            df[c] = df[c].ffill()
    return df


def parse_vendor_products(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name='직거래업체', header=None, engine='openpyxl')
    rows = []
    for _, r in raw.iterrows():
        vals = [None if pd.isna(v) else str(v).strip() for v in r.tolist()]
        if not vals or not vals[0]:
            continue
        vendor = vals[0]
        alias = vals[1] if len(vals)>1 else None
        contact = vals[2] if len(vals)>2 else None
        for p in vals[3:]:
            if p:
                rows.append({'업체': vendor, '구분/방명': alias, '공급처/메모': contact, '상품명': p})
        if not any(vals[3:]):
            rows.append({'업체': vendor, '구분/방명': alias, '공급처/메모': contact, '상품명': ''})
    return pd.DataFrame(rows)


def parse_substitution(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name='대체보고', header=None, engine='openpyxl')
    if raw.shape[1] < 7:
        return pd.DataFrame()
    cal = raw.iloc[1:, 3:7].copy()
    cal.columns = ['일','요일','확인','확인자']
    cal = cal[cal['일'].notna()].reset_index(drop=True)
    return cal


def parse_influencers(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name='인플루언서 리스트', header=None, engine='openpyxl')
    raw = raw.iloc[1:].dropna(how='all').reset_index(drop=True)
    raw.columns = [f'항목{i+1}' for i in range(raw.shape[1])]
    labels = ['분류','날짜','시간','장소/시간2','국가/시간3','팔로워/언어','닉네임/이름','플랫폼/특이사항','링크/기프트','제품/가이드']
    rename = {f'항목{i+1}': labels[i] for i in range(min(len(labels), raw.shape[1]))}
    raw = raw.rename(columns=rename)
    return raw


def parse_sets(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    df = clean_empty(load_sheet('세트상품', path))
    if df.empty:
        return df
    if '세트상품명' in df.columns:
        df['세트상품명'] = df['세트상품명'].ffill()
    return df


def parse_hub_safe(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name='❤️ HUB❤️', header=None, engine='openpyxl')
    forbidden = re.compile(r'(비번|password|wi-?fi\s*password|접속\s*코드|passcode|id\s*:)', re.I)
    rows=[]
    for r in raw.itertuples(index=False, name=None):
        vals=[]
        blocked=False
        for v in r:
            if pd.isna(v):
                vals.append('')
                continue
            s=str(v).strip()
            if forbidden.search(s):
                blocked=True
            vals.append(s)
        if blocked:
            continue
        nonempty=[v for v in vals if v]
        if nonempty:
            rows.append({'내용': '  ·  '.join(nonempty[:4])})
    return pd.DataFrame(rows)


def parse_burn_down(path: Path | str = DEFAULT_EXCEL) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name='소진해야하는 상품', header=None, engine='openpyxl')
    raw = raw.dropna(how='all')
    return pd.DataFrame({'상품명': raw.iloc[:,0].astype(str).tolist()}) if not raw.empty else pd.DataFrame()


def save_preview(name: str, df: pd.DataFrame):
    safe = re.sub(r'[^0-9A-Za-z가-힣_-]+','_',name)
    path = EDIT_DIR / f'{safe}.csv'
    df.to_csv(path, index=False, encoding='utf-8-sig')
    return path


def load_preview(name: str):
    safe = re.sub(r'[^0-9A-Za-z가-힣_-]+','_',name)
    path = EDIT_DIR / f'{safe}.csv'
    if path.exists():
        return pd.read_csv(path)
    return None
