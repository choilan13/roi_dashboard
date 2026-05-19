"""
process_gagang.py
주문서확인처리 xlsx → Google Sheets '전체 가공' 시트 생성
+ database 폴더에 주차별 rawdata 저장

사용법:
  python process_gagang.py                          # 기본 파일 처리 + 시트 전체 교체
  python process_gagang.py 파일경로.xlsx             # 특정 파일 처리
  python process_gagang.py 파일경로.xlsx --append   # 기존 시트에 이어 붙이기
"""

from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from week_utils import get_week_info

# ── 설정 ──────────────────────────────────────────────────────────────────────

BASE            = Path(__file__).parent
CREDENTIALS     = BASE / 'credentials.json'
SPREADSHEET_ID  = '17_oeV41FVchyFYUl6dHQyEk689tkzrA840WfWApFkQY'
REF_SHEET       = '기준'
OUTPUT_SHEET    = '전체 가공'
DB_DIR          = BASE / 'database'

DEFAULT_INPUT   = DB_DIR / '20260519_주문서확인처리_온채팀_공헌이익 확인용.xlsx'

GAGANG_HEADERS = [
    '날짜', '월', '주차', '송장번호', '채널명',
    '채널상품코드', '옵션 바코드', '브랜드', 'SKU명',
    '매출', '배송비', '실제 부과 배송',
    '원가', '수수료%', '수수료',
    '광고비', '리워드 등 추가비용', '공헌이익', '출고수량',
]

# 사방넷 주문서확인처리 xlsx 컬럼 인덱스 (0-based, 39열 고정 포맷)
COLS = {
    'shop':      1,   # 쇼핑몰명(1)
    'barcode':   2,   # 바코드
    'order_dt':  3,   # 주문일시(YYYY-MM-DD HH:MM)
    'chan_code':  8,   # 상품코드(쇼핑몰)
    'sales':    11,   # 공급단가*수량
    'sku':      14,   # 자체상품코드
    'qty':      21,   # 수량
    'ship_raw': 33,   # 배송비(수집)
    'invoice':  38,   # 송장번호
}

UPLOAD_CHUNK = 1500   # Google Sheets 1회 업로드 최대 행수

SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
]


# ── 헬퍼 ──────────────────────────────────────────────────────────────────────

def clean(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ''
    return str(v).strip()


def to_num(v, default: float = 0.0) -> float:
    try:
        f = float(v)
        return default if np.isnan(f) else f
    except (TypeError, ValueError):
        return default


def parse_fee(v) -> float:
    """수수료를 소수로 변환 (0.13 / '13%' / '13.0' 모두 처리)"""
    s = clean(v)
    if not s:
        return 0.0
    if s.endswith('%'):
        try:
            return float(s[:-1]) / 100
        except ValueError:
            return 0.0
    try:
        f = float(s)
        return f / 100 if f > 1.0 else f
    except ValueError:
        return 0.0


def safe_val(v):
    """Google Sheets 업로드용: NaN/Inf → 0"""
    if isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
        return 0
    return v


# ── 기준 시트 로드 ──────────────────────────────────────────────────────────────

def load_reference() -> tuple[dict, dict, dict]:
    """
    Google Sheets '기준' 시트에서 lookup dicts 구성.
    반환:
      exact_map : (sku, chan_code)       → ref_row
      sku_map   : sku                    → list[ref_row]
      esm_map   : (sku, chan_code[:10])  → ref_row  (ESM 앞 10자리)
    """
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_file(str(CREDENTIALS), scopes=SCOPES)
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(REF_SHEET)
    data = ws.get_all_values()
    headers, rows = data[0], data[1:]

    exact_map: dict = {}
    sku_map: dict   = {}
    esm_map: dict   = {}

    for row in rows:
        if len(row) < 9:
            continue
        d = dict(zip(headers, row))

        sku       = d.get('자체상품코드',  '').strip()
        chan_code  = d.get('채널상품코드',  '').strip()
        chan_name  = d.get('채널명',        '').strip()
        brand      = d.get('브랜드',        '').strip()
        alias      = d.get('별칭',          '').strip()
        cost_raw   = d.get('최종원가',      '').strip()
        fee_raw    = d.get('채널별 수수료', '').strip()

        ref_row = {
            '채널명':   chan_name,
            '브랜드':   brand,
            'SKU명':    alias,
            '최종원가': cost_raw,
            '수수료':   fee_raw,
        }

        # 1차 키: 채널상품코드 (100% 채워짐)
        if chan_code:
            exact_map.setdefault(chan_code, ref_row)
            if any(k in chan_name for k in ('ESM', '지마켓', '옥션')):
                esm_map.setdefault(chan_code[:10], ref_row)

        # 2차 키: 자체상품코드 (약 4% 채워짐, fallback용)
        if sku:
            sku_map.setdefault(sku, []).append(ref_row)

    print(f"  기준 시트 로드: {len(rows)}행 → exact {len(exact_map)}, SKU fallback {len(sku_map)}")
    return exact_map, sku_map, esm_map


def lookup_ref(sku: str, chan_code: str, chan_base: str,
               exact_map: dict, sku_map: dict, esm_map: dict) -> dict | None:
    """
    매칭 우선순위:
    1) 채널상품코드 완전 일치
    2) ESM: 채널상품코드 앞 10자리
    3) 자체상품코드(SKU) + 채널명 base 부분 일치
    4) 자체상품코드(SKU) 첫 번째 (채널 불문)
    """
    if chan_code in exact_map:
        return exact_map[chan_code]
    if len(chan_code) >= 10 and chan_code[:10] in esm_map:
        return esm_map[chan_code[:10]]
    if sku in sku_map:
        for ref in sku_map[sku]:
            ref_base = ref['채널명'].split('(')[0].strip()
            if chan_base and (chan_base in ref_base or ref_base in chan_base):
                return ref
        return sku_map[sku][0]
    return None


# ── 배송비 계산 ────────────────────────────────────────────────────────────────

def calc_actual_shipping(chan_base: str, brand: str, sales: float, collected: float) -> float:
    """
    실제 부과 배송비 (양수=판매자 부담, 음수=고객 납부가 실비 초과).
    배송비(수집)=4,000 → -1,750 / 3,500 → -900 / 0 → 2,250(비피젠) or 2,600
    스마트스토어를 포함한 모든 채널에 동일 적용.
    """
    if collected == 4_000:
        return -1_750.0
    if collected == 3_500:
        return -900.0
    return 2_250.0 if brand == '비피젠' else 2_600.0


# ── 주차별 rawdata 저장 ────────────────────────────────────────────────────────

def _save_weekly_rawdata(df: pd.DataFrame):
    """_dt 컬럼으로 주차 분리 → database 폴더에 xlsx 저장"""
    orig_cols = [c for c in df.columns if not c.startswith('_')]

    def wk_key(d):
        if pd.isna(d):
            return 'unknown'
        월, 주차 = get_week_info(d)
        return f'{d.year}_{d.month:02d}월{주차}'

    df = df.copy()
    df['_wk'] = df['_dt'].apply(wk_key)

    for wk, group in df.groupby('_wk'):
        if wk == 'unknown':
            continue
        _, label = wk.split('_', 1)   # e.g. '01월1주'
        fname = DB_DIR / f'2026_{label}_주문확인처리.xlsx'
        group[orig_cols].to_excel(fname, index=False)
        print(f"    → {fname.name}  ({len(group)}행)")


# ── rawdata 처리 ───────────────────────────────────────────────────────────────

def process_file(
    rawdata_path: Path,
    exact_map: dict,
    sku_map: dict,
    esm_map: dict,
) -> pd.DataFrame:
    """주문서확인처리 xlsx → 전체 가공 DataFrame"""
    df = pd.read_excel(rawdata_path)
    print(f"  rawdata: {len(df)}행 × {len(df.columns)}열")

    df['_dt']      = pd.to_datetime(df.iloc[:, COLS['order_dt']], errors='coerce')
    df['_invoice'] = df.iloc[:, COLS['invoice']].astype(str).str.strip()

    # 주차별 rawdata 보관
    print("  주차별 rawdata 저장 중...")
    _save_weekly_rawdata(df)

    result_rows: list = []
    unmatched:   list = []
    invoice_seen: set[str] = set()

    for _, row in df.iterrows():
        dt = row['_dt']
        if pd.isna(dt):
            continue

        invoice = row['_invoice']
        valid_invoice = invoice and invoice not in ('nan', '0', '')
        if not valid_invoice:
            continue  # 송장번호 없음 = 주문 취소 건, 제외
        is_first = invoice not in invoice_seen
        invoice_seen.add(invoice)

        chan_raw  = clean(row.iloc[COLS['shop']])
        chan_base = chan_raw.split('(')[0].strip()
        chan_code = clean(row.iloc[COLS['chan_code']])
        barcode   = clean(row.iloc[COLS['barcode']])
        sku       = clean(row.iloc[COLS['sku']])
        qty       = max(to_num(row.iloc[COLS['qty']], 1), 1)

        sales_raw = to_num(row.iloc[COLS['sales']], 0)
        ship_raw  = to_num(row.iloc[COLS['ship_raw']], 0)

        # 세트 처리: 2행+는 매출·배송비·원가 0 (송장번호 기준, 원가는 첫 행에만 부과)
        sales          = sales_raw if is_first else 0.0
        ship_collected = ship_raw  if is_first else 0.0

        # 기준 시트 매칭
        ref = lookup_ref(sku, chan_code, chan_base, exact_map, sku_map, esm_map)
        if ref is None:
            unmatched.append({'sku': sku, 'chan_code': chan_code, 'chan': chan_raw})
            ref = {'채널명': chan_raw, '브랜드': '', 'SKU명': '', '최종원가': '0', '수수료': '0'}

        chan_full = ref['채널명'] or chan_raw
        brand     = ref['브랜드']
        sku_name  = ref['SKU명']
        cost      = round(to_num(ref['최종원가'], 0) * qty, 2) if is_first else 0.0
        fee_rate  = parse_fee(ref['수수료'])

        actual_ship = calc_actual_shipping(chan_base, brand, sales, ship_collected) if is_first else 0.0
        fee_amt     = round(sales * fee_rate, 2)
        contrib     = round(sales - actual_ship - cost - fee_amt, 2)

        월, 주차 = get_week_info(dt)

        result_rows.append([
            dt.strftime('%Y-%m-%d'),
            월, 주차,
            invoice,
            chan_full, chan_code, barcode,
            brand, sku_name,
            sales, ship_collected, actual_ship,
            cost, fee_rate, fee_amt,
            0, 0,     # 광고비, 리워드 등 추가비용 (수동 입력)
            contrib,
            int(qty),
        ])

    result_df = pd.DataFrame(result_rows, columns=GAGANG_HEADERS)
    print(f"  전체 가공: {len(result_df)}행")

    if unmatched:
        today = datetime.now().strftime('%Y%m%d')
        upath = DB_DIR / f'unmatched_{today}.csv'
        (pd.DataFrame(unmatched).drop_duplicates()
           .to_csv(upath, index=False, encoding='utf-8-sig'))
        print(f"  미매칭 {len(unmatched)}건 → {upath.name}")

    return result_df


# ── Google Sheets 업로드 ────────────────────────────────────────────────────────

def upload_gagang(result_df: pd.DataFrame, append: bool = False):
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_file(str(CREDENTIALS), scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(OUTPUT_SHEET)
        if not append:
            ws.clear()
            print(f"  '{OUTPUT_SHEET}' 시트 초기화")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(
            title=OUTPUT_SHEET,
            rows=len(result_df) + 500,
            cols=len(GAGANG_HEADERS) + 2,
        )
        print(f"  '{OUTPUT_SHEET}' 시트 새로 생성")

    values = [
        [safe_val(v) for v in row]
        for row in result_df.values.tolist()
    ]

    if not append:
        ws.update(values=[GAGANG_HEADERS], range_name='A1', value_input_option='USER_ENTERED')
        start_row = 2
    else:
        start_row = len(ws.get_all_values()) + 1

    total = len(values)
    for i in range(0, total, UPLOAD_CHUNK):
        chunk = values[i : i + UPLOAD_CHUNK]
        rng   = f'A{start_row + i}'
        ws.update(values=chunk, range_name=rng, value_input_option='USER_ENTERED')
        end_row = start_row + i + len(chunk) - 1
        print(f"  업로드: {start_row + i}~{end_row}행")
        if i + UPLOAD_CHUNK < total:
            time.sleep(1.5)

    print(f"  완료: {total}행 → '{OUTPUT_SHEET}'")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='전체 가공 시트 생성')
    parser.add_argument(
        'file', nargs='?', default=str(DEFAULT_INPUT),
        help='주문서확인처리 xlsx 경로 (기본: 20260519 파일)',
    )
    parser.add_argument(
        '--append', action='store_true',
        help='기존 시트에 이어 붙이기 (기본: 시트 전체 교체)',
    )
    args = parser.parse_args()

    rawdata_path = Path(args.file)
    if not rawdata_path.exists():
        print(f'[오류] 파일 없음: {rawdata_path}')
        sys.exit(1)

    print('=' * 55)
    print('  전체 가공 시트 생성')
    print('=' * 55)

    print('\n[1/3] 기준 시트 로드 (Google Sheets)...')
    exact_map, sku_map, esm_map = load_reference()

    print(f'\n[2/3] rawdata 처리: {rawdata_path.name}')
    result_df = process_file(rawdata_path, exact_map, sku_map, esm_map)

    print('\n[3/3] Google Sheets 업로드...')
    upload_gagang(result_df, append=args.append)

    print('\n' + '=' * 55)
    print('  완료')
    print('=' * 55)


if __name__ == '__main__':
    main()
