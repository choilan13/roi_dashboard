"""
patch_gagang.py
'전체 가공' 시트에서 브랜드 / SKU명이 공란인 행만
'기준' 시트의 채널상품코드 매핑을 참고해서 채워 넣기.

다른 셀은 절대 수정하지 않음.
사용법:  python patch_gagang.py
"""
from __future__ import annotations
from pathlib import Path

BASE           = Path(__file__).parent
CREDENTIALS    = BASE / 'credentials.json'
SPREADSHEET_ID = '17_oeV41FVchyFYUl6dHQyEk689tkzrA840WfWApFkQY'
REF_SHEET      = '기준'
GAGANG_SHEET   = '전체 가공'
SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
]

# '전체 가공' 헤더 컬럼 위치 (0-indexed)
# 날짜(0) 월(1) 주차(2) 송장번호(3) 채널명(4)
# 채널상품코드(5) 옵션 바코드(6) 브랜드(7) SKU명(8) ...
COL_CHAN_CODE = 5   # F
COL_CHAN_NAME = 4   # E
COL_BRAND     = 7   # H
COL_SKU       = 8   # I

CHUNK = 500


def col_letter(n: int) -> str:
    """0-based col index → 시트 열 문자 (A=0, B=1, ...)"""
    result = ''
    n += 1
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


# ── 기준 시트 로드 ─────────────────────────────────────────────────────────────

def load_reference():
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_file(str(CREDENTIALS), scopes=SCOPES)
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(SPREADSHEET_ID)
    ws    = sh.worksheet(REF_SHEET)
    data  = ws.get_all_values()
    headers, rows = data[0], data[1:]

    exact_map: dict = {}
    esm_map:   dict = {}

    for row in rows:
        if len(row) < 9:
            continue
        d         = dict(zip(headers, row))
        chan_code  = d.get('채널상품코드', '').strip()
        chan_name  = d.get('채널명',       '').strip()
        brand      = d.get('브랜드',       '').strip()
        alias      = d.get('별칭',         '').strip()

        if not chan_code:
            continue

        ref = {'브랜드': brand, 'SKU명': alias}
        exact_map.setdefault(chan_code, ref)
        if any(k in chan_name for k in ('ESM', '지마켓', '옥션')):
            esm_map.setdefault(chan_code[:10], ref)

    print(f'  기준 시트: {len(rows)}행 → exact {len(exact_map)}, ESM {len(esm_map)}')
    return gc, sh, exact_map, esm_map


# ── 패치 실행 ─────────────────────────────────────────────────────────────────

def patch():
    import gspread

    print('[1/3] 기준 시트 로드...')
    gc, sh, exact_map, esm_map = load_reference()

    print('[2/3] 전체 가공 시트 로드...')
    ws   = sh.worksheet(GAGANG_SHEET)
    data = ws.get_all_values()

    if len(data) < 2:
        print('  데이터 없음')
        return

    rows = data[1:]   # 헤더(row 1) 제외, data는 row 2부터
    print(f'  {len(rows):,}행 로드')

    updates         = []
    matched_rows    = 0
    skipped_no_ref  = 0

    for i, row in enumerate(rows):
        sheet_row = i + 2   # 실제 시트 행 번호 (1-indexed, 헤더=1)

        def get(col):
            return row[col].strip() if len(row) > col else ''

        brand     = get(COL_BRAND)
        sku_name  = get(COL_SKU)

        # 둘 다 값 있으면 skip
        if brand and sku_name:
            continue

        chan_code = get(COL_CHAN_CODE)
        if not chan_code:
            continue

        # 기준 시트 매칭 (채널상품코드 기준)
        ref = None
        if chan_code in exact_map:
            ref = exact_map[chan_code]
        elif len(chan_code) >= 10 and chan_code[:10] in esm_map:
            ref = esm_map[chan_code[:10]]

        if ref is None:
            skipped_no_ref += 1
            continue

        matched_rows += 1
        new_brand = ref['브랜드']
        new_sku   = ref['SKU명']

        # 공란 필드만 업데이트 — 값 있는 필드는 절대 건드리지 않음
        if not brand and new_brand:
            updates.append({
                'range':  f'{col_letter(COL_BRAND)}{sheet_row}',
                'values': [[new_brand]],
            })
        if not sku_name and new_sku:
            updates.append({
                'range':  f'{col_letter(COL_SKU)}{sheet_row}',
                'values': [[new_sku]],
            })

    print(f'  매칭 행: {matched_rows}개 | 미매칭(기준 없음): {skipped_no_ref}개')
    print(f'  업데이트 대상 셀: {len(updates)}개')

    if not updates:
        print('  채울 공란 없음. 종료.')
        return

    print('[3/3] 시트 업데이트 중...')
    for i in range(0, len(updates), CHUNK):
        batch = updates[i:i + CHUNK]
        ws.batch_update(batch, value_input_option='USER_ENTERED')
        done = min(i + CHUNK, len(updates))
        print(f'  {done}/{len(updates)} 셀 완료')

    print(f'\n완료: {len(updates)}개 셀 업데이트 (브랜드/SKU명 공란만)')


if __name__ == '__main__':
    patch()
