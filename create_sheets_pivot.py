"""
create_sheets_pivot.py
전체 가공 → '피봇 대시보드' 시트
채널/브랜드/SKU 피봇 + KPI 카드 (컬러 포함)

사용법:
  python create_sheets_pivot.py
"""
from __future__ import annotations
from pathlib import Path
import numpy as np
import pandas as pd

BASE           = Path(__file__).parent
CREDENTIALS    = BASE / 'credentials.json'
SPREADSHEET_ID = '17_oeV41FVchyFYUl6dHQyEk689tkzrA840WfWApFkQY'
INPUT_SHEET    = '전체 가공'
OUTPUT_SHEET   = '피봇 대시보드'
SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
]

# ── 제외 채널 (이 문자열 포함 채널명 전체 제외) ──────────────────────────────────
EXCLUDE_CHANNELS   = ['쿠팡', '화해']

# ── 채널명 오류 (브랜드명이 채널명 컬럼에 잘못 입력된 경우 — exact match) ────────
INVALID_CHAN_NAMES = ['프로뉴트리션', '하우스윗']
DATA_SOURCE_SHEET  = '데이터 원본'
NATIVE_PIVOT_SHEET = '채널×브랜드×SKU 피봇'

# ── 색상 (RGB 0.0~1.0) ────────────────────────────────────────────────────────

def rgb(r, g, b):
    return {'red': r / 255, 'green': g / 255, 'blue': b / 255}

C_NAVY     = rgb(30,  58,  95)   # #1E3A5F
C_SLATE    = rgb(51,  65,  85)   # #334155
C_GREEN    = rgb(5,  150, 105)   # #059669
C_RED      = rgb(220, 38,  38)   # #DC2626
C_AMBER    = rgb(180, 120,  20)  # 광고비율 강조
C_WHITE    = rgb(255, 255, 255)
C_LGRAY    = rgb(248, 250, 252)  # 행 줄무늬
C_DARK     = rgb(30,  41,  59)   # #1E293B
C_MIDGRAY  = rgb(148, 163, 184)  # #94A3B8
C_GREEN_BG = rgb(209, 250, 229)  # #D1FAE5
C_RED_BG   = rgb(254, 226, 226)  # #FEE2E2
C_AMBER_BG = rgb(255, 251, 235)  # #FFFBEB

# ── 레이아웃 ──────────────────────────────────────────────────────────────────
#
# 피봇당 6열: 이름 / 매출 / 공헌이익 / 공헌이익율 / 광고비 / 광고비율
#
# 컬럼 (0-indexed):
#  0-5  : 채널별 피봇
#  6    : spacer
#  7-12 : 브랜드별 피봇
#  13   : spacer
#  14-19: SKU별 피봇 (상위 50)
#
# 행 (0-indexed):
#  0 : 타이틀 바
#  1 : KPI 라벨
#  2 : KPI 값
#  3 : 구분선
#  4 : 섹션 헤더
#  5 : 컬럼 헤더
#  6+: 데이터

TOTAL_COLS = 20
P_OFFSETS  = [0, 7, 14]
P_LABELS   = ['채널명', '브랜드', 'SKU명']
P_TITLES   = ['채널별 피봇', '브랜드별 피봇', 'SKU별 피봇 (상위 50)']
COL_HDRS   = [
    ('이름',      'LEFT'),
    ('매출',      'RIGHT'),
    ('공헌이익',  'RIGHT'),
    ('공헌이익율','CENTER'),
    ('광고비',    'RIGHT'),
    ('광고비율',  'CENTER'),
]
P_WIDTH    = 6    # 피봇당 열 수
DATA_ROW   = 6

# KPI 그룹: (시작열, 끝열exclusive) — 총 20열
# 총 매출(3) 총 공헌이익(3) 공헌이익율(3) 총 광고비(3) 채널 수(2) 브랜드 수(2) SKU 수(2) 기간(2)
KPI_GROUPS = [(0,3),(3,6),(6,9),(9,12),(12,14),(14,16),(16,18),(18,20)]

# ── 포맷 헬퍼 ─────────────────────────────────────────────────────────────────

def krw(v):
    if pd.isna(v): return '₩0'
    return f'₩{int(v):,}'

def pct(v):
    if pd.isna(v): return '-'
    return f'{v * 100:.1f}%'

def sv(v):
    return {'stringValue': str(v) if v is not None else ''}

def nv(v):
    try:
        f = float(v)
        if np.isnan(f) or np.isinf(f):
            return {'numberValue': 0.0}
        return {'numberValue': f}
    except (TypeError, ValueError):
        return {'numberValue': 0.0}

# ── 셀 빌더 ───────────────────────────────────────────────────────────────────

def cell(uev, bg=None, bold=False, size=10, fg=None,
         halign='LEFT', valign='MIDDLE', num_fmt=None):
    fmt = {
        'horizontalAlignment': halign,
        'verticalAlignment': valign,
        'wrapStrategy': 'CLIP',
        'textFormat': {
            'bold': bold,
            'fontSize': size,
            'foregroundColor': fg or C_DARK,
        },
    }
    if bg:
        fmt['backgroundColor'] = bg
    if num_fmt:
        fmt['numberFormat'] = num_fmt
    return {'userEnteredValue': uev, 'userEnteredFormat': fmt}

def empty(bg=None):
    return cell(sv(''), bg=bg or C_WHITE)

# ── Sheets API 요청 빌더 ──────────────────────────────────────────────────────

def req_write(sid, row, col, rows):
    return {
        'updateCells': {
            'start': {'sheetId': sid, 'rowIndex': row, 'columnIndex': col},
            'rows': [{'values': r} for r in rows],
            'fields': 'userEnteredValue,userEnteredFormat',
        }
    }

def req_merge(sid, r1, r2, c1, c2):
    return {
        'mergeCells': {
            'range': {
                'sheetId': sid,
                'startRowIndex': r1, 'endRowIndex': r2,
                'startColumnIndex': c1, 'endColumnIndex': c2,
            },
            'mergeType': 'MERGE_ALL',
        }
    }

def req_col_w(sid, c1, c2, px):
    return {
        'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS',
                      'startIndex': c1, 'endIndex': c2},
            'properties': {'pixelSize': px},
            'fields': 'pixelSize',
        }
    }

def req_row_h(sid, r1, r2, px):
    return {
        'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'ROWS',
                      'startIndex': r1, 'endIndex': r2},
            'properties': {'pixelSize': px},
            'fields': 'pixelSize',
        }
    }

def req_freeze(sid, rows=0, cols=0):
    return {
        'updateSheetProperties': {
            'properties': {
                'sheetId': sid,
                'gridProperties': {'frozenRowCount': rows, 'frozenColumnCount': cols},
            },
            'fields': 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
        }
    }

def req_border(sid, r1, r2, c1, c2):
    b = {'style': 'SOLID', 'width': 1, 'color': rgb(203, 213, 225)}
    return {
        'updateBorders': {
            'range': {'sheetId': sid,
                      'startRowIndex': r1, 'endRowIndex': r2,
                      'startColumnIndex': c1, 'endColumnIndex': c2},
            'top': b, 'bottom': b, 'left': b, 'right': b,
            'innerHorizontal': b, 'innerVertical': b,
        }
    }

def req_move_sheet(sid, index):
    return {
        'updateSheetProperties': {
            'properties': {'sheetId': sid, 'index': index},
            'fields': 'index',
        }
    }

def req_num_fmt(sid, c1, c2, pattern, fmt_type='PERCENT'):
    return {
        'repeatCell': {
            'range': {
                'sheetId': sid,
                'startRowIndex': 0, 'endRowIndex': 500,
                'startColumnIndex': c1, 'endColumnIndex': c2,
            },
            'cell': {
                'userEnteredFormat': {
                    'numberFormat': {'type': fmt_type, 'pattern': pattern}
                }
            },
            'fields': 'userEnteredFormat.numberFormat',
        }
    }

def req_bg_col(sid, c1, c2, color):
    return {
        'repeatCell': {
            'range': {
                'sheetId': sid,
                'startRowIndex': 0, 'endRowIndex': 500,
                'startColumnIndex': c1, 'endColumnIndex': c2,
            },
            'cell': {'userEnteredFormat': {'backgroundColor': color}},
            'fields': 'userEnteredFormat.backgroundColor',
        }
    }

def get_sheet_pos(sh, title):
    for ws in sh.worksheets():
        if ws.title == title:
            return ws.index
    return None

# ── 데이터 로드 ───────────────────────────────────────────────────────────────

def load_data():
    import gspread
    from google.oauth2.service_account import Credentials

    print('[1/3] 전체 가공 시트 로드...')
    creds = Credentials.from_service_account_file(str(CREDENTIALS), scopes=SCOPES)
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(SPREADSHEET_ID)
    ws    = sh.worksheet(INPUT_SHEET)
    data  = ws.get_all_values()

    if len(data) < 2:
        raise ValueError(f"'{INPUT_SHEET}' 시트가 비어 있습니다.")

    headers, rows = data[0], data[1:]
    df = pd.DataFrame(rows, columns=headers)
    print(f'  {len(df):,}행 로드')

    # 수치 컬럼 변환
    for col in ['매출', '공헌이익', '광고비']:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(',', '').str.strip(),
                errors='coerce',
            ).fillna(0)
        else:
            df[col] = 0

    # 채널명 통합: 'ESM지마켓(housweet23)' → 'ESM지마켓'
    if '채널명' in df.columns:
        df['채널명'] = df['채널명'].str.replace(r'\s*\([^)]*\)', '', regex=True).str.strip()

    # 분류 컬럼 정리
    for col in ['SKU명', '브랜드', '채널명']:
        if col in df.columns:
            df[col] = df[col].replace('', '(미분류)').fillna('(미분류)')

    # 제외 채널 필터 (contains 매칭)
    pattern = '|'.join(EXCLUDE_CHANNELS)
    mask    = df['채널명'].str.contains(pattern, na=False)
    excluded = df.loc[mask, '채널명'].value_counts()
    if not excluded.empty:
        print('  제외 채널:')
        for ch, cnt in excluded.items():
            print(f'    {ch}: {cnt:,}행')
    df = df[~mask].reset_index(drop=True)

    # 채널명 오류 제거 (브랜드명 등 잘못 입력된 채널명 — exact match)
    if INVALID_CHAN_NAMES:
        inv_mask = df['채널명'].isin(INVALID_CHAN_NAMES)
        inv_counts = df.loc[inv_mask, '채널명'].value_counts()
        if not inv_counts.empty:
            print('  채널명 오류 제거 (브랜드명 혼입):')
            for ch, cnt in inv_counts.items():
                print(f'    {ch}: {cnt:,}행')
        df = df[~inv_mask].reset_index(drop=True)

    print(f'  필터 후: {len(df):,}행')

    return sh, df

# ── 피봇 계산 ─────────────────────────────────────────────────────────────────

def make_pivot(df, group_col, top_n=None):
    p = df.groupby(group_col, as_index=False).agg(
        매출=('매출',    'sum'),
        공헌이익=('공헌이익', 'sum'),
        광고비=('광고비',  'sum'),
    )
    p['공헌이익율'] = np.where(p['매출'] != 0, p['공헌이익'] / p['매출'], np.nan)
    p['광고비율']   = np.where(p['매출'] != 0, p['광고비']   / p['매출'], np.nan)
    p = p.sort_values('매출', ascending=False).reset_index(drop=True)
    return p.head(top_n) if top_n else p

# ── 시트 빌드 ─────────────────────────────────────────────────────────────────

def build_sheet(sh, df):
    import gspread

    print('[2/3] 피봇 대시보드 시트 생성...')

    pos = get_sheet_pos(sh, OUTPUT_SHEET)
    try:
        sh.del_worksheet(sh.worksheet(OUTPUT_SHEET))
        print(f"  기존 '{OUTPUT_SHEET}' 삭제")
    except gspread.WorksheetNotFound:
        pass

    pivots  = [make_pivot(df, lc, top_n=(50 if lc == 'SKU명' else None))
               for lc in P_LABELS]
    n_data  = max(len(p) for p in pivots)
    n_rows  = DATA_ROW + n_data + 5

    ws  = sh.add_worksheet(title=OUTPUT_SHEET, rows=n_rows, cols=TOTAL_COLS)
    sid = ws.id
    print(f'  시트 ID: {sid}')

    # ── KPI 계산
    total_s = df['매출'].sum()
    total_c = df['공헌이익'].sum()
    total_r = total_c / total_s if total_s else 0
    total_a = df['광고비'].sum()
    total_ar = total_a / total_s if total_s else 0
    n_chan   = df['채널명'].nunique()
    n_brand  = df['브랜드'].nunique()
    n_sku    = df['SKU명'].nunique()
    try:
        period = f"{df['월'].min()} ~ {df['월'].max()}"
    except Exception:
        period = '전체'

    # KPI: label / value / fg_color — 8개 (KPI_GROUPS와 순서 맞춤)
    KPI_DATA = [
        ('총 매출',     krw(total_s),  C_DARK),
        ('총 공헌이익', krw(total_c),  C_GREEN if total_c >= 0 else C_RED),
        ('공헌이익율',  pct(total_r),  C_GREEN if total_r >= 0 else C_RED),
        ('총 광고비',   krw(total_a),  C_AMBER if total_a > 0 else C_DARK),
        ('채널 수',     f'{n_chan}개',  C_DARK),
        ('브랜드 수',   f'{n_brand}개', C_DARK),
        ('SKU 수',      f'{n_sku}개',   C_DARK),
        ('기간',        period,          C_DARK),
    ]

    reqs = []
    if pos is not None:
        reqs.append(req_move_sheet(sid, pos))

    # ── 행/열 크기 ────────────────────────────────────────────────────────────
    reqs.append(req_freeze(sid, rows=DATA_ROW))
    reqs.append(req_row_h(sid, 0, 1, 44))
    reqs.append(req_row_h(sid, 1, 2, 26))
    reqs.append(req_row_h(sid, 2, 3, 54))
    reqs.append(req_row_h(sid, 3, 4, 6))
    reqs.append(req_row_h(sid, 4, 5, 32))
    reqs.append(req_row_h(sid, 5, 6, 28))
    reqs.append(req_row_h(sid, 6, n_rows, 24))

    col_pxs = [130, 100, 100, 75, 100, 75]  # 6열 너비
    for p_s in P_OFFSETS:
        for i, px in enumerate(col_pxs):
            reqs.append(req_col_w(sid, p_s + i, p_s + i + 1, px))
    reqs.append(req_col_w(sid, 6,  7,  16))
    reqs.append(req_col_w(sid, 13, 14, 16))

    # ── Row 0: 타이틀 ─────────────────────────────────────────────────────────
    title_row = [
        cell(sv(f'📊 ROI 대시보드 2026   |   전체 가공 기준 · {len(df):,}건 · {period}  '
                f'  (제외: {", ".join(EXCLUDE_CHANNELS)})'),
             bg=C_NAVY, bold=True, size=13, fg=C_WHITE, halign='LEFT')
    ] + [empty(C_NAVY)] * (TOTAL_COLS - 1)
    reqs.append(req_write(sid, 0, 0, [title_row]))
    reqs.append(req_merge(sid, 0, 1, 0, TOTAL_COLS))

    # ── Rows 1-2: KPI 카드 ────────────────────────────────────────────────────
    lbl_row = []
    val_row = []
    for (label, val, fg), (c1, c2) in zip(KPI_DATA, KPI_GROUPS):
        width = c2 - c1
        lbl_row += (
            [cell(sv(label), bg=C_SLATE, size=9, fg=C_MIDGRAY,
                  halign='LEFT', valign='BOTTOM')]
            + [empty(C_SLATE)] * (width - 1)
        )
        val_row += (
            [cell(sv(val), bg=C_WHITE, bold=True, size=16, fg=fg,
                  halign='LEFT', valign='MIDDLE')]
            + [empty(C_WHITE)] * (width - 1)
        )
    reqs.append(req_write(sid, 1, 0, [lbl_row, val_row]))
    for c1, c2 in KPI_GROUPS:
        reqs.append(req_merge(sid, 1, 2, c1, c2))
        reqs.append(req_merge(sid, 2, 3, c1, c2))

    # ── Row 3: 구분선 ─────────────────────────────────────────────────────────
    reqs.append(req_write(sid, 3, 0, [[cell(sv(''), bg=C_NAVY)] * TOTAL_COLS]))

    # ── Row 4: 섹션 헤더 ──────────────────────────────────────────────────────
    sec_row = [empty() for _ in range(TOTAL_COLS)]
    for title, p_s in zip(P_TITLES, P_OFFSETS):
        sec_row[p_s] = cell(sv(title), bg=C_SLATE, bold=True, size=11,
                             fg=C_WHITE, halign='LEFT', valign='MIDDLE')
        for c in range(p_s + 1, p_s + P_WIDTH):
            sec_row[c] = empty(C_SLATE)
    reqs.append(req_write(sid, 4, 0, [sec_row]))
    for p_s in P_OFFSETS:
        reqs.append(req_merge(sid, 4, 5, p_s, p_s + P_WIDTH))

    # ── Row 5: 컬럼 헤더 ──────────────────────────────────────────────────────
    hdr_row = [empty() for _ in range(TOTAL_COLS)]
    for p_s in P_OFFSETS:
        for i, (h, align) in enumerate(COL_HDRS):
            hdr_row[p_s + i] = cell(sv(h), bg=C_NAVY, bold=True, size=10,
                                     fg=C_WHITE, halign=align, valign='MIDDLE')
    reqs.append(req_write(sid, 5, 0, [hdr_row]))

    # ── Rows 6+: 피봇 데이터 ─────────────────────────────────────────────────
    FMT_KRW = {'type': 'CURRENCY', 'pattern': '₩#,##0'}
    FMT_PCT = {'type': 'PERCENT',  'pattern': '0.0%'}

    for pv, p_s, lc in zip(pivots, P_OFFSETS, P_LABELS):
        rows_out = []
        for i in range(n_data):
            if i < len(pv):
                r      = pv.iloc[i]
                name   = str(r[lc])
                s      = float(r['매출'])
                ci     = float(r['공헌이익'])
                rt_raw = r['공헌이익율']
                rt     = float(rt_raw) if not pd.isna(rt_raw) else 0.0
                ad     = float(r['광고비'])
                ar_raw = r['광고비율']
                ar     = float(ar_raw) if not pd.isna(ar_raw) else 0.0

                bg0   = C_LGRAY    if i % 2 == 0 else C_WHITE
                ci_bg = C_GREEN_BG if ci >= 0 else C_RED_BG
                rt_bg = C_GREEN_BG if rt >= 0 else C_RED_BG
                ci_fg = C_GREEN    if ci >= 0 else C_RED
                rt_fg = C_GREEN    if rt >= 0 else C_RED
                ar_bg = C_AMBER_BG if ar > 0 else bg0

                rows_out.append([
                    cell(sv(name), bg=bg0,   size=10, halign='LEFT'),
                    cell(nv(s),    bg=bg0,   size=10, halign='RIGHT',  num_fmt=FMT_KRW),
                    cell(nv(ci),   bg=ci_bg, size=10, halign='RIGHT',  fg=ci_fg, num_fmt=FMT_KRW),
                    cell(nv(rt),   bg=rt_bg, size=10, halign='CENTER', fg=rt_fg, num_fmt=FMT_PCT),
                    cell(nv(ad),   bg=bg0,   size=10, halign='RIGHT',  num_fmt=FMT_KRW),
                    cell(nv(ar),   bg=ar_bg, size=10, halign='CENTER', fg=C_AMBER if ar > 0 else C_DARK, num_fmt=FMT_PCT),
                ])
            else:
                rows_out.append([empty() for _ in range(P_WIDTH)])

        reqs.append(req_write(sid, DATA_ROW, p_s, rows_out))
        reqs.append(req_border(sid, DATA_ROW, DATA_ROW + n_data, p_s, p_s + P_WIDTH))

    # ── 배치 실행 ─────────────────────────────────────────────────────────────
    print(f'  batch_update: {len(reqs)}개 요청 전송 중...')
    sh.batch_update({'requests': reqs})
    print(f"  완료: '{OUTPUT_SHEET}'")
    return ws


# ── 네이티브 피봇 (수정 가능) ────────────────────────────────────────────────────
#
# 두 시트 생성:
#   1. '데이터 원본'  — 채널×브랜드×SKU×월 집계 rawdata (스크립트가 관리)
#   2. '채널×브랜드×SKU 피봇' — Google Sheets 네이티브 피봇 테이블 (사용자가 직접 수정)

def build_native_pivot(sh, df):
    import gspread

    print('[네이티브 피봇] 시트 생성 중...')

    # ── 1. 데이터 집계: 채널×브랜드×SKU×월×주차
    agg = df.groupby(['채널명', '브랜드', 'SKU명', '월', '주차'], as_index=False).agg(
        매출=('매출',    'sum'),
        공헌이익=('공헌이익', 'sum'),
        광고비=('광고비',  'sum'),
    )
    agg = agg.sort_values(['채널명', '브랜드', 'SKU명', '월', '주차']).reset_index(drop=True)
    print(f'  집계: {len(agg):,}행 (채널×브랜드×SKU×월×주차 조합)')

    # ── 2. 기존 위치 저장 후 삭제
    pos_data  = get_sheet_pos(sh, DATA_SOURCE_SHEET)
    pos_pivot = get_sheet_pos(sh, NATIVE_PIVOT_SHEET)
    for name in [DATA_SOURCE_SHEET, NATIVE_PIVOT_SHEET]:
        try:
            sh.del_worksheet(sh.worksheet(name))
        except gspread.WorksheetNotFound:
            pass

    # ── 3. '데이터 원본' 시트 생성 + 업로드
    # 컬럼 순서: 채널명(0) 브랜드(1) SKU명(2) 월(3) 주차(4) 매출(5) 공헌이익(6) 광고비(7)
    HEADERS = ['채널명', '브랜드', 'SKU명', '월', '주차', '매출', '공헌이익', '광고비']
    data_ws = sh.add_worksheet(
        title=DATA_SOURCE_SHEET,
        rows=len(agg) + 5,
        cols=len(HEADERS),
    )
    data_sid = data_ws.id

    # 행 변환 (numpy 타입 → Python 기본형)
    def to_native(v):
        if isinstance(v, (np.integer,)):   return int(v)
        if isinstance(v, (np.floating,)):  return 0 if (np.isnan(v) or np.isinf(v)) else float(v)
        if isinstance(v, float) and (np.isnan(v) or np.isinf(v)): return 0
        return v

    rows_data = [
        [to_native(r[h]) for h in HEADERS]
        for _, r in agg.iterrows()
    ]

    CHUNK = 1000
    data_ws.update(values=[HEADERS], range_name='A1', value_input_option='USER_ENTERED')
    for i in range(0, len(rows_data), CHUNK):
        data_ws.update(
            values=rows_data[i:i + CHUNK],
            range_name=f'A{i + 2}',
            value_input_option='USER_ENTERED',
        )
    print(f"  '{DATA_SOURCE_SHEET}' 업로드 완료")

    # 데이터 원본 시트 서식 + 위치 복원
    fmt_reqs = [
        {
            'repeatCell': {
                'range': {
                    'sheetId': data_sid,
                    'startRowIndex': 0, 'endRowIndex': 1,
                    'startColumnIndex': 0, 'endColumnIndex': len(HEADERS),
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': rgb(30, 58, 95),
                        'textFormat': {'bold': True, 'foregroundColor': rgb(255, 255, 255)},
                    }
                },
                'fields': 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat',
            }
        },
        {
            'updateSheetProperties': {
                'properties': {
                    'sheetId': data_sid,
                    'gridProperties': {'frozenRowCount': 1},
                },
                'fields': 'gridProperties.frozenRowCount',
            }
        },
    ]
    if pos_data is not None:
        fmt_reqs.append(req_move_sheet(data_sid, pos_data))

    # ── 4. '채널×브랜드×SKU 피봇' 시트 생성
    pivot_ws  = sh.add_worksheet(title=NATIVE_PIVOT_SHEET, rows=300, cols=20)
    pivot_sid = pivot_ws.id

    # ── 5. 네이티브 피봇 테이블 생성
    # 행 그룹: 채널명 > 브랜드 > SKU명 > 월 > 주차
    # 컬럼 레이아웃 (5 row group → 라벨 5열):
    #  A(0)=채널명  B(1)=브랜드  C(2)=SKU명  D(3)=월  E(4)=주차
    #  F(5)=매출  G(6)=공헌이익  H(7)=공헌이익율  I(8)=광고비  J(9)=광고비율
    pivot_spec = {
        'source': {
            'sheetId': data_sid,
            'startRowIndex': 0,
            'startColumnIndex': 0,
            'endRowIndex': len(agg) + 1,
            'endColumnIndex': len(HEADERS),
        },
        'rows': [
            {'sourceColumnOffset': 0, 'showTotals': True,  'sortOrder': 'DESCENDING', 'valueBucket': {'valuesIndex': 0}},  # 채널명
            {'sourceColumnOffset': 1, 'showTotals': True,  'sortOrder': 'DESCENDING', 'valueBucket': {'valuesIndex': 0}},  # 브랜드
            {'sourceColumnOffset': 2, 'showTotals': False, 'sortOrder': 'DESCENDING', 'valueBucket': {'valuesIndex': 0}},  # SKU명
            {'sourceColumnOffset': 3, 'showTotals': False, 'sortOrder': 'ASCENDING'},   # 월 (1월→5월 순)
            {'sourceColumnOffset': 4, 'showTotals': False, 'sortOrder': 'ASCENDING'},   # 주차 (1주→5주 순)
        ],
        'values': [
            {'summarizeFunction': 'SUM',    'sourceColumnOffset': 5, 'name': '매출'},
            {'summarizeFunction': 'SUM',    'sourceColumnOffset': 6, 'name': '공헌이익'},
            {'summarizeFunction': 'CUSTOM', 'formula': "='공헌이익'/'매출'", 'name': '공헌이익율'},
            {'summarizeFunction': 'SUM',    'sourceColumnOffset': 7, 'name': '광고비'},
            {'summarizeFunction': 'CUSTOM', 'formula': "='광고비'/'매출'",   'name': '광고비율'},
        ],
        'valueLayout': 'HORIZONTAL',
    }

    # 피봇 컬럼 레이아웃 (row group 5개 → 라벨 5열):
    #  A(0)=채널명  B(1)=브랜드  C(2)=SKU명  D(3)=월  E(4)=주차
    #  F(5)=매출  G(6)=공헌이익  H(7)=공헌이익율  I(8)=광고비  J(9)=광고비율

    # 차원별 배경 컬러 (열 단위)
    #  채널명: 파랑 계열  브랜드: 연파랑  SKU명: 회백  월: 초록  주차: 연초록
    pivot_reqs = fmt_reqs + [
        {
            'updateCells': {
                'rows': [{'values': [{'pivotTable': pivot_spec}]}],
                'start': {'sheetId': pivot_sid, 'rowIndex': 0, 'columnIndex': 0},
                'fields': 'pivotTable',
            }
        },
        # ── 차원 열 배경색
        req_bg_col(pivot_sid, 0, 1, rgb(219, 234, 254)),  # 채널명 — blue-100
        req_bg_col(pivot_sid, 1, 2, rgb(239, 246, 255)),  # 브랜드 — blue-50
        req_bg_col(pivot_sid, 2, 3, rgb(248, 250, 252)),  # SKU명  — slate-50
        req_bg_col(pivot_sid, 3, 4, rgb(209, 250, 229)),  # 월     — green-100
        req_bg_col(pivot_sid, 4, 5, rgb(236, 253, 245)),  # 주차   — green-50
        # ── 값 열 숫자 포맷
        req_num_fmt(pivot_sid, 5, 6, '#,##0', 'NUMBER'),   # 매출 (F)
        req_num_fmt(pivot_sid, 6, 7, '#,##0', 'NUMBER'),   # 공헌이익 (G)
        req_num_fmt(pivot_sid, 7, 8, '0.0%',  'PERCENT'),  # 공헌이익율 (H)
        req_num_fmt(pivot_sid, 8, 9, '#,##0', 'NUMBER'),   # 광고비 (I)
        req_num_fmt(pivot_sid, 9,10, '0.0%',  'PERCENT'),  # 광고비율 (J)
    ]
    if pos_pivot is not None:
        pivot_reqs.append(req_move_sheet(pivot_sid, pos_pivot))

    sh.batch_update({'requests': pivot_reqs})
    print(f"  '{NATIVE_PIVOT_SHEET}' 네이티브 피봇 생성 완료")
    print('  ※ 시트 우클릭 → 피봇 테이블 수정 → 월(필드 목록) 드래그로 월별 필터 추가 가능')
    return pivot_ws


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    sh, df = load_data()

    print()
    ws1 = build_sheet(sh, df)
    print()
    ws2 = build_native_pivot(sh, df)

    print('\n완료')
    print('=' * 65)
    print(f'정적 피봇:   https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={ws1.id}')
    print(f'편집 피봇:   https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={ws2.id}')
    print('=' * 65)


if __name__ == '__main__':
    main()
