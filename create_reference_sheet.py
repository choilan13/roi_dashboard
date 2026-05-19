"""
기준 시트 자동 생성 스크립트
사방넷 Excel 파일들을 파싱·통합하여 Google Sheets '기준' 시트를 생성합니다.

사전 준비:
  1. Google Cloud Console에서 Sheets API + Drive API 활성화
  2. 서비스 계정(Service Account) JSON 키 발급 → credentials.json 으로 이 스크립트와 같은 폴더에 저장
  3. 대상 Google Sheets에 서비스 계정 이메일을 편집자로 공유
  4. pip install gspread google-auth pandas openpyxl
"""

import re
import pandas as pd
import numpy as np
from pathlib import Path

BASE = Path(r'c:\Users\2359\Desktop\최란업무파일\05. ROI시트 관리\roi_dashboard\database')
CREDENTIALS_FILE = Path(__file__).parent / 'credentials.json'
SPREADSHEET_ID = '17_oeV41FVchyFYUl6dHQyEk689tkzrA840WfWApFkQY'
TARGET_SHEET = '기준'

SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
]

HEADERS = [
    '자체상품코드', '판매코드', '상품명', '연결상품코드', 'EA',
    '브랜드', '별칭', '채널명', '채널상품코드',
    '원가(+VAT)', '최종원가', '채널별 수수료',
]

# 송신내역에서 메타 컬럼 개수 (번호·품번코드·전체상품코드·모델명·상품명·판매가·등급)
META_COL_COUNT = 7


# ─────────────────── 헬퍼 ───────────────────

def to_str(v) -> str:
    """pandas 값을 안전하게 문자열로 변환 (NaN → '')"""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ''
    s = str(v).strip()
    return '' if s in ('nan', 'None') else s


def parse_fee(v) -> float | None:
    """수수료 값을 소수 비율로 정규화 (0.13, '3.3%', '13.0%' → float)"""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s.endswith('%'):
        try:
            return float(s[:-1]) / 100
        except ValueError:
            return None
    try:
        return float(s)
    except ValueError:
        return None


def match_channel_fee(채널명: str, 수수료_dict: dict) -> str:
    """채널 컬럼명 기준 부분 일치로 수수료 검색. '채널명(아이디)' → '채널명' 으로 기준 매칭"""
    base = 채널명.split('(')[0].strip()
    if base in 수수료_dict:
        v = 수수료_dict[base]
        return '' if v is None else str(round(v, 4))
    for k, v in 수수료_dict.items():
        if k in base or base in k:
            return '' if v is None else str(round(v, 4))
    return ''


def parse_linked_codes(cell_value, default_ea: int) -> list[dict]:
    """
    연결상품코드 셀 파싱.
    형식 예: '100717-0001:1,100718-0002:2'  또는  '100717-0001'
    반환: [{'연결상품코드': str, 'EA': int}, ...]
    """
    if not to_str(cell_value):
        return []
    rows = []
    for part in str(cell_value).split(','):
        part = part.strip()
        if not part:
            continue
        if ':' in part:
            code, ea_s = part.rsplit(':', 1)
            code = code.strip()
            try:
                ea = int(ea_s.strip())
            except ValueError:
                ea = default_ea
        else:
            code = part
            ea = default_ea
        if re.match(r'^\d{6}-\d{4}$', code):
            rows.append({'연결상품코드': code, 'EA': ea})
    return rows


# ─────────────────── Step 1: 송신내역 ───────────────────

def load_송신내역(path: Path) -> pd.DataFrame:
    """온채팀 판매코드만 필터 후 채널별 wide → long 변환"""
    df = pd.read_excel(path, header=1)
    cols = list(df.columns)

    code_col  = cols[1]   # 품번코드 (B열) = 판매코드
    name_col  = cols[4]   # 상품명 (E열)
    chan_cols = cols[META_COL_COUNT:]  # 채널 컬럼 (G열 이후)

    # 온채팀 필터
    mask = df[name_col].astype(str).str.contains('온채팀', na=False)
    온채팀 = df[mask].copy()
    print(f"  송신내역 온채팀 행: {len(온채팀)}행 (전체 {len(df)}행)")

    # Wide → Long
    long = 온채팀.melt(
        id_vars=[code_col, name_col],
        value_vars=chan_cols,
        var_name='채널명',
        value_name='채널상품코드',
    )
    # 채널상품코드 없는 행 제거
    long = long.dropna(subset=['채널상품코드'])
    long = long[long['채널상품코드'].astype(str).str.strip().ne('')]
    long = long.rename(columns={code_col: '판매코드', name_col: '상품명'})
    long['판매코드'] = long['판매코드'].astype(str)
    long['채널상품코드'] = long['채널상품코드'].astype(str)
    return long.reset_index(drop=True)


# ─────────────────── Step 2: 단품대량수정 ───────────────────

def load_단품(path: Path, 온채팀_판매코드: set) -> pd.DataFrame:
    """판매코드 → 연결상품코드 + EA 매핑 (세트 분리 포함)"""
    df = pd.read_excel(path, header=1)
    df = df.iloc[1:].reset_index(drop=True)  # 주의사항 행(2행) 제거

    rows = []
    for _, row in df.iterrows():
        a_val = to_str(row.iloc[0])
        if len(a_val) < 6:
            continue
        판매코드 = a_val[:6]
        if 판매코드 not in 온채팀_판매코드:
            continue

        # M열 기본 EA
        try:
            default_ea = int(row.iloc[12]) if pd.notna(row.iloc[12]) else 1
        except (ValueError, TypeError):
            default_ea = 1

        # J열 연결상품코드 파싱
        for lc in parse_linked_codes(row.iloc[9], default_ea):
            rows.append({'판매코드': 판매코드, **lc})

    if not rows:
        return pd.DataFrame(columns=['판매코드', '연결상품코드', 'EA'])
    return pd.DataFrame(rows).drop_duplicates(subset=['판매코드', '연결상품코드'])


# ─────────────────── Step 3: 다운로드 ───────────────────

def load_다운로드(path: Path) -> tuple[dict, dict]:
    """판매코드 → (SKU, 브랜드명) 매핑"""
    df = pd.read_excel(path, header=2)
    df['판매코드'] = df.iloc[:, 1].astype(str)
    df['SKU']     = df.iloc[:, 2].apply(to_str)
    df['브랜드']   = df.iloc[:, 5].apply(to_str)

    # 판매코드별 첫 번째 유효 SKU / 브랜드 선택
    grp = df.groupby('판매코드').first().reset_index()
    sku_map   = dict(zip(grp['판매코드'], grp['SKU']))
    brand_map = dict(zip(grp['판매코드'], grp['브랜드']))
    return sku_map, brand_map


# ─────────────────── Step 4: 원가표 ───────────────────

def load_원가표(path: Path) -> dict:
    """SKU → {원가, 브랜드, 별칭} 매핑 (원가표 SKU2 = col[5] 기준)"""
    df = pd.read_excel(path, header=4)
    df = df.iloc[1:].reset_index(drop=True)  # 빈 서브헤더 행 제거

    sku_col   = df.columns[5]   # 상품코드.1 (다운로드 자체상품코드와 형식 일치)
    원가_col  = df.columns[20]  # 제품 원가 (VAT 포함)
    브랜드_col = df.columns[3]   # 브랜드
    별칭_col  = df.columns[8]   # 별칭

    df = df.dropna(subset=[sku_col])
    sku = df[sku_col].astype(str)
    return {
        'SKU_원가':  dict(zip(sku, df[원가_col])),
        'SKU_브랜드': dict(zip(sku, df[브랜드_col].apply(to_str))),
        'SKU_별칭':  dict(zip(sku, df[별칭_col].apply(to_str))),
    }


# ─────────────────── Step 5: 쇼핑몰관리 ───────────────────

def load_쇼핑몰(path: Path) -> dict:
    """쇼핑몰명 → 수수료(소수) 매핑"""
    df = pd.read_excel(path, header=1)
    df['수수료_f'] = df['수수료'].apply(parse_fee)
    # 쇼핑몰명별 첫 번째 수수료 사용
    return df.groupby('쇼핑몰명')['수수료_f'].first().to_dict()


# ─────────────────── Step 5b: ★2026 ROI관리 채널별_판매코드 (보조 소스) ───────────────────

def load_roi_관리(path: Path, 온채팀_판매코드: set) -> dict:
    """
    ★2026 ROI관리 파일의 '채널별_판매코드' 시트에서
    판매코드 → {브랜드, 별칭, 원가} 매핑 (온채팀 한정)
    SKU 경로로 원가를 못 구할 때 fallback으로 사용
    """
    import openpyxl
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.worksheets[6]  # 채널별_판매코드 시트 (index 6)

    result = {}
    first = True
    for row in ws.iter_rows(values_only=True):
        if first:   # 헤더 행 건너뜀
            first = False
            continue
        판매코드 = to_str(row[1]) if row[1] is not None else ''
        if not 판매코드 or 판매코드 not in 온채팀_판매코드:
            continue
        if 판매코드 not in result:
            result[판매코드] = {
                '브랜드': to_str(row[6]),
                '별칭':   to_str(row[7]),
                '원가':   row[10] if row[10] is not None else '',
            }
    wb.close()
    return result


# ─────────────────── Step 6: 병합 ───────────────────

def build_reference(
    base_df: pd.DataFrame,
    linked_df: pd.DataFrame,
    sku_map: dict,
    brand_map: dict,
    원가_lookup: dict,
    수수료_map: dict,
    roi_lookup: dict,       # ★2026 ROI관리 보조 소스 (판매코드 → 브랜드/별칭/원가)
) -> pd.DataFrame:
    """모든 소스를 병합하여 기준 시트 DataFrame 생성"""

    # 채널 × 연결상품코드 확장
    if not linked_df.empty:
        merged = base_df.merge(linked_df, on='판매코드', how='left')
    else:
        merged = base_df.copy()
        merged['연결상품코드'] = ''
        merged['EA'] = ''

    # SKU (다운로드 자체상품코드)
    merged['자체상품코드'] = merged['판매코드'].map(sku_map).apply(to_str)

    # 브랜드: 다운로드 직접 참조 → fallback: ROI관리 채널별_판매코드
    merged['브랜드'] = merged['판매코드'].apply(
        lambda k: to_str(brand_map.get(k)) or roi_lookup.get(k, {}).get('브랜드', '')
    )

    # 별칭: 원가표 via SKU → fallback: ROI관리 채널별_판매코드
    sku_별칭 = 원가_lookup['SKU_별칭']
    merged['별칭'] = merged.apply(
        lambda r: (sku_별칭.get(r['자체상품코드'], '') if r['자체상품코드']
                   else roi_lookup.get(r['판매코드'], {}).get('별칭', '')),
        axis=1,
    )

    # 원가: 원가표 via SKU → fallback: ROI관리 채널별_판매코드
    sku_원가 = 원가_lookup['SKU_원가']
    def get_원가(r):
        if r['자체상품코드']:
            v = sku_원가.get(r['자체상품코드'], '')
            if v != '':
                return v
        return roi_lookup.get(r['판매코드'], {}).get('원가', '')
    merged['원가(+VAT)'] = merged.apply(get_원가, axis=1)

    # 수수료
    merged['채널별 수수료'] = merged['채널명'].apply(
        lambda x: match_channel_fee(x, 수수료_map)
    )

    # 컬럼 정리 (최종원가는 Google Sheets 수식으로 처리)
    result = merged[[
        '자체상품코드', '판매코드', '상품명', '연결상품코드', 'EA',
        '브랜드', '별칭', '채널명', '채널상품코드',
        '원가(+VAT)', '채널별 수수료',
    ]].copy()

    # 모든 값을 문자열로 통일 (NaN 제거)
    for col in result.columns:
        result[col] = result[col].apply(to_str)

    return result.reset_index(drop=True)


# ─────────────────── Step 7: Google Sheets 업로드 ───────────────────

def upload_to_sheets(df: pd.DataFrame) -> None:
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(TARGET_SHEET)
        ws.clear()
        print(f"  기존 '{TARGET_SHEET}' 시트 초기화 완료")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=TARGET_SHEET, rows=len(df) + 100, cols=len(HEADERS) + 2)
        print(f"  '{TARGET_SHEET}' 시트 새로 생성")

    n = len(df)

    # 헤더 (A1:L1)
    ws.update(values=[HEADERS], range_name='A1', value_input_option='USER_ENTERED')

    # 데이터 준비 (최종원가 K열 제외: A~J + L)
    data_order = [
        '자체상품코드', '판매코드', '상품명', '연결상품코드', 'EA',
        '브랜드', '별칭', '채널명', '채널상품코드',
        '원가(+VAT)', '채널별 수수료',
    ]
    rows = df[data_order].values.tolist()

    # A~J 업로드 (indices 0~9)
    ws.update(values=[r[:10] for r in rows], range_name='A2', value_input_option='USER_ENTERED')

    # K열: 최종원가 수식 =IF(OR(J="",E=""),"",J*E)
    k_formulas = [
        [f'=IF(OR(J{i}="",E{i}=""),"",J{i}*E{i})']
        for i in range(2, n + 2)
    ]
    ws.update(values=k_formulas, range_name=f'K2:K{n+1}', value_input_option='USER_ENTERED')

    # L열: 채널별 수수료 (index 10)
    ws.update(values=[[r[10]] for r in rows], range_name=f'L2:L{n+1}', value_input_option='USER_ENTERED')

    print(f"  업로드 완료: {n}행 → '{TARGET_SHEET}' 시트")


def save_preview_excel(df: pd.DataFrame) -> None:
    """credentials.json 없을 때 로컬 Excel로 저장"""
    out = Path(__file__).parent / '기준시트_미리보기.xlsx'

    preview = df.copy()
    # 최종원가 컬럼을 K 위치(index 10)에 삽입
    preview.insert(10, '최종원가', preview['원가(+VAT)'].apply(
        lambda x: float(x) * 1 if x else ''
    ))
    preview.columns = HEADERS
    preview.to_excel(out, index=False, engine='openpyxl')
    print(f"  로컬 저장 완료: {out}")


# ─────────────────── Main ───────────────────

def main():
    print("=" * 50)
    print("  기준 시트 생성 시작")
    print("=" * 50)

    print("\n[1/5] 송신내역 로딩 (온채팀 필터 + 채널 long format)...")
    base_df = load_송신내역(BASE / '사방넷상품조회수정_송신내역 (3).xlsx')
    온채팀_판매코드 = set(base_df['판매코드'].unique())
    print(f"  온채팀 판매코드 {len(온채팀_판매코드)}개, 채널 행 {len(base_df)}행")

    print("\n[2/5] 단품대량수정 로딩 (연결상품코드 파싱)...")
    linked_df = load_단품(BASE / '사방넷단품대량수정_수정파일 (10).xlsx', 온채팀_판매코드)
    print(f"  연결상품코드 행 {len(linked_df)}행 (판매코드 {linked_df['판매코드'].nunique()}개)")

    print("\n[3/5] 다운로드 로딩 (판매코드 → SKU/브랜드 매핑)...")
    sku_map, brand_map = load_다운로드(BASE / '사방넷상품조회수정_다운로드 (6).xlsx')
    sku_filled = sum(1 for v in sku_map.values() if v)
    print(f"  SKU 있는 판매코드: {sku_filled}개")

    print("\n[4/5] 원가표 로딩 (SKU → 원가/브랜드/별칭 매핑)...")
    원가_lookup = load_원가표(BASE / '전제품 원가표(실무자 공유용)_____ (1).xlsx')
    print(f"  원가표 SKU 수: {len(원가_lookup['SKU_원가'])}개")

    print("\n[5/5] 쇼핑몰관리 로딩 (채널 수수료 매핑)...")
    수수료_map = load_쇼핑몰(BASE / '쇼핑몰관리(국내)_거래중 쇼핑몰.xlsx')
    print(f"  채널 수수료 매핑: {len(수수료_map)}개 채널")

    print("\n[5b] ★2026 ROI관리 로딩 (브랜드/별칭/원가 보조 소스)...")
    roi_lookup = load_roi_관리(BASE / '★2026_외부채널_ROI관리.xlsx', 온채팀_판매코드)
    print(f"  ROI관리 온채팀 판매코드: {len(roi_lookup)}개")

    print("\n[병합] 데이터 통합 중...")
    result = build_reference(base_df, linked_df, sku_map, brand_map, 원가_lookup, 수수료_map, roi_lookup)
    print(f"  최종 행 수: {len(result)}")
    print("\n  컬럼별 채움 현황:")
    for col in result.columns:
        n = (result[col] != '').sum()
        pct = 100 * n // len(result) if len(result) else 0
        print(f"    {col:20s}: {n:5d}/{len(result)} ({pct}%)")

    print()
    if CREDENTIALS_FILE.exists():
        print("[업로드] Google Sheets에 업로드 중...")
        upload_to_sheets(result)
    else:
        print(f"[주의] credentials.json 파일 없음 → 로컬 Excel로 저장합니다.")
        print(f"       설정 방법은 스크립트 상단 주석 참고")
        save_preview_excel(result)

    print("\n" + "=" * 50)
    print("  완료")
    print("=" * 50)


if __name__ == '__main__':
    main()
