"""
sabangnet_collector.py
사방넷 주문확인처리 자동 수집 스크립트 (Selenium)
5월 3주차부터 주 1회 실행 후 process_gagang.py 호출

사전 설정:
  .env 파일에 추가:
    SB_ID=사방넷아이디
    SB_PW=사방넷비밀번호

  패키지 설치:
    python -m pip install selenium webdriver-manager python-dotenv

사용법:
  python sabangnet_collector.py                     # 이전 주 데이터 수집 (기본)
  python sabangnet_collector.py --month 5 --week 3  # 2026년 5월 3주차 지정
  python sabangnet_collector.py --date-from 2026-05-18 --date-to 2026-05-24
"""

from __future__ import annotations

import argparse
import os
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).parent
DOWNLOAD_DIR = BASE / 'database'

try:
    from dotenv import load_dotenv
    load_dotenv(BASE / '.env')
except ImportError:
    pass

SB_ID = os.getenv('SB_ID', '')
SB_PW = os.getenv('SB_PW', '')


# ── 날짜 계산 ─────────────────────────────────────────────────────────────────

def resolve_date_range(
    month: int | None,
    week_num: int | None,
    date_from: str | None,
    date_to: str | None,
) -> tuple[str, str]:
    import sys
    sys.path.insert(0, str(BASE))
    from week_utils import prev_week_range, week_date_range

    if date_from and date_to:
        return date_from, date_to
    if month and week_num:
        start, end = week_date_range(2026, month, week_num)
    else:
        start, end = prev_week_range()
    return start.strftime('%Y%m%d'), end.strftime('%Y%m%d')


# ── Selenium 수집 ─────────────────────────────────────────────────────────────

def collect(date_from: str, date_to: str) -> Path | None:
    """
    사방넷 주문확인처리에서 xlsx 다운로드 후 파일 경로 반환.
    date_from / date_to: YYYYMMDD 형식
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.keys import Keys
        from webdriver_manager.chrome import ChromeDriverManager
    except ImportError:
        print('[오류] selenium 또는 webdriver-manager 미설치')
        print('      python -m pip install selenium webdriver-manager 실행 후 재시도')
        return None

    if not SB_ID or not SB_PW:
        print('[오류] .env 파일에 SB_ID와 SB_PW를 설정하세요.')
        return None

    opts = webdriver.ChromeOptions()
    opts.add_experimental_option('prefs', {
        'download.default_directory': str(DOWNLOAD_DIR),
        'download.prompt_for_download': False,
        'plugins.always_open_pdf_externally': True,
    })

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts,
    )
    wait = WebDriverWait(driver, 30)

    try:
        # ① 로그인
        driver.get('https://www.sabangnet.co.kr/login/login-main')
        time.sleep(3)

        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[type='text']")
        )).send_keys(SB_ID)
        driver.find_element(By.CSS_SELECTOR, "input[type='password']").send_keys(SB_PW)

        for b in driver.find_elements(By.TAG_NAME, 'button'):
            if '시스템' in b.text:
                b.click()
                break
        time.sleep(5)

        # ② 동시 접속 팝업 처리
        for b in driver.find_elements(By.TAG_NAME, 'button'):
            if b.text.strip() == '확인':
                b.click()
                time.sleep(3)
                break

        # ③ 주문확인처리 메뉴 이동 (서비스 도메인은 sbadmin1~N으로 동적)
        base_url = driver.current_url.split('#')[0].rstrip('/')
        driver.get(base_url + '#/order/order-confirm')
        time.sleep(5)

        # ④ 날짜 타입 → 주문일 (non-login el-select 첫 번째, 옵션 index 2)
        all_selects = driver.find_elements(By.CSS_SELECTOR, '.el-select')
        non_login = [s for s in all_selects if 'select_login' not in s.get_attribute('class')]
        non_login[0].click()
        time.sleep(1)
        visible_opts = [o for o in driver.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item')
                        if o.is_displayed()]
        visible_opts[2].click()  # 주문일
        time.sleep(0.5)

        # ⑤ 날짜 설정 (JavaScript — El-DatePicker 직접 값 주입)
        date_inputs = [i for i in driver.find_elements(By.CSS_SELECTOR, 'input[type="text"]')
                       if len(i.get_attribute('value') or '') == 8
                       and (i.get_attribute('value') or '').isdigit()]

        for di, val in zip(date_inputs[:2], [date_from, date_to]):
            driver.execute_script("arguments[0].value = arguments[1]", di, val)
            driver.execute_script("arguments[0].dispatchEvent(new Event('input',  {bubbles:true}))", di)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles:true}))", di)
            time.sleep(0.3)
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        time.sleep(0.5)
        print(f'  날짜 설정: {date_from} ~ {date_to}')

        # ⑥ 쇼핑몰 선택 (선택사항 IV 오른쪽 "쇼핑몰선택" = non_login[25], placeholder="mall"인 [24] 바로 다음)
        # 옵션 순서: [0]카카오톡스토어 [1]ESM옥션 [2]CJ온스타일 [3]ESM지마켓
        #           [4]11번가 [5]신세계몰(신) [6]스마트스토어(선) [7]현대홈쇼핑(3)
        #           [8]무신사 [9]GS shop [10]카카오톡선물하기
        #           [11]cafe24 [12]웰릿 [13]화해 [14]쿠팡 ← 제외
        TARGET_MALL_INDICES = list(range(11))   # 0~10번 선택

        all_selects_now = driver.find_elements(By.CSS_SELECTOR, '.el-select')
        non_login_now = [s for s in all_selects_now if 'select_login' not in s.get_attribute('class')]
        non_login_now[25].click()  # 선택사항 IV 오른쪽 쇼핑몰선택
        time.sleep(1.5)

        mall_opts = [o for o in driver.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item')
                     if o.is_displayed()]
        print(f'  쇼핑몰 옵션 {len(mall_opts)}개 확인')
        for idx in TARGET_MALL_INDICES:
            if idx < len(mall_opts):
                driver.execute_script("arguments[0].click()", mall_opts[idx])
                time.sleep(0.2)
        print(f'  쇼핑몰 {len(TARGET_MALL_INDICES)}개 선택 완료')

        # 드롭다운 닫기
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
        time.sleep(0.5)

        # ⑦ 검색
        for b in driver.find_elements(By.TAG_NAME, 'button'):
            if b.text.strip() == '검색':
                b.click()
                break
        time.sleep(5)

        rows = driver.find_elements(By.CSS_SELECTOR, '.el-table__body tr')
        print(f'  검색 결과: {len(rows)}행')
        if not rows:
            print('  [경고] 검색 결과 없음')
            return None

        # ⑧ 전체선택
        try:
            header_cb = driver.find_element(
                By.CSS_SELECTOR, '.el-table__header-wrapper .el-checkbox'
            )
            driver.execute_script("arguments[0].click()", header_cb)
            time.sleep(1)
            print('  전체선택 완료')
        except Exception as e:
            print(f'  [경고] 전체선택 실패: {e}')

        # ⑨ 주문서출력양식 → "온채팀_공헌이익 확인용" 선택
        # placeholder="양식" 인 el-select (검색 결과 후 non_login 재탐색)
        all_selects2 = driver.find_elements(By.CSS_SELECTOR, '.el-select')
        non_login2 = [s for s in all_selects2 if 'select_login' not in s.get_attribute('class')]
        form_sel = None
        for sel in non_login2:
            inp = sel.find_elements(By.CSS_SELECTOR, 'input')
            if inp:
                val = inp[0].get_attribute('value') or ''
                ph  = inp[0].get_attribute('placeholder') or ''
                if ph == '양식' or '양식선택' in val:
                    form_sel = sel
                    break
        if form_sel:
            driver.execute_script("arguments[0].scrollIntoView(true)", form_sel)
            time.sleep(0.3)
            driver.execute_script("arguments[0].click()", form_sel)
            time.sleep(1.5)
            for o in driver.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item'):
                if o.is_displayed() and ('공헌' in o.text or '온채' in o.text):
                    driver.execute_script("arguments[0].click()", o)
                    print(f'  양식 선택: {o.text}')
                    time.sleep(0.5)
                    break
        else:
            print('  [경고] 주문서양식 select 미발견')

        # ⑩ 다운로드
        before_files = set(DOWNLOAD_DIR.glob('*.xlsx'))
        for b in driver.find_elements(By.TAG_NAME, 'button'):
            if b.text.strip() == '다운로드':
                b.click()
                print('  다운로드 클릭')
                break
        time.sleep(2)

        # 팝업 확인 처리
        for b in driver.find_elements(By.TAG_NAME, 'button'):
            if b.text.strip() == '확인' and b.is_displayed():
                b.click()
                break

        # 다운로드 완료 대기
        print('  다운로드 대기 중...')
        for _ in range(30):
            time.sleep(2)
            if not (list(DOWNLOAD_DIR.glob('*.crdownload')) + list(DOWNLOAD_DIR.glob('*.part'))):
                break

        after_files = set(DOWNLOAD_DIR.glob('*.xlsx'))
        new_files = after_files - before_files
        if not new_files:
            print('[오류] 다운로드된 xlsx 파일 없음')
            return None

        downloaded = max(new_files, key=lambda f: f.stat().st_mtime)

        # 파일명 정규화 (기존 파일 덮어쓰기)
        date_str = datetime.now().strftime('%Y%m%d')
        new_name = DOWNLOAD_DIR / f'{date_str}_주문서확인처리_온채팀_공헌이익 확인용_{date_from}-{date_to}.xlsx'
        downloaded.replace(new_name)  # replace = 기존 파일 덮어쓰기
        print(f'  다운로드 완료: {new_name.name}')
        return new_name

    except Exception as e:
        print(f'[오류] 사방넷 수집 실패: {e}')
        import traceback
        traceback.print_exc()
        return None
    finally:
        driver.quit()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='사방넷 주문확인처리 자동 수집')
    parser.add_argument('--month',     type=int, default=None, help='수집 월 (예: 5)')
    parser.add_argument('--week',      type=int, default=None, help='수집 주차 (예: 3)')
    parser.add_argument('--date-from', default=None, help='시작일 YYYYMMDD')
    parser.add_argument('--date-to',   default=None, help='종료일 YYYYMMDD')
    args = parser.parse_args()

    date_from, date_to = resolve_date_range(
        args.month, args.week, args.date_from, args.date_to
    )
    print(f'수집 기간: {date_from} ~ {date_to}')

    downloaded = collect(date_from, date_to)
    if downloaded is None:
        print('[중단] 파일 다운로드 실패')
        return

    # 전체 가공 처리 + Google Sheets append
    import sys
    sys.path.insert(0, str(BASE))
    from process_gagang import load_reference, process_file, upload_gagang

    print('\n전체 가공 처리 시작...')
    exact_map, sku_map, esm_map = load_reference()
    result_df = process_file(downloaded, exact_map, sku_map, esm_map)
    upload_gagang(result_df, append=True)
    print('완료')


if __name__ == '__main__':
    main()
