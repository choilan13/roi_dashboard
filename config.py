import os

# ─── 사방넷 API 설정 ───────────────────────────────────────────────────────────
# API 키 발급 후 아래 환경변수를 설정하거나, settings 페이지에서 입력하세요.
# 발급 방법: 사방넷 로그인 > 서비스 신청 > API 서비스
# 문의: fbs_dev@daou.co.kr
SABANGNET_SHOP_ID = os.environ.get("SB_SHOP_ID", "")
SABANGNET_API_KEY = os.environ.get("SB_API_KEY", "")
SABANGNET_BASE_URL = "https://r.sabangnet.co.kr/RTL_API/"

# API 키가 없으면 자동으로 Mock 모드로 동작
API_MODE = "mock" if not SABANGNET_API_KEY else "live"

# ─── 기존 엑셀 파일 경로 ────────────────────────────────────────────────────────
XLSX_PATH = r"C:\Users\2359\Desktop\최란업무파일\사방넷단품대량수정_수정파일_ROI 관리용_260316 (8).xlsx"
XLSX_COST_SHEET = "물류코드"

# ─── 관리 채널 목록 ─────────────────────────────────────────────────────────────
CHANNELS = [
    "무신사",
    "카카오톡스토어",
    "CJ온스타일",
    "ESM지마켓",
    "ESM옥션",
    "11번가",
    "GSshop",
    "화해",
    "네이버스마트스토어",
    "옥션지마켓ESM",
    "카카오기타",
]

# ─── 채널별 기본 수수료율 (DB에서 주별로 덮어쓰기 가능) ───────────────────────────
DEFAULT_COMMISSION_RATES = {
    "무신사":           0.094,  # 실질 ~9.4% (명목 28%)
    "카카오톡스토어":    0.05,
    "CJ온스타일":       0.30,
    "ESM지마켓":        0.12,
    "ESM옥션":          0.12,
    "11번가":           0.13,
    "GSshop":          0.30,
    "화해":             0.15,
    "네이버스마트스토어": 0.05,
    "옥션지마켓ESM":    0.12,
    "카카오기타":        0.05,
}
