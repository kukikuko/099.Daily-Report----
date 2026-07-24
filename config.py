import os
import sys
import re

# --- 경로 관련 설정 ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)          # .env, output 등 실행파일 옆
    BUNDLE_DIR = sys._MEIPASS                            # 번들 리소스 (Template 등)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = BASE_DIR

ENV_PATH = os.path.join(BASE_DIR, '.env')
DATA_FILE_PATH = os.path.join(BASE_DIR, 'weekly_data.json')
TEMPLATE_PATH = os.path.join(BUNDLE_DIR, "Template", "Daily_Report_Template.xlsx")
ADDRESSBOOK_FILENAME = "주소록.csv"

# --- 로깅 및 백업 경로 설정 ---
LOG_DIR = os.path.join(BASE_DIR, 'Logs')
LOG_FILE_PATH = os.path.join(LOG_DIR, 'daily_report.log')
BACKUP_DIR = os.path.join(BASE_DIR, 'Data', 'backups')
if not os.path.exists(BACKUP_DIR):
    os.makedirs(BACKUP_DIR, exist_ok=True)

# --- 정규식 패턴 ---
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
OUTLOOK_TOKEN_PATTERN = re.compile(r"\[\[[^\[\]\r\n]+\]\]|\{[^{}\r\n]+\}")

# --- 엑셀 셀 매핑 정보 ---
CELL_MAP = {
    "report_date": "C7", 
    "department": "I6", 
    "work_location": "I7",
    "author_name": "C8", 
    "headcount": "I8", 
    "work_content": "D12",
    "tomorrow_work": "B18", 
    "notes": "B23",
}

# --- 보고서 인쇄 및 이미지 범위 ---
REPORT_PRINT_AREA = "A1:I27"
REPORT_IMAGE_RANGE = "A2:I27"

# --- 기본 메일 템플릿 및 토큰 ---
DEFAULT_SUBJECT = "Daily Report_[[년]]년 [[월]]월 [[일]]일"
DEFAULT_BODY = """안녕하십니까.
금일([[년]]년 [[월]]월 [[일]]일) 일일 업무 보고 드립니다.

주요 업무 내용은 아래 이미지와 같습니다.
감사합니다."""

OUTLOOK_FRIENDLY_TOKENS = [
    "[[년]]", "[[월]]", "[[일]]",
]

OUTLOOK_TOKEN_ALIASES = {
    "[[년]]": "year",
    "[[월]]": "month",
    "[[일]]": "day",
}
