import os
import sys
import re
import time
import json
import tkinter as tk
from tkinter import filedialog
import win32com.client as win32
from datetime import datetime, date, timedelta
from dotenv import load_dotenv

# ==========================================================
# 1. 설정 및 경로 초기화
# ==========================================================

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ENV_PATH = os.path.join(BASE_DIR, '.env')
# [NEW] 이번 주 근무지 정보를 저장할 파일 경로
DATA_FILE_PATH = os.path.join(BASE_DIR, 'weekly_data.json')
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# ----------------------------------------------------------
# [설정] 환경 설정 파일(.env) 체크
# ----------------------------------------------------------
if not os.path.exists(ENV_PATH):
    print("=" * 60)
    print("⚙️  초기 환경 설정 파일(.env)이 없습니다. 설정을 시작합니다.")
    print("=" * 60)

    while True:
        input_name = input("👉 작성자 성함을 입력해주세요: ").strip()
        if input_name: break
        print("   ⚠️ 이름을 입력해야 합니다.")

    print("\n👉 주로 근무하는 '기본 근무지'를 입력해주세요. (예: 본사, 현장, 판교 등)")
    input_loc = input("   입력 (기본값: 본사): ").strip()
    if not input_loc:
        input_loc = "본사"

    default_dir = os.path.join(BASE_DIR, "Output")
    print(f"\n👉 [Enter]를 누르면 결과물을 저장할 '폴더 선택 창'이 열립니다.")
    input("   [Enter] 입력...")

    root = tk.Tk()
    root.withdraw() 
    root.attributes('-topmost', True)
    selected_dir = filedialog.askdirectory(initialdir=BASE_DIR)
    root.destroy()

    if selected_dir:
        input_dir = selected_dir.replace('/', '\\')
    else:
        input_dir = default_dir
    
    try:
        with open(ENV_PATH, "w", encoding="utf-8") as f:
            f.write(f"AUTHOR_NAME={input_name}\n")
            f.write(f"DEFAULT_WORK_LOCATION={input_loc}\n")
            f.write(f"BASE_OUTPUT_DIR={input_dir}\n")
        print(f"\n✅ 설정 저장 완료!")
    except Exception as e:
        print(f"❌ 설정 저장 실패: {e}")
        sys.exit()

load_dotenv(ENV_PATH)
AUTHOR_NAME = os.getenv("AUTHOR_NAME", "User")
DEPARTMENT = os.getenv("DEPARTMENT", "TE")
EMPLOYEE_ID = os.getenv("EMPLOYEE_ID", "000")
DEFAULT_WORK_LOCATION = os.getenv("DEFAULT_WORK_LOCATION", "본사").strip() or "본사"
BASE_OUTPUT_DIR = os.getenv("BASE_OUTPUT_DIR", os.path.join(BASE_DIR, "Output"))
TEMPLATE_PATH = os.path.join(BASE_DIR, "Template", "Daily_Report_Template.xlsx")

CELL_MAP = {
    "report_date": "C7", "department": "I6", "work_location": "I7",
    "author_name": "C8", "headcount": "I8", "work_content": "D12",
    "tomorrow_work": "B18", "notes": "B23",
}

# ==========================================================
# 2. 유틸리티 함수 (JSON 관리 포함)
# ==========================================================

def get_week_dates():
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    dates = []
    for i in range(5):
        day = monday + timedelta(days=i)
        weekday_kor = ["월", "화", "수", "목", "금", "토", "일"][day.weekday()]
        dates.append(f"{day.day}({weekday_kor})")
    return dates

def get_monday_str():
    """이번 주 월요일 날짜를 문자열로 반환 (주차 식별용)"""
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    return monday.strftime("%Y-%m-%d")

def parse_doc_number(raw_value: str, auto_num: str):
    text = (raw_value or "").strip()
    if not text:
        return auto_num
    if text.isdigit():
        return f"{int(text):03d}"
    if INVALID_FILENAME_CHARS.search(text) or text.endswith(".") or text.endswith(" "):
        return None
    return text

def normalize_weekly_data(raw_data: dict, default_loc: str) -> dict:
    """weekly_data JSON 구조를 강제 정규화하여 반환한다.

    보장하는 불변 조건:
      - locations  : 정확히 5개의 비어있지 않은 str 목록
                     (값 부족·비문자열·빈 문자열 → default_loc으로 채움)
      - holiday_indices : 0~4 범위 int로만 구성된 중복 없는 정렬 목록
                          (범위 이탈·비정수·중복 → 제거)
    week_start 는 변경하지 않는다.
    """
    result = dict(raw_data)

    # locations 정규화
    raw_locs = result.get("locations", [])
    if not isinstance(raw_locs, list):
        raw_locs = []
    locs = []
    for i in range(5):
        val = raw_locs[i] if i < len(raw_locs) else ""
        locs.append(str(val).strip() if isinstance(val, str) and str(val).strip() else default_loc)
    result["locations"] = locs

    # holiday_indices 정규화
    raw_idx = result.get("holiday_indices", [])
    if not isinstance(raw_idx, list):
        raw_idx = []
    h_indices = []
    for x in raw_idx:
        try:
            v = int(x)
            if 0 <= v <= 4:
                h_indices.append(v)
        except (TypeError, ValueError):
            pass
    result["holiday_indices"] = sorted(set(h_indices))

    return result

def load_weekly_data():
    """
    JSON 파일에서 이번 주 근무지 정보를 불러오거나 초기화.
    검증 통과 시 정규화된 데이터 사용, 실패 시 안전 기본값 복구.
    반환값: (data, is_new_week) 튜플 형태
    """
    current_monday = get_monday_str()
    default_data = {
        "week_start": current_monday,
        "locations": [DEFAULT_WORK_LOCATION] * 5,
        "holiday_indices": [],
    }

    if os.path.exists(DATA_FILE_PATH):
        try:
            with open(DATA_FILE_PATH, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if raw.get("week_start") == current_monday:
                return normalize_weekly_data(raw, DEFAULT_WORK_LOCATION), False
        except (json.JSONDecodeError, IOError):
            pass

    return default_data, True

def save_weekly_data(data):
    """근무지 정보를 JSON 파일에 저장"""
    try:
        with open(DATA_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"⚠️ 데이터 저장 실패: {e}")

def get_auto_doc_number(year_dir: str) -> str:
    if not os.path.exists(year_dir): return "001"
    max_num = 0
    pattern = re.compile(r"_(\d{3})\s*@\s*Daily Report")
    for _, _, files in os.walk(year_dir):
        for f in files:
            if not f.lower().endswith(".xlsx"): continue
            match = pattern.search(f)
            if match:
                try:
                    num = int(match.group(1))
                    if num > max_num: max_num = num
                except ValueError: continue
    return f"{max_num + 1:03d}"

# ==========================================================
# 3. 엑셀 및 PDF 생성 함수
# ==========================================================

def generate_excel_draft(report_data: dict, output_xlsx: str, template_path: str, weekly_locations: list) -> bool:
    excel = None
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        
        if not os.path.exists(template_path):
            print(f"🚫 템플릿 없음: {template_path}")
            return False

        wb = excel.Workbooks.Open(template_path)
        ws = wb.Worksheets(1)
        
        # 기본 데이터 매핑
        for key, addr in CELL_MAP.items():
            val = report_data.get(key)
            if val and addr:
                cell = ws.Range(addr)
                cell.Value = str(val)
                if key in ["work_content", "tomorrow_work", "notes"]:
                    cell.WrapText = True
                    cell.VerticalAlignment = -4160
                    cell.HorizontalAlignment = -4131

        # 주간 날짜 (D15~H15)
        week_cells = ["D15", "E15", "F15", "G15", "H15"]
        for cell, value in zip(week_cells, get_week_dates()):
            ws.Range(cell).Value = value

        # =================================================================
        # [수정됨] 주간 근무지 입력 및 색상 처리 로직
        # =================================================================
        location_cells = ["D16", "E16", "F16", "G16", "H16"]
        
        # 빨간색으로 표시할 키워드 목록
        red_keywords = ["공휴일", "휴가", "연차", "휴일"] 

        for i, cell_addr in enumerate(location_cells):
            cell = ws.Range(cell_addr)
            val = weekly_locations[i]
            cell.Value = val
            
            # 입력된 값에 키워드가 포함되어 있는지 확인 (예: "여름휴가"도 빨간색 됨)
            is_red = False
            if val: # 값이 있을 때만 체크
                for keyword in red_keywords:
                    if keyword in str(val): # '휴가' 글자가 포함되면 True
                        is_red = True
                        break
            
            if is_red:
                cell.Font.Color = 255  # 빨간색 (Red)
                # cell.Font.Bold = True  # (선택사항) 굵게 하고 싶으면 주석 해제
            else:
                cell.Font.Color = 0    # 검정색 (Black) - 기본값
        # =================================================================

        wb.SaveAs(output_xlsx)
        wb.Close(SaveChanges=True)
        return True

    except Exception as e:
        print(f"🚫 엑셀 생성 실패: {e}")
        return False
    finally:
        if excel: excel.Quit()

def export_final_reports(xlsx_path: str, pdf_path: str, png_path: str) -> bool:
    if not os.path.exists(xlsx_path): return False
    excel = None
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True # 캡처 위해 보이기 설정
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(xlsx_path)
        ws = wb.Worksheets(1)
        ws.Activate()
        time.sleep(1)

        wb.ExportAsFixedFormat(0, pdf_path)

        rng = ws.Range("A1:I25")
        rng.Select()
        time.sleep(0.5)
        rng.CopyPicture(1, 2)
        
        chart = ws.Shapes.AddChart2()
        chart.Select()
        excel.Selection.Width = rng.Width
        excel.Selection.Height = rng.Height
        chart.Chart.Paste()
        chart.Chart.Export(png_path)
        chart.Delete()

        wb.Close(SaveChanges=False)
        return True
    except Exception as e:
        print(f"🚫 PDF 변환 실패: {e}")
        return False
    finally:
        if excel: excel.Quit()

# ==========================================================
# 4. 메인 실행
# ==========================================================

if __name__ == "__main__":
    print("[1/3] 설정 확인 및 데이터 로드...")
    
    if not os.path.exists(TEMPLATE_PATH):
        print("❌ 템플릿 파일 없음.")
        exit()

    # 1. 이번 주 데이터(JSON) 불러오기 [수정됨]
    weekly_data, is_new_week = load_weekly_data()  # 변수 두 개로 받음
    locations = weekly_data["locations"]

    # 2. 오늘 날짜 확인
    now = datetime.now()
    today_idx = now.weekday() # 0:월 ~ 6:일
    days_kor = ['월', '화', '수', '목', '금', '토', '일']

    # 3. 경로 설정
    year_folder = f"{now.year}_Daily Report"
    month_folder = f"{now.month}월_Daily Report"
    year_output_dir = os.path.join(BASE_OUTPUT_DIR, year_folder)
    final_output_dir = os.path.join(year_output_dir, month_folder)
    if not os.path.exists(final_output_dir): os.makedirs(final_output_dir, exist_ok=True)

    auto_num = get_auto_doc_number(year_output_dir)
    print(f"📂 저장 경로: {final_output_dir}")
    
    while True:
        user_input = input(f"👉 문서 번호 입력 (Enter 시 '{auto_num}'): ")
        user_num = parse_doc_number(user_input, auto_num)
        if user_num is not None:
            break
        print("⚠️ 문서 번호에 사용할 수 없는 문자가 있습니다. (예: < > : \" / \\ | ? *)")

    # -----------------------------------------------------------
    # [핵심 수정] 근무지 입력 로직 (월요일엔 주간 계획 가능)
    # -----------------------------------------------------------
    
    monday_date = now.date() - timedelta(days=now.weekday())

    if 0 <= today_idx <= 4:
        edit_all = False
        
        # [조건 변경] 
        # 1. 오늘이 월요일이거나(today_idx == 0)
        # 2. 이번 주 들어 처음 실행하는 경우(is_new_week == True)
        if today_idx == 0 or is_new_week:
            
            # 메시지를 상황에 맞게 조금 다르게 출력 (선택 사항)
            if today_idx == 0:
                print("\n📅 오늘은 월요일입니다.")
            else:
                print(f"\n📅 이번 주 데이터가 없습니다. (오늘: {days_kor[today_idx]}요일)")

            q = input("👉 이번 주 근무 계획을 일괄 입력하시겠습니까? (y/n, Enter=No): ").strip().lower()
            if q == 'y':
                edit_all = True
        
        # 2) 입력 루프 (전체 수정이면 0~4, 아니면 오늘만)
        range_to_edit = range(5) if edit_all else range(today_idx, today_idx + 1)
        
        for i in range_to_edit:
            day_name = days_kor[i]
            suggestion = locations[i] # 기존 저장값 or 기본값
            
            # 질문 메시지 다르게 표시
            if edit_all:
                prompt_msg = f"👉 {day_name}요일 근무지 (Enter 시 '{suggestion}'): "
            else:
                prompt_msg = f"👉 오늘({day_name}) 근무지 (Enter 시 '{suggestion}'): "

            loc_input = input(prompt_msg).strip()
            
            # 입력값이 있으면 업데이트, 없으면 기존 값 유지
            if loc_input:
                locations[i] = loc_input
        
        # 3) 변경된 리스트 저장
        weekly_data["locations"] = locations
        save_weekly_data(weekly_data)
        print("-" * 50)
        print("💾 근무지 정보가 업데이트되었습니다.")

    # 오늘의 최종 근무지 (엑셀 헤더용)
    if 0 <= today_idx <= 4:
        current_location_header = locations[today_idx]
    else:
        current_location_header = DEFAULT_WORK_LOCATION # 주말엔 기본값

    # -----------------------------------------------------------

    # 파일명 생성
    year_short = now.strftime('%y')
    base_name = f"FMS{year_short}_{DEPARTMENT}{EMPLOYEE_ID}_{user_num} @ Daily Report_{now.strftime('%y%m%d')}"
    output_xlsx = os.path.join(final_output_dir, base_name + ".xlsx")
    output_pdf = os.path.join(final_output_dir, base_name + ".pdf")
    output_png = os.path.join(final_output_dir, base_name + ".png")

    if os.path.exists(output_xlsx):
        overwrite = input(f"⚠️ 이미 동일한 파일이 존재합니다: {base_name}.xlsx\n덮어쓰시겠습니까? (y/n): ").strip().lower()
        if overwrite != "y":
            print("작업을 취소했습니다.")
            exit()

    report_data = {
        "report_date": now.strftime("%Y-%m-%d"),
        "department": DEPARTMENT,
        "work_location": current_location_header, # 상단 헤더
        "author_name": AUTHOR_NAME,
        "headcount": 1,
        "doc_number": user_num,
        "work_content": "", "tomorrow_work": "", "notes": ""
    }

    print("\n[2/3] 엑셀 초안 생성 중...")
    # locations 리스트(5일치)를 넘겨줌
    if generate_excel_draft(report_data, output_xlsx, TEMPLATE_PATH, locations):
        print(f"✅ 엑셀 생성 완료: {output_xlsx}")
        os.startfile(output_xlsx) 
    else:
        print("❌ 오류 발생.")
        exit()

    print("\n" + "="*50)
    print("📝 엑셀 파일 작성 후 '저장(Ctrl+S)' 하고 닫아주세요.")
    print("="*50)
    input("👉 완료 후 [Enter]를 누르면 PDF 변환 시작...")

    print("\n[3/3] PDF/PNG 변환 중...")
    if export_final_reports(output_xlsx, output_pdf, output_png):
        print(f"🎉 완료! PDF: {output_pdf}")
        try: os.startfile(output_pdf)
        except OSError: pass
    else:
        print("⚠️ 변환 실패")
    
    input("\n종료하려면 Enter...")
