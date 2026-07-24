import os
import sys
import re
import time
import json
import csv
import logging
from logging.handlers import RotatingFileHandler
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext
import win32com.client as win32
from datetime import datetime, date, timedelta
from dotenv import load_dotenv
import threading
import pythoncom
import urllib.request
import urllib.parse

# ==========================================================
# 1. 설정 및 전역 변수 초기화
# ==========================================================

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
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# --- 1.1 로깅 설정 ---
LOG_DIR = os.path.join(BASE_DIR, 'Logs')
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE_PATH = os.path.join(LOG_DIR, 'daily_report.log')

logger = logging.getLogger("DailyReport")
logger.setLevel(logging.INFO)

if not logger.handlers:
    file_handler = RotatingFileHandler(
        LOG_FILE_PATH,
        maxBytes=2 * 1024 * 1024,
        backupCount=5,
        encoding='utf-8'
    )
    formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

def open_log_folder():
    """로그 폴더를 파일 탐색기로 엽니다."""
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR, exist_ok=True)
    try:
        os.startfile(LOG_DIR)
        logger.info("사용자가 UI에서 로그 폴더를 열었습니다.")
    except AttributeError:
        import subprocess
        if sys.platform == "darwin":
            subprocess.run(["open", LOG_DIR])
        else:
            subprocess.run(["xdg-open", LOG_DIR])
        logger.info("사용자가 UI에서 로그 폴더를 열었습니다.")
    except Exception:
        logger.exception("로그 폴더 열기 실패")

# 엑셀 셀 매핑 정보
CELL_MAP = {
    "report_date": "C7", "department": "I6", "work_location": "I7",
    "author_name": "C8", "headcount": "I8", "work_content": "D12",
    "tomorrow_work": "B18", "notes": "B23",
}

# 기본 메일 템플릿
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
OUTLOOK_TOKEN_PATTERN = re.compile(r"\[\[[^\[\]\r\n]+\]\]|\{[^{}\r\n]+\}")

# GUI 요소 전역 변수
status_label = None
start_button = None
settings_button = None
outlook_button = None

# ==========================================================
# 2. 유틸리티 및 설정 관련 함수
# ==========================================================

def get_monday_str():
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    return monday.strftime("%Y-%m-%d")

def normalize_default_location(value: str) -> str:
    value = (value or "").strip()
    return value if value else "본사"

def parse_doc_number(raw_value: str, auto_num: str):
    text = (raw_value or "").strip()
    if not text:
        return auto_num
    if text.isdigit():
        return f"{int(text):03d}"
    if INVALID_FILENAME_CHARS.search(text) or text.endswith(".") or text.endswith(" "):
        return None
    return text

def resolve_addressbook_csv_path() -> str:
    candidates = [
        os.path.join(BASE_DIR, ADDRESSBOOK_FILENAME),
        os.path.join(BUNDLE_DIR, ADDRESSBOOK_FILENAME),
    ]
    checked = set()
    for path in candidates:
        key = os.path.normcase(os.path.abspath(path))
        if key in checked:
            continue
        checked.add(key)
        if os.path.exists(path):
            return path
    return candidates[0]

def normalize_recipient_addresses(raw_text: str) -> list:
    recipients = []
    seen = set()
    for token in (raw_text or "").split(";"):
        email = token.strip()
        if not email:
            continue
        email_key = email.lower()
        if email_key in seen:
            continue
        seen.add(email_key)
        recipients.append(email)
    return recipients

def dedupe_email_list(emails: list) -> list:
    recipients = []
    seen = set()
    for email in emails:
        value = str(email).strip()
        if not value:
            continue
        email_key = value.lower()
        if email_key in seen:
            continue
        seen.add(email_key)
        recipients.append(value)
    return recipients

def build_outlook_template_values(report_data: dict) -> dict:
    r_date = report_data.get("report_date", "")
    try:
        dt = datetime.strptime(r_date, "%Y-%m-%d")
        year_str = str(dt.year)
        month_str = str(dt.month)
        day_str = str(dt.day)
        weekday_str = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
    except (ValueError, IndexError):
        year_str, month_str, day_str, weekday_str = "", "", "", ""

    data_map = {
        "report_date": r_date,
        "year": year_str,
        "month": month_str,
        "day": day_str,
        "weekday": weekday_str,
        "department": report_data.get("department", ""),
        "author_name": report_data.get("author_name", ""),
    }
    return {token: str(data_map.get(key, "")) for token, key in OUTLOOK_TOKEN_ALIASES.items()}

def render_outlook_template(template_text: str, template_values: dict) -> str:
    text = template_text or ""
    for token, value in template_values.items():
        text = text.replace(token, value)
    return text

def find_unknown_outlook_tokens(text: str) -> list:
    unknown = []
    seen = set()
    for token in OUTLOOK_TOKEN_PATTERN.findall(text or ""):
        if token in OUTLOOK_TOKEN_ALIASES or token in seen:
            continue
        seen.add(token)
        unknown.append(token)
    return unknown

def insert_token_into_entry(entry_widget, token: str):
    try:
        entry_widget.insert(entry_widget.index(tk.INSERT), token)
    except Exception:
        entry_widget.insert(tk.END, token)
    entry_widget.focus_set()

def insert_token_into_text(text_widget, token: str):
    text_widget.insert(tk.INSERT, token)
    text_widget.focus_set()

def load_addressbook_contacts() -> list:
    csv_path = resolve_addressbook_csv_path()
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"주소록 파일을 찾을 수 없습니다: {csv_path}")

    contacts = []
    seen_emails = set()

    try:
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            if not reader.fieldnames or "이메일" not in reader.fieldnames:
                raise ValueError("주소록 형식 오류: '이메일' 헤더가 필요합니다.")

            for row in reader:
                email = (row.get("이메일", "") or "").strip()
                if not email:
                    continue

                email_key = email.lower()
                if email_key in seen_emails:
                    continue
                seen_emails.add(email_key)

                name = (row.get("이름", "") or "").strip()
                title = (row.get("직책", "") or "").strip()
                dept = (row.get("부서", "") or "").strip()
                company = (row.get("회사", "") or "").strip()

                contacts.append({
                    "name": name,
                    "email": email,
                    "email_lower": email_key,
                    "title": title,
                    "department": dept,
                    "company": company,
                    "label": f"{name} ({email})" if name else email,
                    "search_text": " ".join([name, email, title, dept, company]).lower(),
                })
    except UnicodeDecodeError as e:
        raise ValueError(f"주소록 인코딩 오류: {e}") from e
    except csv.Error as e:
        raise ValueError(f"주소록 CSV 파싱 오류: {e}") from e

    if not contacts:
        raise ValueError("주소록에 사용할 수 있는 이메일이 없습니다.")
    return contacts

def merge_selected_recipients(existing_text: str, selected_emails: list, addressbook_email_keys: set) -> str:
    existing_emails = normalize_recipient_addresses(existing_text)
    external_emails = [email for email in existing_emails if email.lower() not in addressbook_email_keys]
    final_emails = dedupe_email_list((selected_emails or []) + external_emails)
    return ";".join(final_emails)

def ensure_initial_setup(parent):
    """초기 .env가 없으면 설정 마법사 실행. parent는 메인 Tk 인스턴스."""
    if not os.path.exists(ENV_PATH):
        messagebox.showinfo("초기 설정", "⚙️ 초기 환경 설정 파일(.env)이 없습니다.\n설정을 시작합니다.", parent=parent)

        # 1. 기본 정보
        while True:
            input_name = simpledialog.askstring("설정 (1/4)", "👉 작성자 성함을 입력해주세요:", parent=parent)
            if input_name and input_name.strip(): break

        input_dept = simpledialog.askstring("설정 (2/4)", "👉 부서 코드를 입력해주세요 (예: TE):", parent=parent)
        if not input_dept: input_dept = "TE"

        input_id = simpledialog.askstring("설정 (3/4)", "👉 사원번호를 입력해주세요 (예: 044):", parent=parent)
        if not input_id: input_id = "000"
        elif input_id.strip().isdigit(): input_id = f"{int(input_id.strip()):03d}"

        input_loc = simpledialog.askstring("설정 (4/4)", "👉 기본 근무지를 입력해주세요 (예: 본사):", parent=parent)
        input_loc = normalize_default_location(input_loc)

        # 2. 아웃룩 사용 여부 질문
        use_outlook = messagebox.askyesno(
            "설정",
            "📧 '아웃룩 메일 초안 자동 작성' 기능을 사용하시겠습니까?\n(나중에 설정에서 변경 가능)",
            parent=parent
        )
        outlook_val = "True" if use_outlook else "False"

        # 3. 폴더 설정
        messagebox.showinfo("설정", "👉 확인을 누르면 결과물을 저장할 '폴더 선택 창'이 열립니다.", parent=parent)
        default_dir = os.path.join(BASE_DIR, "Output")
        selected_dir = filedialog.askdirectory(initialdir=BASE_DIR, title="결과물 저장 폴더 선택", parent=parent)
        input_dir = selected_dir.replace('/', '\\') if selected_dir else default_dir

        try:
            with open(ENV_PATH, "w", encoding="utf-8") as f:
                f.write(f"AUTHOR_NAME={input_name.strip()}\n")
                f.write(f"DEPARTMENT={input_dept.strip()}\n")
                f.write(f"EMPLOYEE_ID={input_id.strip()}\n")
                f.write(f"DEFAULT_WORK_LOCATION={input_loc}\n")
                f.write(f"BASE_OUTPUT_DIR={input_dir}\n")
                # 아웃룩 관련 설정 초기화
                f.write(f"OUTLOOK_ENABLE={outlook_val}\n")
                f.write(f"OUTLOOK_TO=\n")
                f.write(f"OUTLOOK_CC=\n")
                f.write(f"OUTLOOK_SENDER=\n")
                f.write(f"OUTLOOK_SUBJECT={DEFAULT_SUBJECT}\n")
                f.write(f"OUTLOOK_BODY={json.dumps(DEFAULT_BODY, ensure_ascii=False)}\n")

            messagebox.showinfo("완료", "✅ 설정 저장이 완료되었습니다.", parent=parent)
        except Exception as e:
            messagebox.showerror("오류", f"❌ 설정 저장 실패: {e}", parent=parent)
            sys.exit()

def migrate_env_if_needed():
    """이전 버전 .env에 누락된 필드가 있으면 기본값으로 추가한다."""
    if not os.path.exists(ENV_PATH):
        return

    env_dict = {}
    with open(ENV_PATH, "r", encoding="utf-8") as f:
        for line in f:
            if '=' in line:
                k, v = line.split('=', 1)
                env_dict[k.strip()] = v.strip()

    defaults = {
        "DEPARTMENT": "TE",
        "EMPLOYEE_ID": "000",
        "OUTLOOK_ENABLE": "False",
        "OUTLOOK_TO": "",
        "OUTLOOK_CC": "",
        "OUTLOOK_SENDER": "",
        "OUTLOOK_SUBJECT": DEFAULT_SUBJECT,
        "OUTLOOK_BODY": json.dumps(DEFAULT_BODY, ensure_ascii=False),
    }

    added = []
    for key, default_val in defaults.items():
        if key not in env_dict:
            env_dict[key] = default_val
            added.append(key)

    if added:
        with open(ENV_PATH, "w", encoding="utf-8") as f:
            for k, v in env_dict.items():
                f.write(f"{k}={v}\n")

# --- 데이터 정규화 ---
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

# --- 설정 로드/저장 ---
def load_json_locations():
    current_monday = get_monday_str()
    default_loc = normalize_default_location(os.getenv("DEFAULT_WORK_LOCATION", "본사"))

    if os.path.exists(DATA_FILE_PATH):
        try:
            with open(DATA_FILE_PATH, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if raw.get("week_start") == current_monday:
                return normalize_weekly_data(raw, default_loc)["locations"]
        except (json.JSONDecodeError, IOError): pass
    return [default_loc] * 5

def load_holiday_indices() -> list:
    """이번 주 공휴일 day 인덱스(0=월~4=금) 목록 반환."""
    current_monday = get_monday_str()
    default_loc = normalize_default_location(os.getenv("DEFAULT_WORK_LOCATION", "본사"))

    if os.path.exists(DATA_FILE_PATH):
        try:
            with open(DATA_FILE_PATH, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if raw.get("week_start") == current_monday:
                return normalize_weekly_data(raw, default_loc)["holiday_indices"]
        except (json.JSONDecodeError, IOError): pass
    return []

def save_basic_settings_gui(widgets, window):
    new_name = widgets['name'].get().strip()
    new_dept = widgets['dept'].get().strip()
    new_id = widgets['id'].get().strip()
    new_def_loc = normalize_default_location(widgets['def_loc'].get().strip())
    new_path = widgets['path'].get().strip()
    new_locations = [w.get().strip() for w in widgets['loc_entries']]

    if not all([new_name, new_dept, new_id, new_path]):
        messagebox.showwarning("경고", "필수 항목은 비워둘 수 없습니다.", parent=window)
        return

    try:
        env_dict = {}
        if os.path.exists(ENV_PATH):
            with open(ENV_PATH, "r", encoding="utf-8") as f:
                for line in f:
                    if '=' in line:
                        k, v = line.split('=', 1)
                        env_dict[k.strip()] = v.strip()
        
        env_dict["AUTHOR_NAME"] = new_name
        env_dict["DEPARTMENT"] = new_dept
        env_dict["EMPLOYEE_ID"] = f"{int(new_id):03d}" if new_id.isdigit() else new_id
        env_dict["DEFAULT_WORK_LOCATION"] = new_def_loc
        env_dict["BASE_OUTPUT_DIR"] = new_path

        with open(ENV_PATH, "w", encoding="utf-8") as f:
            for k, v in env_dict.items():
                f.write(f"{k}={v}\n")
        
        load_dotenv(ENV_PATH, override=True)

        current_monday = get_monday_str()
        existing_json = {}
        if os.path.exists(DATA_FILE_PATH):
            try:
                with open(DATA_FILE_PATH, "r", encoding="utf-8") as f:
                    existing_json = json.load(f)
            except (json.JSONDecodeError, IOError): pass
        if existing_json.get("week_start") != current_monday:
            existing_json = {}
        existing_json["week_start"] = current_monday
        existing_json["locations"] = new_locations
        existing_json = normalize_weekly_data(existing_json, new_def_loc)
        with open(DATA_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(existing_json, f, ensure_ascii=False, indent=4)

        messagebox.showinfo("성공", "✅ 기본 설정이 저장되었습니다!", parent=window)
        window.destroy()

    except Exception as e:
        messagebox.showerror("오류", f"저장 중 오류 발생: {e}", parent=window)

def save_outlook_settings_gui(widgets, window):
    enable = "True" if widgets['enable_var'].get() else "False"
    to_addr = ";".join(normalize_recipient_addresses(widgets['to'].get()))
    cc_addr = ";".join(normalize_recipient_addresses(widgets['cc'].get()))
    sender_addr = widgets['sender'].get().strip()
    subject = widgets['subject'].get().strip()
    body = widgets['body'].get("1.0", tk.END).strip()
    unknown_tokens = find_unknown_outlook_tokens(subject) + find_unknown_outlook_tokens(body)
    unknown_tokens = list(dict.fromkeys(unknown_tokens))

    if unknown_tokens:
        allowed_text = ", ".join(OUTLOOK_FRIENDLY_TOKENS)
        unknown_text = ", ".join(unknown_tokens)
        messagebox.showerror(
            "오류",
            f"알 수 없는 변수가 있습니다.\n\n{unknown_text}\n\n사용 가능한 변수:\n{allowed_text}\n\n"
            "변수는 [[...]] 형식만 지원합니다.",
            parent=window
        )
        return

    try:
        env_dict = {}
        if os.path.exists(ENV_PATH):
            with open(ENV_PATH, "r", encoding="utf-8") as f:
                for line in f:
                    if '=' in line:
                        k, v = line.split('=', 1)
                        env_dict[k.strip()] = v.strip()

        env_dict["OUTLOOK_ENABLE"] = enable
        env_dict["OUTLOOK_TO"] = to_addr
        env_dict["OUTLOOK_CC"] = cc_addr
        env_dict["OUTLOOK_SENDER"] = sender_addr
        env_dict["OUTLOOK_SUBJECT"] = subject
        env_dict["OUTLOOK_BODY"] = json.dumps(body, ensure_ascii=False)

        with open(ENV_PATH, "w", encoding="utf-8") as f:
            for k, v in env_dict.items():
                f.write(f"{k}={v}\n")
        
        load_dotenv(ENV_PATH, override=True)
        messagebox.showinfo("성공", "✅ 아웃룩 설정이 저장되었습니다!", parent=window)
        window.destroy()

    except Exception as e:
        messagebox.showerror("오류", f"저장 중 오류 발생: {e}", parent=window)

def open_folder_dialog_gui(entry_widget, window):
    initial = entry_widget.get() or BASE_DIR
    selected_dir = filedialog.askdirectory(initialdir=initial, title="새로운 저장 경로 선택", parent=window)
    if selected_dir:
        win_path = selected_dir.replace('/', '\\')
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, win_path)

def open_addressbook_selector(parent_window, target_entry):
    try:
        contacts = load_addressbook_contacts()
    except Exception as e:
        messagebox.showwarning(
            "주소록 안내",
            f"주소록을 불러오지 못했습니다.\n\n{e}\n\n수동 입력은 계속 사용할 수 있습니다.",
            parent=parent_window
        )
        return

    popup = tk.Toplevel(parent_window)
    popup.title("주소록 선택")
    popup.geometry("540x440")
    popup.resizable(False, False)
    popup.transient(parent_window)
    popup.grab_set()

    tk.Label(popup, text="이름/이메일/직책/부서/회사 검색").pack(anchor="w", padx=10, pady=(10, 2))
    search_var = tk.StringVar()
    search_entry = tk.Entry(popup, textvariable=search_var)
    search_entry.pack(fill="x", padx=10)

    tk.Label(popup, text="여러 명 선택 후 적용하세요.", fg="gray").pack(anchor="w", padx=10, pady=(4, 4))

    listbox = tk.Listbox(popup, selectmode=tk.MULTIPLE)
    listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    selected_email_keys = set()
    addressbook_email_keys = {item["email_lower"] for item in contacts}
    current_entry_emails = normalize_recipient_addresses(target_entry.get())
    for email in current_entry_emails:
        key = email.lower()
        if key in addressbook_email_keys:
            selected_email_keys.add(key)

    current_view = []

    def sync_selection_from_view():
        if not current_view:
            return
        visible_keys = {item["email_lower"] for item in current_view}
        selected_email_keys.difference_update(visible_keys)
        for idx in listbox.curselection():
            if 0 <= idx < len(current_view):
                selected_email_keys.add(current_view[idx]["email_lower"])

    def render_list():
        nonlocal current_view
        keyword = search_var.get().strip().lower()
        if keyword:
            current_view = [item for item in contacts if keyword in item["search_text"]]
        else:
            current_view = contacts[:]

        listbox.delete(0, tk.END)
        for idx, item in enumerate(current_view):
            listbox.insert(tk.END, item["label"])
            if item["email_lower"] in selected_email_keys:
                listbox.selection_set(idx)

    def on_search_change(*_):
        sync_selection_from_view()
        render_list()

    def apply_selection():
        sync_selection_from_view()
        selected_emails = [item["email"] for item in contacts if item["email_lower"] in selected_email_keys]
        merged = merge_selected_recipients(target_entry.get(), selected_emails, addressbook_email_keys)
        target_entry.delete(0, tk.END)
        target_entry.insert(0, merged)
        popup.destroy()

    btn_row = tk.Frame(popup)
    btn_row.pack(fill="x", padx=10, pady=(0, 10))
    tk.Button(btn_row, text="적용", command=apply_selection, bg="#0078D7", fg="white").pack(side="right")
    tk.Button(btn_row, text="취소", command=popup.destroy).pack(side="right", padx=(0, 6))

    search_var.trace_add("write", on_search_change)
    listbox.bind("<<ListboxSelect>>", lambda _evt: sync_selection_from_view())

    render_list()
    search_entry.focus_set()

# --- 설정 창 GUI ---
def open_basic_settings_window(root_window):
    load_dotenv(ENV_PATH, override=True)
    win = tk.Toplevel(root_window)
    win.title("기본 환경 설정")
    win.geometry("450x500")
    win.resizable(False, False)

    widgets = {}

    f_user = tk.LabelFrame(win, text="👤 기본 정보", padx=10, pady=5)
    f_user.pack(fill="x", padx=10, pady=5)
    
    row1 = tk.Frame(f_user); row1.pack(fill="x", pady=2)
    tk.Label(row1, text="이름:", width=6).pack(side="left")
    widgets['name'] = tk.Entry(row1, width=12); widgets['name'].pack(side="left")
    widgets['name'].insert(0, os.getenv("AUTHOR_NAME", ""))
    
    tk.Label(row1, text="부서:", width=6).pack(side="left")
    widgets['dept'] = tk.Entry(row1, width=10); widgets['dept'].pack(side="left")
    widgets['dept'].insert(0, os.getenv("DEPARTMENT", ""))

    row2 = tk.Frame(f_user); row2.pack(fill="x", pady=2)
    tk.Label(row2, text="사번:", width=6).pack(side="left")
    widgets['id'] = tk.Entry(row2, width=12); widgets['id'].pack(side="left")
    widgets['id'].insert(0, os.getenv("EMPLOYEE_ID", ""))

    tk.Label(row2, text="장소:", width=6).pack(side="left")
    widgets['def_loc'] = tk.Entry(row2, width=10); widgets['def_loc'].pack(side="left")
    widgets['def_loc'].insert(0, normalize_default_location(os.getenv("DEFAULT_WORK_LOCATION", "본사")))

    f_sched = tk.LabelFrame(win, text="📅 이번 주 근무지 스케줄", padx=10, pady=5)
    f_sched.pack(fill="x", padx=10, pady=5)
    days = ["월", "화", "수", "목", "금"]
    current_locs = load_json_locations()
    holiday_indices = load_holiday_indices()
    red_kw = ["공휴일", "휴가", "연차", "반차", "휴일"]
    widgets['loc_entries'] = []
    for i, day in enumerate(days):
        row = tk.Frame(f_sched); row.pack(fill="x", pady=2)
        tk.Label(row, text=f"{day}요일:", width=6).pack(side="left")
        val = str(current_locs[i])
        fg = "red" if i in holiday_indices or any(k in val for k in red_kw) else "black"
        e = tk.Entry(row); e.pack(side="left", fill="x", expand=True)
        e.insert(0, val); e.config(fg=fg)
        widgets['loc_entries'].append(e)

    f_path = tk.LabelFrame(win, text="📂 저장 경로", padx=10, pady=5)
    f_path.pack(fill="x", padx=10, pady=5)
    row_p = tk.Frame(f_path); row_p.pack(fill="x")
    widgets['path'] = tk.Entry(row_p); widgets['path'].pack(side="left", fill="x", expand=True)
    widgets['path'].insert(0, os.getenv("BASE_OUTPUT_DIR", ""))
    tk.Button(row_p, text="찾기", command=lambda: open_folder_dialog_gui(widgets['path'], win)).pack(side="right")

    tk.Button(win, text="💾 저장", command=lambda: save_basic_settings_gui(widgets, win), bg="#0078D7", fg="white", height=2).pack(fill="x", padx=10, pady=10)

def open_outlook_settings_window(root_window):
    load_dotenv(ENV_PATH, override=True)
    win = tk.Toplevel(root_window)
    win.title("아웃룩(Outlook) 설정")
    win.geometry("560x760")
    win.resizable(False, False)

    widgets = {}

    bottom_bar = tk.Frame(win)
    bottom_bar.pack(side="bottom", fill="x", padx=10, pady=10)

    f_switch = tk.Frame(win, pady=10)
    f_switch.pack(fill="x", padx=15)
    widgets['enable_var'] = tk.BooleanVar(value=(os.getenv("OUTLOOK_ENABLE", "False") == "True"))
    chk = tk.Checkbutton(
        f_switch,
        text="아웃룩 자동 작성 기능 사용",
        variable=widgets['enable_var'],
        font=("맑은 고딕", 12, "bold")
    )
    chk.pack(anchor="w")

    f_input = tk.LabelFrame(win, text="메일 양식 설정", padx=10, pady=10)
    f_input.pack(fill="both", expand=True, padx=10, pady=5)

    tip_msg = (
        "※ 사용 변수: [[년]], [[월]], [[일]]\n"
        "변수는 [[...]] 형식만 지원합니다."
    )
    tk.Label(f_input, text=tip_msg, fg="blue", font=("맑은 고딕", 9)).pack(anchor="w", pady=(0, 10))

    tk.Label(f_input, text="받는 사람 (여러 명은 ;로 구분):").pack(anchor="w")
    to_row = tk.Frame(f_input); to_row.pack(fill="x", pady=(0, 5))
    widgets['to'] = tk.Entry(to_row); widgets['to'].pack(side="left", fill="x", expand=True)
    widgets['to'].insert(0, os.getenv("OUTLOOK_TO", ""))
    tk.Button(
        to_row,
        text="주소록 선택",
        command=lambda: open_addressbook_selector(win, widgets['to'])
    ).pack(side="left", padx=(6, 0))

    tk.Label(f_input, text="참조 (CC):").pack(anchor="w")
    cc_row = tk.Frame(f_input); cc_row.pack(fill="x", pady=(0, 5))
    widgets['cc'] = tk.Entry(cc_row); widgets['cc'].pack(side="left", fill="x", expand=True)
    widgets['cc'].insert(0, os.getenv("OUTLOOK_CC", ""))
    tk.Button(
        cc_row,
        text="주소록 선택",
        command=lambda: open_addressbook_selector(win, widgets['cc'])
    ).pack(side="left", padx=(6, 0))

    tk.Label(f_input, text="보내는 사람 (SMTP, 선택):").pack(anchor="w")
    widgets['sender'] = tk.Entry(f_input); widgets['sender'].pack(fill="x", pady=(0, 5))
    widgets['sender'].insert(0, os.getenv("OUTLOOK_SENDER", ""))

    tk.Label(f_input, text="메일 제목:").pack(anchor="w")
    widgets['subject'] = tk.Entry(f_input); widgets['subject'].pack(fill="x", pady=(0, 5))
    widgets['subject'].insert(0, os.getenv("OUTLOOK_SUBJECT", DEFAULT_SUBJECT))

    tk.Label(f_input, text="메일 본문 (상단 인사말):").pack(anchor="w")
    widgets['body'] = scrolledtext.ScrolledText(f_input, height=8, font=("맑은 고딕", 9))
    widgets['body'].pack(fill="both", expand=True)
    
    saved_body = os.getenv("OUTLOOK_BODY", "")
    if saved_body:
        try: body_text = json.loads(saved_body)
        except (json.JSONDecodeError, ValueError): body_text = saved_body
    else:
        body_text = DEFAULT_BODY
    
    widgets['body'].insert("1.0", body_text)

    token_box = tk.LabelFrame(f_input, text="변수 삽입 (클릭)", padx=8, pady=6)
    token_box.pack(fill="x", pady=(8, 6))

    def add_token_buttons(container, on_click):
        row = None
        for idx, token in enumerate(OUTLOOK_FRIENDLY_TOKENS):
            if idx % 4 == 0:
                row = tk.Frame(container)
                row.pack(anchor="w", pady=(0, 2))
            tk.Button(row, text=token, width=9, command=lambda t=token: on_click(t)).pack(side="left", padx=(0, 4))

    preview_subject_var = tk.StringVar()
    preview_body_var = tk.StringVar()

    def update_preview(*_):
        preview_data = {
            "report_date": date.today().strftime("%Y-%m-%d"),
            "department": os.getenv("DEPARTMENT", ""),
            "author_name": os.getenv("AUTHOR_NAME", ""),
        }
        template_values = build_outlook_template_values(preview_data)

        subject_preview = render_outlook_template(widgets['subject'].get().strip(), template_values)
        body_preview_raw = render_outlook_template(widgets['body'].get("1.0", tk.END).strip(), template_values)
        body_preview = body_preview_raw.replace("\n", " ")
        if len(body_preview) > 120:
            body_preview = body_preview[:120] + "..."

        preview_subject_var.set(subject_preview if subject_preview else "(제목 미입력)")
        preview_body_var.set(body_preview if body_preview else "(본문 미입력)")

    def insert_subject_token(token):
        insert_token_into_entry(widgets['subject'], token)
        update_preview()

    def insert_body_token(token):
        insert_token_into_text(widgets['body'], token)
        update_preview()

    tk.Label(token_box, text="제목 변수").pack(anchor="w")
    subject_token_frame = tk.Frame(token_box)
    subject_token_frame.pack(anchor="w", pady=(2, 4))
    add_token_buttons(subject_token_frame, insert_subject_token)

    tk.Label(token_box, text="본문 변수").pack(anchor="w")
    body_token_frame = tk.Frame(token_box)
    body_token_frame.pack(anchor="w", pady=(2, 0))
    add_token_buttons(body_token_frame, insert_body_token)

    preview_box = tk.LabelFrame(f_input, text="치환 미리보기", padx=8, pady=6)
    preview_box.pack(fill="x", pady=(0, 4))
    tk.Label(preview_box, text="제목", font=("맑은 고딕", 9, "bold")).pack(anchor="w")
    tk.Label(preview_box, textvariable=preview_subject_var, wraplength=500, justify="left", fg="#1f4e79").pack(anchor="w", fill="x")
    tk.Label(preview_box, text="본문", font=("맑은 고딕", 9, "bold")).pack(anchor="w", pady=(4, 0))
    tk.Label(preview_box, textvariable=preview_body_var, wraplength=500, justify="left", fg="#555555").pack(anchor="w", fill="x")

    widgets['subject'].bind("<KeyRelease>", update_preview)
    widgets['subject'].bind("<FocusOut>", update_preview)
    widgets['body'].bind("<KeyRelease>", update_preview)
    widgets['body'].bind("<FocusOut>", update_preview)
    update_preview()

    tk.Button(
        bottom_bar,
        text="💾 아웃룩 설정 저장",
        command=lambda: save_outlook_settings_gui(widgets, win),
        bg="#28a745",
        fg="white",
        height=2
    ).pack(fill="x")




# ==========================================================
# 3. 데이터 처리 함수
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

def load_weekly_data(default_loc):
    current_monday = get_monday_str()
    default_data = {
        "week_start": current_monday,
        "locations": [default_loc] * 5,
        "holiday_indices": [],
    }
    if os.path.exists(DATA_FILE_PATH):
        try:
            with open(DATA_FILE_PATH, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if raw.get("week_start") == current_monday:
                return normalize_weekly_data(raw, default_loc), False
        except (json.JSONDecodeError, IOError): pass
    return default_data, True

def save_weekly_data(data):
    try:
        with open(DATA_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        logger.info("주간 근무지 데이터 저장 성공")
    except Exception:
        logger.exception("주간 근무지 데이터 저장 실패")

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

def get_holidays_this_week() -> dict:
    """공공데이터포털 API로 이번 주 월~금 공휴일 조회. {date: 공휴일명} 반환. 실패 시 {} 반환."""
    api_key = os.getenv("HOLIDAY_API_KEY", "").strip()
    if not api_key:
        logger.info("공휴일 API 키 미설정으로 조회 건너뜀")
        return {}

    today = date.today()
    monday = today - timedelta(days=today.weekday())
    friday = monday + timedelta(days=4)

    months_to_query = {(monday.year, monday.month), (friday.year, friday.month)}
    base_url = "https://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo"
    holidays = {}

    for year, month in months_to_query:
        try:
            query = urllib.parse.urlencode({
                "serviceKey": api_key,
                "solYear": str(year),
                "solMonth": f"{month:02d}",
                "numOfRows": "50",
                "_type": "json"
            })
            url = f"{base_url}?{query}"
            req = urllib.request.Request(url, headers={"Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            items = data.get("response", {}).get("body", {}).get("items", "")
            if not items:
                continue

            item_list = items.get("item", [])
            if isinstance(item_list, dict):
                item_list = [item_list]

            for item in item_list:
                if item.get("isHoliday", "N") != "Y":
                    continue
                loc_date_str = str(item.get("locdate", ""))
                dat_name = item.get("datName", "공휴일")
                if len(loc_date_str) == 8:
                    try:
                        d = date(int(loc_date_str[:4]), int(loc_date_str[4:6]), int(loc_date_str[6:8]))
                        if monday <= d <= friday:
                            holidays[d] = dat_name
                    except ValueError:
                        pass
            logger.info(f"공휴일 API 조회 성공 ({year}/{month})")
        except Exception:
            logger.exception(f"공휴일 API 조회 실패 ({year}/{month})")

    return holidays

# ==========================================================
# 4. 엑셀/PDF 생성 함수
# ==========================================================

def generate_excel_draft(report_data: dict, output_xlsx: str, template_path: str, weekly_locations: list, holiday_indices: list = None) -> bool:
    excel = None
    wb = None
    logger.info(f"엑셀 초안 생성 시작: {os.path.basename(output_xlsx)}")
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        if not os.path.exists(template_path):
            logger.error(f"엑셀 템플릿 없음: {template_path}")
            return False

        wb = excel.Workbooks.Open(template_path)
        ws = wb.Worksheets(1)

        for key, addr in CELL_MAP.items():
            val = report_data.get(key)
            if val and addr:
                cell = ws.Range(addr)
                cell.Value = str(val)
                if key in ["work_content", "tomorrow_work", "notes"]:
                    cell.WrapText = True; cell.VerticalAlignment = -4160; cell.HorizontalAlignment = -4131

        week_cells = ["D15", "E15", "F15", "G15", "H15"]
        for cell, value in zip(week_cells, get_week_dates()): ws.Range(cell).Value = value

        location_cells = ["D16", "E16", "F16", "G16", "H16"]
        red_keywords = ["공휴일", "휴가", "연차", "반차", "휴일"]
        for i, cell_addr in enumerate(location_cells):
            cell = ws.Range(cell_addr)
            is_holiday = bool(holiday_indices and i in holiday_indices)
            if is_holiday:
                val = "공휴일"
                is_red = True
            else:
                val = weekly_locations[i]
                is_red = False
                if val:
                    for keyword in red_keywords:
                        if keyword in str(val): is_red = True; break
            cell.Value = val
            if is_red: cell.Font.Color = 255
            else: cell.Font.Color = 0

        wb.SaveAs(output_xlsx)
        wb.Close(SaveChanges=False)
        wb = None
        logger.info("엑셀 초안 생성 완료")
        return True
    except Exception:
        logger.exception("엑셀 초안 생성 실패")
        return False
    finally:
        try:
            if wb: wb.Close(SaveChanges=False)
        except Exception: pass
        try:
            if excel: excel.Quit()
        except Exception: pass

def export_final_reports(xlsx_path: str, pdf_path: str, png_path: str) -> bool:
    if not os.path.exists(xlsx_path): return False
    excel = None
    wb = None
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.ScreenUpdating = True

        wb = excel.Workbooks.Open(xlsx_path)
        ws = wb.Worksheets(1)
        ws.Activate()

        time.sleep(2)

        wb.ExportAsFixedFormat(0, pdf_path)

        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1

        rng = ws.Range("A2:I16")
        rng.Select()
        time.sleep(1)

        rng.CopyPicture(1, 2)

        chart = ws.Shapes.AddChart2()
        chart.Select()
        excel.Selection.Width = rng.Width
        excel.Selection.Height = rng.Height
        chart.Chart.Paste()
        time.sleep(1)

        chart.Chart.Export(png_path)
        chart.Delete()

        wb.Close(SaveChanges=False)
        wb = None
        return True
    except Exception as e:
        print(f"🚫 PDF/PNG 변환 실패: {e}")
        return False
    finally:
        try:
            if wb: wb.Close(SaveChanges=False)
        except Exception: pass
        try:
            if excel: excel.Quit()
        except Exception: pass

# ==========================================================
# 5. [핵심] 아웃룩 연동 함수 (변수 분리 적용)
# ==========================================================
def create_outlook_draft(report_data, output_paths) -> str:
    """아웃룩 메일 초안을 생성한다.
    이 함수는 메일을 자동 전송하지 않고 초안 창만 연다.

    반환값:
      "created"  - 초안 생성 성공
      "disabled" - OUTLOOK_ENABLE=False 로 비활성 상태
      "failed"   - 초안 생성 중 오류 발생
    """
    if os.getenv("OUTLOOK_ENABLE", "False") != "True":
        logger.info("아웃룩 기능이 꺼져있어 건너뜁니다.")
        return "disabled"

    output_xlsx, output_pdf, output_png = output_paths
    
    # 1. 설정값 로드 및 치환
    target_email = os.getenv("OUTLOOK_TO", "")
    cc_email = os.getenv("OUTLOOK_CC", "")
    sender_email = os.getenv("OUTLOOK_SENDER", "").strip().lower()
    
    subject_tmpl = os.getenv("OUTLOOK_SUBJECT", DEFAULT_SUBJECT)
    raw_body = os.getenv("OUTLOOK_BODY", "")
    if raw_body:
        try: body_tmpl = json.loads(raw_body)
        except (json.JSONDecodeError, ValueError): body_tmpl = raw_body
    else:
        body_tmpl = DEFAULT_BODY

    template_values = build_outlook_template_values(report_data)
    final_subject = render_outlook_template(subject_tmpl, template_values)
    final_body_text = render_outlook_template(body_tmpl, template_values)

    # 2. 아웃룩 실행
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0) 

        if sender_email:
            selected_account = None
            try:
                accounts = outlook.Session.Accounts
                for i in range(1, int(accounts.Count) + 1):
                    acct = accounts.Item(i)
                    smtp = (getattr(acct, "SmtpAddress", "") or "").strip().lower()
                    if smtp == sender_email:
                        selected_account = acct
                        break
            except Exception as e:
                print(f"발신 계정 조회 실패: {e}")

            if selected_account is not None:
                try:
                    mail.SendUsingAccount = selected_account
                    print(f"발신 계정 설정: {sender_email}")
                except Exception as e:
                    print(f"발신 계정 설정 실패: {e}")
            else:
                print(f"지정한 발신 계정을 찾지 못했습니다: {sender_email}")

        if target_email: mail.To = target_email
        if cc_email: mail.CC = cc_email
        
        mail.Subject = final_subject
        
        if os.path.exists(output_pdf):
            mail.Attachments.Add(output_pdf)
        
        if os.path.exists(output_png):
            attachment = mail.Attachments.Add(output_png)
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "daily_report_img")
        
        html_content = final_body_text.replace("\n", "<br>")
        
        mail.HTMLBody = f"""
        <html>
        <body style="font-family: 'Malgun Gothic', sans-serif; font-size: 11pt;">
            <p>{html_content}</p>
            <br>
            <img src='cid:daily_report_img' width='650'>
        </body>
        </html>
        """
        
        # 안전 정책: 자동 전송(Send) 없이 초안 창만 표시한다.
        mail.Display()
        return "created"

    except Exception as e:
        print(f"아웃룩 연동 실패: {e}")
        return "failed"


# ==========================================================
# 6. 스레드 분리
# ==========================================================

def update_status(message):
    if status_label:
        status_label.after(0, lambda: status_label.config(text=f"상태: {message}"))

# --- Step 1. 엑셀 생성 스레드 ---
def thread_step1_create_excel(root_window, final_output_dir, report_data, locations, output_paths, template_path, holiday_indices=None):
    pythoncom.CoInitialize()
    output_xlsx, output_pdf, output_png = output_paths

    try:
        update_status("2/3 엑셀 초안 생성 중...")
        if generate_excel_draft(report_data, output_xlsx, template_path, locations, holiday_indices):
            update_status("엑셀 생성 완료. 작성 대기 중...")
            try: os.startfile(output_xlsx)
            except OSError: pass
            
            root_window.after(0, lambda: ask_user_to_continue(root_window, output_paths, report_data))
        else:
            update_status("❌ 엑셀 생성 오류 발생")
            reset_buttons()
    except Exception as e:
        update_status(f"에러 발생: {e}")
        reset_buttons()
    finally:
        pythoncom.CoUninitialize()

# --- Step 2. 사용자 확인 팝업 (메인 스레드) ---
def ask_user_to_continue(root_window, output_paths, report_data):
    response = messagebox.askokcancel(
        "작성 확인", 
        "📝 엑셀 파일이 열렸습니다.\n\n내용을 작성하고 '저장(Ctrl+S)'한 뒤,\n[확인]을 누르면 PDF 변환 및 메일 작성을 시작합니다.",
        parent=root_window
    )
    
    if response:
        start_step2_thread(root_window, output_paths, report_data)
    else:
        update_status("작업 취소됨")
        reset_buttons()

# --- Step 3. PDF 변환 및 아웃룩 스레드 ---
def start_step2_thread(root_window, output_paths, report_data):
    t = threading.Thread(target=thread_step2_export_pdf, args=(root_window, output_paths, report_data), daemon=True)
    t.start()

def thread_step2_export_pdf(root_window, output_paths, report_data):
    pythoncom.CoInitialize()
    output_xlsx, output_pdf, output_png = output_paths

    try:
        update_status("3/3 PDF/PNG 변환 중...")
        time.sleep(1)

        if export_final_reports(output_xlsx, output_pdf, output_png):
            update_status("메일 초안 생성 중...")

            outlook_status = create_outlook_draft(report_data, output_paths)
            if outlook_status == "created":
                update_status("🎉 모든 작업 완료! (메일 창 확인)")
            elif outlook_status == "disabled":
                update_status("✅ PDF/PNG 변환 완료 (아웃룩 자동 작성 비활성)")
                root_window.after(
                    0,
                    lambda: messagebox.showinfo(
                        "아웃룩 비활성",
                        "PDF/PNG 생성은 완료되었습니다.\n\n"
                        "아웃룩 자동 작성 기능이 꺼져 있어 메일 초안은 생성하지 않았습니다.\n"
                        "메인 화면의 '📧 아웃룩 설정'에서 사용 체크 후 저장하면 다음부터 자동 생성됩니다.",
                        parent=root_window,
                    ),
                )
                try: os.startfile(output_pdf)
                except OSError: pass
            else:  # "failed"
                update_status("⚠️ 아웃룩 생성 실패 — PDF를 직접 확인하세요")
                root_window.after(
                    0,
                    lambda: messagebox.showwarning(
                        "아웃룩 연동 안내",
                        "PDF/PNG 생성은 완료되었지만 아웃룩 메일 초안을 만들지 못했습니다.\n\n"
                        "Outlook 실행/로그인 상태와 기본 프로필 설정을 확인한 뒤 다시 시도해 주세요.\n"
                        "지금은 열리는 PDF를 첨부해 수동으로 발송할 수 있습니다.",
                        parent=root_window,
                    ),
                )
                try: os.startfile(output_pdf)
                except OSError: pass
        else:
            update_status("⚠️ PDF 변환 실패")
            root_window.after(0, lambda: messagebox.showerror("실패", "PDF 변환 중 오류가 발생했습니다.", parent=root_window))

    except Exception as e:
        update_status(f"에러 발생: {e}")
    finally:
        pythoncom.CoUninitialize()
        reset_buttons()

def reset_buttons():
    if start_button:
        start_button.after(0, lambda: start_button.config(state=tk.NORMAL))
        settings_button.after(0, lambda: settings_button.config(state=tk.NORMAL))
        outlook_button.after(0, lambda: outlook_button.config(state=tk.NORMAL))

# --- 시작 래퍼 함수 ---
def start_automation_wrapper(root_window):
    if not os.path.exists(TEMPLATE_PATH):
        messagebox.showerror("오류", "❌ 템플릿 파일이 없습니다.")
        return

    load_dotenv(ENV_PATH, override=True)
    AUTHOR_NAME = os.getenv("AUTHOR_NAME", "User")
    DEPARTMENT = os.getenv("DEPARTMENT", "TE")
    EMPLOYEE_ID = os.getenv("EMPLOYEE_ID", "000")
    DEFAULT_WORK_LOCATION = normalize_default_location(os.getenv("DEFAULT_WORK_LOCATION", "본사"))
    BASE_OUTPUT_DIR = os.getenv("BASE_OUTPUT_DIR", os.path.join(BASE_DIR, "Output"))

    weekly_data, is_new_week = load_weekly_data(DEFAULT_WORK_LOCATION)
    locations = weekly_data["locations"]
    holiday_indices = weekly_data["holiday_indices"]

    now = datetime.now()
    today_idx = now.weekday()
    days_kor = ['월', '화', '수', '목', '금', '토', '일']

    if today_idx >= 5:
        day_name = days_kor[today_idx]
        if not messagebox.askyesno(
            "주말 확인",
            f"📅 오늘은 {day_name}요일(주말)입니다.\n계속 진행하시겠습니까?",
            parent=root_window
        ):
            return

    year_folder = f"{now.year}_Daily Report"
    month_folder = f"{now.month}월_Daily Report"
    year_output_dir = os.path.join(BASE_OUTPUT_DIR, year_folder)
    final_output_dir = os.path.join(year_output_dir, month_folder)
    if not os.path.exists(final_output_dir): os.makedirs(final_output_dir, exist_ok=True)

    auto_num = get_auto_doc_number(year_output_dir)

    doc_input = simpledialog.askstring(
        "문서 번호",
        f"👉 문서 번호를 입력하세요 (미입력 시 '{auto_num}'):",
        parent=root_window,
    )
    if doc_input is None:
        update_status("작업 취소됨")
        return

    user_num = parse_doc_number(doc_input, auto_num)
    if user_num is None:
        messagebox.showerror(
            "입력 오류",
            "문서 번호에 사용할 수 없는 문자가 포함되어 있습니다.\n예: < > : \" / \\ | ? *",
            parent=root_window,
        )
        return

    if 0 <= today_idx <= 4:
        edit_all = False
        if today_idx == 0 or is_new_week:
            msg = "📅 오늘은 월요일입니다." if today_idx == 0 else f"📅 이번 주 데이터가 없습니다. (오늘: {days_kor[today_idx]}요일)"
            if messagebox.askyesno("근무 계획", f"{msg}\n👉 이번 주 근무 계획을 일괄 입력하시겠습니까?", parent=root_window):
                edit_all = True

        # 공휴일 조회: 이번 주 데이터에 holiday_indices가 없으면 항상 조회
        if not holiday_indices:
            monday_date = (now - timedelta(days=today_idx)).date()
            api_key = os.getenv("HOLIDAY_API_KEY", "").strip()
            if not api_key:
                update_status("공휴일 API 키 미설정 — 조회 생략")
            else:
                update_status("공휴일 조회 중...")
                holidays_map = get_holidays_this_week()
                holiday_indices = []
                for i in range(5):
                    d = monday_date + timedelta(days=i)
                    if d in holidays_map:
                        if edit_all:
                            locations[i] = holidays_map[d]
                        holiday_indices.append(i)
                weekly_data["holiday_indices"] = holiday_indices
                update_status("준비됨")

        range_to_edit = range(5) if edit_all else range(today_idx, today_idx + 1)

        for i in range_to_edit:
            day_name = days_kor[i]
            suggestion = locations[i]
            prompt_msg = f"{day_name}요일 근무지 입력 (기본값: {suggestion})" if edit_all else f"오늘({day_name}) 근무지 입력 (기본값: {suggestion})"

            loc_input = simpledialog.askstring("근무지 입력", prompt_msg + "\n(휴가인 경우 '휴가' 입력)", initialvalue=suggestion, parent=root_window)

            if loc_input:
                locations[i] = loc_input

        weekly_data["locations"] = locations
        save_weekly_data(weekly_data)

    current_location_header = locations[today_idx] if 0 <= today_idx <= 4 else DEFAULT_WORK_LOCATION

    date_str = now.strftime('%y%m%d')
    year_short = now.strftime('%y')
    
    base_name = f"FMS{year_short}_{DEPARTMENT}{EMPLOYEE_ID}_{user_num} @ Daily Report_{date_str}"
    
    output_xlsx = os.path.join(final_output_dir, base_name + ".xlsx")
    output_pdf = os.path.join(final_output_dir, base_name + ".pdf")
    output_png = os.path.join(final_output_dir, base_name + ".png")

    if os.path.exists(output_xlsx):
        if not messagebox.askyesno(
            "파일 중복",
            f"⚠️ 이미 동일한 파일이 존재합니다.\n\n{base_name}.xlsx\n\n덮어쓰시겠습니까?",
            parent=root_window
        ):
            return

    report_data = {
        "report_date": now.strftime("%Y-%m-%d"),
        "department": DEPARTMENT,
        "work_location": current_location_header,
        "author_name": AUTHOR_NAME,
        "headcount": 1,
        "doc_number": user_num,
        "work_content": "", "tomorrow_work": "", "notes": ""
    }

    start_button.config(state=tk.DISABLED)
    settings_button.config(state=tk.DISABLED)
    outlook_button.config(state=tk.DISABLED)
    update_status("1/3 작업 준비 중...")

    t = threading.Thread(
        target=thread_step1_create_excel,
        args=(root_window, final_output_dir, report_data, locations, (output_xlsx, output_pdf, output_png), TEMPLATE_PATH, holiday_indices),
        daemon=True
    )
    t.start()


# ==========================================================
# 7. 메인 GUI 구성
# ==========================================================
def main_gui():
    global status_label, start_button, settings_button, outlook_button

    logger.info("Daily Report Automation 프로그램 시작")

    root = tk.Tk()
    root.withdraw()

    def on_closing():
        logger.info("Daily Report Automation 프로그램 종료")
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    ensure_initial_setup(root)
    migrate_env_if_needed()

    root.title("데일리 리포트 자동화")
    root.geometry("450x340")
    root.resizable(False, False)
    root.deiconify()

    title_font = ("맑은 고딕", 16, "bold")
    btn_font = ("맑은 고딕", 10)
    status_font = ("맑은 고딕", 9)

    tk.Label(root, text="Daily Report Automation", font=title_font, fg="#333333").pack(pady=(25, 5))
    tk.Label(root, text="작업을 선택하세요.", font=("맑은 고딕", 10), fg="#666666").pack(pady=(0, 20))

    btn_frame = tk.Frame(root)
    btn_frame.pack(fill="both", padx=30, expand=True)

    # 1. 기본 설정 (왼쪽)
    settings_button = tk.Button(btn_frame, text="⚙️ 기본 설정", command=lambda: open_basic_settings_window(root), font=btn_font, height=2)
    settings_button.grid(row=0, column=0, sticky="ew", padx=2, pady=5)

    # 2. 아웃룩 설정 (오른쪽)
    outlook_button = tk.Button(btn_frame, text="📧 아웃룩 설정", command=lambda: open_outlook_settings_window(root), font=btn_font, height=2, bg="#e8f0fe")
    outlook_button.grid(row=0, column=1, sticky="ew", padx=2, pady=5)

    # 3. 시작 버튼 (아래 꽉 채우기)
    start_button = tk.Button(btn_frame, text="▶ 리포트 생성 시작", command=lambda: start_automation_wrapper(root), font=("맑은 고딕", 11, "bold"), bg="#0078D7", fg="white", height=2)
    start_button.grid(row=1, column=0, columnspan=2, sticky="ew", padx=2, pady=6)

    # 4. 로그 폴더 열기 버튼
    log_button = tk.Button(btn_frame, text="📋 로그 폴더 열기", command=open_log_folder, font=btn_font, height=1)
    log_button.grid(row=2, column=0, columnspan=2, sticky="ew", padx=2, pady=4)

    btn_frame.columnconfigure(0, weight=1)
    btn_frame.columnconfigure(1, weight=1)

    status_label = tk.Label(root, text="상태: 준비됨", font=status_font, fg="gray", anchor="w", relief=tk.SUNKEN, bd=1)
    status_label.pack(side="bottom", fill="x", padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main_gui()
