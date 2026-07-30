import os
import sys
import threading
import time
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox, simpledialog

import pythoncom
from dotenv import load_dotenv

from config import BASE_DIR, ENV_PATH, TEMPLATE_PATH
from models.report_data import DailyReportData
from repositories.weekly_data_repository import (
    get_auto_doc_number,
    load_weekly_data,
    normalize_default_location,
    save_weekly_data,
)
from services.excel_service import export_final_reports, generate_excel_draft
from services.holiday_service import get_holidays_this_week
from services.outlook_service import create_outlook_draft
from ui.initial_setup_window import ensure_initial_setup
from ui.outlook_settings_window import open_outlook_settings_window
from ui.settings_window import open_settings_window
from utils.file_utils import ensure_dir_writable, is_file_locked
from utils.logger_utils import logger, open_log_folder
from utils.path_utils import parse_doc_number, sanitize_filename_part

# GUI 전역 참조 변수
status_label = None
start_button = None
settings_button = None
outlook_button = None

def update_status(message):
    if status_label:
        status_label.after(0, lambda: status_label.config(text=f"상태: {message}"))

def reset_buttons():
    if start_button:
        start_button.after(0, lambda: start_button.config(state=tk.NORMAL))
        settings_button.after(0, lambda: settings_button.config(state=tk.NORMAL))
        outlook_button.after(0, lambda: outlook_button.config(state=tk.NORMAL))

# --- Step 1. 엑셀 생성 스레드 ---
def thread_step1_create_excel(root_window, final_output_dir, report_data, locations, output_paths, template_path, holiday_indices=None):
    pythoncom.CoInitialize()
    output_xlsx, output_pdf, output_png = output_paths

    try:
        update_status("2/3 엑셀 초안 생성 중...")
        if generate_excel_draft(report_data, output_xlsx, template_path, locations, holiday_indices):
            update_status("엑셀 생성 완료. 작성 대기 중...")
            try:
                os.startfile(output_xlsx)
            except OSError:
                logger.exception("Excel 파일 자동 열기 실패")
            
            root_window.after(0, lambda: ask_user_to_continue(root_window, output_paths, report_data))
        else:
            update_status("❌ 엑셀 생성 오류 발생")
            reset_buttons()
    except Exception:
        logger.exception("Excel 초안 생성 스레드 에러 발생")
        update_status("에러 발생")
        reset_buttons()
    finally:
        pythoncom.CoUninitialize()

# --- Step 2. 사용자 확인 팝업 (메인 스레드) ---
def ask_user_to_continue(root_window, output_paths, report_data):
    output_xlsx, output_pdf, output_png = output_paths
    while True:
        response = messagebox.askokcancel(
            "작성 확인", 
            "📝 Excel 파일이 열렸습니다.\n\n내용을 작성한 후 저장(Ctrl+S)하고,\nExcel 파일을 완전히 닫은 뒤 [확인]을 눌러주세요.",
            parent=root_window
        )
        
        if not response:
            logger.info("사용자가 엑셀 편집 후 작업을 취소함")
            update_status("작업 취소됨")
            reset_buttons()
            return

        if is_file_locked(output_xlsx):
            logger.warning(f"Excel 파일이 열려 있어 변환 보류: {os.path.basename(output_xlsx)}")
            messagebox.showwarning(
                "파일 열림 경고",
                "⚠️ Excel 파일이 아직 열려 있습니다.\n\nExcel을 저장하고 완전히 닫은 뒤 다시 [확인]을 눌러주세요.",
                parent=root_window
            )
            continue
        break

    start_step2_thread(root_window, output_paths, report_data)

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

        result = export_final_reports(output_xlsx, output_pdf, output_png)

        if result.is_full_success or result.is_partial_success:
            if result.is_partial_success:
                logger.warning("PNG 캡처는 실패했으나 PDF 생성이 완료되어 부분 성공으로 진행합니다.")

            update_status("메일 초안 생성 중...")

            outlook_status = create_outlook_draft(report_data, output_paths)
            if outlook_status == "created":
                if result.is_partial_success:
                    update_status("🎉 작업 완료! (PNG 본문 이미지 제외, 메일 창 확인)")
                else:
                    update_status("🎉 모든 작업 완료! (메일 창 확인)")
            elif outlook_status == "disabled":
                update_status("✅ PDF 변환 완료 (아웃룩 자동 작성 비활성)")
                root_window.after(
                    0,
                    lambda: messagebox.showinfo(
                        "아웃룩 비활성",
                        "PDF 생성은 완료되었습니다.\n\n"
                        "아웃룩 자동 작성 기능이 꺼져 있어 메일 초안은 생성하지 않았습니다.\n"
                        "메인 화면의 '📧 아웃룩 설정'에서 사용 체크 후 저장하면 다음부터 자동 생성됩니다.",
                        parent=root_window,
                    ),
                )
                try:
                    os.startfile(output_pdf)
                except OSError:
                    pass
            else:  # "failed"
                update_status("⚠️ 아웃룩 생성 실패 — PDF를 직접 확인하세요")
                root_window.after(
                    0,
                    lambda: messagebox.showwarning(
                        "아웃룩 연동 안내",
                        "PDF 생성은 완료되었지만 아웃룩 메일 초안을 만들지 못했습니다.\n\n"
                        "Outlook 실행/로그인 상태와 기본 프로필 설정을 확인한 뒤 다시 시도해 주세요.\n"
                        "지금은 열리는 PDF를 첨부해 수동으로 발송할 수 있습니다.",
                        parent=root_window,
                    ),
                )
                try:
                    os.startfile(output_pdf)
                except OSError:
                    pass
        else:
            update_status("⚠️ PDF 변환 실패")
            root_window.after(0, lambda: messagebox.showerror("실패", "PDF 변환 중 오류가 발생했습니다.", parent=root_window))

    except Exception:
        logger.exception("PDF/PNG 및 아웃룩 연동 스레드 예외 발생")
        update_status("에러 발생")
    finally:
        pythoncom.CoUninitialize()
        reset_buttons()

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
    if not os.path.exists(final_output_dir):
        os.makedirs(final_output_dir, exist_ok=True)

    if not ensure_dir_writable(final_output_dir):
        messagebox.showerror(
            "권한 오류",
            "⚠️ 출력 폴더에 파일을 생성/작성할 권한이 없습니다.\n\n"
            f"경로: {final_output_dir}\n"
            "폴더 권한 또는 저장 경로 설정을 확인해 주세요.",
            parent=root_window
        )
        update_status("작업 중단됨 (권한 부족)")
        return

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
    
    safe_dept = sanitize_filename_part(DEPARTMENT, "TE")
    safe_emp_id = sanitize_filename_part(EMPLOYEE_ID, "000")
    safe_user_num = sanitize_filename_part(user_num, auto_num)

    base_name = f"FMS{year_short}_{safe_dept}{safe_emp_id}_{safe_user_num} @ Daily Report_{date_str}"
    
    output_xlsx = os.path.abspath(os.path.normpath(os.path.join(final_output_dir, base_name + ".xlsx")))
    output_pdf = os.path.abspath(os.path.normpath(os.path.join(final_output_dir, base_name + ".pdf")))
    output_png = os.path.abspath(os.path.normpath(os.path.join(final_output_dir, base_name + ".png")))

    if is_file_locked(output_xlsx):
        logger.warning(f"생성 대상 엑셀 파일이 이미 열려 있음: {os.path.basename(output_xlsx)}")
        messagebox.showwarning(
            "파일 열림 경고",
            f"⚠️ 생성 대상 엑셀 파일이 이미 열려 있습니다.\n\n"
            f"파일명: {os.path.basename(output_xlsx)}\n\n"
            "Excel을 저장하고 완전히 닫은 뒤 다시 시도해 주세요.",
            parent=root_window
        )
        update_status("작업 취소됨 (파일 열림)")
        return

    existing_files = [f for f in [output_xlsx, output_pdf, output_png] if os.path.exists(f)]
    if existing_files:
        existing_names = "\n".join(f"- {os.path.basename(f)}" for f in existing_files)
        if not messagebox.askyesno(
            "파일 중복 경고",
            f"⚠️ 이미 다음 파일(들)이 존재합니다.\n\n{existing_names}\n\n덮어쓰시겠습니까?",
            parent=root_window
        ):
            logger.info("사용자에 의해 중복 파일 덮어쓰기 취소됨")
            update_status("작업 취소됨")
            return

    report_obj = DailyReportData(
        report_date=now.strftime("%Y-%m-%d"),
        department=DEPARTMENT,
        work_location=current_location_header,
        author_name=AUTHOR_NAME,
        headcount=1,
        doc_number=user_num,
        employee_id=EMPLOYEE_ID,
        work_content="",
        tomorrow_work="",
        notes="",
    )

    val_errors = report_obj.validate()
    if val_errors:
        messagebox.showerror(
            "입력 오류",
            "\n".join(val_errors),
            parent=root_window,
        )
        return

    report_data = report_obj.to_dict()


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


def main_gui():

    global status_label, start_button, settings_button, outlook_button

    logger.info("Daily Report Automation 프로그램 시작")

    root = tk.Tk()
    root.withdraw()

    if not ensure_initial_setup(root):
        logger.info("사용자에 의해 초기 설정 마법사가 취소되어 프로그램을 종료합니다.")
        root.destroy()
        sys.exit(0)

    def on_closing():
        if start_button and str(start_button["state"]) == str(tk.DISABLED):
            if not messagebox.askyesno(
                "작업 중",
                "현재 보고서 생성 작업이 진행 중입니다.\n작업이 끝난 뒤 종료하는 것이 안전합니다.\n\n그래도 종료하시겠습니까?",
                parent=root,
            ):
                return
        logger.info("Daily Report Automation 프로그램 종료")
        root.destroy()


    root.protocol("WM_DELETE_WINDOW", on_closing)

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
    btn_frame.pack(pady=10)

    start_button = tk.Button(
        btn_frame,
        text="🚀 데일리 리포트 작성",
        font=("맑은 고딕", 11, "bold"),
        bg="#007bff",
        fg="white",
        width=26,
        height=2,
        command=lambda: start_automation_wrapper(root)
    )
    start_button.pack(pady=5)

    settings_button = tk.Button(
        btn_frame,
        text="⚙️ 기본 설정",
        font=btn_font,
        width=26,
        command=lambda: open_settings_window(root)
    )
    settings_button.pack(pady=3)

    outlook_button = tk.Button(
        btn_frame,
        text="📧 아웃룩 설정",
        font=btn_font,
        width=26,
        command=lambda: open_outlook_settings_window(root)
    )
    outlook_button.pack(pady=3)

    tk.Button(
        btn_frame,
        text="📁 로그 폴더 열기",
        font=("맑은 고딕", 9),
        fg="#555555",
        width=26,
        command=open_log_folder
    ).pack(pady=3)

    status_label = tk.Label(root, text="상태: 준비됨", font=status_font, fg="blue")
    status_label.pack(side="bottom", pady=15)

    root.mainloop()
