import os
import sys
import json
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from dotenv import load_dotenv

from config import BASE_DIR, ENV_PATH, DEFAULT_SUBJECT, DEFAULT_BODY
from utils.file_utils import save_env_dict_atomically
from repositories.weekly_data_repository import normalize_default_location
from repositories.settings_repository import migrate_env_if_needed

def ensure_initial_setup(parent: tk.Tk) -> bool:
    """초기 .env 가 없으면 설정 마법사를 실행한다.
    
    반환값:
        bool: True - 정상 준비 완료 (.env 존재함 또는 마법사 완수)
              False - 사용자가 마법사를 취소함
    """
    if os.path.exists(ENV_PATH):
        return True

    messagebox.showinfo("초기 설정", "⚙️ 초기 환경 설정 파일(.env)이 없습니다.\n설정을 시작합니다.", parent=parent)

    # 1. 작성자 이름
    input_name = simpledialog.askstring("설정 (1/4)", "👉 작성자 성함을 입력해주세요:", parent=parent)
    if input_name is None or not input_name.strip():
        return False

    # 2. 부서 코드
    input_dept = simpledialog.askstring("설정 (2/4)", "👉 부서 코드를 입력해주세요 (예: TE):", parent=parent)
    if input_dept is None:
        return False
    input_dept = input_dept.strip() if input_dept.strip() else "TE"

    # 3. 사원번호
    input_id = simpledialog.askstring("설정 (3/4)", "👉 사원번호를 입력해주세요 (예: 007):", parent=parent)
    if input_id is None:
        return False
    input_id_str = input_id.strip()
    if not input_id_str:
        input_id_str = "000"
    elif input_id_str.isdigit():
        input_id_str = f"{int(input_id_str):03d}"

    # 4. 기본 근무지
    input_loc = simpledialog.askstring("설정 (4/4)", "👉 기본 근무지를 입력해주세요 (예: 본사):", parent=parent)
    if input_loc is None:
        return False
    input_loc = normalize_default_location(input_loc)

    # 5. 아웃룩 사용 여부
    use_outlook = messagebox.askyesno(
        "설정",
        "📧 '아웃룩 메일 초안 자동 작성' 기능을 사용하시겠습니까?\n(나중에 설정에서 변경 가능)",
        parent=parent
    )
    outlook_val = "True" if use_outlook else "False"

    # 6. 결과물 저장 폴더 선택
    messagebox.showinfo("설정", "👉 확인을 누르면 결과물을 저장할 '폴더 선택 창'이 열립니다.", parent=parent)
    default_dir = os.path.join(BASE_DIR, "Output")
    selected_dir = filedialog.askdirectory(initialdir=BASE_DIR, title="결과물 저장 폴더 선택", parent=parent)
    input_dir = selected_dir.replace('/', '\\') if selected_dir else default_dir

    try:
        env_dict = {
            "AUTHOR_NAME": input_name.strip(),
            "DEPARTMENT": input_dept,
            "EMPLOYEE_ID": input_id_str,
            "DEFAULT_WORK_LOCATION": input_loc,
            "BASE_OUTPUT_DIR": input_dir,
            "OUTLOOK_ENABLE": outlook_val,
            "OUTLOOK_TO": "",
            "OUTLOOK_CC": "",
            "OUTLOOK_SENDER": "",
            "OUTLOOK_SUBJECT": DEFAULT_SUBJECT,
            "OUTLOOK_BODY": json.dumps(DEFAULT_BODY, ensure_ascii=False),
        }

        save_env_dict_atomically(env_dict, ENV_PATH)
        migrate_env_if_needed()
        load_dotenv(ENV_PATH, override=True)

        messagebox.showinfo("완료", "✅ 초기 환경 설정 저장이 완료되었습니다.", parent=parent)
        return True

    except Exception as e:
        messagebox.showerror("오류", f"❌ 설정 저장 실패: {e}", parent=parent)
        return False
