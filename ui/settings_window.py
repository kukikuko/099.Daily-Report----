import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from dotenv import load_dotenv

from config import ENV_PATH, DATA_FILE_PATH, BASE_DIR
from utils.file_utils import save_env_dict_atomically
from repositories.weekly_data_repository import (
    get_monday_str, normalize_default_location, normalize_weekly_data, load_json_locations
)
from repositories.settings_repository import load_env_dict

def save_basic_settings_gui(widgets: dict, window: tk.Toplevel):
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
        env_dict = load_env_dict(ENV_PATH)

        env_dict["AUTHOR_NAME"] = new_name
        env_dict["DEPARTMENT"] = new_dept
        env_dict["EMPLOYEE_ID"] = f"{int(new_id):03d}" if new_id.isdigit() else new_id
        env_dict["DEFAULT_WORK_LOCATION"] = new_def_loc
        env_dict["BASE_OUTPUT_DIR"] = new_path

        save_env_dict_atomically(env_dict, ENV_PATH)
        load_dotenv(ENV_PATH, override=True)

        current_monday = get_monday_str()
        existing_json = {}
        if os.path.exists(DATA_FILE_PATH):
            try:
                with open(DATA_FILE_PATH, "r", encoding="utf-8") as f:
                    existing_json = json.load(f)
            except (json.JSONDecodeError, IOError):
                pass

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

def open_settings_window(root: tk.Tk):
    load_dotenv(ENV_PATH, override=True)

    win = tk.Toplevel(root)
    win.title("기본 설정")
    win.geometry("450x580")
    win.resizable(False, False)

    widgets = {}

    tk.Label(win, text="[ 사용자 기본 정보 ]", font=("맑은 고딕", 10, "bold")).pack(anchor="w", padx=20, pady=(15, 5))

    frame_info = tk.Frame(win)
    frame_info.pack(fill="x", padx=20, pady=5)

    labels = ["작성자 이름:", "부서명 (예: TE):", "사번 (예: 007):", "기본 근무지:"]
    keys = ['name', 'dept', 'id', 'def_loc']
    defaults = [
        os.getenv("AUTHOR_NAME", ""),
        os.getenv("DEPARTMENT", "TE"),
        os.getenv("EMPLOYEE_ID", "000"),
        normalize_default_location(os.getenv("DEFAULT_WORK_LOCATION", "본사"))
    ]

    for i, (lbl, key, default) in enumerate(zip(labels, keys, defaults)):
        tk.Label(frame_info, text=lbl, width=15, anchor="w").grid(row=i, column=0, pady=3)
        ent = tk.Entry(frame_info, width=28)
        ent.insert(0, default)
        ent.grid(row=i, column=1, pady=3)
        widgets[key] = ent

    tk.Label(win, text="[ 출력 결과 저장 경로 ]", font=("맑은 고딕", 10, "bold")).pack(anchor="w", padx=20, pady=(15, 5))

    frame_path = tk.Frame(win)
    frame_path.pack(fill="x", padx=20, pady=5)

    default_path = os.getenv("BASE_OUTPUT_DIR", os.path.join(BASE_DIR, "Output"))
    ent_path = tk.Entry(frame_path, width=32)
    ent_path.insert(0, default_path)
    ent_path.pack(side="left", padx=(0, 5))
    widgets['path'] = ent_path

    def browse_folder():
        folder = filedialog.askdirectory(parent=win, initialdir=ent_path.get())
        if folder:
            ent_path.delete(0, tk.END)
            ent_path.insert(0, folder)

    tk.Button(frame_path, text="찾아보기", command=browse_folder).pack(side="left")

    tk.Label(win, text="[ 이번 주 근무지 계획 (월~금) ]", font=("맑은 고딕", 10, "bold")).pack(anchor="w", padx=20, pady=(15, 5))

    frame_locs = tk.Frame(win)
    frame_locs.pack(fill="x", padx=20, pady=5)

    days = ["월요일:", "화요일:", "수요일:", "목요일:", "금요일:"]
    current_locations = load_json_locations()

    widgets['loc_entries'] = []
    for i, day in enumerate(days):
        tk.Label(frame_locs, text=day, width=15, anchor="w").grid(row=i, column=0, pady=2)
        ent = tk.Entry(frame_locs, width=28)
        ent.insert(0, current_locations[i])
        ent.grid(row=i, column=1, pady=2)
        widgets['loc_entries'].append(ent)

    btn_frame = tk.Frame(win)
    btn_frame.pack(fill="x", padx=20, pady=20)

    tk.Button(
        btn_frame, text="저장", font=("맑은 고딕", 10, "bold"),
        bg="#007bff", fg="white", command=lambda: save_basic_settings_gui(widgets, win)
    ).pack(fill="x")
