import os
import json
import tkinter as tk
from tkinter import messagebox, scrolledtext
from dotenv import load_dotenv

from config import (
    ENV_PATH, DEFAULT_SUBJECT, DEFAULT_BODY, OUTLOOK_FRIENDLY_TOKENS
)
from utils.logger_utils import logger
from utils.file_utils import save_env_dict_atomically
from services.outlook_service import (
    normalize_recipient_addresses, find_unknown_outlook_tokens
)
from repositories.settings_repository import load_env_dict
from repositories.addressbook_repository import (
    load_addressbook_contacts, merge_selected_recipients
)

def insert_token_into_entry(entry_widget, token: str):
    try:
        entry_widget.insert(entry_widget.index(tk.INSERT), token)
    except Exception:
        entry_widget.insert(tk.END, token)
    entry_widget.focus_set()

def insert_token_into_text(text_widget, token: str):
    text_widget.insert(tk.INSERT, token)
    text_widget.focus_set()

def open_addressbook_window(parent_window, target_entry_widget):
    try:
        contacts = load_addressbook_contacts()
    except Exception:
        logger.exception("주소록 로드 중 예외 발생")
        messagebox.showerror("주소록 오류", "⚠️ 주소록 파일(주소록.csv)을 읽을 수 없습니다.\n파일 존재 여부 및 형식/인코딩을 확인해 주세요.", parent=parent_window)
        return

    addressbook_email_keys = {c["email_lower"] for c in contacts}

    win = tk.Toplevel(parent_window)
    win.title("주소록 선택")
    win.geometry("520x460")
    win.resizable(False, False)

    tk.Label(win, text="[ 주소록 검색 및 추가 ]", font=("맑은 고딕", 10, "bold")).pack(anchor="w", padx=15, pady=(12, 4))

    frame_search = tk.Frame(win)
    frame_search.pack(fill="x", padx=15, pady=2)

    tk.Label(frame_search, text="검색어:").pack(side="left")
    search_entry = tk.Entry(frame_search, width=32)
    search_entry.pack(side="left", padx=5)
    search_entry.focus_set()

    list_frame = tk.Frame(win)
    list_frame.pack(fill="both", expand=True, padx=15, pady=6)

    scrollbar = tk.Scrollbar(list_frame, orient="vertical")
    listbox = tk.Listbox(
        list_frame,
        selectmode=tk.MULTIPLE,
        yscrollcommand=scrollbar.set,
        font=("맑은 고딕", 9),
        exportselection=False
    )
    scrollbar.config(command=listbox.yview)

    scrollbar.pack(side="right", fill="y")
    listbox.pack(side="left", fill="both", expand=True)

    filtered_contacts = list(contacts)

    def populate_list(items):
        listbox.delete(0, tk.END)
        for idx, item in enumerate(items):
            dept_str = f" [{item['department']}]" if item['department'] else ""
            title_str = f" {item['title']}" if item['title'] else ""
            display_text = f"{item['name']}{title_str}{dept_str} - {item['email']}"
            listbox.insert(tk.END, display_text)

            if item["email_lower"] in current_selected_keys:
                listbox.selection_set(idx)

    existing_text = target_entry_widget.get()
    initial_existing = normalize_recipient_addresses(existing_text)
    current_selected_keys = {e.lower() for e in initial_existing if e.lower() in addressbook_email_keys}

    populate_list(filtered_contacts)

    def on_search_change(*args):
        nonlocal filtered_contacts
        query = search_entry.get().strip().lower()

        for idx in listbox.curselection():
            if idx < len(filtered_contacts):
                current_selected_keys.add(filtered_contacts[idx]["email_lower"])

        if not query:
            filtered_contacts = list(contacts)
        else:
            filtered_contacts = [c for c in contacts if query in c["search_text"]]

        populate_list(filtered_contacts)

    search_entry.bind("<KeyRelease>", on_search_change)

    btn_frame = tk.Frame(win)
    btn_frame.pack(fill="x", padx=15, pady=(4, 12))

    def apply_selection():
        for idx in listbox.curselection():
            if idx < len(filtered_contacts):
                current_selected_keys.add(filtered_contacts[idx]["email_lower"])

        selected_emails = []
        for c in contacts:
            if c["email_lower"] in current_selected_keys:
                selected_emails.append(c["email"])

        final_text = merge_selected_recipients(
            target_entry_widget.get(),
            selected_emails,
            addressbook_email_keys
        )

        target_entry_widget.delete(0, tk.END)
        target_entry_widget.insert(0, final_text)
        win.destroy()

    tk.Button(btn_frame, text="적용", font=("맑은 고딕", 9, "bold"), bg="#28a745", fg="white", width=12, command=apply_selection).pack(side="right", padx=3)
    tk.Button(btn_frame, text="취소", width=10, command=win.destroy).pack(side="right", padx=3)

def save_outlook_settings_gui(widgets: dict, window: tk.Toplevel):
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
        env_dict = load_env_dict(ENV_PATH)

        env_dict["OUTLOOK_ENABLE"] = enable
        env_dict["OUTLOOK_TO"] = to_addr
        env_dict["OUTLOOK_CC"] = cc_addr
        env_dict["OUTLOOK_SENDER"] = sender_addr
        env_dict["OUTLOOK_SUBJECT"] = subject
        env_dict["OUTLOOK_BODY"] = json.dumps(body, ensure_ascii=False)

        save_env_dict_atomically(env_dict, ENV_PATH)
        load_dotenv(ENV_PATH, override=True)

        logger.info("아웃룩 설정 저장 성공")
        messagebox.showinfo("성공", "✅ 아웃룩 설정이 저장되었습니다!", parent=window)
        window.destroy()

    except Exception:
        logger.exception("아웃룩 설정 저장 중 예외 발생")
        messagebox.showerror("오류", "❌ 아웃룩 설정 저장 중 오류가 발생했습니다.\n자세한 내용은 로그를 확인해주세요.", parent=window)

def open_outlook_settings_window(root: tk.Tk):
    load_dotenv(ENV_PATH, override=True)

    win = tk.Toplevel(root)
    win.title("아웃룩 자동화 설정")
    win.geometry("520x620")
    win.resizable(False, False)

    widgets = {}

    frame_top = tk.Frame(win)
    frame_top.pack(fill="x", padx=20, pady=(15, 5))

    enable_var = tk.BooleanVar(value=(os.getenv("OUTLOOK_ENABLE", "False") == "True"))
    widgets['enable_var'] = enable_var

    chk_enable = tk.Checkbutton(
        frame_top,
        text="아웃룩 메일 초안 자동 생성 사용",
        variable=enable_var,
        font=("맑은 고딕", 10, "bold"),
        fg="#0056b3"
    )
    chk_enable.pack(anchor="w")

    tk.Label(win, text="[ 발신 / 수신 설정 ]", font=("맑은 고딕", 10, "bold")).pack(anchor="w", padx=20, pady=(10, 5))

    frame_addrs = tk.Frame(win)
    frame_addrs.pack(fill="x", padx=20, pady=5)

    tk.Label(frame_addrs, text="수신자 (To):", width=15, anchor="w").grid(row=0, column=0, pady=3)
    ent_to = tk.Entry(frame_addrs, width=32)
    ent_to.insert(0, os.getenv("OUTLOOK_TO", ""))
    ent_to.grid(row=0, column=1, pady=3)
    widgets['to'] = ent_to
    tk.Button(frame_addrs, text="주소록", command=lambda: open_addressbook_window(win, ent_to)).grid(row=0, column=2, padx=5, pady=3)

    tk.Label(frame_addrs, text="참조 (CC):", width=15, anchor="w").grid(row=1, column=0, pady=3)
    ent_cc = tk.Entry(frame_addrs, width=32)
    ent_cc.insert(0, os.getenv("OUTLOOK_CC", ""))
    ent_cc.grid(row=1, column=1, pady=3)
    widgets['cc'] = ent_cc
    tk.Button(frame_addrs, text="주소록", command=lambda: open_addressbook_window(win, ent_cc)).grid(row=1, column=2, padx=5, pady=3)

    tk.Label(frame_addrs, text="보내는 계정 (선택):", width=15, anchor="w").grid(row=2, column=0, pady=3)
    ent_sender = tk.Entry(frame_addrs, width=32)
    ent_sender.insert(0, os.getenv("OUTLOOK_SENDER", ""))
    ent_sender.grid(row=2, column=1, pady=3)
    widgets['sender'] = ent_sender

    tk.Label(win, text="[ 메인 제목 및 본문 템플릿 ]", font=("맑은 고딕", 10, "bold")).pack(anchor="w", padx=20, pady=(10, 5))

    frame_tmpl = tk.Frame(win)
    frame_tmpl.pack(fill="x", padx=20, pady=5)

    tk.Label(frame_tmpl, text="제목:", width=6, anchor="w").grid(row=0, column=0, pady=3)
    ent_subject = tk.Entry(frame_tmpl, width=48)
    ent_subject.insert(0, os.getenv("OUTLOOK_SUBJECT", DEFAULT_SUBJECT))
    ent_subject.grid(row=0, column=1, pady=3, sticky="w")
    widgets['subject'] = ent_subject

    frame_tokens = tk.Frame(win)
    frame_tokens.pack(fill="x", padx=20, pady=(2, 5))
    tk.Label(frame_tokens, text="사용 가능 변수: ").pack(side="left")

    for token in OUTLOOK_FRIENDLY_TOKENS:
        btn = tk.Button(
            frame_tokens,
            text=token,
            font=("맑은 고딕", 8),
            relief="groove",
            command=lambda t=token: insert_token_into_text(widgets['body'], t)
        )
        btn.pack(side="left", padx=2)

    frame_body = tk.Frame(win)
    frame_body.pack(fill="both", expand=True, padx=20, pady=5)

    tk.Label(frame_body, text="본문 텍스트:").pack(anchor="w")
    txt_body = scrolledtext.ScrolledText(frame_body, height=8, font=("맑은 고딕", 9))

    raw_body_env = os.getenv("OUTLOOK_BODY", "")
    if raw_body_env:
        try:
            body_text = json.loads(raw_body_env)
        except (json.JSONDecodeError, ValueError):
            body_text = raw_body_env
    else:
        body_text = DEFAULT_BODY

    txt_body.insert("1.0", body_text)
    txt_body.pack(fill="both", expand=True, pady=2)
    widgets['body'] = txt_body

    btn_frame = tk.Frame(win)
    btn_frame.pack(fill="x", padx=20, pady=15)

    tk.Button(
        btn_frame,
        text="저장",
        font=("맑은 고딕", 10, "bold"),
        bg="#007bff",
        fg="white",
        command=lambda: save_outlook_settings_gui(widgets, win)
    ).pack(fill="x")
