import csv
import os

from services.outlook_service import dedupe_email_list, normalize_recipient_addresses
from utils.path_utils import resolve_addressbook_csv_path


def load_addressbook_contacts() -> list:
    csv_path = resolve_addressbook_csv_path()
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"주소록 파일을 찾을 수 없습니다: {csv_path}")

    contacts = []
    seen_emails = set()

    encodings = ["utf-8-sig", "cp949", "utf-8", "euc-kr"]
    last_err = None

    for enc in encodings:
        try:
            with open(csv_path, "r", encoding=enc, newline="") as f:
                reader = csv.DictReader(f)
                if not reader.fieldnames or "이메일" not in reader.fieldnames:
                    last_err = ValueError("주소록 형식 오류: '이메일' 헤더가 필요합니다.")
                    continue

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
                break
        except (UnicodeDecodeError, csv.Error) as e:
            last_err = e
            continue

    if not contacts:
        if last_err:
            raise ValueError(f"주소록 읽기 실패: {last_err}")
        raise ValueError("주소록에 사용할 수 있는 이메일이 없습니다.")
    return contacts

def merge_selected_recipients(existing_text: str, selected_emails: list, addressbook_email_keys: set) -> str:
    existing_emails = normalize_recipient_addresses(existing_text)
    external_emails = [email for email in existing_emails if email.lower() not in addressbook_email_keys]
    final_emails = dedupe_email_list((selected_emails or []) + external_emails)
    return ";".join(final_emails)
