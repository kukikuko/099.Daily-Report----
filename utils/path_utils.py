import os

from config import ADDRESSBOOK_FILENAME, BASE_DIR, BUNDLE_DIR, INVALID_FILENAME_CHARS


def sanitize_filename_part(value: str, fallback: str = "default") -> str:
    """파일명 구성 요소에 사용할 수 없는 문자를 안전하게 제거/치환한다."""
    text = str(value or "").strip()
    cleaned = INVALID_FILENAME_CHARS.sub("_", text)
    cleaned = cleaned.strip(". ")
    return cleaned if cleaned else fallback

def parse_doc_number(user_input: str, default_val: str) -> str:
    """문서 번호 입력값을 파싱하고 검증한다. 유효하지 않은 경우 None 반환."""
    if not user_input or not user_input.strip():
        return default_val
    text = user_input.strip()
    if INVALID_FILENAME_CHARS.search(text):
        return None
    return text

def resolve_addressbook_csv_path() -> str:
    """주소록 CSV 파일 경로를 탐색하여 반환한다."""
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
