import os
import json
import re
from datetime import date, timedelta
from config import DATA_FILE_PATH, ENV_PATH
from utils.logger_utils import logger
from dotenv import load_dotenv

def get_monday_str() -> str:
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    return monday.strftime("%Y-%m-%d")

def normalize_default_location(value: str) -> str:
    text = str(value or "").strip()
    return text if text else "본사"

def normalize_weekly_data(raw_data: dict, default_loc: str) -> dict:
    """weekly_data JSON 구조를 강제 정규화하여 반환한다.

    보장하는 불변 조건:
      - locations  : 정확히 5개의 비어있지 않은 str 목록
      - holiday_indices : 0~4 범위 int로만 구성된 중복 없는 정렬 목록
    """
    result = dict(raw_data)

    raw_locs = result.get("locations", [])
    if not isinstance(raw_locs, list):
        raw_locs = []
    locs = []
    for i in range(5):
        val = raw_locs[i] if i < len(raw_locs) else ""
        locs.append(str(val).strip() if isinstance(val, str) and str(val).strip() else default_loc)
    result["locations"] = locs

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

def load_weekly_data(default_loc: str):
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
        except (json.JSONDecodeError, IOError):
            pass
    return default_data, True

def save_weekly_data(data: dict):
    try:
        with open(DATA_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        logger.info("주간 근무지 데이터 저장 성공")
    except Exception:
        logger.exception("주간 근무지 데이터 저장 실패")

def load_json_locations() -> list:
    current_monday = get_monday_str()
    load_dotenv(ENV_PATH, override=True)
    default_loc = normalize_default_location(os.getenv("DEFAULT_WORK_LOCATION", "본사"))

    if os.path.exists(DATA_FILE_PATH):
        try:
            with open(DATA_FILE_PATH, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if raw.get("week_start") == current_monday:
                return normalize_weekly_data(raw, default_loc)["locations"]
        except (json.JSONDecodeError, IOError):
            pass
    return [default_loc] * 5

def load_holiday_indices() -> list:
    current_monday = get_monday_str()
    load_dotenv(ENV_PATH, override=True)
    default_loc = normalize_default_location(os.getenv("DEFAULT_WORK_LOCATION", "본사"))

    if os.path.exists(DATA_FILE_PATH):
        try:
            with open(DATA_FILE_PATH, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if raw.get("week_start") == current_monday:
                return normalize_weekly_data(raw, default_loc)["holiday_indices"]
        except (json.JSONDecodeError, IOError):
            pass
    return []

def get_auto_doc_number(year_dir: str) -> str:
    if not os.path.exists(year_dir):
        return "001"
    max_num = 0
    pattern = re.compile(r"_(\d{3})\s*@\s*Daily Report")
    for _, _, files in os.walk(year_dir):
        for f in files:
            if not f.lower().endswith(".xlsx"):
                continue
            match = pattern.search(f)
            if match:
                try:
                    num = int(match.group(1))
                    if num > max_num:
                        max_num = num
                except ValueError:
                    continue
    return f"{max_num + 1:03d}"
