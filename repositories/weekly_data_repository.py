import json
import os
import re
from datetime import date, datetime, timedelta

from dotenv import load_dotenv

from config import BACKUP_DIR, DATA_FILE_PATH, ENV_PATH
from utils.logger_utils import logger


def get_monday_str() -> str:
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    return monday.strftime("%Y-%m-%d")


def normalize_default_location(value: str) -> str:
    text = str(value or "").strip()
    return text if text else "본사"


def save_json_atomically(data: dict, destination: str) -> None:
    target_dir = os.path.dirname(destination)
    if target_dir:
        os.makedirs(target_dir, exist_ok=True)

    tmp_path = f"{destination}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as file:
        json.dump(
            data,
            file,
            ensure_ascii=False,
            indent=4,
        )

    os.replace(tmp_path, destination)


def backup_corrupt_data_file(filepath: str) -> None:
    if not os.path.isfile(filepath):
        return

    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"weekly_data_corrupt_{timestamp}.json"
        backup_path = os.path.join(BACKUP_DIR, backup_filename)
        os.replace(filepath, backup_path)
        logger.warning(
            "손상된 데이터 파일을 백업하고 기본 데이터로 복구함: backup_path=%s",
            backup_path,
        )
    except Exception:
        logger.exception("손상 데이터 파일 백업 실패")


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
            logger.warning("주간 데이터 파일 손상/읽기 오류 발생")
            backup_corrupt_data_file(DATA_FILE_PATH)

    return default_data, True


def save_weekly_data(data: dict):
    try:
        save_json_atomically(data, DATA_FILE_PATH)
        logger.info("주간 근무지 데이터 원자적 저장 성공")
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
            logger.warning("주간 근무지 로드 중 데이터 손상/읽기 오류 발생")
            backup_corrupt_data_file(DATA_FILE_PATH)

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
            logger.warning("공휴일 인덱스 로드 중 데이터 손상/읽기 오류 발생")
            backup_corrupt_data_file(DATA_FILE_PATH)

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
