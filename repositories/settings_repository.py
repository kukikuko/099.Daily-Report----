import os
import json
import shutil
from datetime import datetime
from config import BASE_DIR, ENV_PATH, DEFAULT_SUBJECT, DEFAULT_BODY
from utils.logger_utils import logger
from utils.file_utils import save_env_dict_atomically

def load_env_dict(env_path: str = ENV_PATH) -> dict:
    env_dict = {}
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                if '=' in line:
                    k, v = line.split('=', 1)
                    env_dict[k.strip()] = v.strip()
    return env_dict

def migrate_env_if_needed():
    """이전 버전 .env에 누락된 필드가 있으면 백업 후 기본값으로 추가한다."""
    if not os.path.exists(ENV_PATH):
        return

    env_dict = load_env_dict(ENV_PATH)

    defaults = {
        "DEPARTMENT": "TE",
        "EMPLOYEE_ID": "000",
        "DEFAULT_WORK_LOCATION": "본사",
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
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = os.path.join(BASE_DIR, f"env_backup_{timestamp}.env")
            shutil.copy2(ENV_PATH, backup_path)
            logger.info(f".env 사전 백업 생성 완료: {os.path.basename(backup_path)}")
        except Exception:
            logger.exception(".env 사전 백업 생성 중 예외 발생")

        save_env_dict_atomically(env_dict, ENV_PATH)
        logger.info(f".env 마이그레이션 완료 (추가 항목: {added})")
