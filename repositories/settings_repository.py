import json
import os
import shutil
from datetime import datetime

from config import BACKUP_DIR, DEFAULT_BODY, DEFAULT_SUBJECT, ENV_PATH
from utils.file_utils import save_env_dict_atomically
from utils.logger_utils import logger


def load_env_dict(env_path: str = ENV_PATH) -> dict:
    env_dict = {}

    if not os.path.exists(env_path):
        logger.info("환경 설정 파일 없음")
        return env_dict

    try:
        with open(env_path, "r", encoding="utf-8") as file:
            for raw_line in file:
                line = raw_line.strip()

                if not line or line.startswith("#") or "=" not in line:
                    continue

                key, value = line.split("=", 1)
                env_dict[key.strip()] = value.strip()

        logger.info(
            "환경 설정 로드 완료: key_count=%d",
            len(env_dict),
        )
        return env_dict

    except Exception:
        logger.exception("환경 설정 로드 실패")
        raise


def migrate_env_if_needed():
    """이전 버전 .env에 누락된 필드가 있으면 백업 후 기본값으로 추가한다."""
    if not os.path.exists(ENV_PATH):
        return

    env_dict = load_env_dict(ENV_PATH)

    defaults = {
        "AUTHOR_NAME": "",
        "DEPARTMENT": "TE",
        "EMPLOYEE_ID": "000",
        "DEFAULT_WORK_LOCATION": "본사",
        "BASE_OUTPUT_DIR": "",
        "HOLIDAY_API_KEY": "",
        "OUTLOOK_ENABLE": "False",
        "OUTLOOK_TO": "",
        "OUTLOOK_CC": "",
        "OUTLOOK_SENDER": "",
        "OUTLOOK_SUBJECT": DEFAULT_SUBJECT,
        "OUTLOOK_BODY": json.dumps(
            DEFAULT_BODY,
            ensure_ascii=False,
        ),
    }

    added = []
    for key, default_val in defaults.items():
        if key not in env_dict:
            env_dict[key] = default_val
            added.append(key)

    if added:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            os.makedirs(BACKUP_DIR, exist_ok=True)
            backup_path = os.path.join(BACKUP_DIR, f"env_backup_{timestamp}.env")
            shutil.copy2(ENV_PATH, backup_path)
            logger.info(
                ".env 사전 백업 생성 완료: %s",
                os.path.basename(backup_path),
            )
        except Exception:
            logger.exception(".env 사전 백업 생성 중 예외 발생")

        save_env_dict_atomically(env_dict, ENV_PATH)
        logger.info(
            ".env 마이그레이션 완료 (추가 항목 수: %d)",
            len(added),
        )
