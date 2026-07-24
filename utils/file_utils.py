import os
from config import ENV_PATH
from utils.logger_utils import logger

def is_file_locked(filepath: str) -> bool:
    """파일이 다른 프로세스에 의해 사용 중(잠김)인지 확인한다."""
    if not os.path.exists(filepath):
        return False
    try:
        os.rename(filepath, filepath)
        return False
    except OSError:
        return True

def save_env_dict_atomically(env_dict: dict, env_path: str = ENV_PATH):
    """임시 파일(.env.tmp)을 활용하여 .env 파일을 안전하게 원자적 저장합니다."""
    tmp_path = f"{env_path}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        for k, v in env_dict.items():
            f.write(f"{k}={v}\n")
    os.replace(tmp_path, env_path)
