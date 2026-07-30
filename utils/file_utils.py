import os
import uuid

import pywintypes
import win32con
import win32file

from config import ENV_PATH
from utils.logger_utils import logger


def ensure_dir_writable(dir_path: str) -> bool:
    """지정한 디렉터리에 쓰기 권한이 존재하는지 사전 검사한다."""
    if not dir_path:
        return False
    norm_path = os.path.abspath(os.path.normpath(dir_path))
    if not os.path.exists(norm_path):
        try:
            os.makedirs(norm_path, exist_ok=True)
        except OSError:
            logger.exception(f"출력 디렉터리 생성 실패: {norm_path}")
            return False

    test_file = os.path.join(norm_path, f".write_test_{uuid.uuid4().hex}.tmp")
    try:
        with open(test_file, "w", encoding="utf-8") as f:
            f.write("write_permission_check")
        if os.path.exists(test_file):
            os.remove(test_file)
        return True
    except (OSError, IOError):
        logger.exception(f"디렉터리 쓰기 권한 없음: {norm_path}")
        return False


def is_file_locked(filepath: str) -> bool:
    """파일이 다른 프로세스(사용자의 Excel 창 등)에 의해 사용 중(잠김)인지 확인한다.

    OneDrive 및 한글 경로 환경에서도 win32file 독점 핸들 시도로 점유 상태를 정확하게 감지한다.
    """
    if not filepath or not os.path.isfile(filepath):
        return False

    norm_path = os.path.abspath(os.path.normpath(filepath))

    try:
        handle = win32file.CreateFile(
            norm_path,
            win32con.GENERIC_READ | win32con.GENERIC_WRITE,
            0,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None,
        )
        handle.Close()
        return False
    except pywintypes.error:
        return True
    except Exception:
        try:
            os.rename(norm_path, norm_path)
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
