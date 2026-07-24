import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from config import LOG_DIR, LOG_FILE_PATH

if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR, exist_ok=True)

logger = logging.getLogger("DailyReport")
logger.setLevel(logging.INFO)

if not logger.handlers:
    file_handler = RotatingFileHandler(
        LOG_FILE_PATH,
        maxBytes=2 * 1024 * 1024,
        backupCount=5,
        encoding='utf-8'
    )
    formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

def open_log_folder():
    """로그 폴더를 파일 탐색기로 엽니다."""
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR, exist_ok=True)
    try:
        os.startfile(LOG_DIR)
        logger.info("사용자가 UI에서 로그 폴더를 열었습니다.")
    except AttributeError:
        import subprocess
        if sys.platform == "darwin":
            subprocess.run(["open", LOG_DIR])
        else:
            subprocess.run(["xdg-open", LOG_DIR])
        logger.info("사용자가 UI에서 로그 폴더를 열었습니다.")
    except Exception:
        logger.exception("로그 폴더 열기 실패")
