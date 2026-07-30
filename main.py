from repositories.settings_repository import migrate_env_if_needed
from ui.main_window import main_gui


def main():
    # 실행 시 누락된 .env 키 마이그레이션 및 사전 백업
    migrate_env_if_needed()
    # 메인 GUI 실행
    main_gui()

if __name__ == "__main__":
    main()