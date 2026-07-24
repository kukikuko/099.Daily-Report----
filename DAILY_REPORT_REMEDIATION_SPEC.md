# Daily Report Automation 보완 개발 지시서

> 대상 저장소: `kukikuko/099.Daily-Report----`  
> 기준 브랜치: `main`  
> 문서 목적: 현재 구현된 Phase 0~5 결과를 점검하고, 남은 보안·회귀·안정성 문제를 AI 코딩 에이전트가 순서대로 수정하도록 지시한다.  
> 대상 환경: Windows 10/11, Python 3.14, Microsoft Excel Desktop, Classic Outlook  
> 중요 원칙: 기존 기능을 보존하고, 단계별 테스트가 끝난 뒤 다음 작업으로 이동한다.

---

# 1. AI 코딩 에이전트 공통 지침

다음 규칙을 반드시 준수한다.

1. 이 문서의 작업 순서를 지킨다.
2. 한 번에 하나의 작업 그룹만 수정한다.
3. 수정 전 관련 파일을 읽고 현재 동작을 설명한다.
4. 기존 기능을 삭제하거나 변경하기 전에 영향 범위를 기록한다.
5. 수정 후 반드시 문법 검사와 관련 테스트를 실행한다.
6. 테스트하지 않은 기능을 완료로 표시하지 않는다.
7. Outlook 메일을 자동 전송하지 않는다.
8. `mail.Send()`를 프로젝트 어디에도 추가하지 않는다.
9. 사용자가 실행 중인 다른 Excel 인스턴스를 종료하지 않는다.
10. Excel 자동화는 `win32.DispatchEx("Excel.Application")`를 유지한다.
11. UI 스레드에서 Excel, PDF, Outlook 같은 장시간 작업을 수행하지 않는다.
12. Tkinter 위젯은 메인 스레드에서만 수정한다.
13. 기술적인 예외 정보는 로그에 기록하고 사용자에게는 이해하기 쉬운 메시지를 표시한다.
14. `.env`, API Key, 메일 주소, 메일 본문 전체를 로그에 출력하지 않는다.
15. 기존 `.env`와 `weekly_data.json`의 하위 호환성을 유지한다.
16. 현재 모듈 구조를 다시 단일 파일로 합치지 않는다.
17. 불필요한 대규모 리팩터링을 추가로 진행하지 않는다.
18. 작업 종료 후 변경 파일, 테스트 결과, 남은 위험을 보고한다.

---

# 2. 현재 프로젝트 구조

현재 저장소는 다음과 같이 모듈화되어 있다.

```text
DailyReport/
├─ main.py
├─ config.py
├─ legacy_daily_report.py
├─ requirements-current.txt
├─ models/
│  └─ report_data.py
├─ repositories/
│  ├─ addressbook_repository.py
│  ├─ settings_repository.py
│  └─ weekly_data_repository.py
├─ services/
│  ├─ excel_service.py
│  ├─ holiday_service.py
│  └─ outlook_service.py
├─ ui/
│  ├─ main_window.py
│  ├─ outlook_settings_window.py
│  └─ settings_window.py
├─ utils/
│  ├─ file_utils.py
│  ├─ logger_utils.py
│  └─ path_utils.py
├─ Template/
│  └─ Daily_Report_Template.xlsx
├─ 주소록.csv
└─ .gitignore
```

현재 `main.py`는 실행 진입점 역할만 수행한다.

```python
from repositories.settings_repository import migrate_env_if_needed
from ui.main_window import main_gui


def main():
    migrate_env_if_needed()
    main_gui()


if __name__ == "__main__":
    main()
```

---

# 3. 현재 완료된 기능

다음 기능은 이미 구현되어 있으므로 삭제하거나 이전 방식으로 되돌리지 않는다.

- RotatingFileHandler 기반 로그
- 로그 파일 2MB 제한
- 로그 파일 최대 5개 백업
- UTF-8 로그
- 로그 폴더 열기 버튼
- 프로그램 시작 및 종료 로그
- Excel `DispatchEx` 독립 인스턴스
- Excel COM 객체 `finally` 정리
- 파일명 금지 문자 치환
- Excel/PDF/PNG 중복 파일 확인
- Excel 편집 완료 후 파일 잠금 확인
- `.env` 원자적 저장
- `.env` 마이그레이션 전 백업
- 주간 근무지 데이터 정규화
- 공휴일 API 호출
- Outlook 수신자 중복 제거
- Outlook 자동 전송 금지
- 기존 단일 파일의 기능별 모듈 분리
- `DailyReportData` 데이터 클래스 기본 구현

---

# 4. 수정 우선순위

## P0: 보안 및 기능 회귀

1. `.env` 백업 파일 Git 제외
2. 최초 실행 설정 마법사 복구
3. UI 예외 처리와 로그 통일
4. 설정 로드 로그 추가
5. 최소 실행 환경 파일 추가

## P1: Excel 및 파일 안정화

1. 파일 잠금 검사 개선
2. 출력 경로 쓰기 권한 검사
3. PDF 인쇄 영역 지정
4. PNG 범위 확대
5. PDF와 PNG 결과 분리

## P2: Outlook 안정화

1. HTML 이스케이프
2. PNG 조건부 본문 삽입
3. PDF 필수 정책 적용
4. Outlook 기본 서명 유지
5. 발신 계정 이메일 로그 제거
6. 수신자 형식 기본 검증

## P3: 테스트와 문서

1. 단위 테스트 추가
2. 통합 테스트 체크리스트
3. 최소 `requirements.txt`
4. `.env.example`
5. `README.md`
6. 현재 Phase 상태 문서화

## P4: 다음 기능

P0~P3가 완료된 뒤에만 Daily Report 직접 입력 화면인 Phase 6을 시작한다.

---

# 5. 작업 1: 환경 파일 보안 강화

## 목적

마이그레이션 과정에서 생성되는 `.env` 백업 파일이 공개 Git 저장소에 커밋되는 것을 방지한다.

## 관련 파일

- `.gitignore`
- `repositories/settings_repository.py`
- `config.py`

## 현재 문제

현재 백업 파일은 프로젝트 루트에 다음 이름으로 생성된다.

```text
env_backup_YYYYMMDD_HHMMSS.env
```

`.gitignore`에는 `.env`만 포함되어 있고 백업 파일 패턴은 없다.

## 구현 요구사항

### 5.1 `.gitignore` 수정

다음 항목을 추가한다.

```gitignore
# Environment files and backups
.env
.env.*
!.env.example
env_backup_*.env
*.env.bak
*.tmp

# Runtime backups
Data/backups/
```

기존 `.gitignore` 항목은 유지한다.

### 5.2 백업 경로 변경

프로젝트 루트가 아니라 다음 경로를 사용한다.

```text
Data/backups/env_backup_YYYYMMDD_HHMMSS.env
```

`config.py`에 다음 상수를 추가한다.

```python
DATA_DIR = os.path.join(BASE_DIR, "Data")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
```

`settings_repository.py`에서 백업 전에 폴더를 생성한다.

```python
os.makedirs(BACKUP_DIR, exist_ok=True)
```

### 5.3 백업 로그

백업 파일의 전체 경로나 환경값을 로그에 출력하지 않는다.

허용:

```python
logger.info(".env 사전 백업 생성 완료")
```

금지:

```python
logger.info("환경값: %s", env_dict)
logger.info("API KEY: %s", api_key)
```

## 검증

```powershell
git check-ignore -v .env
git check-ignore -v env_backup_20260724_120000.env
git check-ignore -v Data/backups/test.env
```

## 완료 조건

- `.env`가 Git에서 제외된다.
- `.env` 백업 파일도 Git에서 제외된다.
- 백업 파일은 `Data/backups`에 생성된다.
- 환경값이 로그에 출력되지 않는다.

---

# 6. 작업 2: 최초 실행 설정 마법사 복구

## 목적

신규 PC 또는 `.env`가 없는 환경에서 기본값으로 바로 실행되는 회귀 문제를 해결한다.

## 관련 파일

- `main.py`
- `ui/main_window.py`
- 신규 권장 파일: `ui/initial_setup_window.py`
- `repositories/settings_repository.py`

## 현재 문제

`main.py`는 `.env`가 없어도 `migrate_env_if_needed()` 이후 바로 GUI를 실행한다.

현재 구현에서는 신규 사용자가 작성자, 부서, 사번, 저장 경로를 입력하지 않고 프로그램을 실행할 수 있다.

## 구현 요구사항

### 6.1 초기 설정 함수 분리

신규 파일을 생성한다.

```text
ui/initial_setup_window.py
```

다음 함수를 제공한다.

```python
def ensure_initial_setup(parent: tk.Tk) -> bool:
    """
    .env가 없으면 초기 설정 마법사를 실행한다.

    반환:
        True: 설정 완료
        False: 사용자가 취소함
    """
```

### 6.2 초기 설정 필드

- 작성자 이름
- 부서 코드
- 사원번호
- 기본 근무지
- 결과 저장 폴더
- Outlook 초안 기능 사용 여부

### 6.3 설정 저장

기존 `save_env_dict_atomically()`을 재사용한다.

저장 키:

```dotenv
AUTHOR_NAME=
DEPARTMENT=TE
EMPLOYEE_ID=000
DEFAULT_WORK_LOCATION=본사
BASE_OUTPUT_DIR=

OUTLOOK_ENABLE=False
OUTLOOK_TO=
OUTLOOK_CC=
OUTLOOK_SENDER=
OUTLOOK_SUBJECT=Daily Report_[[년]]년 [[월]]월 [[일]]일
OUTLOOK_BODY=

HOLIDAY_API_KEY=
```

### 6.4 실행 순서

권장 구조:

```python
def main():
    root = tk.Tk()
    root.withdraw()

    if not ensure_initial_setup(root):
        root.destroy()
        return

    migrate_env_if_needed()
    load_dotenv(ENV_PATH, override=True)

    main_gui(root)
```

또는 `main_gui()`가 root를 생성하는 기존 구조를 유지하려면 초기 설정과 root 생성을 한 곳에서 일관되게 관리한다.

중복으로 `tk.Tk()`를 생성하지 않는다.

### 6.5 취소 처리

사용자가 초기 설정을 취소하면 프로그램을 정상 종료한다.

기본값으로 강제 진행하지 않는다.

## 검증

1. 기존 `.env`를 임시 이동한다.
2. 프로그램을 실행한다.
3. 초기 설정 마법사가 표시되는지 확인한다.
4. 설정 완료 후 `.env`가 생성되는지 확인한다.
5. 프로그램 재실행 시 마법사가 다시 나오지 않는지 확인한다.
6. 설정 취소 시 프로그램이 종료되는지 확인한다.

## 완료 조건

- `.env`가 없으면 초기 설정 화면이 표시된다.
- 필수 입력을 완료해야 메인 화면으로 이동한다.
- 기존 사용자의 `.env`가 있으면 기존처럼 바로 실행된다.
- 초기 설정 취소 시 오류 없이 종료된다.

---

# 7. 작업 3: UI 오류 처리 및 로깅 통일

## 목적

기술 예외 메시지를 사용자에게 그대로 노출하지 않고, 모든 예외를 로그에서 추적할 수 있게 한다.

## 관련 파일

- `ui/settings_window.py`
- `ui/outlook_settings_window.py`
- `ui/main_window.py`
- `repositories/addressbook_repository.py`
- `utils/logger_utils.py`

## 현재 문제

일부 UI 코드가 다음처럼 예외 원문을 그대로 표시한다.

```python
except Exception as e:
    messagebox.showerror("오류", f"저장 중 오류 발생: {e}")
```

주소록 오류도 `str(e)`를 그대로 표시한다.

## 구현 요구사항

### 7.1 예외 처리 표준

다음 패턴을 사용한다.

```python
except Exception:
    logger.exception("기본 설정 저장 실패")
    messagebox.showerror(
        "설정 저장 실패",
        "설정을 저장하지 못했습니다.\n로그 파일을 확인해 주세요.",
        parent=window,
    )
```

### 7.2 사용자 메시지 표준

| 상황 | 사용자 메시지 |
|---|---|
| 기본 설정 저장 실패 | 설정을 저장하지 못했습니다. 로그 파일을 확인해 주세요. |
| Outlook 설정 저장 실패 | Outlook 설정을 저장하지 못했습니다. 로그 파일을 확인해 주세요. |
| 주소록 로드 실패 | 주소록을 불러오지 못했습니다. CSV 파일을 확인해 주세요. |
| Excel 생성 실패 | Excel 보고서를 생성하지 못했습니다. 로그 파일을 확인해 주세요. |
| PDF 변환 실패 | PDF를 생성하지 못했습니다. Excel 파일 상태를 확인해 주세요. |
| Outlook 초안 실패 | Outlook 초안을 만들지 못했습니다. Outlook 로그인 상태를 확인해 주세요. |

### 7.3 사용자 취소는 오류로 기록하지 않음

사용자 취소는 `logger.info()` 또는 `logger.warning()`으로 기록한다.

### 7.4 `pass` 제거

다음 종류의 `except ...: pass`를 검토한다.

- JSON 읽기
- Excel 파일 열기
- PDF 열기
- 주소록 읽기
- 데이터 저장

복구 가능한 오류는 최소한 `logger.warning()`을 남긴다.

## 검증

의도적으로 다음 오류를 발생시킨다.

- 읽기 전용 `.env`
- 잘못된 주소록 CSV
- 존재하지 않는 저장 경로
- 손상된 `weekly_data.json`

로그에 Traceback이 남고 사용자에게는 단순 메시지만 표시되는지 확인한다.

## 완료 조건

- 사용자 화면에 Python 또는 COM 예외 원문이 노출되지 않는다.
- 모든 예상하지 못한 예외는 `logger.exception()`으로 기록된다.
- 사용자 취소와 시스템 오류가 구분된다.

---

# 8. 작업 4: 설정 로드 로그 추가

## 목적

환경 설정 파일을 정상적으로 읽었는지 확인하되 민감정보는 기록하지 않는다.

## 관련 파일

- `repositories/settings_repository.py`

## 구현 요구사항

`load_env_dict()`가 파일 존재 여부, 로드 성공 여부, 키 개수만 기록한다.

```python
def load_env_dict(env_path: str = ENV_PATH) -> dict:
    env_dict = {}

    if not os.path.exists(env_path):
        logger.info("환경 설정 파일 없음")
        return env_dict

    try:
        with open(env_path, "r", encoding="utf-8") as file:
            for line in file:
                if "=" not in line:
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
```

## 금지

- `env_dict` 전체 출력
- Outlook 수신자 출력
- 발신 계정 이메일 출력
- `HOLIDAY_API_KEY` 출력
- 메일 본문 출력

## 완료 조건

- 정상 로드 시 키 개수 로그가 남는다.
- 파일이 없을 때 오류가 아닌 정보 로그가 남는다.
- 설정값 자체는 로그에 남지 않는다.

---

# 9. 작업 5: 파일 잠금 검사 개선

## 목적

Excel 파일이 열려 있는 상태에서 PDF 변환을 시작하지 않도록 한다.

## 관련 파일

- `utils/file_utils.py`
- `ui/main_window.py`

## 현재 문제

현재 구현은 다음 방식을 사용한다.

```python
os.rename(filepath, filepath)
```

이 방식은 Office 잠금 상태를 안정적으로 판별하지 못할 수 있다.

## 구현 요구사항

### 9.1 잠금 검사 함수

```python
def is_file_locked(filepath: str) -> bool:
    if not os.path.exists(filepath):
        return False

    try:
        with open(filepath, "a+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return True
```

### 9.2 실제 Excel 열기 예외 병행

파일 잠금 검사 결과가 `False`여도 `excel.Workbooks.Open()` 실패를 별도로 처리한다.

사용자 메시지:

```text
Excel 보고서 파일을 열 수 없습니다.
파일이 열려 있거나 동기화 중인지 확인해 주세요.
```

### 9.3 재시도 제한

현재 확인 창 반복 방식은 유지할 수 있다.

단, 무한 자동 재시도는 구현하지 않는다.

## 검증

- Excel에서 대상 파일을 연 상태
- Excel에서 파일을 저장 중인 상태
- OneDrive 동기화 중인 파일
- 파일을 닫은 상태

## 완료 조건

- 열린 Excel 파일을 대부분 감지한다.
- 감지 실패 시에도 Workbook 열기 예외가 사용자 메시지와 로그로 처리된다.
- 사용자가 파일을 닫으면 다시 진행할 수 있다.

---

# 10. 작업 6: 출력 경로 쓰기 권한 검사

## 목적

Excel COM 실행 전에 저장 폴더에 파일을 생성할 수 있는지 확인한다.

## 관련 파일

- `utils/file_utils.py`
- `ui/main_window.py`

## 신규 함수

```python
def ensure_writable_directory(directory: str) -> tuple[bool, str]:
    """
    디렉터리를 생성하고 임시 파일 쓰기 테스트를 수행한다.

    반환:
        (True, ""): 쓰기 가능
        (False, message): 쓰기 불가
    """
```

권장 구현:

1. `os.makedirs(directory, exist_ok=True)`
2. 임시 파일 생성
3. 임시 파일 삭제
4. 실패 시 로그 기록

## 사용자 메시지

```text
선택한 폴더에 파일을 저장할 수 없습니다.
저장 경로와 폴더 권한을 확인해 주세요.
```

## 완료 조건

- Excel 실행 전에 쓰기 권한을 확인한다.
- 권한이 없으면 COM 자동화를 시작하지 않는다.
- 테스트용 임시 파일은 항상 정리된다.

---

# 11. 작업 7: PDF 및 PNG 출력 범위 수정

## 목적

금일 업무, 익일 업무, 특이사항이 PDF와 PNG에 모두 포함되도록 한다.

## 관련 파일

- `config.py`
- `services/excel_service.py`

## 현재 문제

셀 매핑:

```text
금일 업무: D12
익일 업무: B18
특이사항: B23
```

현재 PNG 범위:

```python
ws.Range("A2:I16")
```

따라서 익일 업무와 특이사항이 PNG에 포함되지 않는다.

## 구현 요구사항

### 11.1 상수 추가

`config.py`:

```python
REPORT_PRINT_AREA = "$A$2:$I$24"
REPORT_IMAGE_RANGE = "A2:I24"
```

실제 템플릿을 확인하여 마지막 행이 24가 아니라면 정확한 범위로 조정한다.

### 11.2 PDF 페이지 설정

```python
ws.PageSetup.PrintArea = REPORT_PRINT_AREA
ws.PageSetup.Zoom = False
ws.PageSetup.FitToPagesWide = 1
ws.PageSetup.FitToPagesTall = 1
```

### 11.3 PDF 내보내기

```python
ws.ExportAsFixedFormat(
    Type=0,
    Filename=pdf_path,
    Quality=0,
    IncludeDocProperties=True,
    IgnorePrintAreas=False,
    OpenAfterPublish=False,
)
```

### 11.4 PNG 범위

```python
rng = ws.Range(REPORT_IMAGE_RANGE)
```

### 11.5 Chart 정리

임시 Chart는 별도 변수로 보관하고 `finally`에서 삭제를 시도한다.

```python
chart = None
```

### 11.6 파일 생성 확인

PDF와 PNG 내보내기 후 다음을 확인한다.

```python
os.path.exists(path)
os.path.getsize(path) > 0
```

## 검증

- 금일 업무 입력
- 익일 업무 입력
- 특이사항 입력
- 긴 여러 줄 텍스트 입력
- PDF 1페이지 출력 확인
- PNG 전체 내용 확인

## 완료 조건

- PDF에 모든 보고서 항목이 포함된다.
- PNG에 익일 업무와 특이사항이 포함된다.
- 출력 파일 크기가 0이면 실패로 처리한다.

---

# 12. 작업 8: PDF 및 PNG 결과 분리

## 목적

PDF는 생성됐지만 PNG가 실패한 경우 전체 작업을 실패로 처리하지 않는다.

## 관련 파일

- `models/report_data.py` 또는 신규 `models/export_result.py`
- `services/excel_service.py`
- `ui/main_window.py`
- `services/outlook_service.py`

## 신규 결과 모델

```python
from dataclasses import dataclass


@dataclass(slots=True)
class ExportResult:
    pdf_created: bool
    png_created: bool
    pdf_path: str
    png_path: str
    error_message: str = ""

    @property
    def is_success(self) -> bool:
        return self.pdf_created
```

PDF는 메일 첨부의 필수 결과이고 PNG는 본문 이미지용 선택 결과로 취급한다.

## 동작 정책

| PDF | PNG | 결과 |
|---|---|---|
| 성공 | 성공 | 정상 완료 |
| 성공 | 실패 | 메일 초안은 생성하되 본문 이미지 제외 |
| 실패 | 성공 | 전체 실패 |
| 실패 | 실패 | 전체 실패 |

## 완료 조건

- PNG 생성 실패가 PDF 생성 성공을 무효화하지 않는다.
- UI에 부분 성공 상태가 표시된다.
- Outlook은 PNG가 없으면 이미지 없이 생성된다.

---

# 13. 작업 9: Outlook HTML 및 첨부 안정화

## 목적

HTML 특수문자, PNG 누락, PDF 누락으로 인해 메일 본문이 깨지는 문제를 해결한다.

## 관련 파일

- `services/outlook_service.py`

## 구현 요구사항

### 13.1 HTML 이스케이프

```python
from html import escape

html_content = escape(final_body_text).replace("\n", "<br>")
```

### 13.2 PDF 필수 정책

PDF가 없거나 크기가 0이면 초안 생성을 중단한다.

```python
if not os.path.exists(output_pdf) or os.path.getsize(output_pdf) == 0:
    logger.error("Outlook 초안 생성 중단: PDF 파일 없음")
    return "failed"
```

### 13.3 PNG 조건부 삽입

```python
image_html = ""

if os.path.exists(output_png) and os.path.getsize(output_png) > 0:
    attachment = mail.Attachments.Add(output_png)
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "daily_report_img",
    )
    image_html = "<img src='cid:daily_report_img' width='650'>"
```

PNG가 없으면 `<img>` 태그를 생성하지 않는다.

### 13.4 수신자 없는 경우

수신자가 없어도 초안 생성을 허용할 수 있다.

다만 다음 로그만 남긴다.

```python
logger.warning("Outlook 수신자가 설정되지 않았습니다.")
```

주소를 로그에 출력하지 않는다.

## 완료 조건

- 본문에 `<`, `>`, `&`가 있어도 메일 HTML이 깨지지 않는다.
- PDF가 없으면 초안 생성을 중단한다.
- PNG가 없으면 깨진 이미지가 표시되지 않는다.
- 수신자 주소가 로그에 노출되지 않는다.

---

# 14. 작업 10: Outlook 기본 서명 유지

## 목적

사용자의 기존 Outlook 기본 서명을 보고서 본문 뒤에 유지한다.

## 관련 파일

- `services/outlook_service.py`

## 구현 방식

권장 순서:

1. MailItem 생성
2. 수신자, 제목, 첨부 설정
3. `mail.Display()` 호출
4. Outlook이 삽입한 기존 `mail.HTMLBody` 읽기
5. 보고서 HTML을 서명 앞에 삽입
6. 메일 창은 그대로 표시

예시:

```python
mail.Display()

signature_html = mail.HTMLBody or ""

mail.HTMLBody = f"""
<html>
<body style="font-family:'Malgun Gothic', sans-serif; font-size:11pt;">
    <p>{html_content}</p>
    <br>
    {image_html}
    <br><br>
    {signature_html}
</body>
</html>
"""
```

## 주의사항

- `mail.Display()`를 두 번 호출하지 않는다.
- Outlook 환경에 따라 서명이 비어 있어도 오류로 처리하지 않는다.
- 고정된 장시간 `sleep()`은 피한다.
- 필요하면 짧은 제한 시간 안에서 `HTMLBody` 확인을 반복한다.
- 절대 `mail.Send()`를 호출하지 않는다.

## 검증

- 기본 서명 설정된 Outlook
- 기본 서명 없는 Outlook
- 이미지 포함
- 이미지 없음
- 여러 발신 계정

## 완료 조건

- 기본 서명이 유지된다.
- 서명이 없어도 보고서 본문은 정상 표시된다.
- 메일은 자동 전송되지 않는다.

---

# 15. 작업 11: 발신 계정 로그 개인정보 제거

## 목적

발신 계정의 전체 이메일 주소가 로그에 기록되는 것을 방지한다.

## 관련 파일

- `services/outlook_service.py`

## 현재 금지 예시

```python
logger.info(f"지정 발신 계정 적용 완료: {sender_email}")
logger.warning(f"지정한 발신 계정을 찾지 못함: {sender_email}")
```

## 변경 예시

```python
logger.info("지정 발신 계정 적용 완료")
logger.warning("지정한 발신 계정을 찾지 못해 기본 계정을 사용합니다.")
```

## 완료 조건

- 발신·수신 이메일 주소가 로그에 직접 출력되지 않는다.
- 계정 적용 성공·실패 여부는 확인할 수 있다.

---

# 16. 작업 12: 최소 의존성 파일 생성

## 목적

현재 PC 전체 패키지 목록과 실제 프로젝트 의존성을 분리한다.

## 관련 파일

- `requirements-current.txt`
- 신규 `requirements.txt`
- 선택 `requirements-dev.txt`

## 정책

`requirements-current.txt`는 환경 백업용으로 유지한다.

신규 `requirements.txt`:

```text
python-dotenv==1.2.1
pywin32==311
```

PyInstaller를 운영 의존성에 포함할지 개발 의존성으로 분리할지 결정한다.

권장 `requirements-dev.txt`:

```text
-r requirements.txt
pyinstaller==6.17.0
pytest
ruff
```

Python 3.14에서 각 패키지가 실제 설치 가능한지 확인한다.

## 검증

새 가상환경에서:

```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements-dev.txt
python -m compileall .
```

## 완료 조건

- 불필요한 AI, 웹 서버, 브라우저 자동화 패키지가 운영 의존성에서 제거된다.
- 깨끗한 가상환경에서 프로그램 모듈을 import할 수 있다.

---

# 17. 작업 13: `.env.example` 생성

## 목적

민감정보 없이 필요한 설정 키를 문서화한다.

## 신규 파일

```text
.env.example
```

## 내용

```dotenv
AUTHOR_NAME=
DEPARTMENT=TE
EMPLOYEE_ID=000
DEFAULT_WORK_LOCATION=본사
BASE_OUTPUT_DIR=

OUTLOOK_ENABLE=False
OUTLOOK_TO=
OUTLOOK_CC=
OUTLOOK_SENDER=
OUTLOOK_SUBJECT=Daily Report_[[년]]년 [[월]]월 [[일]]일
OUTLOOK_BODY=

HOLIDAY_API_KEY=
```

## 요구사항

- 실제 이메일 주소 금지
- 실제 API Key 금지
- 실제 사용자 이름 금지
- 실제 출력 경로 금지

## 완료 조건

- 신규 개발자가 필요한 키를 파악할 수 있다.
- `.env.example`은 Git에 포함된다.
- 실제 `.env`는 Git에 포함되지 않는다.

---

# 18. 작업 14: README 작성

## 목적

프로그램의 설치, 실행, 설정, 제한사항을 저장소에서 확인할 수 있도록 한다.

## 신규 파일

```text
README.md
```

## 필수 내용

1. 프로그램 개요
2. 주요 기능
3. 시스템 요구사항
4. Python 설치 및 가상환경
5. 의존성 설치
6. `.env.example`을 `.env`로 복사하는 방법
7. Excel 템플릿 위치
8. 주소록 CSV 형식
9. 프로그램 실행
10. 보고서 생성 흐름
11. Outlook은 Classic Outlook만 지원한다는 안내
12. 메일은 자동 전송되지 않는다는 안내
13. 로그 위치
14. PyInstaller 빌드 방법
15. 알려진 제한사항
16. 개발 Phase 상태

## 주소록 CSV 예시

```csv
이름,이메일,직책,부서,회사
홍길동,user@example.com,대리,개발팀,회사명
```

## 완료 조건

- 새 개발자가 README만 보고 프로그램을 실행할 수 있다.
- 민감정보가 포함되지 않는다.

---

# 19. 작업 15: 단위 테스트 추가

## 목적

모듈화 이후 발생할 수 있는 회귀를 자동으로 감지한다.

## 신규 구조

```text
tests/
├─ test_path_utils.py
├─ test_outlook_template.py
├─ test_weekly_data_repository.py
├─ test_settings_repository.py
└─ test_report_data.py
```

## 테스트 대상

### `utils/path_utils.py`

- 금지 문자 치환
- 앞뒤 공백 제거
- 마지막 점 제거
- 빈 문자열 fallback
- 숫자 문서 번호 처리

### `services/outlook_service.py`

- 수신자 분리
- 중복 제거
- 변수 치환
- 알 수 없는 변수 탐지
- 미치환 변수 경고 대상 확인

COM 객체 생성 함수 자체는 단위 테스트에서 실행하지 않는다.

### `repositories/weekly_data_repository.py`

- locations 5개 보정
- holiday_indices 정규화
- 범위 밖 인덱스 제거
- 중복 제거
- 새 주 기본값

### `repositories/settings_repository.py`

- 환경 설정 파싱
- 누락 키 추가
- 기존 알 수 없는 키 유지
- 원자적 저장

### `models/report_data.py`

- 필수 필드 검증
- headcount 검증
- dict 변환
- 문자열 정리

## 테스트 명령

```powershell
python -m pytest -q
python -m compileall .
ruff check .
```

## 완료 조건

- 모든 테스트가 통과한다.
- Excel 및 Outlook COM 없이 단위 테스트를 실행할 수 있다.
- 임시 파일 테스트는 pytest의 `tmp_path`를 사용한다.

---

# 20. 작업 16: 수동 통합 테스트

## 목적

Excel과 Outlook COM 동작은 자동 테스트만으로 확인할 수 없으므로 실제 환경에서 검증한다.

## Excel 테스트

- [ ] 다른 Excel 문서가 열린 상태에서 보고서 생성
- [ ] 다른 Excel 문서가 종료되지 않음
- [ ] 보고서 Excel을 연 상태에서 변환 차단
- [ ] 보고서 저장 후 닫으면 변환 진행
- [ ] Excel 템플릿 누락
- [ ] 출력 폴더 권한 없음
- [ ] OneDrive 경로
- [ ] 한글 경로
- [ ] PDF 전체 범위
- [ ] PNG 전체 범위
- [ ] 익일 업무와 특이사항 포함

## Outlook 테스트

- [ ] Outlook 실행 중
- [ ] Outlook 미실행
- [ ] 발신 계정 1개
- [ ] 발신 계정 여러 개
- [ ] 지정 발신 계정 일치
- [ ] 지정 발신 계정 불일치
- [ ] 기본 서명 있음
- [ ] 기본 서명 없음
- [ ] PNG 있음
- [ ] PNG 없음
- [ ] PDF 없음
- [ ] 수신자 없음
- [ ] Outlook 기능 비활성
- [ ] 메일 자동 전송되지 않음

## 데이터 테스트

- [ ] `.env` 없음
- [ ] `.env` 이전 버전
- [ ] `.env` 백업 생성
- [ ] `weekly_data.json` 손상
- [ ] 공휴일 API Key 없음
- [ ] 공휴일 API 실패
- [ ] 주소록 CSV 손상
- [ ] 주말 실행

## 완료 조건

수동 테스트 결과를 다음 파일에 기록한다.

```text
docs/MANUAL_TEST_RESULTS.md
```

각 항목에 다음 상태를 사용한다.

```text
PASS
FAIL
BLOCKED
NOT TESTED
```

---

# 21. 작업 17: Phase 상태 문서 정정

## 목적

현재 저장소의 실제 구현 수준과 문서 표현을 일치시킨다.

## 신규 또는 수정 파일

```text
docs/DEVELOPMENT_STATUS.md
```

## 권장 상태

| Phase | 상태 | 설명 |
|---|---|---|
| Phase 0 | 부분 완료 | 백업과 저장소 구성 완료, 기준선 수동 테스트 보완 필요 |
| Phase 1 | 진행 중 | 핵심 로깅 완료, 일부 UI 예외 처리 보완 필요 |
| Phase 2 | 진행 중 | DispatchEx 적용, 권한 및 통합 테스트 필요 |
| Phase 3 | 진행 중 | 파일명·중복 구현, 잠금 검사 개선 필요 |
| Phase 4 | 진행 중 | PDF 및 PNG 범위 수정 필요 |
| Phase 5 | 진행 중 | Outlook HTML·서명 처리 필요 |
| Phase 6 | 미시작 | 앞 단계 안정화 후 진행 |

`완료`는 테스트 결과까지 확인된 경우에만 사용한다.

---

# 22. 작업 완료 후에만 Phase 6 시작

다음 조건이 모두 충족된 후 Daily Report 직접 입력 화면 개발을 시작한다.

- [ ] 환경 백업 파일 Git 제외
- [ ] 최초 실행 설정 복구
- [ ] UI 예외 처리 통일
- [ ] 파일 잠금 검사 개선
- [ ] PDF 전체 범위 출력
- [ ] PNG 전체 범위 출력
- [ ] Outlook HTML 이스케이프
- [ ] PNG 조건부 삽입
- [ ] Outlook 기본 서명 유지
- [ ] 최소 requirements 생성
- [ ] `.env.example` 생성
- [ ] README 생성
- [ ] 단위 테스트 통과
- [ ] 수동 Excel/Outlook 테스트 기록

---

# 23. 전체 검증 명령

AI 에이전트는 코드 변경 후 다음 명령을 실행한다.

```powershell
python -m compileall .
python -m pytest -q
ruff check .
```

Git 상태도 확인한다.

```powershell
git status --short
git diff --check
```

민감 파일 포함 여부를 검사한다.

```powershell
git ls-files | findstr /i ".env"
git ls-files | findstr /i "env_backup"
```

`.env.example`만 나타나는 것이 정상이다.

---

# 24. AI 작업 보고 형식

각 작업 그룹이 끝날 때 다음 형식으로 보고한다.

```markdown
## 작업 결과

### 작업 범위
- 작업 1: 환경 파일 보안 강화

### 변경 파일
- `.gitignore`
- `config.py`
- `repositories/settings_repository.py`

### 구현 내용
- `.env` 백업 경로를 `Data/backups`로 변경
- 환경 백업 파일 Git 제외 패턴 추가
- 민감정보 로그 출력 방지

### 테스트
- `python -m compileall .`: PASS
- `python -m pytest -q`: PASS
- `ruff check .`: PASS
- `git check-ignore`: PASS

### 수동 확인
- `.env` 백업 생성: PASS
- 백업 파일 Git 제외: PASS

### 남은 문제
- 없음

### 다음 작업
- 작업 2: 최초 실행 설정 마법사 복구
```

---

# 25. 첫 번째 AI 작업 지시

AI 코딩 에이전트는 다음 작업부터 시작한다.

```text
대상 저장소의 main 브랜치를 기준으로 작업한다.

1. 현재 git status와 프로젝트 파일 구조를 확인한다.
2. `.gitignore`, config.py, repositories/settings_repository.py를 읽는다.
3. 실제 `.env`나 백업 파일이 Git 추적 중인지 확인한다.
4. `.gitignore`에 환경 백업 제외 규칙을 추가한다.
5. `.env` 백업 경로를 Data/backups로 변경한다.
6. 환경값이나 이메일 주소가 로그에 출력되지 않도록 확인한다.
7. .env.example을 생성한다.
8. python -m compileall .을 실행한다.
9. 가능한 테스트를 실행한다.
10. 변경 파일, 테스트 결과, 민감정보 포함 여부를 보고한다.

이 작업이 완료되기 전에는 최초 실행 설정 화면이나 Outlook 코드를 수정하지 않는다.
```

---

# 26. 두 번째 AI 작업 지시

첫 번째 작업 완료 후 다음을 수행한다.

```text
1. 기존 초기 설정 기능이 legacy_daily_report.py에 어떻게 구현되어 있는지 확인한다.
2. 현재 모듈 구조에 맞게 ui/initial_setup_window.py로 이전한다.
3. .env가 없을 때만 초기 설정 마법사를 표시한다.
4. 사용자가 취소하면 프로그램을 정상 종료한다.
5. 설정 완료 후 migrate_env_if_needed를 실행한다.
6. Tk 루트 인스턴스가 중복 생성되지 않도록 한다.
7. 기존 .env가 있는 사용자는 기존 실행 흐름을 유지한다.
8. compileall과 관련 테스트를 실행한다.
9. 신규 PC 시나리오를 수동 테스트한다.
10. 변경 및 검증 결과를 보고한다.
```

---

# 27. 세 번째 AI 작업 지시

두 번째 작업 완료 후 다음을 수행한다.

```text
1. ui/settings_window.py와 ui/outlook_settings_window.py의 예외 처리를 검토한다.
2. 사용자에게 예외 원문을 노출하는 코드를 제거한다.
3. 모든 예상하지 못한 오류를 logger.exception으로 기록한다.
4. 주소록 로드 오류도 동일한 정책으로 변경한다.
5. 설정 로드 성공 로그를 추가하되 값은 출력하지 않는다.
6. 발신·수신 이메일 주소가 로그에 출력되지 않도록 수정한다.
7. 의도적 오류 테스트를 실행한다.
8. 로그에 Traceback이 기록되는지 확인한다.
9. 사용자 메시지가 이해하기 쉬운지 확인한다.
10. 결과를 보고한다.
```

---

# 28. 네 번째 AI 작업 지시

세 번째 작업 완료 후 다음을 수행한다.

```text
1. utils/file_utils.py의 파일 잠금 방식을 개선한다.
2. 출력 폴더 쓰기 권한 검사 함수를 추가한다.
3. ui/main_window.py에서 Excel 생성 전에 쓰기 가능 여부를 확인한다.
4. 열린 Excel 파일을 감지하고 사용자에게 닫도록 안내한다.
5. Workbooks.Open 실패도 별도 로그와 메시지로 처리한다.
6. OneDrive와 한글 경로를 고려한다.
7. compileall과 단위 테스트를 실행한다.
8. 실제 Excel 열린 상태를 수동 테스트한다.
9. 결과를 보고한다.
```

---

# 29. 다섯 번째 AI 작업 지시

네 번째 작업 완료 후 다음을 수행한다.

```text
1. config.py에 REPORT_PRINT_AREA와 REPORT_IMAGE_RANGE를 추가한다.
2. services/excel_service.py의 PDF 인쇄 범위를 명시한다.
3. PNG 범위를 보고서 전체로 확대한다.
4. PDF와 PNG 생성 결과를 ExportResult로 분리한다.
5. PDF 성공/PNG 실패 상황을 부분 성공으로 처리한다.
6. 임시 Chart 정리를 보장한다.
7. 출력 파일 존재와 크기를 검증한다.
8. 전체 보고서 내용이 포함되는지 수동 확인한다.
9. 테스트 결과를 보고한다.
```

---

# 30. 여섯 번째 AI 작업 지시

다섯 번째 작업 완료 후 다음을 수행한다.

```text
1. services/outlook_service.py에 HTML 이스케이프를 적용한다.
2. PDF가 없으면 초안 생성을 중단한다.
3. PNG가 있을 때만 본문 이미지와 Content-ID를 추가한다.
4. Outlook 기본 서명을 유지한다.
5. 이메일 주소를 로그에 출력하지 않는다.
6. mail.Send 호출이 없는지 전체 저장소를 검색한다.
7. PNG 없음, 서명 없음, 수신자 없음 상황을 테스트한다.
8. 메일이 자동 전송되지 않는지 확인한다.
9. 결과를 보고한다.
```

---

# 31. 최종 완료 기준

다음 조건을 모두 만족해야 안정화 작업 완료로 판단한다.

- [ ] `.env`와 백업 파일이 Git에 포함되지 않는다.
- [ ] `.env.example`이 제공된다.
- [ ] 신규 PC에서 초기 설정 마법사가 동작한다.
- [ ] 모든 UI 예외가 로그에 기록된다.
- [ ] 사용자에게 Python 예외 원문이 표시되지 않는다.
- [ ] 설정값과 이메일 주소가 로그에 노출되지 않는다.
- [ ] 열린 Excel 파일이 감지된다.
- [ ] 출력 경로 권한이 사전 검사된다.
- [ ] PDF에 전체 보고서가 포함된다.
- [ ] PNG에 전체 보고서가 포함된다.
- [ ] PDF 성공과 PNG 실패가 구분된다.
- [ ] Outlook 본문 HTML이 안전하게 생성된다.
- [ ] PNG가 없을 때 깨진 이미지가 표시되지 않는다.
- [ ] Outlook 기본 서명이 유지된다.
- [ ] 메일이 자동 전송되지 않는다.
- [ ] 최소 `requirements.txt`가 존재한다.
- [ ] README가 존재한다.
- [ ] 단위 테스트가 통과한다.
- [ ] 수동 Excel/Outlook 테스트 결과가 기록된다.
- [ ] 실제 Phase 상태 문서가 갱신된다.

이 조건을 만족한 이후 Phase 6 Daily Report 직접 입력 화면 개발을 시작한다.
