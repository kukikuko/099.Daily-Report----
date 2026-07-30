# Daily Report Automation 저장소 종합 개선 지시서

> 대상 저장소: `kukikuko/099.Daily-Report----`  
> 기준 브랜치: `main`  
> 기준 커밋: `b78fdcf6f81a385910f7a68556141af8a3185c54`  
> 저장소 공개 상태: Public  
> 목적: 최신 저장소 검토에서 확인된 보안 문제, 실행 오류, 구조 문제, 테스트·배포 누락 사항을 AI 코딩 에이전트가 우선순위에 따라 개선하도록 지시한다.  
> 대상 환경: Windows 10/11, Python 3.14, Microsoft Excel Desktop, Classic Outlook

---

# 1. 작업 원칙

AI 코딩 에이전트는 다음 원칙을 반드시 준수한다.

1. P0 작업을 모두 끝내기 전에는 신규 기능을 개발하지 않는다.
2. 한 번에 하나의 작업 그룹만 수정한다.
3. 수정 전 관련 파일과 호출 흐름을 먼저 확인한다.
4. 현재 모듈 구조를 단일 파일로 되돌리지 않는다.
5. Outlook 메일을 자동 전송하지 않는다.
6. 저장소 전체에 `mail.Send()` 또는 `.Send()` 호출을 추가하지 않는다.
7. Excel은 `win32.DispatchEx("Excel.Application")`를 유지한다.
8. 사용자가 열어 둔 다른 Excel 인스턴스를 종료하지 않는다.
9. 기술 예외는 로그에 기록하고 사용자에게는 일반적인 안내만 표시한다.
10. 임직원 이름, 이메일, API Key, 메일 본문 전체를 로그에 기록하지 않는다.
11. 실제 업무 데이터와 개인정보를 공개 저장소에 커밋하지 않는다.
12. 변경 후 `compileall`, `pytest`, `ruff`, `git diff --check`를 실행한다.
13. Windows COM 기능은 실제 Excel·Outlook 환경에서 수동 테스트한다.
14. 테스트하지 않은 항목을 완료로 표시하지 않는다.
15. 각 작업이 끝날 때 변경 파일, 테스트 결과, 남은 위험을 보고한다.

---

# 2. 현재 상태 요약

현재 저장소에는 다음 개선사항이 반영되어 있다.

- RotatingFileHandler 기반 로그
- Excel 독립 COM 인스턴스
- Excel/PDF/PNG 중복 파일 확인
- 파일명 금지 문자 정리
- 출력 폴더 쓰기 권한 검사
- 파일 잠금 검사
- `.env` 백업
- `.env` 백업 파일 Git 제외
- 최초 실행 설정 마법사
- PDF와 PNG 결과 분리
- PDF 성공·PNG 실패 부분 성공 처리
- Outlook 기본 서명 유지 시도
- Outlook HTML escape 적용 시도
- 코드 모듈화

그러나 다음 문제가 남아 있다.

- 공개 저장소에 실제 임직원 주소록 존재
- Outlook 서비스에 정의되지 않은 이름과 중복 함수 존재
- PDF 첨부 실패를 성공으로 처리
- 실제 실행 데이터가 Git에 포함됨
- 오래된 실행 코드가 루트에 중복 존재
- 테스트, README, 최소 의존성, CI 없음
- 초기 설정과 환경 마이그레이션의 오류 처리 미완료
- PDF 한 페이지 출력 보장 없음

---

# 3. 우선순위

## P0 — 보안 및 실행 차단 문제

1. 공개 저장소 개인정보 제거
2. Outlook 서비스 실행 오류 수정
3. PDF 첨부 실패 처리
4. 실제 실행 데이터 Git 추적 중단

## P1 — 안정화 및 배포 기반

1. 오래된 실행 코드 정리
2. 최소 의존성 파일 추가
3. README 작성
4. 테스트·Ruff·GitHub Actions 추가
5. 환경 설정 처리 보완
6. Excel PDF 출력 보완
7. 파일 잠금 검사 강화

## P2 — 구조 개선

1. `DailyReportData` 실제 사용
2. 주간 데이터 저장소 통합
3. UI 전역 상태 제거
4. 작업 중 종료 안전 처리
5. PyInstaller 빌드 구조 정리

## P3 — 신규 기능

P0~P2 완료 후에만 진행한다.

- Daily Report 직접 입력 화면
- 빠른 생성 / Excel 편집 모드
- 자동 임시 저장
- 전일 업무 이어쓰기
- 최근 보고서 목록
- 주소록 관리 UI
- 공휴일 API Key 설정 UI

---

# 4. 작업 1: 공개 저장소 개인정보 제거

## 현재 문제

저장소가 Public 상태인데 `주소록.csv`에 실제 임직원 정보가 포함되어 있다.

포함 정보:

- 실명
- 회사 이메일
- 직책
- 부서
- 회사명

파일을 현재 커밋에서 삭제하는 것만으로는 충분하지 않다. 과거 Git 이력에서도 계속 조회할 수 있기 때문이다.

## 작업 순서

### 4.1 저장소 공개 범위 변경

가장 먼저 저장소를 Private으로 변경한다.

```text
Settings
→ General
→ Danger Zone
→ Change repository visibility
→ Make private
```

저장소를 계속 Public으로 유지해야 한다면 이력 제거를 반드시 수행한다.

### 4.2 실제 주소록 추적 중단

`.gitignore`에 추가한다.

```gitignore
# Private runtime data
주소록.csv
weekly_data.json
```

현재 추적 중인 파일 제거:

```powershell
git rm --cached "주소록.csv"
git rm --cached weekly_data.json
```

### 4.3 샘플 주소록 생성

신규 파일:

```text
주소록.example.csv
```

내용:

```csv
이름,이메일,직책,부서,회사
홍길동,user@example.com,대리,개발팀,회사명
김예시,sample@example.com,사원,기술팀,예시회사
```

실제 이름, 이메일, 회사 내부 정보는 포함하지 않는다.

### 4.4 Git 이력 제거

```powershell
python -m pip install git-filter-repo
git filter-repo --path "주소록.csv" --invert-paths
git push origin --force --all
git push origin --force --tags
```

주의:

- 다른 사용자가 저장소를 사용 중이면 사전에 공지한다.
- 이력 재작성 후 기존 clone은 다시 받아야 한다.
- 작업 전 저장소를 백업한다.

### 4.5 추가 개인정보 검사

다음 파일을 점검한다.

```text
Template/Daily_Report_Template.xlsx
.env.example
DAILY_REPORT_DEVELOPMENT_SPEC.md
DAILY_REPORT_REMEDIATION_SPEC.md
requirements-current.txt
legacy_daily_report.py
main_cmd.py
```

검색 명령:

```powershell
git grep -n "@fmstec.co.kr"
git grep -n "HOLIDAY_API_KEY"
git grep -n "OUTLOOK_TO"
git grep -n "OUTLOOK_SENDER"
```

## 완료 조건

- 실제 `주소록.csv`가 현재 저장소에서 제거됨
- 실제 주소록이 Git 이력에서도 제거됨
- `.gitignore`에 주소록과 실행 데이터 제외 규칙 존재
- 샘플 주소록만 저장소에 포함
- Public 저장소에 실제 임직원 정보 없음

---

# 5. 작업 2: Outlook 서비스 실행 오류 수정

## 대상 파일

```text
services/outlook_service.py
```

## 현재 문제

다음 함수가 중복 정의되어 있다.

```text
normalize_recipient_addresses
render_outlook_template
```

다음 이름은 정의되지 않았거나 import되지 않았다.

```text
html
raw_tokens
OUTLOOK_FRIENDLY_TOKENS
```

발생 가능한 오류:

```text
NameError: name 'html' is not defined
NameError: name 'raw_tokens' is not defined
NameError: name 'OUTLOOK_FRIENDLY_TOKENS' is not defined
```

## 요구사항

파일을 다음 순서로 재정리한다.

```text
1. import
2. normalize_recipient_addresses
3. dedupe_email_list
4. build_outlook_template_values
5. find_unknown_outlook_tokens
6. render_outlook_template
7. build_safe_html_text
8. create_outlook_draft
```

각 함수는 한 번만 정의한다.

## 5.1 import 정리

```python
import json
import os
from datetime import datetime
from html import escape

import win32com.client as win32

from config import (
    DEFAULT_BODY,
    DEFAULT_SUBJECT,
    OUTLOOK_TOKEN_ALIASES,
    OUTLOOK_TOKEN_PATTERN,
)
from utils.logger_utils import logger
```

## 5.2 수신자 정규화

```python
def normalize_recipient_addresses(raw_addresses) -> list[str]:
    if not raw_addresses:
        return []

    if isinstance(raw_addresses, (list, tuple, set)):
        candidates = raw_addresses
    else:
        candidates = str(raw_addresses).replace(",", ";").split(";")

    recipients = []
    seen = set()

    for item in candidates:
        email = str(item).strip()
        if not email:
            continue

        key = email.lower()
        if key in seen:
            continue

        seen.add(key)
        recipients.append(email)

    return recipients
```

```python
def dedupe_email_list(emails: list[str]) -> list[str]:
    return normalize_recipient_addresses(emails)
```

## 5.3 알 수 없는 토큰 검사

```python
def find_unknown_outlook_tokens(text: str) -> list[str]:
    unknown = []
    seen = set()

    for token in OUTLOOK_TOKEN_PATTERN.findall(text or ""):
        if token in OUTLOOK_TOKEN_ALIASES:
            continue

        if token in seen:
            continue

        seen.add(token)
        unknown.append(token)

    return unknown
```

다음 이름은 제거한다.

```text
raw_tokens
OUTLOOK_FRIENDLY_TOKENS
```

`OUTLOOK_FRIENDLY_TOKENS`는 설정 UI에서 사용 가능한 변수 목록을 보여줄 때만 사용한다.

## 5.4 템플릿 렌더링

```python
def render_outlook_template(
    template_text: str,
    template_values: dict[str, str],
) -> str:
    result = str(template_text or "")

    for token, value in template_values.items():
        replacement = "" if value is None else str(value)
        result = result.replace(token, replacement)

        alias_name = OUTLOOK_TOKEN_ALIASES.get(token)
        if alias_name:
            result = result.replace(
                f"{{{alias_name}}}",
                replacement,
            )

    unknown_tokens = find_unknown_outlook_tokens(result)
    if unknown_tokens:
        logger.warning(
            "메일 템플릿에 미치환 토큰이 남아 있음: count=%d",
            len(unknown_tokens),
        )

    return result
```

## 5.5 HTML 안전 처리

```python
def build_safe_html_text(text: str) -> str:
    return escape(text or "").replace("\n", "<br>")
```

사용:

```python
escaped_body = build_safe_html_text(final_body_text)
```

## 5.6 Outlook 초안 생성 순서

```text
1. Outlook 기능 활성화 확인
2. PDF 존재와 크기 확인
3. 수신자·참조 정규화
4. 제목과 본문 렌더링
5. Outlook Application 생성
6. MailItem 생성
7. 발신 계정 적용
8. 수신자·참조·제목 적용
9. PDF 첨부
10. PNG 조건부 첨부
11. mail.Display()
12. 기존 Outlook 서명 읽기
13. 보고서 HTML과 서명 병합
14. created 반환
```

## 5.7 PDF 필수 첨부

PDF 첨부는 `mail.Display()` 전에 수행한다.

```python
try:
    mail.Attachments.Add(norm_pdf)
except Exception:
    logger.exception("PDF 파일 메일 첨부 실패")
    return "failed"
```

PDF가 없거나 0바이트이면 초안 생성을 시작하지 않는다.

```python
if not os.path.isfile(norm_pdf):
    logger.error("Outlook 초안 생성 중단: PDF 없음")
    return "failed"

if os.path.getsize(norm_pdf) <= 0:
    logger.error("Outlook 초안 생성 중단: PDF가 0바이트")
    return "failed"
```

## 5.8 PNG 선택 첨부

```python
image_html = ""

if os.path.isfile(norm_png) and os.path.getsize(norm_png) > 0:
    try:
        attachment = mail.Attachments.Add(norm_png)
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
            "daily_report_img",
        )
        image_html = "<br><br><img src='cid:daily_report_img'>"
    except Exception:
        logger.exception("PNG 본문 이미지 첨부 실패")
```

PNG 실패는 Outlook 초안 전체 실패로 처리하지 않는다.

## 5.9 Outlook 서명 유지

```python
mail.Display()

signature_html = mail.HTMLBody or ""

mail.HTMLBody = (
    "<div style=\"font-family:'맑은 고딕',Arial,sans-serif;"
    "font-size:10pt;\">"
    f"{escaped_body}"
    f"{image_html}"
    "</div>"
    "<br>"
    f"{signature_html}"
)
```

## 5.10 자동 전송 금지 확인

```powershell
git grep -n "mail.Send"
git grep -n "\.Send()"
```

결과가 없어야 한다.

## 완료 조건

- 중복 함수 없음
- 정의되지 않은 이름 없음
- Outlook 설정 저장 정상
- Outlook 초안 생성 정상
- PDF 첨부 정상
- PNG 없을 때 깨진 이미지 없음
- 기본 Outlook 서명 유지
- 자동 전송 없음

---

# 6. 작업 3: 실제 실행 데이터 저장 위치 변경

## 대상 파일

```text
config.py
repositories/weekly_data_repository.py
ui/settings_window.py
```

## 권장 구조

```text
Data/
├─ weekly_data.json
├─ drafts/
└─ backups/
```

`config.py`:

```python
DATA_DIR = os.path.join(BASE_DIR, "Data")
DATA_FILE_PATH = os.path.join(DATA_DIR, "weekly_data.json")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
```

실제 저장 전에:

```python
os.makedirs(DATA_DIR, exist_ok=True)
```

## 원자적 저장

```python
def save_json_atomically(
    data: dict,
    destination: str,
) -> None:
    tmp_path = f"{destination}.tmp"

    with open(tmp_path, "w", encoding="utf-8") as file:
        json.dump(
            data,
            file,
            ensure_ascii=False,
            indent=4,
        )

    os.replace(tmp_path, destination)
```

## 손상 파일 백업

JSON 파싱 실패 시:

```text
Data/backups/weekly_data_corrupt_YYYYMMDD_HHMMSS.json
```

으로 백업한 후 기본값을 사용한다.

## 완료 조건

- 루트 `weekly_data.json` 추적 해제
- 실행 데이터가 `Data/weekly_data.json`에 저장
- JSON 저장이 원자적으로 수행
- 손상 JSON이 백업됨
- 복구 상황이 WARNING 로그에 기록됨

---

# 7. 작업 4: 오래된 실행 코드 정리

## 현재 파일

```text
main.py
main_cmd.py
legacy_daily_report.py
```

## 문제

`main_cmd.py`와 `legacy_daily_report.py`는 이전 방식의 코드를 포함한다.

포함된 문제:

- `.env` 직접 저장
- `print()` 기반 오류 처리
- `win32.Dispatch("Excel.Application")`
- 오래된 PNG 범위
- 현재 모듈 구조와 다른 흐름

실수로 실행하거나 빌드 대상에 포함하면 현재 코드와 다른 동작을 한다.

## 권장 처리

### 선택 A — 삭제

```powershell
git rm main_cmd.py
git rm legacy_daily_report.py
```

### 선택 B — 문서 보관

```text
docs/archive/main_cmd_legacy.py.txt
docs/archive/legacy_daily_report.py.txt
```

확장자를 `.txt`로 변경하여 실행 대상에서 제외한다.

## 진입점

실제 실행 진입점은 `main.py` 하나만 유지한다.

## 완료 조건

- 루트에 실행 가능한 오래된 Python 진입점 없음
- PyInstaller가 `main.py`만 빌드
- README에 공식 진입점 명시

---

# 8. 작업 5: 환경 설정 처리 개선

## 대상 파일

```text
config.py
repositories/settings_repository.py
ui/initial_setup_window.py
```

## 8.1 config import 부작용 제거

`config.py`에서는 경로만 선언한다.

```python
DATA_DIR = os.path.join(BASE_DIR, "Data")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
```

폴더 생성은 실제 저장·백업 시점에만 한다.

## 8.2 환경 설정 로드 개선

```python
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
```

설정값 자체는 로그에 기록하지 않는다.

## 8.3 마이그레이션 기본 키

```python
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
```

## 8.4 초기 설정 오류 처리

상단:

```python
from utils.logger_utils import logger
```

예외 처리:

```python
except Exception:
    logger.exception("초기 환경 설정 저장 실패")
    messagebox.showerror(
        "초기 설정 실패",
        "초기 설정을 저장하지 못했습니다.\n"
        "로그 파일을 확인해 주세요.",
        parent=parent,
    )
    return False
```

초기 설정 딕셔너리에 추가:

```python
"HOLIDAY_API_KEY": "",
```

## 완료 조건

- config import 시 폴더 생성 없음
- 설정 로드 성공·실패 로그 존재
- 설정값 로그 노출 없음
- 누락 기본 키가 모두 마이그레이션됨
- 초기 설정 사용자 화면에 예외 원문 노출 없음

---

# 9. 작업 6: Excel 내보내기 안정화

## 대상 파일

```text
services/excel_service.py
```

## 9.1 Excel SaveAs 형식 명시

```python
wb.SaveAs(
    norm_output_xlsx,
    FileFormat=51,
)
```

`51`은 `.xlsx` 형식이다.

## 9.2 PDF 페이지 맞춤 설정

```python
ws.PageSetup.PrintArea = REPORT_PRINT_AREA
ws.PageSetup.Zoom = False
ws.PageSetup.FitToPagesWide = 1
ws.PageSetup.FitToPagesTall = 1
```

Worksheet 기준으로 내보낸다.

```python
ws.ExportAsFixedFormat(
    Type=0,
    Filename=norm_pdf_path,
    Quality=0,
    IncludeDocProperties=True,
    IgnorePrintAreas=False,
    OpenAfterPublish=False,
)
```

## 9.3 사용자용 오류 메시지

다음 코드는 피한다.

```python
result.error_message = str(exc)
```

대신:

```python
result.error_message = (
    "PDF/PNG 내보내기 중 오류가 발생했습니다."
)
```

상세 COM 오류는 `logger.exception()`에만 기록한다.

## 9.4 ExportResult 파일 분리

신규 권장 파일:

```text
models/export_result.py
```

```python
from dataclasses import dataclass


@dataclass(slots=True)
class ExportResult:
    pdf_success: bool = False
    png_success: bool = False
    pdf_path: str = ""
    png_path: str = ""
    error_message: str = ""

    @property
    def is_full_success(self) -> bool:
        return self.pdf_success and self.png_success

    @property
    def is_partial_success(self) -> bool:
        return self.pdf_success and not self.png_success
```

## 완료 조건

- `.xlsx` 형식 명시
- PDF 전체 범위가 한 페이지에 출력
- COM 예외 원문이 UI로 전달되지 않음
- ExportResult가 별도 모델로 분리
- PDF·PNG 0바이트 검사 유지

---

# 10. 작업 7: 파일 잠금 검사 강화

## 대상 파일

```text
utils/file_utils.py
```

## Windows 권장 구현

```python
import os

import pywintypes
import win32con
import win32file


def is_file_locked(filepath: str) -> bool:
    if not os.path.isfile(filepath):
        return False

    try:
        handle = win32file.CreateFile(
            filepath,
            win32con.GENERIC_READ
            | win32con.GENERIC_WRITE,
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
```

## 완료 조건

- 열린 Excel 파일 감지
- 파일을 닫으면 정상 진행
- OneDrive 경로 수동 테스트
- 한글 경로 수동 테스트

---

# 11. 작업 8: 최소 의존성 파일 작성

## 신규 파일

```text
requirements.txt
requirements-dev.txt
```

## requirements.txt

```text
python-dotenv==1.2.1
pywin32==311
```

## requirements-dev.txt

```text
-r requirements.txt
pyinstaller==6.17.0
pytest
pytest-mock
ruff
```

`requirements-current.txt`는 운영 설치에 사용하지 않는다.

권장 처리:

- `docs/environment/requirements-current.txt`로 이동
- 또는 삭제
- README에 운영 의존성 파일이 아님을 명시

---

# 12. 작업 9: README 작성

## 신규 파일

```text
README.md
```

## 필수 내용

1. 프로젝트 개요
2. 주요 기능
3. 지원 운영체제
4. Excel Desktop 요구사항
5. Classic Outlook 요구사항
6. Python 버전
7. 가상환경 생성
8. 의존성 설치
9. `.env.example` 설정
10. `주소록.example.csv` 사용법
11. 프로그램 실행
12. 보고서 생성 흐름
13. PDF·PNG 출력 설명
14. Outlook 초안은 자동 발송되지 않음
15. 로그 위치
16. 실행 데이터 위치
17. PyInstaller 빌드
18. 알려진 제한사항
19. 개인정보 취급 주의
20. 테스트 실행 방법

## 실행 예시

```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python main.py
```

---

# 13. 작업 10: Ruff와 pytest 설정

## 신규 파일

```text
pyproject.toml
```

```toml
[tool.ruff]
target-version = "py314"
line-length = 120

[tool.ruff.lint]
select = [
    "E",
    "F",
    "I",
]

[tool.pytest.ini_options]
testpaths = ["tests"]
python_files = ["test_*.py"]
```

`F821` 규칙은 정의되지 않은 이름을 검출한다.

---

# 14. 작업 11: 단위 테스트 추가

## 신규 구조

```text
tests/
├─ test_outlook_service.py
├─ test_path_utils.py
├─ test_weekly_data_repository.py
├─ test_settings_repository.py
├─ test_report_data.py
└─ test_file_utils.py
```

## Outlook 서비스 테스트

- 수신자 분리
- 쉼표·세미콜론 처리
- 대소문자 중복 제거
- 템플릿 변수 치환
- 미지정 변수 탐지
- HTML escape
- PDF 없음 처리
- PNG 없음 처리

COM 객체는 mock한다.

## report_data 보완

`from_dict()`의 `headcount` 변환은 잘못된 값에서 예외가 발생할 수 있다.

```python
try:
    headcount = int(data.get("headcount", 1))
except (TypeError, ValueError):
    headcount = 1
```

---

# 15. 작업 12: GitHub Actions CI 추가

## 신규 파일

```text
.github/workflows/test.yml
```

```yaml
name: Test

on:
  push:
  pull_request:

jobs:
  test:
    runs-on: windows-latest

    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.14"

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          python -m pip install -r requirements-dev.txt

      - name: Compile
        run: python -m compileall .

      - name: Ruff
        run: ruff check .

      - name: Pytest
        run: python -m pytest -q
```

Excel과 Outlook COM 통합 테스트는 CI에서 실행하지 않는다.

---

# 16. 작업 13: DailyReportData 실제 사용

## 현재 문제

`models/report_data.py`는 존재하지만 실제 실행 흐름은 dict를 직접 생성한다.

## 권장 적용

```python
report = DailyReportData(
    report_date=now.strftime("%Y-%m-%d"),
    department=department,
    employee_id=employee_id,
    work_location=current_location,
    author_name=author_name,
    headcount=1,
    doc_number=user_num,
)
```

```python
errors = report.validate()
if errors:
    messagebox.showerror(
        "입력 오류",
        "\n".join(errors),
        parent=root_window,
    )
    return
```

```python
report_data = report.to_dict()
```

---

# 17. 작업 14: UI 구조와 종료 처리

## 현재 문제

다음 위젯이 전역 변수로 관리된다.

```text
status_label
start_button
settings_button
outlook_button
```

백그라운드 스레드는 daemon으로 실행된다.

작업 중 창을 닫으면 Excel COM 정리가 끝나기 전에 프로세스가 종료될 수 있다.

## 권장 구조

```python
class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.is_busy = False
        self.close_requested = False
```

작업 중 종료 확인:

```python
def on_closing(self):
    if self.is_busy:
        close = messagebox.askyesno(
            "작업 중",
            "현재 보고서 생성 작업이 진행 중입니다.\n"
            "작업이 끝난 뒤 종료하는 것이 안전합니다.\n\n"
            "그래도 종료하시겠습니까?",
            parent=self.root,
        )
        if not close:
            return

    self.root.destroy()
```

권장 사항:

- daemon thread 제거 검토
- 스레드 완료 이벤트 사용
- 종료 요청 시 신규 작업 차단
- COM 작업이 끝나면 정상 종료

---

# 18. 작업 15: PyInstaller 배포 구조

## 신규 파일

```text
build.ps1
DailyReport.spec
```

## 포함 대상

- `Template/Daily_Report_Template.xlsx`
- 필요 시 `주소록.example.csv`

## 제외 대상

- `.env`
- 실제 `주소록.csv`
- `weekly_data.json`
- `Logs/`
- `Output/`
- `Data/`
- 개발 명세서
- tests

## 빌드 검증

- 신규 PC 실행
- `.env` 없음 초기 설정
- 템플릿 로드
- 로그 생성
- Excel 생성
- PDF 생성
- Outlook 초안 생성

---

# 19. 수동 통합 테스트

## 개인정보 및 Git

- [ ] 저장소가 Private 또는 개인정보 이력 제거 완료
- [ ] 실제 주소록 Git 추적 없음
- [ ] `weekly_data.json` Git 추적 없음
- [ ] `.env` Git 추적 없음
- [ ] 샘플 파일만 공개

## Excel

- [ ] 다른 Excel 문서가 열린 상태
- [ ] 다른 Excel 문서가 종료되지 않음
- [ ] 보고서 Excel이 열린 상태에서 변환 차단
- [ ] PDF 전체 내용 포함
- [ ] PDF 한 페이지
- [ ] PNG 전체 내용 포함
- [ ] OneDrive 경로
- [ ] 한글 경로
- [ ] 출력 경로 권한 없음
- [ ] 템플릿 누락

## Outlook

- [ ] Outlook 설정 저장
- [ ] 주소록 선택
- [ ] PDF + PNG
- [ ] PDF만 존재
- [ ] PDF 없음
- [ ] PNG 0바이트
- [ ] 기본 서명 있음
- [ ] 기본 서명 없음
- [ ] 지정 발신 계정 일치
- [ ] 지정 발신 계정 불일치
- [ ] 수신자 없음
- [ ] Outlook 기능 비활성
- [ ] 자동 전송되지 않음

## 설정 및 데이터

- [ ] `.env` 없음
- [ ] 이전 버전 `.env`
- [ ] 손상된 `.env`
- [ ] 손상된 `weekly_data.json`
- [ ] 공휴일 API Key 없음
- [ ] 공휴일 API 실패
- [ ] 주소록 CSV 손상
- [ ] 주말 실행

---

# 20. 권장 작업 순서

## 1차 작업 — 보안

```text
1. 저장소를 Private으로 변경한다.
2. 주소록.csv와 weekly_data.json의 Git 추적을 중단한다.
3. 주소록.example.csv를 생성한다.
4. 주소록.csv를 Git 이력에서 제거한다.
5. 저장소 전체 개인정보를 검색한다.
6. 결과를 보고한다.
```

## 2차 작업 — Outlook 긴급 수정

```text
1. services/outlook_service.py 중복 함수를 제거한다.
2. html, raw_tokens, OUTLOOK_FRIENDLY_TOKENS 오류를 수정한다.
3. PDF 첨부 실패 시 failed를 반환한다.
4. PNG 조건부 첨부를 유지한다.
5. Outlook 서명을 유지한다.
6. Outlook 순수 함수 테스트를 추가한다.
7. ruff, pytest, compileall을 실행한다.
8. 결과를 보고한다.
```

## 3차 작업 — 실행 데이터와 환경 설정

```text
1. Data 디렉터리 구조를 적용한다.
2. weekly_data.json을 Data로 이동한다.
3. JSON 원자적 저장과 손상 백업을 추가한다.
4. config import 부작용을 제거한다.
5. 환경 설정 로드와 마이그레이션을 보완한다.
6. 초기 설정 예외 처리를 수정한다.
7. 결과를 보고한다.
```

## 4차 작업 — Excel 안정화

```text
1. SaveAs FileFormat=51을 적용한다.
2. PDF 페이지 맞춤을 적용한다.
3. Worksheet 기준 PDF 내보내기를 적용한다.
4. COM 오류 원문이 UI로 전달되지 않게 한다.
5. 실제 Excel에서 수동 테스트한다.
6. 결과를 보고한다.
```

## 5차 작업 — 저장소 품질

```text
1. main_cmd.py와 legacy_daily_report.py를 정리한다.
2. requirements 파일을 추가한다.
3. README를 작성한다.
4. pyproject.toml을 추가한다.
5. tests를 추가한다.
6. GitHub Actions를 추가한다.
7. 모든 자동 검사를 통과시킨다.
8. 결과를 보고한다.
```

## 6차 작업 — 구조 개선

```text
1. DailyReportData를 실제 실행 흐름에 사용한다.
2. UI를 MainWindow 클래스로 정리한다.
3. 작업 중 종료 안전 처리를 구현한다.
4. PyInstaller 배포 구조를 추가한다.
5. 수동 통합 테스트 결과를 기록한다.
```

---

# 21. AI 작업 결과 보고 형식

```markdown
## 작업 결과

### 작업 범위
- Outlook 서비스 긴급 수정

### 변경 파일
- `services/outlook_service.py`
- `tests/test_outlook_service.py`

### 구현 내용
- 중복 함수 제거
- 정의되지 않은 이름 제거
- HTML escape 함수 정리
- PDF 첨부 실패 처리
- PNG 조건부 첨부
- Outlook 기본 서명 유지

### 자동 검증
- `python -m compileall .`: PASS
- `ruff check .`: PASS
- `python -m pytest -q`: PASS
- `git diff --check`: PASS

### 코드 검색
- `mail.Send`: 없음
- 정의되지 않은 이름: 없음
- 중복 함수: 없음

### 수동 테스트
- Outlook 설정 저장: PASS / NOT TESTED
- PDF 첨부: PASS / NOT TESTED
- PNG 조건부 이미지: PASS / NOT TESTED
- 기본 서명 유지: PASS / NOT TESTED
- 자동 전송 없음: PASS / NOT TESTED

### 남은 문제
- Classic Outlook 실제 환경 확인 필요

### 다음 작업
- 실행 데이터 저장 위치 및 환경 설정 처리
```

---

# 22. 최종 완료 조건

## 보안

- [ ] 실제 주소록이 현재 저장소에서 제거됨
- [ ] 실제 주소록이 Git 이력에서 제거됨
- [ ] 실제 임직원 정보가 공개 저장소에 없음
- [ ] `weekly_data.json`이 Git에서 제외됨
- [ ] `.env`가 Git에서 제외됨
- [ ] 샘플 주소록만 저장소에 존재

## Outlook

- [ ] 중복 함수 제거
- [ ] 정의되지 않은 이름 제거
- [ ] Outlook 설정 저장 정상
- [ ] Outlook 초안 생성 정상
- [ ] PDF 첨부 실패가 성공 처리되지 않음
- [ ] PNG 없음 상태 정상
- [ ] 기본 서명 유지
- [ ] 자동 전송 없음

## Excel

- [ ] SaveAs 형식 명시
- [ ] PDF 전체 범위
- [ ] PDF 한 페이지
- [ ] PNG 전체 범위
- [ ] 다른 Excel 인스턴스 보호
- [ ] 열린 보고서 파일 감지

## 저장소 품질

- [ ] 공식 진입점은 `main.py` 하나
- [ ] 최소 `requirements.txt` 존재
- [ ] `requirements-dev.txt` 존재
- [ ] README 존재
- [ ] pyproject.toml 존재
- [ ] tests 존재
- [ ] GitHub Actions 존재
- [ ] Ruff 통과
- [ ] pytest 통과
- [ ] compileall 통과

## 배포

- [ ] PyInstaller 빌드 성공
- [ ] 신규 PC 초기 설정 성공
- [ ] Excel 생성 성공
- [ ] PDF 생성 성공
- [ ] Outlook 초안 성공
- [ ] 로그 생성 성공

위 조건을 충족한 이후 Daily Report 직접 입력 화면과 편의 기능 개발을 시작한다.
