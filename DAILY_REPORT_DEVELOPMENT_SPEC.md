# Daily Report Automation 개발 실행 명세서

> 문서 목적: AI 코딩 에이전트가 기존 Python/Tkinter 기반 Daily Report 자동화 프로그램을 분석하고, 단계적으로 안정화·기능 개선·테스트·배포할 수 있도록 제공하는 실행형 개발 문서  
> 기준 소스: `Pasted code(1).py`  
> 대상 환경: Windows 10/11, Python 3.11 이상, Microsoft Excel Desktop, Classic Outlook  
> 기본 언어: Python  
> UI 프레임워크: Tkinter  
> Excel/Outlook 연동: `pywin32` COM  
> 배포 방식: PyInstaller

---

## 1. AI 개발 에이전트 실행 지침

이 문서를 읽는 AI 개발 에이전트는 아래 원칙을 반드시 지킨다.

1. 기존 기능을 유지하면서 단계적으로 수정한다.
2. 한 번에 전체 코드를 새로 작성하지 않는다.
3. 각 단계가 완료될 때마다 문법 검사와 관련 기능 테스트를 수행한다.
4. 테스트되지 않은 기능을 완료로 표시하지 않는다.
5. Excel 또는 Outlook COM 객체는 예외 발생 여부와 관계없이 정리한다.
6. Outlook 메일은 자동 전송하지 않는다.
7. 사용자의 기존 Excel 창이나 문서를 종료하지 않는다.
8. 설정 파일과 기존 주간 근무지 데이터의 하위 호환성을 유지한다.
9. 사용자 입력값을 파일명, HTML, 경로에 직접 사용하기 전에 검증한다.
10. 오류를 숨기지 말고 로그 파일에 예외 정보와 호출 스택을 기록한다.
11. UI 스레드에서 장시간 작업을 수행하지 않는다.
12. Tkinter 위젯은 메인 스레드에서만 수정한다.
13. 기능 단위로 커밋 가능한 크기의 변경을 수행한다.
14. 각 단계 종료 후 이 문서의 체크리스트를 갱신한다.
15. 요구사항이 모호할 경우 기존 동작을 보존하는 방향으로 구현한다.

### 절대 금지 사항

- `mail.Send()`를 호출하지 않는다.
- 기존 Excel 프로세스에 연결될 수 있는 방식으로 COM 자동화를 구현하지 않는다.
- 사용자가 열어둔 다른 Excel Application을 종료하지 않는다.
- `.env`에 비밀번호, API 비밀값 또는 메일 본문 원문을 로그로 출력하지 않는다.
- 예외를 `except Exception: pass` 형태로 무시하지 않는다.
- UI 작업을 백그라운드 스레드에서 직접 수행하지 않는다.
- 테스트 없이 대규모 리팩터링을 먼저 수행하지 않는다.
- 기존 보고서 파일을 사용자 확인 없이 덮어쓰지 않는다.

---

## 2. 프로젝트 목표

기존 프로그램은 다음 흐름을 제공한다.

```text
설정 확인
→ 문서 번호 및 근무지 입력
→ Excel 템플릿 생성
→ 사용자가 Excel에서 직접 작성
→ PDF/PNG 변환
→ Outlook 메일 초안 생성
```

최종 목표는 다음 두 가지 생성 방식을 지원하는 Daily Report 프로그램이다.

### 빠른 생성 모드

```text
프로그램에서 업무 내용 입력
→ Excel 생성
→ PDF/PNG 생성
→ Outlook 메일 초안 생성
```

### Excel 편집 모드

```text
프로그램에서 기본 내용 입력
→ Excel 생성 및 열기
→ 사용자가 Excel 추가 편집
→ 저장 후 Excel 종료
→ PDF/PNG 생성
→ Outlook 메일 초안 생성
```

---

## 3. 현재 구현 기능

기존 소스에 다음 기능이 구현되어 있다.

- `.env` 기반 사용자 설정
- 최초 실행 설정 마법사
- 작성자, 부서, 사번, 근무지, 저장 경로 설정
- 주간 근무지 JSON 저장
- 공휴일 API 연동
- 문서 번호 자동 증가
- Excel 템플릿 셀 입력
- PDF 내보내기
- 보고서 범위 PNG 생성
- Outlook 수신자 및 참조 설정
- CSV 주소록 검색 및 선택
- Outlook 메일 제목·본문 변수 치환
- PDF 첨부
- PNG 본문 삽입
- Outlook 초안 창 표시
- 백그라운드 스레드를 이용한 COM 작업
- PyInstaller 실행 환경 경로 처리

기존 기능은 삭제하지 않고 개선한다.

---

## 4. 목표 프로젝트 구조

안정화가 완료된 후 아래 구조로 점진적으로 분리한다.

```text
DailyReport/
├─ main.py
├─ config.py
├─ requirements.txt
├─ README.md
├─ models/
│  ├─ __init__.py
│  └─ report_data.py
├─ services/
│  ├─ __init__.py
│  ├─ excel_service.py
│  ├─ outlook_service.py
│  └─ holiday_service.py
├─ repositories/
│  ├─ __init__.py
│  ├─ settings_repository.py
│  ├─ weekly_data_repository.py
│  └─ draft_repository.py
├─ ui/
│  ├─ __init__.py
│  ├─ main_window.py
│  ├─ report_input_window.py
│  ├─ settings_window.py
│  └─ outlook_settings_window.py
├─ utils/
│  ├─ __init__.py
│  ├─ file_utils.py
│  ├─ com_utils.py
│  └─ logging_config.py
├─ Template/
│  └─ Daily_Report_Template.xlsx
├─ Data/
│  ├─ weekly_data.json
│  └─ drafts/
├─ Logs/
├─ Output/
├─ 주소록.csv
└─ .env
```

### 리팩터링 순서

1. 기존 단일 파일에서 안정화 작업을 먼저 수행한다.
2. 안정화 테스트가 통과한 이후 데이터 모델을 분리한다.
3. 다음으로 Excel, Outlook, 공휴일 서비스를 분리한다.
4. 저장소 계층을 분리한다.
5. 마지막에 UI 파일을 분리한다.
6. 각 분리 단계에서 프로그램이 실행되는지 확인한다.

---

## 5. 데이터 모델

`dict` 중심 데이터 전달을 줄이고 아래 데이터 클래스를 사용한다.

```python
from dataclasses import dataclass
from datetime import date


@dataclass(slots=True)
class DailyReportData:
    report_date: date
    department: str
    employee_id: str
    work_location: str
    author_name: str
    headcount: int
    work_content: str
    tomorrow_work: str
    notes: str
    doc_number: str
```

### 검증 규칙

- `author_name`: 빈 값 금지
- `department`: 빈 값 금지
- `employee_id`: 빈 값 금지
- `work_location`: 빈 값 금지
- `headcount`: 1 이상의 정수
- `work_content`: 빠른 생성 시 필수
- `report_date`: 유효한 날짜
- `doc_number`: 파일명 금지 문자를 포함하지 않아야 함
- 모든 텍스트는 앞뒤 공백을 제거한 후 사용

---

## 6. 개발 단계

# Phase 0. 기준선 확보

## 목적

수정 전 프로그램이 실행 가능한지 확인하고 회귀 테스트 기준을 확보한다.

## 작업

- [ ] 기존 파일을 `legacy_daily_report.py`로 백업한다.
- [ ] Python 버전을 기록한다.
- [ ] 설치된 패키지 목록을 기록한다.
- [ ] 기존 파일에 대해 문법 검사를 수행한다.
- [ ] 테스트용 `.env` 예제를 만든다.
- [ ] 테스트용 Excel 템플릿 경로를 확인한다.
- [ ] 현재 정상 동작 기능을 수동 테스트 목록으로 기록한다.
- [ ] Git 저장소가 없다면 초기화한다.
- [ ] `.gitignore`를 생성한다.

## 권장 `.gitignore`

```gitignore
__pycache__/
*.py[cod]
.venv/
venv/
.env
Output/
Logs/
Data/drafts/
build/
dist/
*.spec
```

## 검증 명령

```powershell
python -m py_compile "Pasted code(1).py"
python -m pip freeze > requirements-current.txt
```

## 완료 조건

- 기존 코드가 문법 검사를 통과한다.
- 기존 실행 파일과 템플릿의 위치가 확인된다.
- 수정 전 동작을 재현할 수 있다.

---

# Phase 1. 로깅 및 공통 오류 처리

## 목적

후속 작업 중 발생하는 오류 원인을 파일에서 추적할 수 있도록 한다.

## 작업

- [ ] `logging`과 `RotatingFileHandler`를 적용한다.
- [ ] 로그 폴더가 없으면 자동 생성한다.
- [ ] 로그 파일 최대 크기는 2MB로 설정한다.
- [ ] 백업 로그는 최대 5개 보관한다.
- [ ] 로그 인코딩은 UTF-8로 설정한다.
- [ ] 프로그램 시작·종료 로그를 추가한다.
- [ ] 설정 로드 결과를 기록한다.
- [ ] Excel 생성·변환 시작과 종료를 기록한다.
- [ ] Outlook 초안 생성 시작과 결과를 기록한다.
- [ ] 공휴일 API 성공·실패를 기록한다.
- [ ] `print()` 중심 오류 처리를 `logger`로 변경한다.
- [ ] 예외 발생 시 `logger.exception()`을 사용한다.
- [ ] GUI에 표시할 사용자 메시지와 로그용 상세 오류를 분리한다.
- [ ] 메인 화면에 `로그 폴더 열기` 버튼을 추가한다.

## 로그에 기록하면 안 되는 값

- Outlook 본문 전체
- 메일 수신자 전체 목록
- API Key
- 사용자의 개인 메모 전체
- 환경 파일 전체 내용

## 완료 조건

- 오류 발생 시 `Logs/daily_report.log`가 생성된다.
- 호출 스택이 로그에 기록된다.
- 사용자는 로그 폴더를 UI에서 열 수 있다.
- 콘솔이 없는 PyInstaller 실행 환경에서도 오류 추적이 가능하다.

---

# Phase 2. Excel COM 안정화

## 목적

기존 Excel 세션에 영향을 주지 않고 보고서 생성과 변환을 수행한다.

## 필수 변경

### 2.1 독립 Excel 인스턴스 사용

다음 코드를 사용하지 않는다.

```python
win32.Dispatch("Excel.Application")
```

다음 방식으로 변경한다.

```python
win32.DispatchEx("Excel.Application")
```

적용 대상:

- Excel 초안 생성
- PDF 내보내기
- PNG 이미지 생성

### 2.2 공통 Excel Application 생성 함수

```python
def create_excel_application(*, visible: bool):
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    excel.ScreenUpdating = visible
    return excel
```

### 2.3 COM 객체 정리

- Workbook을 먼저 닫는다.
- Excel Application을 마지막에 종료한다.
- 모든 종료 코드는 `finally`에서 실행한다.
- 종료 중 발생한 예외는 로그로 남기되 원래 예외를 덮어쓰지 않는다.
- COM 객체 참조를 오래 유지하지 않는다.

### 2.4 템플릿 존재 여부

템플릿이 없을 경우:

1. 작업을 시작하지 않는다.
2. 예상 경로를 사용자에게 표시한다.
3. 로그에 실제 확인 경로를 기록한다.

### 2.5 셀 값 입력 조건 수정

기존의 다음 조건은 사용하지 않는다.

```python
if val and addr:
```

`0`과 빈 문자열도 의도적으로 셀에 반영할 수 있도록 다음 원칙을 적용한다.

```python
if addr:
    cell.Value = "" if val is None else str(val)
```

### 2.6 출력 디렉터리

- Excel 저장 전에 출력 디렉터리를 생성한다.
- 해당 디렉터리에 쓰기 권한이 있는지 사전 확인한다.
- 권한이 없으면 명확한 사용자 메시지를 표시한다.

## 완료 조건

- 사용자가 열어둔 다른 Excel 문서가 종료되지 않는다.
- Excel 생성 중 예외가 발생해도 `EXCEL.EXE` 프로세스가 불필요하게 남지 않는다.
- `headcount=0` 같은 값도 코드상 의도대로 처리할 수 있다.
- 템플릿 누락 시 원인을 확인할 수 있다.

---

# Phase 3. 파일명, 중복 및 잠금 처리

## 목적

잘못된 파일명과 열린 파일로 인한 저장·변환 실패를 방지한다.

## 작업

### 3.1 파일명 정리 함수

다음 공통 함수를 작성한다.

```python
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_filename_part(value: str, fallback: str) -> str:
    text = str(value or "").strip()
    text = INVALID_FILENAME_CHARS.sub("_", text)
    text = text.rstrip(". ")
    return text or fallback
```

적용 대상:

- 부서
- 사번
- 문서 번호
- 작성자 이름을 향후 파일명에 사용하는 경우
- 사용자 지정 접두어나 제목

### 3.2 중복 파일 검사

다음 세 파일을 모두 검사한다.

- Excel
- PDF
- PNG

하나라도 존재하면 기존 파일 목록을 표시하고 덮어쓰기 여부를 확인한다.

### 3.3 원자적 또는 안전한 저장

가능하면 임시 파일에 저장 후 최종 파일명으로 이동한다.

```text
report.xlsx.tmp
→ 저장 성공
→ report.xlsx로 교체
```

Excel COM 특성상 임시 확장자가 문제가 되면 동일 폴더의 임시 `.xlsx` 파일을 사용한다.

### 3.4 파일 잠금 검사

보고서 Excel 파일이 열려 있으면 PDF 변환을 시작하지 않는다.

권장 함수:

```python
def is_file_locked(path: str) -> bool:
    if not os.path.exists(path):
        return False

    try:
        with open(path, "a+b"):
            return False
    except PermissionError:
        return True
```

주의:

- Windows와 Office 잠금 특성을 고려한다.
- 잠금 여부만으로 완전한 상태를 보장할 수 없으므로 실제 Workbook 열기 예외도 처리한다.
- 잠금 상태라면 사용자가 파일을 닫은 후 다시 시도할 수 있도록 한다.

### 3.5 사용자 안내 수정

Excel 편집 모드 안내 문구:

```text
Excel 파일이 열렸습니다.

내용을 작성한 후 저장(Ctrl+S)하고,
Excel 파일을 완전히 닫은 뒤 [확인]을 눌러주세요.
```

## 완료 조건

- 잘못된 파일명 문자 때문에 저장이 실패하지 않는다.
- 기존 Excel/PDF/PNG 파일이 있는 경우 사용자 동의 없이 덮어쓰지 않는다.
- Excel 파일이 열린 상태에서 PDF 변환을 진행하지 않는다.
- 파일 잠금 오류 메시지가 이해하기 쉽게 표시된다.

---

# Phase 4. PDF 및 PNG 생성 개선

## 목적

보고서 전체 내용이 PDF와 메일 본문 이미지에 포함되도록 한다.

## 작업

### 4.1 출력 범위 상수화

```python
REPORT_PRINT_AREA = "$A$2:$I$24"
REPORT_IMAGE_RANGE = "A2:I24"
```

실제 템플릿 구조를 확인하여 필요하면 값을 조정한다.

### 4.2 페이지 설정

```python
ws.PageSetup.PrintArea = REPORT_PRINT_AREA
ws.PageSetup.Zoom = False
ws.PageSetup.FitToPagesWide = 1
ws.PageSetup.FitToPagesTall = 1
```

### 4.3 PDF 내보내기

Worksheet 또는 Workbook 중 보고서 범위가 정확하게 나오는 방식을 선택한다.

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

### 4.4 PNG 생성

- 지정 범위를 `CopyPicture`로 복사한다.
- 임시 Chart를 생성한다.
- 붙여넣기 후 PNG로 내보낸다.
- 내보내기 성공 여부와 실제 파일 존재 여부를 확인한다.
- 임시 Chart는 반드시 삭제한다.
- PNG 생성 실패가 PDF 생성 성공을 무효화하지 않도록 결과를 구분한다.

### 4.5 개별 결과 모델

가능하면 단순 `bool` 대신 결과 객체를 사용한다.

```python
@dataclass(slots=True)
class ExportResult:
    pdf_created: bool
    png_created: bool
    pdf_path: str
    png_path: str
    message: str = ""
```

## 완료 조건

- PDF에 보고서 전체 범위가 포함된다.
- PNG에도 금일 업무, 익일 계획, 특이사항이 포함된다.
- PDF 성공·PNG 실패를 구분할 수 있다.
- PNG가 없을 때 Outlook 본문에 깨진 이미지가 표시되지 않는다.

---

# Phase 5. Outlook 초안 안정화

## 목적

메일 본문, 첨부 파일, 기본 서명을 안정적으로 구성하되 자동 전송하지 않는다.

## 작업

### 5.1 HTML 이스케이프

```python
from html import escape

html_content = escape(final_body_text).replace("\n", "<br>")
```

### 5.2 PNG 조건부 첨부

PNG 파일이 실제 존재하고 크기가 0보다 클 때만:

- 첨부 파일 추가
- Content-ID 설정
- `<img>` 태그 삽입

PNG가 없으면 이미지 태그를 생성하지 않는다.

### 5.3 PDF 첨부

PDF 파일이 존재하면 첨부한다.

PDF가 없으면:

- 사용자에게 경고한다.
- 메일 초안 생성을 중단할지, 첨부 없이 생성할지 정책을 명시한다.
- 기본 정책은 PDF가 없으면 초안 생성을 중단하는 것으로 한다.

### 5.4 Outlook 기본 서명 유지

권장 순서:

1. `mail.Display()`를 호출한다.
2. Outlook이 기본 서명을 삽입하도록 한다.
3. 잠시 후 기존 `mail.HTMLBody`를 읽는다.
4. 생성된 보고서 본문을 기존 서명 앞에 삽입한다.

주의:

- 서명 초기화에 필요한 대기 시간을 과도하게 늘리지 않는다.
- 대기 대신 COM 이벤트 또는 반복 확인이 가능하면 더 안정적인 방식을 사용한다.
- 환경별로 서명이 비어 있을 수 있으므로 빈 값도 정상 처리한다.

### 5.5 발신 계정 선택

- 설정한 SMTP 주소와 Outlook Account의 SMTP 주소를 소문자로 비교한다.
- 일치 계정이 없으면 기본 계정을 사용한다.
- 사용자에게 지정 계정을 찾지 못했다는 안내를 표시한다.
- 계정을 찾지 못한 사실은 로그에 남긴다.

### 5.6 수신자 검증

- 세미콜론 기준으로 분리한다.
- 중복 주소를 제거한다.
- 최소 형식 검사를 수행한다.
- 형식 검사는 지나치게 엄격한 정규식보다 기본 구조 검사를 사용한다.
- 주소가 없더라도 초안 생성은 허용할 수 있으나 사용자에게 안내한다.

### 5.7 메일 전송 금지

다음 코드가 프로젝트 어디에도 없어야 한다.

```python
mail.Send()
```

최종 동작은 반드시 다음이어야 한다.

```python
mail.Display()
```

## 완료 조건

- Outlook 기본 서명이 유지된다.
- 본문에 `<`, `>`, `&`가 있어도 HTML이 깨지지 않는다.
- PNG가 없으면 깨진 본문 이미지가 표시되지 않는다.
- 지정 발신 계정이 없을 때 프로그램이 종료되지 않는다.
- 메일은 초안 창만 열리고 자동 전송되지 않는다.

---

# Phase 6. Daily Report 입력 화면

## 목적

사용자가 Excel을 직접 열지 않고 프로그램에서 보고서를 작성할 수 있도록 한다.

## 화면 필드

| 필드 | 위젯 | 필수 | 기본값 |
|---|---|---:|---|
| 보고일자 | 날짜 Entry | 예 | 오늘 |
| 작성자 | Entry | 예 | 설정값 |
| 부서 | Entry | 예 | 설정값 |
| 사번 | Entry | 예 | 설정값 |
| 근무지 | Combobox/Entry | 예 | 금일 주간 근무지 |
| 투입 인원 | Spinbox | 예 | 1 |
| 문서 번호 | Entry | 예 | 자동 번호 |
| 금일 업무 | ScrolledText | 예 | 빈 값 |
| 익일 업무 | ScrolledText | 아니오 | 빈 값 |
| 특이사항 | ScrolledText | 아니오 | 빈 값 |

## 버튼

- 임시 저장
- 임시 저장 불러오기
- 빠른 생성
- Excel 편집 후 생성
- 입력 초기화
- 출력 폴더 열기
- 최근 보고서
- 닫기

## UI 요구사항

- 최소 창 크기를 지정한다.
- 업무 입력란은 창 크기에 따라 확장되도록 한다.
- 긴 작업 중 버튼을 비활성화한다.
- 상태 표시줄에 현재 단계를 표시한다.
- UI 스레드를 차단하지 않는다.
- 생성 중 창을 닫을 경우 안전하게 작업을 종료하거나 경고한다.
- 키보드 Tab 이동 순서를 자연스럽게 구성한다.
- `Ctrl+S`로 임시 저장할 수 있도록 한다.
- `Ctrl+Enter`는 빠른 생성을 실행할 수 있으나, 실수 방지 확인을 고려한다.

## 입력 검증

- 금일 업무가 비어 있으면 빠른 생성 불가
- 투입 인원은 1 이상의 정수
- 날짜는 `YYYY-MM-DD`
- 저장 경로가 존재하지 않으면 생성 시도
- 저장 경로에 쓰기 권한 확인
- 문서 번호와 파일명 구성값 정리
- 지나치게 긴 업무 내용은 경고하되 임의로 자르지 않는다

## 완료 조건

- 프로그램에서 전체 보고서 내용을 작성할 수 있다.
- 입력값이 Excel 지정 셀에 반영된다.
- 빠른 생성과 Excel 편집 모드를 선택할 수 있다.
- 잘못된 입력은 생성 전에 차단된다.

---

# Phase 7. 임시 저장 및 전일 보고서 활용

## 목적

작성 중 데이터 손실을 줄이고 반복 입력을 최소화한다.

## 임시 저장 경로

```text
Data/drafts/YYYY-MM-DD.json
```

## 저장 필드

- 보고일자
- 작성자
- 부서
- 사번
- 근무지
- 투입 인원
- 문서 번호
- 금일 업무
- 익일 업무
- 특이사항
- 마지막 저장 시간
- 프로그램 버전

## 작업

- [ ] 수동 임시 저장
- [ ] 일정 시간마다 자동 저장
- [ ] 프로그램 시작 시 미완료 초안 감지
- [ ] 복구 여부 확인
- [ ] 보고서 생성 성공 후 초안 삭제 또는 완료 상태 표시
- [ ] 손상된 JSON은 백업 후 새 초안으로 시작
- [ ] 전일 보고서 선택 기능
- [ ] 전일의 `익일 업무`를 오늘의 `금일 업무`로 복사
- [ ] 복사 전에 기존 입력 덮어쓰기 여부 확인

## 완료 조건

- 비정상 종료 후 입력 내용을 복구할 수 있다.
- 전일 계획을 금일 업무로 복사할 수 있다.
- 손상된 초안 때문에 프로그램 전체가 실행 실패하지 않는다.

---

# Phase 8. 최근 보고서 목록

## 목적

생성된 보고서를 프로그램에서 쉽게 조회하고 열 수 있도록 한다.

## 표시 정보

- 보고일자
- 문서 번호
- 작성자
- Excel 존재 여부
- PDF 존재 여부
- 생성 경로
- 수정 일시

## 기능

- Excel 열기
- PDF 열기
- 폴더 열기
- 파일명 복사
- 최근 보고서 새로고침
- 날짜 또는 문서 번호 검색

## 데이터 수집 방식

1차 버전에서는 출력 폴더를 스캔한다.  
향후 필요하면 보고서 인덱스 JSON 또는 SQLite로 변경한다.

## 완료 조건

- 최근 보고서가 최신순으로 표시된다.
- 존재하지 않는 파일 버튼은 비활성화된다.
- 파일이 삭제되어도 프로그램이 중단되지 않는다.

---

# Phase 9. 주간 근무지 및 공휴일 기능 개선

## 목적

기존 주간 근무지 기능을 유지하면서 관리 편의성과 오류 복구를 강화한다.

## 작업

- [ ] 월~금 근무지를 한 화면에서 수정한다.
- [ ] 공휴일을 별도 상태로 표시한다.
- [ ] 연차, 반차, 휴가 키워드 표시 규칙을 유지한다.
- [ ] 기본 근무지 일괄 적용 버튼을 추가한다.
- [ ] 이전 주 근무지 복사 기능을 검토한다.
- [ ] `weekly_data.json` 손상 시 백업 후 기본값으로 복구한다.
- [ ] API Key가 없으면 기능을 조용히 건너뛰되 상태 메시지를 표시한다.
- [ ] 공휴일 API 실패가 보고서 생성을 막지 않도록 한다.
- [ ] API 응답 구조 검증을 강화한다.
- [ ] 네트워크 타임아웃은 명시적으로 설정한다.

## 완료 조건

- API 장애 시에도 수동 근무지 입력으로 보고서를 생성할 수 있다.
- 손상된 주간 데이터가 자동 복구된다.
- 공휴일과 휴가 표시가 Excel에 정상 반영된다.

---

# Phase 10. 코드 모듈화

## 목적

단일 파일 구조를 기능별 모듈로 분리한다.

## 분리 순서

### 10.1 `config.py`

포함 내용:

- `BASE_DIR`
- `BUNDLE_DIR`
- 파일 경로
- 셀 매핑
- 출력 범위
- 기본 메일 템플릿
- 정규식
- 프로그램 버전

### 10.2 `models/report_data.py`

포함 내용:

- `DailyReportData`
- `ExportResult`
- 필요 시 `GenerationResult`

### 10.3 `utils/logging_config.py`

포함 내용:

- 로그 초기화
- 로그 폴더 생성
- RotatingFileHandler

### 10.4 `utils/file_utils.py`

포함 내용:

- 파일명 정리
- 중복 검사
- 파일 잠금 검사
- 쓰기 권한 검사
- 파일 열기

### 10.5 `services/excel_service.py`

포함 내용:

- Excel Application 생성
- Excel 보고서 작성
- PDF 생성
- PNG 생성
- COM 정리

### 10.6 `services/outlook_service.py`

포함 내용:

- Outlook Application 생성
- 발신 계정 검색
- 수신자 정리
- HTML 본문 생성
- 첨부 파일 처리
- 초안 표시

### 10.7 `services/holiday_service.py`

포함 내용:

- 공휴일 API 요청
- 응답 검증
- 주간 공휴일 변환

### 10.8 저장소 계층

포함 내용:

- `.env` 읽기·쓰기
- 주간 데이터 읽기·쓰기
- 임시 저장 읽기·쓰기

### 10.9 UI 계층

포함 내용:

- 메인 창
- 보고서 입력 창
- 기본 설정 창
- Outlook 설정 창
- 최근 보고서 창

## 모듈화 완료 조건

- 순환 import가 없다.
- 전역 Tkinter 위젯 변수가 제거되거나 최소화된다.
- 서비스는 Tkinter 위젯을 직접 참조하지 않는다.
- UI는 서비스의 결과를 받아 사용자 메시지를 표시한다.
- 각 모듈을 독립적으로 import할 수 있다.

---

# Phase 11. 테스트

## 11.1 자동 테스트 대상

가능한 함수는 `pytest`로 테스트한다.

### 단위 테스트

- `sanitize_filename_part`
- 문서 번호 파싱
- 수신자 정규화 및 중복 제거
- Outlook 변수 치환
- 알 수 없는 Outlook 변수 탐지
- 주간 데이터 정규화
- 날짜 및 파일명 생성
- 임시 저장 직렬화·역직렬화
- 공휴일 API 응답 파싱
- 설정 마이그레이션

### 권장 테스트 구조

```text
tests/
├─ test_file_utils.py
├─ test_template_renderer.py
├─ test_weekly_data.py
├─ test_draft_repository.py
└─ test_holiday_service.py
```

## 11.2 수동 통합 테스트

### Excel

- [ ] Excel이 실행되지 않은 상태
- [ ] 다른 Excel 문서가 열린 상태
- [ ] 보고서 Excel이 열린 상태
- [ ] 보고서 저장 후 닫은 상태
- [ ] 템플릿 누락
- [ ] 저장 권한 없음
- [ ] 한글 및 공백 포함 경로
- [ ] OneDrive 경로
- [ ] 네트워크 공유 경로
- [ ] Excel 변환 중 강제 종료

### Outlook

- [ ] Outlook 실행 중
- [ ] Outlook 미실행
- [ ] 로그인된 계정 1개
- [ ] 로그인된 계정 여러 개
- [ ] 지정 발신 계정 일치
- [ ] 지정 발신 계정 불일치
- [ ] 기본 서명 있음
- [ ] 기본 서명 없음
- [ ] PDF 있음
- [ ] PNG 있음
- [ ] PNG 없음
- [ ] 수신자 없음
- [ ] Outlook 기능 비활성

### 데이터

- [ ] `.env` 없음
- [ ] 이전 버전 `.env`
- [ ] `weekly_data.json` 없음
- [ ] `weekly_data.json` 손상
- [ ] 초안 JSON 손상
- [ ] 새로운 주 시작
- [ ] 공휴일 API Key 없음
- [ ] 공휴일 API 실패
- [ ] 주말 실행

## 11.3 테스트 명령

```powershell
python -m compileall .
python -m pytest -q
```

가능하면 정적 검사도 추가한다.

```powershell
python -m pip install ruff
ruff check .
```

## 완료 조건

- 자동 테스트가 모두 통과한다.
- 핵심 통합 테스트 결과가 문서화된다.
- 실패한 테스트를 숨기지 않는다.
- 테스트 실패 상태에서 배포 파일을 생성하지 않는다.

---

# Phase 12. PyInstaller 배포

## 목적

Python이 설치되지 않은 Windows PC에서도 실행할 수 있도록 한다.

## 권장 방식

초기 테스트는 `--onedir`로 수행한다.  
안정화 후 필요하면 `--onefile`을 검토한다.

### Onedir 예시

```powershell
pyinstaller `
  --noconsole `
  --onedir `
  --name DailyReport `
  --add-data "Template;Template" `
  --add-data "주소록.csv;." `
  main.py
```

### Onefile 예시

```powershell
pyinstaller `
  --noconsole `
  --onefile `
  --name DailyReport `
  --add-data "Template;Template" `
  --add-data "주소록.csv;." `
  main.py
```

## 주의사항

- `.env`는 실행 파일에 내장하지 않는다.
- 사용자 수정이 필요한 주소록은 외부 파일로 두는 방식을 우선 검토한다.
- 템플릿을 내장할지 외부 수정 가능 파일로 둘지 정책을 결정한다.
- 프로그램 로그와 데이터는 실행 파일 내부가 아니라 실행 파일 옆 또는 사용자 데이터 폴더에 저장한다.
- 회사 보안 정책에 따라 실행 파일 서명을 검토한다.
- 백신 오탐 가능성을 확인한다.

## 배포 전 검사

- [ ] 깨끗한 Windows 10 PC
- [ ] 깨끗한 Windows 11 PC
- [ ] Python 미설치 PC
- [ ] Excel/Outlook 설치 PC
- [ ] 일반 사용자 권한
- [ ] 한글 사용자명 경로
- [ ] OneDrive 바탕화면
- [ ] 사내 보안 프로그램 실행 환경

## 완료 조건

- Python 없이 프로그램이 실행된다.
- 최초 설정이 정상 동작한다.
- 템플릿을 찾을 수 있다.
- Excel, PDF, PNG가 생성된다.
- Outlook 초안이 열린다.
- 로그와 데이터가 재실행 후 유지된다.

---

## 7. 상태 및 진행률 표시

상태 메시지는 아래 표준을 따른다.

```text
준비됨
1/5 입력값 확인 중...
2/5 Excel 보고서 생성 중...
3/5 PDF 생성 중...
4/5 이미지 생성 중...
5/5 Outlook 초안 생성 중...
완료
```

부분 성공 예시:

```text
PDF 생성 완료, PNG 생성 실패
보고서 생성 완료, Outlook 기능 비활성
보고서 생성 완료, Outlook 초안 생성 실패
```

성공과 실패를 하나의 `bool`로만 표현하지 말고 단계별 결과를 유지한다.

---

## 8. 사용자 오류 메시지 표준

| 내부 오류 | 사용자 메시지 |
|---|---|
| Excel COM 생성 실패 | Microsoft Excel이 설치되어 있는지 확인해 주세요. |
| Workbook 열기 실패 | 템플릿 또는 보고서 파일을 열 수 없습니다. 파일이 사용 중인지 확인해 주세요. |
| 파일 잠금 | 보고서 Excel 파일을 닫은 후 다시 시도해 주세요. |
| 저장 권한 없음 | 선택한 폴더에 파일을 저장할 권한이 없습니다. |
| 템플릿 없음 | Daily Report 템플릿 파일을 찾을 수 없습니다. |
| PDF 내보내기 실패 | Excel 인쇄 설정과 보고서 파일 상태를 확인해 주세요. |
| Outlook COM 실패 | Classic Outlook이 설치되고 로그인되어 있는지 확인해 주세요. |
| 발신 계정 없음 | 지정한 발신 계정을 찾지 못해 기본 계정으로 초안을 생성합니다. |
| 공휴일 API 실패 | 공휴일을 조회하지 못했습니다. 근무지를 직접 확인해 주세요. |
| JSON 손상 | 저장 데이터가 손상되어 기본값으로 복구했습니다. |
| 알 수 없는 오류 | 작업 중 오류가 발생했습니다. 로그 파일을 확인해 주세요. |

기술 예외 메시지는 로그에만 상세히 기록한다.

---

## 9. 설정 키

기존 설정 키와의 호환성을 유지한다.

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

향후 추가 가능한 설정:

```dotenv
REPORT_MODE=quick
AUTO_SAVE_INTERVAL_SECONDS=60
OPEN_PDF_AFTER_EXPORT=False
KEEP_DRAFT_AFTER_SUCCESS=False
LOG_LEVEL=INFO
```

### 설정 파일 저장 원칙

- 기존 알 수 없는 설정 키를 삭제하지 않는다.
- 값을 수정할 때 전체 파일을 안전하게 다시 쓴다.
- 저장 전 백업 파일 생성을 검토한다.
- 멀티라인 본문은 JSON 인코딩을 유지한다.
- 비어 있는 값과 누락된 값을 구분한다.

---

## 10. Excel 셀 매핑

현재 기본 매핑은 다음과 같다.

```python
CELL_MAP = {
    "report_date": "C7",
    "department": "I6",
    "work_location": "I7",
    "author_name": "C8",
    "headcount": "I8",
    "work_content": "D12",
    "tomorrow_work": "B18",
    "notes": "B23",
}
```

### 매핑 정책

- 셀 주소는 코드 여러 곳에 중복 작성하지 않는다.
- 향후 JSON 또는 설정 파일로 분리할 수 있도록 한 곳에서 관리한다.
- 템플릿 버전이 바뀌면 매핑 버전도 함께 기록한다.
- 병합 셀인 경우 왼쪽 위 셀 주소를 사용한다.
- 업무 내용 셀은 줄 바꿈, 상단 정렬, 왼쪽 정렬을 적용한다.
- 필요하면 행 높이 자동 조정을 검토한다.
- 텍스트가 지나치게 길 때 페이지가 깨지지 않는지 테스트한다.

---

## 11. 결과 파일명 규칙

기본 규칙:

```text
FMS{YY}_{DEPARTMENT}{EMPLOYEE_ID}_{DOC_NO} @ Daily Report_{YYMMDD}
```

예시:

```text
FMS26_TE044_001 @ Daily Report_260724.xlsx
FMS26_TE044_001 @ Daily Report_260724.pdf
FMS26_TE044_001 @ Daily Report_260724.png
```

### 규칙

- 모든 구성값을 파일명 정리 함수에 통과시킨다.
- 자동 문서 번호는 연도 폴더 전체에서 계산한다.
- 숫자 문서 번호는 기본 3자리로 표시한다.
- 사용자 지정 문자 문서 번호는 허용하되 금지 문자는 치환한다.
- 동일 번호가 이미 있으면 사용자 확인을 받는다.
- 문서 번호 자동 증가의 동시 실행 문제는 초기 버전에서 경고로 처리하고 향후 잠금 파일을 검토한다.

---

## 12. 성능 및 안정성 요구사항

- 일반적인 보고서 1건 생성 작업은 UI가 응답 가능한 상태를 유지해야 한다.
- 네트워크 요청 타임아웃은 5초 이내를 기본으로 한다.
- Excel/Outlook 작업에는 단계별 상태 메시지를 표시한다.
- 무한 대기 루프를 사용하지 않는다.
- 파일 잠금 재시도에는 최대 횟수 또는 제한 시간을 둔다.
- 자동 저장은 UI 입력을 방해하지 않아야 한다.
- 프로그램 종료 시 실행 중 작업 상태를 확인한다.
- 중복 생성 버튼 클릭을 방지한다.
- 작업 완료 또는 실패 후 버튼 상태를 반드시 복구한다.

---

## 13. 보안 요구사항

- 메일 자동 발송 금지
- API Key 로그 출력 금지
- 메일 본문 전체 로그 출력 금지
- 사용자 입력 HTML 이스케이프
- 파일명 및 경로 입력 검증
- 외부 명령 실행 최소화
- 주소록 CSV 값은 코드로 실행하지 않음
- 환경 파일을 실행 파일에 포함하지 않음
- 임시 파일에 민감 정보가 남지 않도록 관리
- Outlook 첨부 파일은 생성된 보고서 경로만 허용

---

## 14. AI 작업 보고 형식

AI 에이전트는 각 작업이 끝날 때 아래 형식으로 결과를 보고한다.

```markdown
## 작업 결과

### 변경한 파일
- `services/excel_service.py`
- `utils/file_utils.py`

### 구현 내용
- Excel 독립 인스턴스 적용
- 파일 잠금 검사 추가

### 검증
- `python -m compileall .` 통과
- `python -m pytest -q` 통과
- 다른 Excel 문서를 연 상태에서 수동 테스트 통과

### 남은 문제
- Outlook 기본 서명은 Office 버전별 추가 확인 필요

### 다음 작업
- PDF/PNG 출력 범위 개선
```

### 작업 단위 원칙

하나의 작업에서 아래 항목을 과도하게 동시에 변경하지 않는다.

- UI 전체 재작성
- 서비스 전체 분리
- 배포 설정 변경
- 데이터 형식 변경

변경 범위가 크면 하위 작업으로 분리한다.

---

## 15. 구현 우선순위

### P0: 반드시 먼저 수행

- [ ] 기준선 백업 및 문법 검사
- [ ] 로그 시스템
- [ ] `DispatchEx` 적용
- [ ] COM 객체 정리
- [ ] 파일 잠금 처리
- [ ] 파일명 검증
- [ ] 중복 파일 검사
- [ ] PDF/PNG 출력 범위 수정
- [ ] Outlook HTML 이스케이프
- [ ] Outlook 자동 전송 금지 확인

### P1: 핵심 사용자 기능

- [ ] Daily Report 직접 입력 화면
- [ ] 빠른 생성 모드
- [ ] Excel 편집 모드
- [ ] 입력 검증
- [ ] 단계별 상태 표시
- [ ] 임시 저장 및 복구

### P2: 편의 기능

- [ ] 전일 업무 가져오기
- [ ] 최근 보고서 목록
- [ ] 주간 근무지 관리 개선
- [ ] 로그 폴더 및 출력 폴더 열기
- [ ] 주소록 관리 개선

### P3: 구조 및 배포

- [ ] 모듈 분리
- [ ] 자동 테스트 확대
- [ ] PyInstaller 배포
- [ ] 깨끗한 PC 테스트
- [ ] 사용자 설명서

---

## 16. 최종 완료 기준

다음 조건을 모두 만족해야 1차 버전을 완료로 판단한다.

- [ ] 프로그램에서 보고서 내용을 직접 입력할 수 있다.
- [ ] 입력 내용이 Excel 템플릿에 정확히 반영된다.
- [ ] 빠른 생성 모드가 정상 동작한다.
- [ ] Excel 편집 모드가 정상 동작한다.
- [ ] PDF에 전체 보고서가 포함된다.
- [ ] PNG에 전체 보고서가 포함된다.
- [ ] Outlook 메일 초안에 PDF가 첨부된다.
- [ ] PNG가 있을 때만 본문 이미지가 표시된다.
- [ ] Outlook 기본 서명이 유지된다.
- [ ] 메일은 자동 전송되지 않는다.
- [ ] 사용자의 다른 Excel 문서가 종료되지 않는다.
- [ ] 열린 보고서 파일을 감지하고 변환을 차단한다.
- [ ] 오류 발생 시 로그 파일에서 원인을 확인할 수 있다.
- [ ] 설정 및 주간 데이터가 이전 버전과 호환된다.
- [ ] 자동 테스트가 통과한다.
- [ ] Python 미설치 PC에서 배포본이 실행된다.
- [ ] Windows 10과 Windows 11에서 핵심 기능을 검증했다.

---

## 17. 권장 AI 실행 프롬프트

아래 프롬프트를 AI 코딩 도구에 함께 제공한다.

```text
첨부된 기존 Python 소스와 DAILY_REPORT_DEVELOPMENT_SPEC.md를 기준으로 개발을 진행해라.

규칙:
1. 문서의 Phase 순서대로 작업한다.
2. 한 번에 하나의 Phase 또는 작은 하위 작업만 구현한다.
3. 기존 기능을 삭제하거나 동작을 변경하기 전에 영향 범위를 설명한다.
4. 변경 후 반드시 문법 검사와 관련 테스트를 실행한다.
5. 테스트 결과와 변경 파일을 보고한다.
6. Outlook 메일은 절대 자동 전송하지 않는다.
7. 사용자가 실행 중인 다른 Excel 프로세스에 영향을 주지 않는다.
8. 막히는 부분이 있으면 임의로 기능을 삭제하지 말고 원인과 대안을 기록한다.
9. 우선 Phase 0과 Phase 1부터 수행한다.
10. 각 Phase 완료 후 다음 Phase로 넘어가기 전에 현재 상태를 요약한다.
```

---

## 18. 첫 번째 작업 지시

AI 에이전트의 첫 작업은 다음과 같다.

```text
1. 기존 소스를 백업한다.
2. 현재 소스의 문법 검사를 수행한다.
3. 프로젝트 폴더 구조와 의존성을 분석한다.
4. logging_config를 추가한다.
5. 기존 print 기반 오류 출력을 logger로 교체하되 기능 동작은 변경하지 않는다.
6. 프로그램 시작·종료, Excel, Outlook, 공휴일 API 작업 로그를 추가한다.
7. 변경 후 compileall을 실행한다.
8. 변경 파일, 테스트 결과, 남은 위험을 보고한다.
```

이 작업이 검증되기 전에는 UI 재작성이나 전체 모듈화를 시작하지 않는다.
