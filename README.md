# Daily Report Automation (일일 업무 보고서 자동화 시스템)

Microsoft Excel Desktop 및 Classic Outlook 환경과 연동하여 Daily Report 엑셀 문서 생성, PDF/PNG 변환, Outlook 메일 초안 작성을 자동화하는 Windows 전용 프로그램입니다.

---

## 주요 기능
- **Excel 보고서 초안 자동 입력**: 지정된 템플릿(`Template/Daily_Report_Template.xlsx`)에 작성자, 근무지, 일일 업무 내역 및 주간 근무지를 반영하여 `.xlsx` 생성.
- **고품질 PDF 및 PNG 내보내기**: Excel COM API를 활용해 한 페이지에 딱 맞춘 PDF 문서 및 고화질 PNG 이미지를 추출.
- **Classic Outlook 초안 작성**: 수신자/참조 정규화, 템플릿 치환, 본문 CID 이미지 삽입 및 기존 아웃룩 서명(Signature) 유지.
- **안전한 데이터 관리**: 개인정보 유출 방지 조치, 원자적(Atomic) JSON 저장 및 손상 시 자동 백업 지원.

---

## System Requirements
1. **OS**: Windows 10 / Windows 11 (64-bit)
2. **Python**: Python 3.14 권장
3. **Microsoft Excel**: Desktop 용 엑셀 애플리케이션 설치 필수 (Classic COM 연동)
4. **Microsoft Outlook**: Classic Outlook 애플리케이션 설치 및 계정 로그인 필수 (New Outlook 미지원)

---

## Quick Start (설치 및 실행)

### 1. 가상환경 생성 및 활성화
```powershell
python -m venv .venv
.\.venv\Scripts\activate
```

### 2. 의존성 설치
- **운영 실행용**:
  ```powershell
  python -m pip install -r requirements.txt
  ```
- **개발 및 테스트/빌드용**:
  ```powershell
  python -m pip install -r requirements-dev.txt
  ```

### 3. 환경 설정 (.env)
처음 실행 시 마법사 화면이 나타나 자동으로 `.env` 파일이 설정됩니다. 수동 설정 시 `.env.example`을 복사하여 사용 가능합니다:
```powershell
copy .env.example .env
```

### 4. 주소록 관리 (선택 사항)
프로젝트 루트 디렉토리에 `주소록.csv` 파일이 존재할 경우 주소록에서 수신자 이메일을 선택할 수 있습니다.
`주소록.example.csv`를 참고하여 작성하세요:
```csv
이름,이메일,직책,부서,회사
홍길동,user@example.com,대리,개발팀,회사명
```

### 5. 프로그램 실행
공식 진입점은 `main.py`입니다:
```powershell
python main.py
```

---

## 보고서 생성 및 처리 흐름
1. **입력 검증**: 작성자명, 부서명, 근무지, 문서 번호 유효성 검사.
2. **Excel 생성**: `win32.DispatchEx("Excel.Application")` 기반 독립 COM 인스턴스로 초안 작성 및 SaveAs(FileFormat=51).
3. **PDF/PNG 변환**:
   - PDF: 한 페이지 출력 맞춤(`FitToPagesWide = 1`, `FitToPagesTall = 1`) 보장.
   - PNG: 보고서 본문 지정 범위(A2:I27) 이미지 변환.
4. **Outlook 초안 작성**:
   - `mail.Display()`로 아웃룩 메일 작성 창을 띄우고 서명을 유지합니다.
   - **⚠️ 절대 메일을 자동 발송(`mail.Send()`)하지 않으며, 사용자가 확인 후 수동으로 보냅니다.**

---

## 데이터 및 로그 위치
- **실행 데이터**: `Data/weekly_data.json` (주간 근무지 및 공휴일 데이터, `.gitignore`로 관리)
- **손상 백업**: `Data/backups/weekly_data_corrupt_*.json`
- **로그 파일**: `Logs/daily_report.log` (RotatingFileHandler 기반 로그 기록)

---

## PyInstaller 배포 빌드
```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```
빌드 결과물은 `dist/DailyReport.exe`로 생성됩니다.

---

## 테스트 및 코드 품질 검증
```powershell
# 린트 검사
ruff check .

# 단체 테스트 실행
python -m pytest -q

# 구문 검사
python -m compileall .
```

---

## 알려진 제한사항 및 개인정보 주의
- **Classic Outlook 필수**: Microsoft Store 전용 "New Outlook"은 COM API를 지원하지 않으므로 Classic Outlook을 사용해야 합니다.
- **개인정보 보안 주의**: `주소록.csv`, `.env`, `weekly_data.json`에는 실제 임직원 개인정보 및 이메일 주소가 포함되므로 절대 Public Git 저장소에 커밋하지 마십시오.
