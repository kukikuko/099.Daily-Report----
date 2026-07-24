from dataclasses import dataclass, asdict
from datetime import date
from utils.path_utils import sanitize_filename_part

@dataclass(slots=True)
class DailyReportData:
    report_date: str
    department: str
    work_location: str
    author_name: str
    headcount: int = 1
    doc_number: str = "001"
    employee_id: str = "000"
    work_content: str = ""
    tomorrow_work: str = ""
    notes: str = ""

    def validate(self) -> list:
        """데이터 유효성을 검사하여 에러 메시지 리스트를 반환한다."""
        errors = []
        if not self.author_name or not self.author_name.strip():
            errors.append("작성자 이름(author_name)은 필수입니다.")
        if not self.department or not self.department.strip():
            errors.append("부서명(department)은 필수입니다.")
        if not self.work_location or not self.work_location.strip():
            errors.append("근무지(work_location)는 필수입니다.")
        if self.headcount < 1:
            errors.append("인원 수(headcount)는 1 이상이어야 합니다.")
        return errors

    def to_dict(self) -> dict:
        """dict 형태로 변환한다."""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'DailyReportData':
        """dict 데이터로부터 DailyReportData 객체를 생성한다."""
        return cls(
            report_date=str(data.get("report_date", "")).strip(),
            department=str(data.get("department", "")).strip(),
            work_location=str(data.get("work_location", "")).strip(),
            author_name=str(data.get("author_name", "")).strip(),
            headcount=int(data.get("headcount", 1)),
            doc_number=str(data.get("doc_number", "001")).strip(),
            employee_id=str(data.get("employee_id", "000")).strip(),
            work_content=str(data.get("work_content", "")).strip(),
            tomorrow_work=str(data.get("tomorrow_work", "")).strip(),
            notes=str(data.get("notes", "")).strip(),
        )
