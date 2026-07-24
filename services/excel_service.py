import os
import time
from datetime import date, timedelta
import win32com.client as win32
from config import CELL_MAP
from utils.logger_utils import logger

def get_week_dates():
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    dates = []
    for i in range(5):
        day = monday + timedelta(days=i)
        weekday_kor = ["월", "화", "수", "목", "금", "토", "일"][day.weekday()]
        dates.append(f"{day.day}({weekday_kor})")
    return dates

def create_excel_application(visible=False):
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    return excel

def generate_excel_draft(report_data: dict, output_xlsx: str, template_path: str, weekly_locations: list, holiday_indices: list = None) -> bool:
    if not os.path.exists(template_path):
        logger.error(f"템플릿 파일을 찾을 수 없습니다: {template_path}")
        return False

    excel = None
    wb = None
    logger.info("독립 Excel 인스턴스로 초안 생성 시작")
    try:
        excel = create_excel_application(visible=False)
        wb = excel.Workbooks.Open(template_path)
        ws = wb.Worksheets(1)

        for key, addr in CELL_MAP.items():
            if addr:
                val = report_data.get(key)
                cell = ws.Range(addr)
                cell.Value = "" if val is None else str(val)
                if key in ["work_content", "tomorrow_work", "notes"]:
                    cell.WrapText = True
                    cell.VerticalAlignment = -4160
                    cell.HorizontalAlignment = -4131

        week_cells = ["D15", "E15", "F15", "G15", "H15"]
        for cell, value in zip(week_cells, get_week_dates()):
            ws.Range(cell).Value = value

        location_cells = ["D16", "E16", "F16", "G16", "H16"]
        red_keywords = ["공휴일", "휴가", "연차", "반차", "휴일"]
        for i, cell_addr in enumerate(location_cells):
            cell = ws.Range(cell_addr)
            is_holiday = bool(holiday_indices and i in holiday_indices)
            if is_holiday:
                val = "공휴일"
                is_red = True
            else:
                val = weekly_locations[i] if i < len(weekly_locations) else ""
                is_red = False
                if val:
                    for keyword in red_keywords:
                        if keyword in str(val):
                            is_red = True
                            break
            cell.Value = val
            cell.Font.Color = 255 if is_red else 0

        wb.SaveAs(output_xlsx)
        wb.Close(SaveChanges=False)
        wb = None
        logger.info("엑셀 초안 생성 완료")
        return True
    except Exception:
        logger.exception("엑셀 초안 생성 실패")
        return False
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                logger.exception("Workbook 닫기 중 예외 발생")
        if excel:
            try:
                excel.Quit()
            except Exception:
                logger.exception("Excel Application 종료 중 예외 발생")

def export_final_reports(xlsx_path: str, pdf_path: str, png_path: str) -> bool:
    if not os.path.exists(xlsx_path):
        logger.error(f"PDF/PNG 변환 실패: 대상 엑셀 파일 없음 ({xlsx_path})")
        return False

    excel = None
    wb = None
    logger.info(f"독립 Excel 인스턴스로 PDF/PNG 변환 시작: {os.path.basename(xlsx_path)}")
    try:
        excel = create_excel_application(visible=True)
        wb = excel.Workbooks.Open(xlsx_path)
        ws = wb.Worksheets(1)
        ws.Activate()

        time.sleep(2)

        wb.ExportAsFixedFormat(0, pdf_path)
        logger.info("PDF 내보내기 성공")

        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1

        rng = ws.Range("A2:I16")
        rng.Select()
        time.sleep(1)

        rng.CopyPicture(1, 2)

        chart = ws.Shapes.AddChart2()
        chart.Select()
        excel.Selection.Width = rng.Width
        excel.Selection.Height = rng.Height
        chart.Chart.Paste()
        time.sleep(1)

        chart.Chart.Export(png_path)
        chart.Delete()
        logger.info("PNG 내보내기 성공")

        wb.Close(SaveChanges=False)
        wb = None
        return True
    except Exception:
        logger.exception("PDF/PNG 변환 실패")
        return False
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                logger.exception("Workbook 닫기 중 예외 발생")
        if excel:
            try:
                excel.Quit()
            except Exception:
                logger.exception("Excel Application 종료 중 예외 발생")
