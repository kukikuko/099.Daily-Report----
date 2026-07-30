import os
import time
from datetime import date, timedelta

import win32com.client as win32

from config import CELL_MAP, REPORT_IMAGE_RANGE, REPORT_PRINT_AREA
from models.export_result import ExportResult
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


def generate_excel_draft(
    report_data: dict,
    output_xlsx: str,
    template_path: str,
    weekly_locations: list,
    holiday_indices: list = None,
) -> bool:
    norm_template_path = os.path.abspath(os.path.normpath(template_path)) if template_path else ""
    norm_output_xlsx = os.path.abspath(os.path.normpath(output_xlsx)) if output_xlsx else ""

    if not os.path.exists(norm_template_path):
        logger.error(f"템플릿 파일을 찾을 수 없습니다: {norm_template_path}")
        return False

    excel = None
    wb = None
    logger.info("독립 Excel 인스턴스로 초안 생성 시작")
    try:
        excel = create_excel_application(visible=False)
        try:
            wb = excel.Workbooks.Open(norm_template_path)
        except Exception:
            logger.exception(f"엑셀 템플릿 Workbooks.Open 실패 ({os.path.basename(norm_template_path)})")
            raise

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
                is_red = any(k in str(val) for k in red_keywords)

            cell.Value = val
            if is_red:
                cell.Font.Color = 255
            else:
                cell.Font.ColorIndex = -4105

        wb.SaveAs(norm_output_xlsx, FileFormat=51)
        logger.info(f"Excel 파일 저장 완료: {os.path.basename(norm_output_xlsx)}")
        wb.Close(SaveChanges=False)
        wb = None
        return True
    except Exception:
        logger.exception("Excel 초안 생성 중 예외 발생")
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


def export_final_reports(xlsx_path: str, pdf_path: str, png_path: str) -> ExportResult:
    norm_xlsx_path = os.path.abspath(os.path.normpath(xlsx_path)) if xlsx_path else ""
    norm_pdf_path = os.path.abspath(os.path.normpath(pdf_path)) if pdf_path else ""
    norm_png_path = os.path.abspath(os.path.normpath(png_path)) if png_path else ""

    result = ExportResult(pdf_path=norm_pdf_path, png_path=norm_png_path)

    if not os.path.exists(norm_xlsx_path):
        err_msg = f"대상 엑셀 파일 없음 ({norm_xlsx_path})"
        logger.error(f"PDF/PNG 변환 실패: {err_msg}")
        result.error_message = "대상 엑셀 파일을 찾을 수 없습니다."
        return result

    excel = None
    wb = None
    chart_shape = None

    logger.info(f"독립 Excel 인스턴스로 PDF/PNG 변환 시작: {os.path.basename(norm_xlsx_path)}")
    try:
        excel = create_excel_application(visible=True)
        try:
            wb = excel.Workbooks.Open(norm_xlsx_path)
        except Exception:
            logger.exception(f"변환 대상 엑셀 Workbooks.Open 실패 ({os.path.basename(norm_xlsx_path)})")
            result.error_message = "엑셀 파일을 여는 중 오류가 발생했습니다."
            return result

        ws = wb.Worksheets(1)
        ws.Activate()
        time.sleep(1)

        # 1. PDF 내보내기 (인쇄 범위 및 페이지 맞춤 지정)
        try:
            ws.PageSetup.PrintArea = REPORT_PRINT_AREA
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = 1

            ws.ExportAsFixedFormat(
                Type=0,
                Filename=norm_pdf_path,
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            if os.path.exists(norm_pdf_path) and os.path.getsize(norm_pdf_path) > 0:
                result.pdf_success = True
                logger.info(f"PDF 내보내기 성공 (영역: {REPORT_PRINT_AREA})")
            else:
                logger.error("PDF 파일이 생성을 완료하지 못했거나 0바이트입니다.")
        except Exception:
            logger.exception("PDF 내보내기 중 예외 발생")

        # 2. PNG 이미지 내보내기 (범위 전체 확대 A2:I27 & Chart 안전 삭제)
        try:
            excel.ActiveWindow.ScrollRow = 1
            excel.ActiveWindow.ScrollColumn = 1

            rng = ws.Range(REPORT_IMAGE_RANGE)
            rng.Select()
            time.sleep(1)

            rng.CopyPicture(1, 2)

            chart_shape = ws.Shapes.AddChart2()
            chart_shape.Select()
            excel.Selection.Width = rng.Width
            excel.Selection.Height = rng.Height
            chart_shape.Chart.Paste()
            time.sleep(1)

            chart_shape.Chart.Export(norm_png_path)

            if os.path.exists(norm_png_path) and os.path.getsize(norm_png_path) > 0:
                result.png_success = True
                logger.info(f"PNG 내보내기 성공 (영역: {REPORT_IMAGE_RANGE})")
            else:
                logger.error("PNG 파일이 생성을 완료하지 못했거나 0바이트입니다.")
        except Exception:
            logger.exception("PNG 내보내기 중 예외 발생")
        finally:
            if chart_shape is not None:
                try:
                    chart_shape.Delete()
                except Exception:
                    logger.exception("임시 Chart Shape 삭제 중 예외 발생")

        if not result.is_full_success and not result.pdf_success:
            result.error_message = "PDF/PNG 내보내기 중 오류가 발생했습니다."

        wb.Close(SaveChanges=False)
        wb = None
        return result

    except Exception:
        logger.exception("PDF/PNG 내보내기 종합 예외 발생")
        result.error_message = "PDF/PNG 내보내기 중 오류가 발생했습니다."
        return result
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
