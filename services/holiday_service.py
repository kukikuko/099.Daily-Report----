import os
import json
import urllib.request
import urllib.parse
from datetime import datetime, date, timedelta
from utils.logger_utils import logger

def get_holidays_this_week() -> dict:
    """공공데이터포럼 API를 호출하여 이번 주(월~금) 내의 공휴일 정보를 조회한다.

    반환값:
        dict: { date_object: "공휴일명" } 형태의 딕셔너리
    """
    api_key = os.getenv("HOLIDAY_API_KEY", "").strip()
    if not api_key:
        logger.info("공휴일 API 키가 설정되지 않아 공휴일 조회를 생략합니다.")
        return {}

    today = date.today()
    monday = today - timedelta(days=today.weekday())

    target_dates = [monday + timedelta(days=i) for i in range(5)]
    year_months = set((d.year, d.month) for d in target_dates)

    holidays = {}

    for year, month in year_months:
        url = (
            "http://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo"
            f"?serviceKey={urllib.parse.quote_plus(api_key)}"
            f"&solYear={year}"
            f"&solMonth={month:02d}"
            "&_type=json"
            "&numOfRows=30"
        )
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=5) as response:
                if response.status != 200:
                    logger.warning(f"공휴일 API 응답 실패 (상태 코드: {response.status})")
                    continue
                data = json.loads(response.read().decode("utf-8"))

            header = data.get("response", {}).get("header", {})
            if header.get("resultCode") != "00":
                logger.warning(f"공휴일 API 오류: {header.get('resultMsg')}")
                continue

            items = data.get("response", {}).get("body", {}).get("items", {})
            if not items:
                continue

            item_list = items.get("item", [])
            if isinstance(item_list, dict):
                item_list = [item_list]

            for item in item_list:
                if item.get("isHoliday") == "Y":
                    locdate_str = str(item.get("locdate", ""))
                    date_name = item.get("dateName", "공휴일")
                    try:
                        h_date = datetime.strptime(locdate_str, "%Y%m%d").date()
                        if h_date in target_dates:
                            holidays[h_date] = date_name
                            logger.info(f"이번 주 공휴일 감지: {h_date} ({date_name})")
                    except ValueError:
                        pass

        except Exception:
            logger.exception(f"공휴일 API 조회 중 예외 발생 ({year}년 {month}월)")

    return holidays
