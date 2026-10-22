# =====================================================================
# l07 · API 특강 2일차 · 최종 결과물 · 주말 문화생활 추천표
# =====================================================================
# 하는 일
#   API 세 개에서 받은 표를 엑셀 한 파일에 시트 세 개로 저장합니다.
#
#   시트   | 어디서                  | 무엇을
#   날씨   | Open-Meteo (키 없음)    | 이번 토·일 서울 최고·최저 기온
#   영화   | KOBIS  (l03 함수)       | 지난 주말 박스오피스 상위 5편
#   공연   | KOPIS  (l05 함수)       | 이번 주말 서울 공연 10개
#
# 준비: .env에 KOBIS_KEY, KOPIS_KEY
#       pip install requests pandas openpyxl python-dotenv xmltodict
# 실행: python l07_실습_주말추천표.py
# 결과: 주말_문화생활_추천표.xlsx
#       키가 없으면 영화·공연 시트는 빈 표로 저장됩니다 (날씨는 키 없이 채워짐)
# =====================================================================

import os
from datetime import date, timedelta
import requests
import pandas as pd

# 앞에서 만든 함수를 다른 파일에서 불러옴
#   from 파일이름(.py 빼고) import 함수이름
#   불러오는 파일의 if __name__ == "__main__": 아래는 실행되지 않음
from l03_2교시_박스오피스함수 import get_boxoffice
from l05_실습_공연목록_xmltodict import get_performances

os.chdir(os.path.dirname(os.path.abspath(__file__)))  # 이 파일이 있는 폴더에 저장


# ---------------------------------------------------------------------
# 날씨 함수 · 1일차 4단계 그대로, 받는 항목만 daily(하루 단위)로
# ---------------------------------------------------------------------
W_URL = "https://api.open-meteo.com/v1/forecast"
DAILY = "temperature_2m_max,temperature_2m_min"   # 하루 최고기온, 최저기온 (쉼표로 여러 개)

def get_weather(sat, sun):
    """sat~sun 기간의 서울 최고·최저 기온 표를 돌려줌."""
    params = {"latitude": 37.5665,           # 서울
              "longitude": 126.978,
              "daily": DAILY,
              "timezone": "Asia/Seoul",
              "start_date": str(sat),        # 날씨 API는 "2026-10-24" 모양 → str(날짜)
              "end_date": str(sun)}
    res = requests.get(W_URL, params=params, timeout=5)
    # 응답의 "daily"는 {"time": [...], "temperature_2m_max": [...], ...}
    #   딕셔너리의 값이 리스트면 키가 열 이름, 리스트가 그 열의 값이 됨
    df = pd.DataFrame(res.json()["daily"])
    return df.rename(columns={
        "time": "날짜",
        "temperature_2m_max": "최고기온",
        "temperature_2m_min": "최저기온"})


# ---------------------------------------------------------------------
# 날짜 계산 · 영화는 '지난' 주말, 공연과 날씨는 '이번' 주말
# ---------------------------------------------------------------------
today = date.today()
sat = today + timedelta(days=(5 - today.weekday()) % 7)        # 이번 토요일
sun = sat + timedelta(days=1)                                   # 이번 일요일
last_sun = today - timedelta(days=today.weekday() + 1)          # 지난 일요일 (박스오피스 집계 완료)
print("이번 주말:", sat, "~", sun, "/ 박스오피스 기준:", last_sun)


# ---------------------------------------------------------------------
# 세 API 요청 · 함수 덕분에 한 줄씩
# ---------------------------------------------------------------------
weather = get_weather(sat, sun)
movies = get_boxoffice(last_sun.strftime("%Y%m%d"))             # KOBIS는 "YYYYMMDD"
shows = get_performances(sat.strftime("%Y%m%d"),
                         sun.strftime("%Y%m%d"))                # KOPIS도 "YYYYMMDD"

movies = movies.head(5)    # 상위 5편만
shows = shows.head(10)     # 10개만


# ---------------------------------------------------------------------
# 시트 세 개로 저장 · pd.ExcelWriter
# ---------------------------------------------------------------------
#   with pd.ExcelWriter(파일) as w:   엑셀 파일 하나를 열어 둠
#       표.to_excel(w, sheet_name=…)   부를 때마다 시트가 하나씩 생김
#   with 블록이 끝나면 저장됨
#
#   주의: to_excel("파일.xlsx")를 세 번 부르면 같은 파일을 덮어써서 마지막 표만 남음
with pd.ExcelWriter("주말_문화생활_추천표.xlsx") as w:
    weather.to_excel(w, sheet_name="날씨", index=False)
    movies.to_excel(w, sheet_name="영화", index=False)
    shows.to_excel(w, sheet_name="공연", index=False)

print(f"추천표 저장 완료 → 주말_문화생활_추천표.xlsx")
print(f"  시트 날씨 {len(weather)}행, 영화 {len(movies)}행, 공연 {len(shows)}행")
