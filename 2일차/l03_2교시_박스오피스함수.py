# =====================================================================
# l03 · API 특강 2일차 · 2교시 · get_boxoffice 함수
# =====================================================================
# 하는 일
#   l02 노트북에서 한 단계씩 해 본 KOBIS 요청과 표 정리를
#   get_boxoffice(날짜) 함수 하나로 묶었습니다.
#   날짜만 넣으면 그 주말 박스오피스 표(DataFrame)가 돌아옵니다.
#   실패하면 빈 표가 돌아오므로 뒤의 코드가 멈추지 않습니다.
#
#   1일차 get_summary와 같은 구조입니다.
#     get_summary(제목)   → 요약문(글자),  실패하면 "설명 없음"
#     get_boxoffice(날짜) → 영화 표(DataFrame), 실패하면 빈 표
#
# 준비: .env에 KOBIS_KEY (l01 참고)
#       pip install requests pandas openpyxl python-dotenv
# 실행: python l03_2교시_박스오피스함수.py
# 결과: 지난 주말 상위 영화를 출력하고, 관객 10만 이상을 주말_박스오피스.xlsx로 저장
#       l07_실습_주말추천표.py가 이 파일의 get_boxoffice를 불러다 씀
# =====================================================================

import os
from datetime import date, timedelta
import requests
import pandas as pd
from dotenv import load_dotenv

os.chdir(os.path.dirname(os.path.abspath(__file__)))  # 이 파일이 있는 폴더에서 .env를 읽고 엑셀을 저장

# ---------------------------------------------------------------------
# 설정값 (대문자 = 바뀌지 않는 값)
# ---------------------------------------------------------------------
load_dotenv()                            # .env 파일 읽기
KOBIS_KEY = os.getenv("KOBIS_KEY")       # 키는 코드에 적지 않고 .env에서 꺼냄
URL = ("http://www.kobis.or.kr/kobisopenapi/webservice/rest"
       "/boxoffice/searchWeeklyBoxOfficeList.json")   # 주간/주말 박스오피스 주소

# 영어 열 이름 → 한글 열 이름 (rename에 넘길 딕셔너리)
NAMES = {"rank": "순위", "movieNm": "영화명",
         "openDt": "개봉일", "audiCnt": "관객수"}


def get_boxoffice(target_dt):
    """target_dt("YYYYMMDD")가 속한 주말(금~일)의 박스오피스 표를 돌려줌."""
    # ① 주소 만들기 : 1일차와 같은 params에 key 한 줄이 늘었을 뿐
    params = {"key": KOBIS_KEY,          # 발급받은 키
              "targetDt": target_dt,     # 조회 날짜 (지난 일요일)
              "weekGb": "1"}             # 1 = 주말(금~일)
    try:
        # ② 보내기
        res = requests.get(URL, params=params, timeout=5)
        # ③ 확인하기 (1) : 4xx·5xx면 예외 → except
        res.raise_for_status()
    except requests.exceptions.RequestException:
        return pd.DataFrame()            # 실패하면 빈 표

    data = res.json()                    # JSON → 딕셔너리

    # ③ 확인하기 (2) : 키가 틀리면 상태 코드는 200인데 내용이 faultInfo 안내
    if "faultInfo" in data:
        print("KOBIS 오류:", data["faultInfo"]["message"])
        return pd.DataFrame()

    # ④ 꺼내기 : data → boxOfficeResult → weeklyBoxOfficeList (딕셔너리가 든 리스트)
    movies = data["boxOfficeResult"]["weeklyBoxOfficeList"]

    # 표로 정리
    df = pd.DataFrame(movies)                              # 리스트 → 표 (열 18개쯤)
    df = df[["rank", "movieNm", "openDt", "audiCnt"]]      # 필요한 4개 열만 (대괄호 두 겹)
    df = df.rename(columns=NAMES)                          # 열 이름을 한글로
    df["순위"] = df["순위"].astype(int)                     # "1" → 1 (글자 → 숫자)
    df["관객수"] = df["관객수"].astype(int)                 # 숫자로 바꿔야 크기 비교·정렬이 맞음
    return df


# ---------------------------------------------------------------------
# 이 파일을 직접 실행했을 때만 아래가 실행됨
#   다른 파일(l07)이 import로 get_boxoffice만 가져갈 때는 실행되지 않음
# ---------------------------------------------------------------------
if __name__ == "__main__":
    # 지난 일요일 : 오늘에서 (요일 번호 + 1)일을 빼면 됨. 월 0, 화 1 … 일 6
    today = date.today()
    last_sunday = today - timedelta(days=today.weekday() + 1)
    target_dt = last_sunday.strftime("%Y%m%d")   # 날짜 → "20261018" 모양의 글자
    print("조회 날짜(지난 일요일):", target_dt)

    df = get_boxoffice(target_dt)
    if len(df) == 0:
        print("받은 영화가 없습니다. .env의 KOBIS_KEY를 확인하세요.")
    else:
        print(df.head(5))

        # 거르기 : 관객수가 10만 이상인 행만 남김
        #   df["관객수"] >= 100000  → 행마다 True/False
        #   df[그 True/False]       → True인 행만 남긴 새 표
        big = df[df["관객수"] >= 100000]
        # 정렬 : 최근 개봉작부터 ("2026-10-14" 같은 글자도 날짜 순서대로 정렬됨)
        big = big.sort_values("개봉일", ascending=False)
        big.to_excel("주말_박스오피스.xlsx", index=False)
        print(f"\n관객 10만 이상 {len(big)}편 저장 완료 → 주말_박스오피스.xlsx")
