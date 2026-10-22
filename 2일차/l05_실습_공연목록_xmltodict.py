# =====================================================================
# l05 · API 특강 2일차 · 실습 예시 정답 ② xmltodict
# =====================================================================
# 과제
#   get_performances(시작일, 종료일, 장르)를 만듭니다.
#   - 받는 값   : start, end ("YYYYMMDD"), genre (안 주면 None)
#   - 요청      : 서울(signgucode 11), rows 100
#   - 장르      : genre가 있을 때만 shcate 파라미터에 넣음
#   - 돌려주는 값: DataFrame · 공연명, 기간, 공연장, 장르
#   - 정렬      : 공연 시작일이 이른 것부터
#   - 실패하면   : 빈 표 pd.DataFrame()
#
#   l04 노트북의 KOPIS 요청을 l03 get_boxoffice처럼 함수로 묶은 것입니다.
#
# 준비: .env에 KOPIS_KEY
#       pip install requests pandas openpyxl python-dotenv xmltodict
# 실행: python l05_실습_공연목록_xmltodict.py
# 결과: 서울_주말_공연.xlsx 생성
#       l07_실습_주말추천표.py가 이 파일의 get_performances를 불러다 씀
# =====================================================================

import os
from datetime import date, timedelta
import requests
import xmltodict                  # XML 글자 → 딕셔너리
import pandas as pd
from dotenv import load_dotenv

os.chdir(os.path.dirname(os.path.abspath(__file__)))  # 이 파일이 있는 폴더에서 .env를 읽고 엑셀을 저장

load_dotenv()
KOPIS_KEY = os.getenv("KOPIS_KEY")
URL = ("http://www.kopis.or.kr"
       "/openApi/restful/pblprfr")   # 공연 목록 주소

# 장르 코드 (KOPIS 문서) · 도전 과제에서 이름으로 고를 때 씀
GENRES = {"연극": "AAAA", "뮤지컬": "GGGA", "클래식": "CCCA"}


def get_performances(start, end, genre=None):
    """start~end 기간의 서울 공연 표를 돌려줌. genre는 장르 코드(예: "GGGA")."""
    #   genre=None : 부를 때 장르를 안 주면 자동으로 None이 됨
    #     get_performances("20261024", "20261025")           → 전체 장르
    #     get_performances("20261024", "20261025", "GGGA")   → 뮤지컬만

    # ① 주소 만들기
    params = {"service": KOPIS_KEY,          # KOPIS는 키 이름이 service
              "stdate": start, "eddate": end,
              "cpage": 1, "rows": 100,       # 첫 페이지, 최대 100개
              "signgucode": "11"}            # 11 = 서울
    if genre is not None:                    # 장르를 줬을 때만
        params["shcate"] = genre             # 딕셔너리에 새 짝 추가

    try:
        # ② 보내기
        res = requests.get(URL, params=params, timeout=5)
        # ③ 확인하기
        res.raise_for_status()
    except requests.exceptions.RequestException:
        return pd.DataFrame()

    # ④ 꺼내기 : XML → 딕셔너리
    #   res.content(바이트)를 넣으면 XML에 적힌 UTF-8로 읽어 한글이 깨지지 않음
    data = xmltodict.parse(res.content)
    if not data["dbs"]:                      # 결과가 0건이면 <dbs/>만 와서 None
        return pd.DataFrame()
    items = data["dbs"]["db"]                # <dbs> 안의 <db>(공연) 목록
    if isinstance(items, dict):              # 한 건이면 딕셔너리로 오므로
        items = [items]                      # 리스트로 감싸기
    if "returncode" in items[0]:             # 키가 틀리면 공연 대신 오류 안내가 옴
        print("KOPIS 오류:", items[0]["errmsg"])
        return pd.DataFrame()

    # 표로 정리
    df = pd.DataFrame(items)                                   # 딕셔너리 하나가 한 행
    df = df.sort_values("prfpdfrom")                           # 시작일이 이른 것부터 ("2026.10.01" 글자도 순서대로)
    df["기간"] = df["prfpdfrom"] + " ~ " + df["prfpdto"]        # 두 열을 글자로 이어 새 열
    df = df.rename(columns={"prfnm": "공연명",                  # 영어 열 이름 → 한글
                            "fcltynm": "공연장",
                            "genrenm": "장르"})
    return df[["공연명", "기간", "공연장", "장르"]]              # 네 열만 골라 돌려줌


# 이 파일을 직접 실행했을 때만 실행 (l07이 import할 때는 실행 안 됨)
if __name__ == "__main__":
    # 이번 토요일 : 토요일 번호 5에서 오늘 요일 번호를 뺀 만큼 더함 (% 7 → 일요일이면 6일 뒤)
    today = date.today()
    sat = today + timedelta(days=(5 - today.weekday()) % 7)
    sun = sat + timedelta(days=1)
    print("이번 주말:", sat, "~", sun)

    df = get_performances(sat.strftime("%Y%m%d"), sun.strftime("%Y%m%d"))

    # 도전 과제 ① : 장르를 골라 받으려면 위 줄 대신 아래 세 줄을 쓰세요
    # name = input("장르(연극/뮤지컬/클래식): ")
    # code = GENRES.get(name)          # 없는 이름이면 None → 전체 장르
    # df = get_performances(sat.strftime("%Y%m%d"), sun.strftime("%Y%m%d"), code)

    if len(df) == 0:
        print("받은 공연이 없습니다. .env의 KOPIS_KEY를 확인하세요.")
    else:
        print(df.head())
    df.to_excel("서울_주말_공연.xlsx", index=False)
    print(f"\n{len(df)}개 공연 저장 완료 → 서울_주말_공연.xlsx")
