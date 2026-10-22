# =====================================================================
# l06 · API 특강 2일차 · 실습 예시 정답 ① ElementTree
# =====================================================================
# l05(xmltodict)와 같은 함수를 ElementTree로 만들었습니다.
# 다른 곳은 딱 두 군데입니다.  ★ 표시를 찾아 l05와 비교해 보세요.
#   ★1  import 한 줄
#   ★2  ④ 꺼내기 (XML 읽기) 부분 · isinstance가 필요 없음
#
# ElementTree는 파이썬 기본 도구라 설치가 필요 없습니다.
#
# 실행: python l06_실습_공연목록_ElementTree.py
# 결과: 서울_주말_공연_et.xlsx 생성 (l05 결과와 같은 표)
# =====================================================================

import os
from datetime import date, timedelta
import requests
import xml.etree.ElementTree as ET     # ★1 xmltodict 대신
import pandas as pd
from dotenv import load_dotenv

os.chdir(os.path.dirname(os.path.abspath(__file__)))  # 이 파일이 있는 폴더에서 .env를 읽고 엑셀을 저장

load_dotenv()
KOPIS_KEY = os.getenv("KOPIS_KEY")
URL = ("http://www.kopis.or.kr"
       "/openApi/restful/pblprfr")


def get_performances(start, end, genre=None):
    """l05와 같은 함수. XML을 ElementTree로 읽음."""
    # ① 주소 만들기 ~ ③ 확인하기 : l05와 같음
    params = {"service": KOPIS_KEY,
              "stdate": start, "eddate": end,
              "cpage": 1, "rows": 100,
              "signgucode": "11"}
    if genre is not None:
        params["shcate"] = genre
    try:
        res = requests.get(URL, params=params, timeout=5)
        res.raise_for_status()
    except requests.exceptions.RequestException:
        return pd.DataFrame()

    # ★2 ④ 꺼내기 : 여기부터 다름 --------------------------------
    root = ET.fromstring(res.content)          # XML 글자 → 나무 구조. root는 맨 바깥 <dbs>
    items = []
    for db in root.findall("db"):              # <db>(공연 하나)를 모두 찾아 하나씩
        # <db> 안의 태그들로 딕셔너리 하나 만들기
        #   c.tag  = 태그 이름 (예: "prfnm")
        #   c.text = 태그 내용 (예: "공연 A")
        items.append({c.tag: c.text for c in db})
    # findall은 0건이어도, 1건이어도 항상 리스트 → isinstance가 필요 없음
    # ------------------------------------------------------------

    if len(items) == 0:                        # 결과 0건
        return pd.DataFrame()
    if "returncode" in items[0]:               # 키 오류 안내
        print("KOPIS 오류:", items[0]["errmsg"])
        return pd.DataFrame()

    # 표로 정리 : l05와 같음
    df = pd.DataFrame(items)
    df = df.sort_values("prfpdfrom")
    df["기간"] = df["prfpdfrom"] + " ~ " + df["prfpdto"]
    df = df.rename(columns={"prfnm": "공연명",
                            "fcltynm": "공연장",
                            "genrenm": "장르"})
    return df[["공연명", "기간", "공연장", "장르"]]


if __name__ == "__main__":
    today = date.today()
    sat = today + timedelta(days=(5 - today.weekday()) % 7)
    sun = sat + timedelta(days=1)

    df = get_performances(sat.strftime("%Y%m%d"), sun.strftime("%Y%m%d"))
    df.to_excel("서울_주말_공연_et.xlsx", index=False)
    print(f"{len(df)}개 공연 저장 완료 → 서울_주말_공연_et.xlsx")
