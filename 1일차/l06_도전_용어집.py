# =====================================================================
# l06 · API 특강 1일차 · 도전 과제 · 용어집 업그레이드
# =====================================================================
# l05 정답에 두 가지를 더합니다.
#   도전 ① 요약은 첫 문장만    → first_sentence 함수
#   도전 ② 맨 위에 오늘 서울 기온 → 2교시 날씨 4단계 + pd.concat
#
# 실행: python l06_도전_용어집.py
# 결과: 같은 폴더에 영화공연_용어집_도전.xlsx 생성
# =====================================================================

import os
import time
import requests
import pandas as pd

os.chdir(os.path.dirname(os.path.abspath(__file__)))  # 이 파일이 있는 폴더에서 읽고 저장

BASE = "https://ko.wikipedia.org/api/rest_v1/page/summary/"
HEADERS = {"User-Agent": "APIClass/1.0 (me@example.com)"}

def get_summary(title):
    """l05와 같은 함수. 위키백과 요약문, 실패하면 "설명 없음"."""
    url = BASE + title.replace(" ", "_")
    try:
        res = requests.get(url, headers=HEADERS, timeout=5)
        res.raise_for_status()
    except requests.exceptions.RequestException:
        return "설명 없음"
    return res.json().get("extract", "설명 없음")


# ---------------------------------------------------------------------
# 도전 ① 첫 문장만 남기기
# ---------------------------------------------------------------------
def first_sentence(text):
    """". "(마침표+띄어쓰기) 앞까지만 잘라 마침표를 다시 붙임."""
    if ". " in text:                         # 문장이 두 개 이상이면
        # split(". ") : ". "를 기준으로 잘라 리스트로 → [첫 문장, 둘째 문장, …]
        return text.split(". ")[0] + "."     # 첫 번째 조각 + 잘려 나간 마침표
    return text                              # 한 문장뿐이면 그대로

# 예: first_sentence("뮤지컬은 공연이다. 영국에서 발전했다.") → "뮤지컬은 공연이다."


# ---------------------------------------------------------------------
# 도전 ② 오늘 서울 기온 받기 (1일차 2교시의 4단계 그대로)
# ---------------------------------------------------------------------
W_URL = "https://api.open-meteo.com/v1/forecast"                    # ① 주소 만들기
params = {"latitude": 37.5665, "longitude": 126.978,
          "current": "temperature_2m", "timezone": "Asia/Seoul"}
res = requests.get(W_URL, params=params, timeout=5)                # ② 보내기
if res.status_code == 200:                                          # ③ 확인하기
    temp = res.json()["current"]["temperature_2m"]                  # ④ 꺼내기
else:
    temp = "?"   # 실패해도 프로그램은 계속
print("오늘 서울 기온:", temp, "°C")


# ---------------------------------------------------------------------
# 용어집 만들기 (l05와 같고, get_summary 한 줄만 first_sentence로 감쌈)
# ---------------------------------------------------------------------
df = pd.read_excel("영화공연용어.xlsx")
summaries = []
lengths = []
for term in df["용어"]:
    text = first_sentence(get_summary(term))   # ← 요약을 받아 첫 문장만
    summaries.append(text)
    lengths.append(len(text))
    time.sleep(0.5)
df["설명"] = summaries
df["글자수"] = lengths
df = df.sort_values("글자수", ascending=False)

# 기온 행을 표 맨 위에 붙이기
#   행 하나짜리 표를 만들고, pd.concat([위 표, 아래 표])로 위아래로 이어 붙임
#   이 행에는 "글자수"가 없으므로 그 칸은 빈 값(NaN)이 됨
weather = pd.DataFrame([{"용어": "오늘 서울 기온",
                         "설명": f"현재 기온 {temp}°C"}])
df = pd.concat([weather, df], ignore_index=True)   # ignore_index=True : 행 번호를 0부터 새로 매김

df.to_excel("영화공연_용어집_도전.xlsx", index=False)
print(df.head())
print(f"\n{len(df)}행 저장 완료 → 영화공연_용어집_도전.xlsx")
