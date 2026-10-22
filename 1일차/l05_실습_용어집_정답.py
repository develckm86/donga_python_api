# =====================================================================
# l05 · API 특강 1일차 · 실습 예시 정답 · 영화·공연 용어집
# =====================================================================
# 하는 일
#   영화공연용어.xlsx의 용어 15개를 하나씩 위키백과에 요청해서
#   "설명"(요약문)과 "글자수" 열을 붙이고, 글자수가 큰 순서로 정렬해
#   영화공연_용어집.xlsx로 저장합니다.
#
# 흐름
#   [1부] get_summary 함수  : 요청 4단계 (l03에서 만든 함수 그대로)
#   [2부] pandas            : 엑셀 읽기 → 반복 → 열 붙이기 → 정렬 → 저장
#
# 준비: pip install requests pandas openpyxl
# 실행: python l05_실습_용어집_정답.py   (VS Code에서는 오른쪽 위 ▷ 버튼)
# 결과: 같은 폴더에 영화공연_용어집.xlsx 생성
# =====================================================================

import os           # 파일 경로를 다루는 기본 도구
import time         # 요청 사이에 쉬는 시간(sleep)을 줄 때 씀
import requests     # 서버에 요청을 보내는 라이브러리 (파이썬용 브라우저)
import pandas as pd # 표(DataFrame)와 엑셀을 다루는 라이브러리

# 터미널이 어느 폴더에 있든, 이 .py 파일이 있는 폴더에서 엑셀을 읽고 저장하도록 위치를 옮김
#   __file__ : 지금 실행 중인 이 파일의 경로
os.chdir(os.path.dirname(os.path.abspath(__file__)))


# ---------------------------------------------------------------------
# [1부] get_summary 함수 · 요청 4단계를 하나로 묶음
# ---------------------------------------------------------------------
# 대문자 변수 = 프로그램 내내 바뀌지 않는 설정값이라는 표시 (파이썬 관례)
BASE = "https://ko.wikipedia.org/api/rest_v1/page/summary/"   # 이 뒤에 문서 제목을 붙임
HEADERS = {"User-Agent": "APIClass/1.0 (me@example.com)"}     # 내 이름표. 위키백과는 이름표를 요구함

def get_summary(title):
    """제목 하나를 받아 위키백과 요약문을 돌려줌. 실패하면 "설명 없음"."""
    # ① 주소 만들기 : 위키백과 주소는 띄어쓰기 대신 _ 를 씀
    url = BASE + title.replace(" ", "_")
    try:
        # ② 보내기 : timeout=5 → 5초 안에 응답이 없으면 기다리지 않고 except로
        res = requests.get(url, headers=HEADERS, timeout=5)
        # ③ 확인하기 : 404(없는 문서) 같은 4xx·5xx면 예외를 일으켜 except로
        res.raise_for_status()
    except requests.exceptions.RequestException:
        # 없는 문서, 시간 초과, 인터넷 끊김… 어떤 실패든 여기로 모임
        return "설명 없음"
    # ④ 꺼내기 : JSON → 딕셔너리로 바꾼 뒤 "extract"(요약문)를 꺼냄
    #   .get(키, 기본값) : 키가 없으면 오류 대신 기본값을 돌려줌
    return res.json().get("extract", "설명 없음")


# ---------------------------------------------------------------------
# [2부] pandas · 읽기 → 반복 → 열 붙이기 → 정렬 → 저장
# ---------------------------------------------------------------------
# ① 엑셀 읽기 : 첫 줄(머리글)이 열 이름이 됨. 여기서는 "용어" 열 하나
df = pd.read_excel("영화공연용어.xlsx")
print(f"용어 {len(df)}개를 읽었습니다. 위키백과에 요청을 시작합니다...\n")

# ②③ 반복하며 두 리스트 채우기
summaries = []   # 요약문을 모을 빈 리스트
lengths = []     # 글자 수를 모을 빈 리스트
for term in df["용어"]:            # "용어" 열의 값을 위에서부터 하나씩 꺼냄
    text = get_summary(term)       # 요청 4단계를 한 번 실행
    summaries.append(text)         # 요약문을 리스트 끝에 추가
    lengths.append(len(text))      # len(글자) = 글자 수
    print(f"{term:<12} → {len(text)}글자")
    time.sleep(0.5)                # 서버에 부담을 주지 않도록 0.5초 쉼

# ④ 열 붙이기 : 리스트 길이(15) = 표의 행 수(15) 여야 함
df["설명"] = summaries
df["글자수"] = lengths

# ⑤ 정렬 : 글자수가 큰 것부터. sort_values는 새 표를 돌려주므로 df = 로 '다시 담기'
df = df.sort_values("글자수", ascending=False)

# ⑥ 저장 : index=False → 왼쪽 번호 열(0, 1, 2…) 없이 저장
df.to_excel("영화공연_용어집.xlsx", index=False)
print(f"\n{len(df)}개 용어 저장 완료 → 영화공연_용어집.xlsx")
