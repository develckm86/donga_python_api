# API 특강 · 수업 파일 안내

파일 이름 앞의 `l01`, `l02`…가 **수업 순서**입니다. 날마다 l01부터 차례로 엽니다.

- `.ipynb` 노트북 · VS Code에서 열고 셀을 하나씩 **Shift+Enter**로 실행
- `.py` 파일 · 오른쪽 위 ▷ 버튼 또는 터미널에서 `python 파일이름.py`

## 처음 한 번 · 설치

```bash
pip install -r requirements.txt
```

노트북을 열면 오른쪽 위 **커널**을 이 폴더의 `.venv`(또는 수업용 가상환경)로 고릅니다.

## 1일차 · API로 데이터 받아 오기 (키 필요 없음)

| 순서 | 파일 | 교시 | 내용 |
|---|---|---|---|
| l01 | `l01_1교시_JSON기초.ipynb` | 1교시 | 요청과 응답, API, URL 읽기, JSON → 딕셔너리 |
| l02 | `l02_2교시_날씨요청.ipynb` | 2교시 | requests, 응답 객체, 상태 코드, params, **요청 4단계** |
| l03 | `l03_3교시_위키백과.ipynb` | 3교시 | User-Agent, try·except 오류 처리, `get_summary` 함수 |
| l04 | `l04_실습_pandas정리.ipynb` | 실습 | DataFrame, 열 붙이기, 엑셀 저장·읽기, 정렬, 실습 과제 |
| l05 | `l05_실습_용어집_정답.py` | 실습 | 실습 예시 정답 → `영화공연_용어집.xlsx` |
| l06 | `l06_도전_용어집.py` | 도전 | 첫 문장만, 맨 위에 오늘 기온 → `영화공연_용어집_도전.xlsx` |
| - | `영화공연용어.xlsx` | 실습 | 실습 재료 (용어 15개) |

## 2일차 · 주말 문화생활 추천표 (키 필요)

**시작 전**: `2일차/.env.example`을 복사해 이름을 `.env`로 바꾸고 KOBIS·KOPIS 키를 적습니다. (자세한 방법은 l01)

| 순서 | 파일 | 교시 | 내용 |
|---|---|---|---|
| l01 | `l01_1교시_API키와env.ipynb` | 1교시 | API 키, 공공데이터, `.env`로 키 숨기기, 키 오류 |
| l02 | `l02_2교시_KOBIS영화.ipynb` | 2교시 | 날짜 계산, KOBIS 요청, 열 고르기·rename·astype, 거르기 |
| l03 | `l03_2교시_박스오피스함수.py` | 2교시 | `get_boxoffice(날짜)` 함수 |
| l04 | `l04_3교시_XML과KOPIS.ipynb` | 3교시 | XML 두 방법(ElementTree, xmltodict), 이번 주말 날짜, KOPIS 요청 |
| l05 | `l05_실습_공연목록_xmltodict.py` | 실습 | 과제 예시 정답 `get_performances` (xmltodict) |
| l06 | `l06_실습_공연목록_ElementTree.py` | 실습 | 같은 함수를 ElementTree로 (l05와 비교) |
| l07 | `l07_실습_주말추천표.py` | 실습 | **최종 결과물** → `주말_문화생활_추천표.xlsx` (시트 3개) |

l07은 l03과 l05의 함수를 불러다 쓰므로 세 파일이 같은 폴더에 있어야 합니다.

## 요청 4단계 (이틀 내내 같은 모양)

| 단계 | 1일차 | 2일차에 더한 것 |
|---|---|---|
| ① 주소 만들기 | `url` + `params` 딕셔너리 | 키 (`.env` → `os.getenv`), 날짜 (`datetime`) |
| ② 보내기 | `requests.get(url, params=params, timeout=5)` | 그대로 |
| ③ 확인하기 | `raise_for_status()` + try·except | 응답 안의 키 오류 (`faultInfo`, `returncode`) |
| ④ 꺼내기 | `res.json()` → 딕셔너리 | XML → ElementTree 또는 xmltodict |
"# donga_python_api" 
