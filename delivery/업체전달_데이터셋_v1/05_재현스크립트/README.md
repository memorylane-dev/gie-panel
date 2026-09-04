# 재현 스크립트

이 데이터셋을 생성한 스크립트입니다. 3주기 이후 데이터 갱신 시 그대로 재사용합니다.

## 실행
```bash
uv venv --python 3.12 .venv
uv pip install --python .venv/bin/python pandas pyreadstat openpyxl
.venv/bin/python build.py
```

## 파일
- `common.py`   — 원본 폴더 탐색 (macOS 한글 파일명 NFD 정규화 처리 포함)
- `taxonomy.py` — 대분류·소주제 매핑, 하위척도 분리, 역문항, 결측 보정 규칙, 데이터 이슈 목록
- `build.py`    — 정제·집계·산출 전 과정

## 신규 주기 추가 절차
1. `build.py`의 `WAVES`에 `("W3", 3, 2029, "3주기")` 추가
2. `SRC`에 3주기 SPSS 파일 경로 추가
3. `taxonomy.py` 및 변수매칭 시트에 3주기 변수명 컬럼 추가 후 `dim` 생성부 확장
4. 재실행 — 집계·헤드라인·분포표가 자동 확장됨

## 입력 원본
- `시계열 시각화 데이터 클리닝/데이터_1주기/*.sav`
- `시계열 시각화 데이터 클리닝/데이터_2주기/*.sav`
- `시계열 시각화 데이터 클리닝/경기학교교육실태조사_12주기_변수매칭_최종.xlsx` (시트 `6_최종연계목록`)
