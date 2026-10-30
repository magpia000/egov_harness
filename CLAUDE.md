# EGDI 순위 변동 보고서 하네스

UN 전자정부발전지수(EGDI) 6개 연도(2014~2024, 격년) 데이터를 하나의 엑셀 보고서로 취합한다.
보고서는 세 가지 질문에 답해야 한다.

1. 국가별로 2년마다 순위가 몇 계단씩 바뀌었는가
2. 10년간 가장 성장한 국가와 가장 후퇴한 국가는 어디인가
3. 2024년 순위를 기준으로 고정했을 때 과거 순위는 어떻게 변해왔는가

## 폴더 구조

```
input_data/     EGOV_DATA_2014.csv ~ EGOV_DATA_2024.csv  (원본, 절대 수정 금지)
template/       EGOV_DATA_template.xlsx                  (출력 양식)
output/         EGOV_DATA_취합.xlsx                       (결과물)
scripts/        build_report.py (생성)
evals/          verify.py (검증)
```

## 입력 데이터

- 연도별 CSV 1개, 각 193개국 × 8열
- 열: Survey Year, Country Name, E-Government Rank, E-Government Index, E-Participation Index, Online Service Index, Human Capital Index, Telecommunication Infrastructure Index
- 알려진 문제
  - 국가명 끝에 공백이 붙어 있다 (`"Denmark "`). 매칭 전 반드시 `strip()` 한다.
  - 2020 파일은 순위순 정렬이 아니고, 국가명에 따옴표가 있으며, 소수 자릿수가 다르다.
  - 2018, 2022에 동순위가 있다. 원본 순위를 그대로 쓰고 재계산하지 않는다.
  - 2024 파일에는 동순위가 없다(193개국 순위 전부 고유, 확인됨). 따라서 "행 순서는 2024 순위 오름차순" 규칙에 동순위 타이브레이크가 필요한 상황은 발생하지 않는다. 향후 데이터 갱신으로 2024 동순위가 생기면 정렬 기준을 먼저 정하고 진행한다.

## 출력 사양

시트명 `EGOV_DATA_취합`. 헤더는 2~3행 2단 병합, 데이터는 4~196행(193개국).
행 순서는 2024 순위 오름차순으로 고정한다.
1행은 사용하지 않고 비워 둔다(템플릿 기준, 제목·구분용이 아닌 여백).

| 열 | 2행 헤더 (병합) | 3행 헤더 | 값 |
|---|---|---|---|
| B | Country Name | | 공백 제거한 국가명 |
| C | A Decade | | 2014 순위 − 2024 순위 |
| D~I | E-Government Rank | 2024, 2022, 2020, 2018, 2016, 2014 | 각 연도 순위 |
| J~N | Rank Change (2-year) | 22→24, 20→22, 18→20, 16→18, 14→16 | 이전 연도 순위 − 다음 연도 순위 |
| O~S | 지수 5종 (각각 2~3행 병합) | | 2024년 값만 |

O~S 순서: E-Government Index, E-Participation Index, Online Service Index, Human Capital Index, Telecommunication Infrastructure Index

원본 양식은 지수가 J~N에 있었으나, 가독성을 위해 변동 블록을 순위 옆(J~N)에 두고 지수를 O~S로 옮겼다.

### 표현 규칙

- 변동값 부호: 양수 = 순위 상승, 음수 = 하락, 0 = 변동 없음
- C, J~N은 숫자로 저장하고 표시 형식만 `"▲"0;"▼"0;0` 로 바꾼다 (변동 없음은 `-` 대신 `0`으로 표시 — 빈 셀처럼 오해되는 것을 방지)
- C열: 0 기준 양방향 데이터 막대 (상승 초록, 하락 빨강)
- J~N: 글자색만 적용 (▲ 초록, ▼ 빨강, 0은 회색). 배경색은 쓰지 않는다
- D~I: 서식 없이 숫자만
- 틀 고정 C4, 3행 헤더에 자동 필터

## 작업 규칙

- 순위·변동값을 대화 중에 직접 계산해 답하지 않는다. 모든 값은 `scripts/build_report.py`가 만든다.
- 국가 매칭은 행 위치가 아니라 공백 제거한 국가명으로 한다.
- `build_report.py`를 실행하거나 수정한 뒤에는 반드시 `scripts/verify.py`를 실행한다.
- 검증이 하나라도 실패하면 완료로 보고하지 않고, 원인을 고친 뒤 다시 검증한다.
- 출력 사양을 바꿔야 할 것 같으면 임의로 바꾸지 말고 먼저 묻는다.

## 검증 기준 (verify.py)

1. 데이터 행이 정확히 193개이고 국가명 중복이 없다
2. D~I의 모든 순위가 해당 연도 원본 CSV와 일치한다
3. 행이 2024 순위 오름차순이다
4. C = I − D, J~N = 인접 연도 순위 차이와 일치한다
5. O~S가 2024 원본 지수와 일치한다
6. 빈 셀이 없다
7. 결과 요약: C열 최댓값 국가(최대 성장), 최솟값 국가(최대 후퇴)를 출력한다

## 실행

```
python scripts/build_report.py
python evals/verify.py
```