# EGDI 순위 변동 보고서 하네스

UN 전자정부발전지수(EGDI) 6개 연도(2014~2024, 격년) CSV를 하나의 엑셀 보고서로
취합하는 작업 하네스. 이 저장소를 새 환경(새 PC, 새 Claude Code 세션 등)에
그대로 옮겨 놓고 아래 순서대로 따라 하면 동일한 결과를 재현할 수 있다.

이 저장소의 "스펙 원본"은 [CLAUDE.md](CLAUDE.md)다. 이 README는 환경 설정과
재현 절차만 다루며, 출력 셀 단위 사양·표현 규칙·검증 기준 등 상세 내용은
CLAUDE.md를 참조한다.

## 사전 요구 사항

- Python 3.12 이상
- [uv](https://docs.astral.sh/uv/) (권장) — 이미 `pyproject.toml` / `uv.lock`이
  있으므로 가장 빠르게 동일 환경을 재현할 수 있다. uv가 없으면 표준
  `venv` + `pip`로도 가능하다.

## 환경 설정

### 방법 A — uv (권장, 이 프로젝트가 실제로 만들어진 방식)

```bash
uv init --bare --python 3.12 && uv python pin 3.12
uv add openpyxl pandas
```

- `uv init --bare --python 3.12`: `pyproject.toml`만 생성하고(샘플 `main.py`,
  `README.md`, `.gitignore` 등 템플릿 파일은 만들지 않음), 동시에 이 프로젝트가
  쓸 Python 버전을 3.12로 고정한 `.python-version` 파일을 만든다
  (`--no-pin-python`을 주지 않는 한 `--python`에 지정한 버전이 자동으로 pin된다).
- `uv add openpyxl pandas`: 의존성을 `pyproject.toml`에 추가하고, `.python-version`에
  맞춰(필요하면 3.12를 새로 내려받아) `.venv`를 만들어 설치하며 `uv.lock`을
  생성/갱신한다.

이미 `pyproject.toml` / `uv.lock` / `.python-version`이 저장소에 포함된 상태
(이 저장소를 그대로 복제한 경우)라면 위 과정을 새로 할 필요 없이 `uv sync`
한 번으로 동일한 `.venv`(Python 3.12)를 재현할 수 있다.

### 방법 B — venv + pip (uv가 설치 안되어 있는 경우)

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -e .
```

> 현재 `scripts/build_report.py`와 `evals/verify.py`는 `openpyxl`만 사용한다.
> `pandas`는 `pyproject.toml`에 선언만 되어 있고 실제 코드에서는 쓰이지 않는다.

## 폴더 구조

```
input_data/     EGOV_DATA_2014.csv ~ EGOV_DATA_2024.csv  (원본, 절대 수정 금지)
template/       EGOV_DATA_template.xlsx                  (출력 양식)
output/         EGOV_DATA_취합.xlsx                       (결과물, 스크립트가 생성)
scripts/        build_report.py  (보고서 생성 스크립트)
evals/          verify.py        (검증 스크립트)
CLAUDE.md       출력 사양 · 작업 규칙 · 검증 기준 (스펙 원본)
```

`input_data/`와 `template/`은 저장소에 이미 포함되어 있어야 한다(원본 데이터이므로
새로 만들지 않는다). `output/`은 비어 있거나 없어도 되며, 실행 시 자동 생성된다.

## 실행

```bash
# 1. 보고서 생성
python scripts/build_report.py

# 2. 검증 (반드시 1번 다음에 실행)
python evals/verify.py
```

uv를 썼다면 `uv run python scripts/build_report.py` / `uv run python evals/verify.py`
로 실행해도 된다. Windows에서 가상환경을 activate하지 않았다면
`.venv\Scripts\python.exe scripts\build_report.py` 식으로 venv의 python을
직접 지정한다.

`verify.py`가 하나라도 실패하면 완료된 상태가 아니다. 원인을 고치고
`build_report.py`부터 다시 실행한 뒤 재검증한다 (CLAUDE.md의 작업 규칙 참고).

### 자주 만나는 오류

- **`PermissionError: ... EGOV_DATA_취합.xlsx`**: 결과 파일을 Excel에서 열어 둔
  상태에서 `build_report.py`를 실행하면 저장이 거부된다. Excel에서 파일을 닫고
  다시 실행한다.
- **콘솔에 한글/▲▼→ 문자가 깨져 보임**: Windows 콘솔 코드페이지(cp949) 때문에
  터미널 출력만 깨지는 것이며, 저장된 xlsx 파일의 실제 데이터는 UTF-8/유니코드로
  정상이다. `build_report.py`, `evals/verify.py` 모두 시작 시
  `sys.stdout.reconfigure(encoding="utf-8")`로 출력 인코딩을 맞춰 두었다.

## 이 하네스가 동작하는 방식 (하네스 엔지니어링 관점)

이 프로젝트는 "스펙 → 생성 → 검증"을 분리해서 재현 가능하게 만든 구조다.

1. **CLAUDE.md가 유일한 스펙 원본이다.** 출력 시트의 셀 단위 레이아웃(B~S열),
   표시 형식(`"▲"0;"▼"0;0`), 조건부 서식 규칙, 검증 기준 7가지가 모두 여기에
   적혀 있다. Claude Code는 이 파일을 세션 시작 시 자동으로 읽고 작업 규칙으로
   삼는다. 사람이 직접 읽어도 동일하게 적용 가능한 문서여야 한다.
2. **`scripts/build_report.py`가 유일한 계산/생성 지점이다.** 순위·변동값·지수를
   사람이 대화 중에 암산하거나 추정하지 않고, 이 스크립트가 원본 CSV를 읽어
   계산한다. 국가 매칭은 행 위치가 아니라 `strip()`한 국가명 기준이다(CSV마다
   국가명 끝 공백, 2020년 파일의 정렬 순서·따옴표·소수 자릿수 차이 등 알려진
   문제가 있기 때문).
3. **`evals/verify.py`가 유일한 합격 기준이다.** CLAUDE.md의 검증 기준 1~7을
   그대로 구현한다 (행 개수/중복, 연도별 순위 일치, 정렬 순서, 변동값 계산식,
   지수 일치, 빈 셀 여부, 최대 성장/후퇴 국가 요약). `build_report.py`를 고친
   뒤에는 반드시 재실행해서 전부 통과하는지 확인한다.

새 환경에서 이 작업을 다시 수행하려면: (1) CLAUDE.md를 읽고 사양을 확인 →
(2) `build_report.py` 실행 → (3) `verify.py` 실행 → (4) 실패 시 원인을 고치고
2~3번 반복, 의 루프를 그대로 따라 하면 된다.

## 구현 노트

- **C열(A Decade) 양방향 데이터 막대**: "0 기준 양방향, 상승 초록·하락 빨강"
  데이터 막대는 Excel 2010+에서 도입된 `x14` 확장 기능이며, `openpyxl`
  (현재 3.1.5 기준)은 이를 생성하는 API를 제공하지 않는다. 이 때문에
  `build_report.py`는 `openpyxl`로 저장을 마친 뒤, 생성된 `.xlsx`(zip)
  내부의 워크시트 XML을 직접 열어 `x14:dataBar` 확장 블록을 주입하는
  후처리(`inject_bidirectional_databar` 함수)를 수행한다. 템플릿 구조가
  바뀌거나(시트 관계 경로 변경 등) openpyxl 버전이 올라가면 이 함수의
  관계(rels) 파싱 로직을 다시 점검해야 한다.
- **J~N(2년 변동) 글자색**: 배경색 없이 글자색만 바꾸는 조건부 서식은
  `openpyxl.formatting.rule.FormulaRule`로 충분히 구현 가능해 별도 후처리가
  필요 없다.
- **빈 데이터 재검산 금지**: 검증 스크립트는 값을 다시 계산해 생성 스크립트와
  "같은 로직을 두 번 짜는" 방식이 아니라, 생성 스크립트가 만든 결과를 원본
  CSV와 직접 대조하는 방식으로 짜여 있다. 생성 스크립트에 버그가 있어도
  검증 스크립트가 같은 버그를 반복하지 않도록 하기 위함이다.

## 참고: 최근 실행 결과 요약

마지막으로 확인된 실행에서는 10년간(2014→2024) 가장 많이 순위가 오른 국가는
**Ukraine (+57)**, 가장 많이 내려간 국가는 **Venezuela (−64)**였다. 이 값은
`input_data/`가 바뀌면 함께 바뀌므로 고정값이 아니라 항상 스크립트 실행
결과를 기준으로 삼는다.

## 프롬프트 기록 (하네스 엔지니어링 로그)

이 저장소는 Claude Code와의 아래 대화 순서를 그대로 따라가며 만들어졌다.
새 환경에서 같은 과정을 재현하고 싶다면 CLAUDE.md를 준비해 둔 상태에서
아래 프롬프트를 순서대로 입력하면 된다(일부는 결과를 보고 사람이 판단해
방향을 바꾼 지점이라 중간에 사람의 확인이 필요하다).

1. > CLAUDE.md를 읽고, 이 프로젝트에서 네가 반드시 지켜야 할 규칙 3가지를 말해줘.
2. > CLAUDE.md의 검증 기준 1~7을 구현한 evals/verify.py를 만들어줘. 아직 output 파일이 없으니, 실행하면 "파일 없음"으로 실패하는 게 정상이야. 만든 뒤 한 번 실행해서 실패하는 것까지 보여줘.
3. > CLAUDE.md
   > 이 파일에는 모순이 없는지 확인해줘
4. > 명시되지 않은 경계조건을 명시하도록 CLAUDE.md에 적어줘
5. > CLAUDE.md 출력 사양대로 scripts/build_report.py를 만들어줘. 만든 뒤 실행하고 evals/verify.py로 검증해서, 모든 항목이 통과할 때까지 스스로 고쳐줘. 통과하면 검증 결과 요약을 보여줘.
6. > (결과 파일에서 A Decade 열이 "-"로 보이는 국가들 스크린샷 첨부) output\EGOV_DATA_취합.xlsx 결과 파일을 열어서 확인해 보면 ADecade 내용에 -로 표현 된 것을 0으로 표현해줘
