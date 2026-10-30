"""EGDI 취합 보고서 검증 스크립트 (CLAUDE.md 검증 기준 1~7).

output/EGOV_DATA_취합.xlsx 를 input_data/*.csv 원본과 대조하여 검증한다.
실패 항목이 하나라도 있으면 종료 코드 1로 끝낸다.
"""

import csv
import sys
from pathlib import Path

import openpyxl

if sys.stdout.encoding is None or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = ROOT / "input_data"
OUTPUT_PATH = ROOT / "output" / "EGOV_DATA_취합.xlsx"
SHEET_NAME = "EGOV_DATA_취합"

YEARS = [2024, 2022, 2020, 2018, 2016, 2014]
RANK_COLS = ["D", "E", "F", "G", "H", "I"]  # 2024, 2022, 2020, 2018, 2016, 2014
CHANGE_COLS = ["J", "K", "L", "M", "N"]  # 22->24, 20->22, 18->20, 16->18, 14->16
CHANGE_YEAR_PAIRS = [(2022, 2024), (2020, 2022), (2018, 2020), (2016, 2018), (2014, 2016)]
INDEX_COLS = ["O", "P", "Q", "R", "S"]
INDEX_NAMES = [
    "E-Government Index",
    "E-Participation Index",
    "Online Service Index",
    "Human Capital Index",
    "Telecommunication Infrastructure Index",
]
FIRST_DATA_ROW = 4
LAST_DATA_ROW = 196
EXPECTED_COUNTRY_COUNT = 193

failures = []
warnings = []


def fail(msg):
    failures.append(msg)
    print(f"[FAIL] {msg}")


def ok(msg):
    print(f"[ OK ] {msg}")


def load_csv(year):
    path = INPUT_DIR / f"EGOV_DATA_{year}.csv"
    rows = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = row["Country Name"].strip()
            rows[name] = {
                "rank": int(float(row["E-Government Rank"])),
                "egov_index": float(row["E-Government Index"]),
                "eparticipation_index": float(row["E-Participation Index"]),
                "online_service_index": float(row["Online Service Index"]),
                "human_capital_index": float(row["Human Capital Index"]),
                "telecom_index": float(row["Telecommunication Infrastructure Index"]),
            }
    return rows


def main():
    if not OUTPUT_PATH.exists():
        fail(f"출력 파일이 존재하지 않음: {OUTPUT_PATH}")
        print_summary_and_exit()
        return

    wb = openpyxl.load_workbook(OUTPUT_PATH, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        fail(f"시트 '{SHEET_NAME}' 가 없음 (존재하는 시트: {wb.sheetnames})")
        print_summary_and_exit()
        return
    ws = wb[SHEET_NAME]

    csv_data = {}
    for year in YEARS:
        try:
            csv_data[year] = load_csv(year)
        except FileNotFoundError:
            fail(f"원본 CSV 없음: EGOV_DATA_{year}.csv")
    if len(csv_data) < len(YEARS):
        print_summary_and_exit()
        return

    rows = []
    for r in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
        country = ws[f"B{r}"].value
        if country is not None:
            country = str(country).strip()
        rows.append({"row": r, "country": country})

    # 1. 데이터 행이 정확히 193개이고 국가명 중복이 없다
    non_empty = [r for r in rows if r["country"]]
    if len(non_empty) != EXPECTED_COUNTRY_COUNT:
        fail(f"데이터 행 개수 불일치: {len(non_empty)}개 (기대: {EXPECTED_COUNTRY_COUNT}개)")
    else:
        ok(f"데이터 행 개수 = {EXPECTED_COUNTRY_COUNT}")

    names = [r["country"] for r in non_empty]
    dupes = sorted({n for n in names if names.count(n) > 1})
    if dupes:
        fail(f"국가명 중복: {dupes}")
    else:
        ok("국가명 중복 없음")

    # 6. 빈 셀이 없다 (B~S, 데이터 행 범위 전체)
    empty_cells = []
    all_cols = ["B", "C"] + RANK_COLS + CHANGE_COLS + INDEX_COLS
    for r in rows:
        if r["row"] > FIRST_DATA_ROW - 1:
            for col in all_cols:
                val = ws[f"{col}{r['row']}"].value
                if val is None or (isinstance(val, str) and val.strip() == ""):
                    empty_cells.append(f"{col}{r['row']}")
    if empty_cells:
        fail(f"빈 셀 발견 ({len(empty_cells)}개): {empty_cells[:20]}{' ...' if len(empty_cells) > 20 else ''}")
    else:
        ok("빈 셀 없음")

    # 3. 행이 2024 순위 오름차순이다
    rank_2024_seq = []
    for r in rows:
        if r["country"]:
            v = ws[f"D{r['row']}"].value
            rank_2024_seq.append(v)
    if any(v is None for v in rank_2024_seq):
        fail("2024 순위(D열)에 빈 값이 있어 정렬 검증 불가")
    else:
        nums = [float(v) for v in rank_2024_seq]
        if nums != sorted(nums):
            fail("행이 2024 순위 오름차순이 아님")
        else:
            ok("행이 2024 순위 오름차순")

    # 2. D~I의 모든 순위가 해당 연도 원본 CSV와 일치한다
    rank_mismatches = []
    missing_countries = []
    for r in rows:
        country = r["country"]
        if not country:
            continue
        for col, year in zip(RANK_COLS, YEARS):
            year_rows = csv_data[year]
            if country not in year_rows:
                missing_countries.append((country, year))
                continue
            expected = year_rows[country]["rank"]
            actual = ws[f"{col}{r['row']}"].value
            if actual is None or int(float(actual)) != expected:
                rank_mismatches.append(f"{country} {year}: 출력={actual} 원본={expected}")
    if missing_countries:
        fail(f"원본 CSV에서 매칭되지 않은 국가-연도 조합 ({len(missing_countries)}개): {missing_countries[:10]}")
    if rank_mismatches:
        fail(f"순위 불일치 ({len(rank_mismatches)}건): {rank_mismatches[:10]}")
    elif not missing_countries:
        ok("D~I 순위가 원본 CSV와 모두 일치")

    # 4. C = I - D, J~N = 인접 연도 순위 차이와 일치
    calc_mismatches_c = []
    calc_mismatches_jn = []
    for r in rows:
        country = r["country"]
        if not country:
            continue
        rank_by_col = {}
        for col in RANK_COLS:
            v = ws[f"{col}{r['row']}"].value
            rank_by_col[col] = float(v) if v is not None else None

        d_val = ws[f"C{r['row']}"].value
        if rank_by_col["D"] is not None and rank_by_col["I"] is not None:
            expected_c = rank_by_col["I"] - rank_by_col["D"]
            if d_val is None or float(d_val) != expected_c:
                calc_mismatches_c.append(f"{country}: C={d_val} 기대값={expected_c}")

        for change_col, (year_from, year_to) in zip(CHANGE_COLS, CHANGE_YEAR_PAIRS):
            col_from = RANK_COLS[YEARS.index(year_from)]
            col_to = RANK_COLS[YEARS.index(year_to)]
            rank_from = rank_by_col[col_from]
            rank_to = rank_by_col[col_to]
            actual_change = ws[f"{change_col}{r['row']}"].value
            if rank_from is not None and rank_to is not None:
                expected_change = rank_from - rank_to
                if actual_change is None or float(actual_change) != expected_change:
                    calc_mismatches_jn.append(
                        f"{country} {change_col}({year_from}->{year_to}): 출력={actual_change} 기대값={expected_change}"
                    )
    if calc_mismatches_c:
        fail(f"C열(A Decade) 계산 불일치 ({len(calc_mismatches_c)}건): {calc_mismatches_c[:10]}")
    else:
        ok("C열(A Decade) = I - D 계산 일치")
    if calc_mismatches_jn:
        fail(f"J~N(2년 변동) 계산 불일치 ({len(calc_mismatches_jn)}건): {calc_mismatches_jn[:10]}")
    else:
        ok("J~N 2년 변동 계산 일치")

    # 5. O~S가 2024 원본 지수와 일치한다
    index_keys = [
        "egov_index",
        "eparticipation_index",
        "online_service_index",
        "human_capital_index",
        "telecom_index",
    ]
    index_mismatches = []
    data_2024 = csv_data[2024]
    for r in rows:
        country = r["country"]
        if not country or country not in data_2024:
            continue
        expected_vals = data_2024[country]
        for col, key in zip(INDEX_COLS, index_keys):
            expected = expected_vals[key]
            actual = ws[f"{col}{r['row']}"].value
            if actual is None or abs(float(actual) - expected) > 1e-9:
                index_mismatches.append(f"{country} {col}: 출력={actual} 원본={expected}")
    if index_mismatches:
        fail(f"O~S 지수 불일치 ({len(index_mismatches)}건): {index_mismatches[:10]}")
    else:
        ok("O~S 지수(2024)가 원본과 일치")

    # 7. 결과 요약: C열 최댓값/최솟값 국가
    c_values = []
    for r in rows:
        if r["country"]:
            v = ws[f"C{r['row']}"].value
            if v is not None:
                c_values.append((r["country"], float(v)))
    if c_values:
        best = max(c_values, key=lambda x: x[1])
        worst = min(c_values, key=lambda x: x[1])
        print(f"[요약] 10년간 최대 성장: {best[0]} ({best[1]:+.0f})")
        print(f"[요약] 10년간 최대 후퇴: {worst[0]} ({worst[1]:+.0f})")
    else:
        warnings.append("C열 값이 없어 요약 통계를 낼 수 없음")

    print_summary_and_exit()


def print_summary_and_exit():
    print()
    if failures:
        print(f"검증 실패: {len(failures)}건")
        for f in failures:
            print(f" - {f}")
        sys.exit(1)
    else:
        print("검증 통과: 모든 항목 성공")
        sys.exit(0)


if __name__ == "__main__":
    main()
