"""EGDI 순위 변동 보고서 생성 스크립트.

input_data/*.csv 6개 연도를 template/EGOV_DATA_template.xlsx 양식에 맞춰
output/EGOV_DATA_취합.xlsx 로 취합한다. CLAUDE.md의 "출력 사양"을 그대로 따른다.
"""

import csv
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import FormulaRule

if sys.stdout.encoding is None or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = ROOT / "input_data"
TEMPLATE_PATH = ROOT / "template" / "EGOV_DATA_template.xlsx"
OUTPUT_PATH = ROOT / "output" / "EGOV_DATA_취합.xlsx"
SHEET_NAME = "EGOV_DATA_취합"

YEARS = [2024, 2022, 2020, 2018, 2016, 2014]
RANK_COLS = ["D", "E", "F", "G", "H", "I"]
CHANGE_COLS = ["J", "K", "L", "M", "N"]
CHANGE_YEAR_PAIRS = [(2022, 2024), (2020, 2022), (2018, 2020), (2016, 2018), (2014, 2016)]
CHANGE_LABELS = ["22→24", "20→22", "18→20", "16→18", "14→16"]
INDEX_COLS = ["O", "P", "Q", "R", "S"]
INDEX_NAMES = [
    "E-Government Index",
    "E-Participation Index",
    "Online Service Index",
    "Human Capital Index",
    "Telecommunication Infrastructure Index",
]
INDEX_KEYS = [
    "egov_index",
    "eparticipation_index",
    "online_service_index",
    "human_capital_index",
    "telecom_index",
]
FIRST_DATA_ROW = 4
LAST_DATA_ROW = 196

SIGN_NUMBER_FORMAT = '"▲"0;"▼"0;0'
GREEN = "FF008000"
RED = "FFFF0000"
GRAY = "FF808080"


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


def rebuild_header(ws):
    # J~N: 기존 템플릿은 지수 5종이 J~N에 2행 단일열 병합으로 들어가 있었다.
    # 이를 해제하고 "Rank Change (2-year)"를 D~I처럼 2행에 가로 병합한다.
    for col in CHANGE_COLS:
        merge_range = f"{col}2:{col}3"
        if merge_range in {str(r) for r in ws.merged_cells.ranges}:
            ws.unmerge_cells(merge_range)
        ws[f"{col}2"] = None
        ws[f"{col}3"] = None

    ws["J2"] = "Rank Change (2-year)"
    ws.merge_cells("J2:N2")
    for col, label in zip(CHANGE_COLS, CHANGE_LABELS):
        ws[f"{col}3"] = label

    # O~S: 지수 5종을 각각 2~3행 병합 단일 헤더로 구성한다.
    for col, name in zip(INDEX_COLS, INDEX_NAMES):
        merge_range = f"{col}2:{col}3"
        if merge_range in {str(r) for r in ws.merged_cells.ranges}:
            ws.unmerge_cells(merge_range)
        ws.merge_cells(merge_range)
        ws[f"{col}2"] = name

    # B, C, D~I 헤더는 템플릿 원본 그대로 사용 (Country Name / A Decade / E-Government Rank + 연도).
    # 혹시 비어 있을 경우를 대비해 명시적으로 채워 둔다.
    ws["B2"] = "Country Name"
    ws["C2"] = "A Decade"
    ws["D2"] = "E-Government Rank"
    for col, year in zip(RANK_COLS, YEARS):
        ws[f"{col}3"] = year


def write_data(ws, data):
    countries_sorted = sorted(data[2024].items(), key=lambda kv: kv[1]["rank"])
    assert len(countries_sorted) == LAST_DATA_ROW - FIRST_DATA_ROW + 1

    for offset, (country, info_2024) in enumerate(countries_sorted):
        row = FIRST_DATA_ROW + offset
        ws[f"B{row}"] = country

        rank_by_year = {year: data[year][country]["rank"] for year in YEARS}
        ws[f"C{row}"] = rank_by_year[2014] - rank_by_year[2024]

        for col, year in zip(RANK_COLS, YEARS):
            ws[f"{col}{row}"] = rank_by_year[year]

        for col, (year_from, year_to) in zip(CHANGE_COLS, CHANGE_YEAR_PAIRS):
            ws[f"{col}{row}"] = rank_by_year[year_from] - rank_by_year[year_to]

        for col, key in zip(INDEX_COLS, INDEX_KEYS):
            ws[f"{col}{row}"] = data[2024][country][key]


def apply_formatting(ws):
    rank_change_range = f"J{FIRST_DATA_ROW}:N{LAST_DATA_ROW}"
    c_range = f"C{FIRST_DATA_ROW}:C{LAST_DATA_ROW}"

    for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
        ws[f"C{row}"].number_format = SIGN_NUMBER_FORMAT
        for col in CHANGE_COLS:
            ws[f"{col}{row}"].number_format = SIGN_NUMBER_FORMAT
        for col in RANK_COLS:
            ws[f"{col}{row}"].number_format = "General"

    ws.conditional_formatting.add(
        rank_change_range,
        FormulaRule(formula=[f"J{FIRST_DATA_ROW}>0"], font=openpyxl.styles.Font(color=GREEN)),
    )
    ws.conditional_formatting.add(
        rank_change_range,
        FormulaRule(formula=[f"J{FIRST_DATA_ROW}<0"], font=openpyxl.styles.Font(color=RED)),
    )
    ws.conditional_formatting.add(
        rank_change_range,
        FormulaRule(formula=[f"J{FIRST_DATA_ROW}=0"], font=openpyxl.styles.Font(color=GRAY)),
    )

    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"B3:S{LAST_DATA_ROW}"

    return c_range


X14_NS_ATTRS = (
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
)
X14_EXT_BLOCK_TEMPLATE = """<extLst><ext uri="{{78C0D931-6437-407d-A8EE-F0AADF4582F6}}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{{B1F1A2C1-0001-4B1A-9C1A-000000000001}}"><x14:dataBar minLength="0" maxLength="100" gradient="0" border="0" negativeBarBorderColorSameAsPositive="0" axisPosition="automatic"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF008000"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>{sqref}</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>"""


def inject_bidirectional_databar(xlsx_path, sheet_name, sqref, fill_color="FF008000"):
    """C열에 0 기준 양방향 데이터 막대(상승 초록, 하락 빨강)를 적용한다.

    openpyxl은 Excel 2010+ 방식의 양방향(음수 축) 데이터 막대(x14 확장)를
    생성하는 API를 제공하지 않으므로, 저장된 xlsx를 직접 후처리해 주입한다.
    """
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        names = zin.namelist()
        wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
        contents = {n: zin.read(n) for n in names}

    import re

    sheet_rid_map = {}
    for m in re.finditer(r"<sheet\b[^>]*/>", wb_xml):
        tag = m.group(0)
        name_m = re.search(r'name="([^"]*)"', tag)
        rid_m = re.search(r'r:id="(rId\d+)"', tag)
        if name_m and rid_m:
            sheet_rid_map[name_m.group(1)] = rid_m.group(1)
    if sheet_name not in sheet_rid_map:
        raise RuntimeError(f"시트 '{sheet_name}' 를 workbook.xml에서 찾을 수 없음")
    target_rid = sheet_rid_map[sheet_name]

    rels_xml = contents["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rid_target = None
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels_xml):
        tag = m.group(0)
        id_m = re.search(r'Id="([^"]*)"', tag)
        target_m = re.search(r'Target="([^"]*)"', tag)
        if id_m and target_m and id_m.group(1) == target_rid:
            rid_target = target_m.group(1)
            break
    if rid_target is None:
        raise RuntimeError("워크시트 관계(rels)를 찾을 수 없음")
    sheet_part = rid_target.lstrip("/") if rid_target.startswith("/") else "xl/" + rid_target

    sheet_xml = contents[sheet_part].decode("utf-8")

    ext_block = X14_EXT_BLOCK_TEMPLATE.format(sqref=sqref)

    if "</worksheet>" not in sheet_xml:
        raise RuntimeError("워크시트 XML 구조를 인식할 수 없음")

    if "<extLst>" in sheet_xml:
        sheet_xml = sheet_xml.replace("</extLst>", ext_block.replace("<extLst>", "").replace("</extLst>", "") + "</extLst>", 1)
    else:
        sheet_xml = sheet_xml.replace("</worksheet>", ext_block + "</worksheet>")

    root_tag_match = re.search(r"<worksheet[^>]*>", sheet_xml)
    root_tag = root_tag_match.group(0)
    if "xmlns:mc=" not in root_tag:
        new_root_tag = root_tag[:-1] + X14_NS_ATTRS + ">"
        sheet_xml = sheet_xml.replace(root_tag, new_root_tag, 1)

    contents[sheet_part] = sheet_xml.encode("utf-8")

    tmp_path = xlsx_path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])
    shutil.move(str(tmp_path), str(xlsx_path))


def main():
    print("원본 CSV 로딩 중...")
    data = {year: load_csv(year) for year in YEARS}
    for year in YEARS:
        print(f"  {year}: {len(data[year])}개국")

    print("템플릿 로딩 중...")
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    if SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(f"템플릿에 시트 '{SHEET_NAME}' 가 없음 (존재: {wb.sheetnames})")
    ws = wb[SHEET_NAME]

    print("헤더 재구성 중...")
    rebuild_header(ws)

    print("데이터 작성 중...")
    write_data(ws, data)

    print("서식 적용 중...")
    c_range = apply_formatting(ws)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"저장 완료: {OUTPUT_PATH}")

    print("양방향 데이터 막대(C열) 후처리 중...")
    try:
        inject_bidirectional_databar(OUTPUT_PATH, SHEET_NAME, c_range)
        print("  완료")
    except Exception as e:
        print(f"  [경고] 데이터 막대 후처리 실패 (데이터/다른 서식에는 영향 없음): {e}")

    # 후처리 후 파일이 정상적으로 열리는지 확인.
    openpyxl.load_workbook(OUTPUT_PATH)
    print("최종 파일 재로딩 확인 완료")


if __name__ == "__main__":
    main()
