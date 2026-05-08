from __future__ import annotations

import math
import sys
import zipfile
from pathlib import Path
from xml.etree.ElementTree import Element, SubElement, tostring

import numpy as np
import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

import main

OUTPUT_FILE = BASE_DIR / "data" / "phan_tich_chi_tieu_4_sheets.xlsx"
CSV_OUTPUT_DIR = BASE_DIR / "data" / "csv_exports"


def column_name(column_index: int) -> str:
    name = ""
    column_index += 1
    while column_index:
        column_index, remainder = divmod(column_index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return pd.isna(value) if not isinstance(value, (list, tuple, dict)) else False


def normalize_cell_value(value: object) -> object:
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, np.datetime64):
        return pd.Timestamp(value).strftime("%Y-%m-%d")
    if isinstance(value, pd.Period):
        return str(value)
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        return float(value)
    return value


def dataframe_rows(data_frame: pd.DataFrame) -> list[list[object]]:
    frame = data_frame.copy()
    frame.columns = [str(column) for column in frame.columns]
    rows: list[list[object]] = [frame.columns.tolist()]
    for row in frame.itertuples(index=False, name=None):
        rows.append([normalize_cell_value(value) for value in row])
    return rows


def make_sheet_xml(rows: list[list[object]]) -> bytes:
    worksheet = Element(
        "worksheet",
        {
            "xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    sheet_data = SubElement(worksheet, "sheetData")

    for row_index, row_values in enumerate(rows, start=1):
        row_element = SubElement(sheet_data, "row", {"r": str(row_index)})
        for column_index, raw_value in enumerate(row_values):
            if is_blank(raw_value):
                continue

            value = normalize_cell_value(raw_value)
            cell_reference = f"{column_name(column_index)}{row_index}"
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cell = SubElement(row_element, "c", {"r": cell_reference})
                SubElement(cell, "v").text = str(value)
            else:
                cell = SubElement(row_element, "c", {"r": cell_reference, "t": "inlineStr"})
                inline_string = SubElement(cell, "is")
                SubElement(inline_string, "t").text = str(value)

    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + tostring(
        worksheet,
        encoding="utf-8",
    )


def make_stats_frame(cleaned_data: pd.DataFrame, cleaning_info: dict[str, int]) -> pd.DataFrame:
    monthly_totals = main.summarize_by_month(cleaned_data)
    forecast_totals, forecast_slope = main.forecast_monthly_expenses(monthly_totals)
    evaluation = main.evaluate_forecast_model(monthly_totals)
    unusual_expenses = main.detect_unusual_expenses(cleaned_data)
    top_expenses = main.top_expenses(cleaned_data)

    rows: list[dict[str, object]] = [
        {"section": "Lam sach du lieu", "metric": "So dong ban dau", "value": cleaning_info["original_rows"]},
        {
            "section": "Lam sach du lieu",
            "metric": "So dong trung lap da xoa",
            "value": cleaning_info["duplicate_rows_removed"],
        },
        {
            "section": "Lam sach du lieu",
            "metric": "So dong loi du lieu da xoa",
            "value": cleaning_info["invalid_rows_removed"],
        },
        {
            "section": "Lam sach du lieu",
            "metric": "So dong khong phai debit da loai",
            "value": cleaning_info["non_expense_rows_removed"],
        },
        {
            "section": "Lam sach du lieu",
            "metric": "So dong chuyen khoan/tra the da loai",
            "value": cleaning_info["transfer_rows_removed"],
        },
        {
            "section": "Lam sach du lieu",
            "metric": "So dong hop le sau lam sach",
            "value": cleaning_info["clean_rows"],
        },
        {"section": "Tong quan", "metric": "So giao dich", "value": len(cleaned_data)},
        {"section": "Tong quan", "metric": "Tong chi tieu", "value": float(cleaned_data["amount"].sum())},
        {"section": "Tong quan", "metric": "Chi tieu trung binh", "value": float(np.mean(cleaned_data["amount"]))},
        {"section": "Tong quan", "metric": "Trung vi chi tieu", "value": float(np.median(cleaned_data["amount"]))},
        {"section": "Tong quan", "metric": "Giao dich lon nhat", "value": float(np.max(cleaned_data["amount"]))},
        {"section": "Tong quan", "metric": "Giao dich nho nhat", "value": float(np.min(cleaned_data["amount"]))},
        {"section": "Bat thuong", "metric": "So giao dich bat thuong", "value": len(unusual_expenses)},
        {"section": "Du bao", "metric": "Do doc xu huong moi thang", "value": forecast_slope},
    ]

    if evaluation["mae"] is not None:
        rows.extend(
            [
                {"section": "Danh gia du bao", "metric": "MAE", "value": evaluation["mae"]},
                {"section": "Danh gia du bao", "metric": "RMSE", "value": evaluation["rmse"]},
                {"section": "Danh gia du bao", "metric": "R2", "value": evaluation["r2"]},
            ]
        )

    for row in monthly_totals.itertuples(index=False):
        rows.append({"section": "Chi tieu theo thang", "metric": row.month, "value": float(row.amount)})

    for row in forecast_totals.itertuples(index=False):
        rows.append({"section": "Du bao 3 thang tiep theo", "metric": row.month, "value": float(row.forecast_amount)})

    for row in top_expenses.itertuples(index=False):
        rows.append(
            {
                "section": "Top giao dich",
                "metric": f"{row.date.strftime('%Y-%m-%d')} | {row.category} | {row.description}",
                "value": float(row.amount),
            }
        )

    return pd.DataFrame(rows)


def make_category_analysis_frame(cleaned_data: pd.DataFrame) -> pd.DataFrame:
    category_totals = main.summarize_by_category(cleaned_data)
    category_totals["comment"] = category_totals["category"].map(main.category_comment)
    return category_totals


def write_xlsx(output_file: Path, sheets: list[tuple[str, list[list[object]]]]) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
{sheet_entries}
  </sheets>
</workbook>
""".format(
        sheet_entries="\n".join(
            f'    <sheet name="{name}" sheetId="{index}" r:id="rId{index}"/>'
            for index, (name, _) in enumerate(sheets, start=1)
        )
    )

    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{worksheet_relationships}
  <Relationship Id="rId{style_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
""".format(
        worksheet_relationships="\n".join(
            f'  <Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
            for index in range(1, len(sheets) + 1)
        ),
        style_id=len(sheets) + 1,
    )

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{worksheet_overrides}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
""".format(
        worksheet_overrides="\n".join(
            f'  <Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for index in range(1, len(sheets) + 1)
        )
    )

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>
"""

    with zipfile.ZipFile(output_file, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types)
        workbook.writestr("_rels/.rels", root_rels)
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        workbook.writestr("xl/styles.xml", styles)
        for index, (_, rows) in enumerate(sheets, start=1):
            workbook.writestr(f"xl/worksheets/sheet{index}.xml", make_sheet_xml(rows))


def main_export() -> None:
    data_file = main.DATA_FILE if main.DATA_FILE.exists() else main.FALLBACK_DATA_FILE
    raw_data = main.read_transaction_file(data_file)
    cleaned_data, cleaning_info = main.load_expenses(data_file)
    stats = make_stats_frame(cleaned_data, cleaning_info)
    category_analysis = make_category_analysis_frame(cleaned_data)

    write_xlsx(
        OUTPUT_FILE,
        [
            ("dữ liệu gốc", dataframe_rows(raw_data)),
            ("dữ liệu đã làm sạch", dataframe_rows(cleaned_data)),
            ("thống kê", dataframe_rows(stats)),
            ("phân tích danh mục", dataframe_rows(category_analysis)),
        ],
    )
    CSV_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    raw_data.to_csv(CSV_OUTPUT_DIR / "sheet1_du_lieu_goc.csv", index=False, encoding="utf-8-sig")
    cleaned_data.to_csv(CSV_OUTPUT_DIR / "sheet2_du_lieu_da_lam_sach.csv", index=False, encoding="utf-8-sig")
    stats.to_csv(CSV_OUTPUT_DIR / "sheet3_thong_ke.csv", index=False, encoding="utf-8-sig")
    category_analysis.to_csv(CSV_OUTPUT_DIR / "sheet4_phan_tich_danh_muc.csv", index=False, encoding="utf-8-sig")
    print(OUTPUT_FILE)
    print(CSV_OUTPUT_DIR)


if __name__ == "__main__":
    main_export()
