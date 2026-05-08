from __future__ import annotations

import tkinter as tk
import zipfile
from pathlib import Path
from tkinter import ttk
import xml.etree.ElementTree as ET

import matplotlib
import numpy as np
import pandas as pd

matplotlib.use("TkAgg")

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "personal_transactions_dashboard_ready (2).xlsx"
FALLBACK_DATA_FILE = BASE_DIR / "data" / "expenses.csv"
REPORT_FILE = BASE_DIR / "report.txt"
EXCEL_REPORT_FILE = BASE_DIR / "data" / "phan_tich_chi_tieu_4_sheets.xlsx"
CHART_DIR = BASE_DIR / "charts"
BAR_CHART_FILE = CHART_DIR / "expense_bar_chart.png"
PIE_CHART_FILE = CHART_DIR / "expense_pie_chart.png"
LINE_CHART_FILE = CHART_DIR / "expense_line_chart.png"


def excel_column_index(cell_reference: str) -> int:
    column_name = "".join(character for character in cell_reference if character.isalpha())
    column_index = 0
    for character in column_name:
        column_index = column_index * 26 + ord(character.upper()) - ord("A") + 1
    return column_index - 1


def read_xlsx_first_sheet(file_path: Path) -> pd.DataFrame:
    # Parser nhe cho workbook don gian, tranh phu thuoc openpyxl khi moi truong chua cai package nay.
    namespace = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(file_path) as workbook:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            for shared_item in shared_root.findall("a:si", namespace):
                shared_strings.append("".join(text.text or "" for text in shared_item.findall(".//a:t", namespace)))

        sheet_root = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
        rows: list[list[object]] = []
        for row in sheet_root.findall("a:sheetData/a:row", namespace):
            row_values: list[object] = []
            for cell in row.findall("a:c", namespace):
                cell_index = excel_column_index(cell.attrib.get("r", "A1"))
                while len(row_values) <= cell_index:
                    row_values.append("")

                cell_value = cell.find("a:v", namespace)
                inline_value = cell.find("a:is/a:t", namespace)
                value: object = "" if cell_value is None else cell_value.text or ""
                if cell.attrib.get("t") == "s" and value != "":
                    value = shared_strings[int(value)]
                elif cell.attrib.get("t") == "inlineStr" and inline_value is not None:
                    value = inline_value.text or ""
                row_values[cell_index] = value
            rows.append(row_values)

    if not rows:
        return pd.DataFrame()

    width = max(len(row) for row in rows)
    normalized_rows = [row + [""] * (width - len(row)) for row in rows]
    return pd.DataFrame(normalized_rows[1:], columns=normalized_rows[0])


def read_transaction_file(file_path: Path) -> pd.DataFrame:
    if file_path.suffix.lower() == ".xlsx":
        try:
            return pd.read_excel(file_path)
        except ImportError:
            return read_xlsx_first_sheet(file_path)
    return pd.read_csv(file_path, encoding="utf-8-sig")


def normalize_column_name(column_name: object) -> str:
    return str(column_name).strip().lower().replace(" ", "_")


def parse_transaction_dates(date_series: pd.Series) -> pd.Series:
    numeric_dates = pd.to_numeric(date_series, errors="coerce")
    parsed_dates = pd.Series(pd.NaT, index=date_series.index, dtype="datetime64[ns]")
    numeric_mask = numeric_dates.notna()
    parsed_dates.loc[numeric_mask] = pd.to_datetime(
        numeric_dates.loc[numeric_mask],
        unit="D",
        origin="1899-12-30",
        errors="coerce",
    )
    parsed_dates.loc[~numeric_mask] = pd.to_datetime(date_series.loc[~numeric_mask], errors="coerce")
    return parsed_dates


def load_expenses(file_path: Path) -> tuple[pd.DataFrame, dict[str, int]]:
    # Doc du lieu tu CSV/Excel vao DataFrame de bat dau quy trinh phan tich.
    data_frame = read_transaction_file(file_path)
    data_frame = data_frame.rename(columns={column: normalize_column_name(column) for column in data_frame.columns})
    data_frame = data_frame.rename(columns={"payment_method": "account_name"})

    # Ghi nhận quy mo ban dau de bao cao so dong bi loai trong qua trinh lam sach.
    original_rows = len(data_frame)
    data_frame = data_frame.drop_duplicates().copy()
    duplicate_rows_removed = original_rows - len(data_frame)

    # Chuyen cac cot chinh sang kieu phu hop; gia tri loi se thanh NaT/NaN de de loc bo.
    required_columns = {"date", "amount", "category"}
    missing_columns = sorted(required_columns - set(data_frame.columns))
    if missing_columns:
        raise ValueError(f"Thieu cot bat buoc trong tep du lieu: {', '.join(missing_columns)}")

    data_frame["date"] = parse_transaction_dates(data_frame["date"])
    data_frame["amount"] = pd.to_numeric(data_frame["amount"], errors="coerce")
    invalid_rows = int(data_frame[["date", "amount", "category"]].isna().any(axis=1).sum())

    # Chuan hoa chuoi truoc khi phan tich de tranh category bi tach thanh nhieu nhom do khoang trang.
    data_frame["category"] = data_frame["category"].astype(str).str.strip()
    if "description" not in data_frame.columns:
        data_frame["description"] = ""
    if "transaction_type" not in data_frame.columns:
        data_frame["transaction_type"] = "debit"
    if "account_name" not in data_frame.columns:
        data_frame["account_name"] = ""
    data_frame["description"] = data_frame["description"].astype(str).str.strip()
    data_frame["transaction_type"] = data_frame["transaction_type"].astype(str).str.strip().str.lower()
    data_frame["account_name"] = data_frame["account_name"].astype(str).str.strip()
    data_frame = data_frame.dropna(subset=["date", "amount", "category"])
    non_expense_rows_removed = int((data_frame["transaction_type"] != "debit").sum())
    data_frame = data_frame[data_frame["transaction_type"] == "debit"].copy()
    transfer_categories = {"credit card payment"}
    transfer_rows_removed = int(data_frame["category"].str.lower().isin(transfer_categories).sum())
    data_frame = data_frame[~data_frame["category"].str.lower().isin(transfer_categories)].copy()
    negative_rows_removed = int((data_frame["amount"] < 0).sum())
    data_frame = data_frame[data_frame["amount"] >= 0].copy()
    data_frame["month"] = data_frame["date"].dt.strftime("%Y-%m")
    data_frame = data_frame.sort_values("date").reset_index(drop=True)

    cleaning_info = {
        "original_rows": original_rows,
        "duplicate_rows_removed": duplicate_rows_removed,
        "invalid_rows_removed": invalid_rows,
        "non_expense_rows_removed": non_expense_rows_removed,
        "transfer_rows_removed": transfer_rows_removed,
        "negative_rows_removed": negative_rows_removed,
        "clean_rows": len(data_frame),
    }
    return data_frame, cleaning_info


def format_currency(value: float) -> str:
    return f"${value:,.2f}"


def category_comment(category: str) -> str:
    comments = {
        "An uong": "Chi cho nhu cau thiet yeu hang ngay, can theo doi de tranh mua sam ngau hung.",
        "Di chuyen": "Anh huong truc tiep den chi phi di lai, co the toi uu bang cach gom lich trinh.",
        "Nha o": "Khoan muc co tinh co dinh, can duoc uu tien trong ke hoach ngan sach.",
        "Giai tri": "Chi phi linh hoat, de vuot ngan sach neu khong dat gioi han.",
        "Hoc tap": "Khoan dau tu cho phat trien ban than, nen duy tri o muc hop ly.",
        "Suc khoe": "Quan trong nhung phat sinh khong deu, can co quy du phong.",
        "Mua sam": "De phat sinh mua sam cam tinh, nen can nhac truoc khi chi.",
        "Shopping": "Nhom mua sam de phat sinh theo cam hung, nen dat gioi han theo thang.",
        "Mortgage & Rent": "Khoan nha o co tinh co dinh va thuong chiem ty trong lon trong ngan sach.",
        "Restaurants": "Chi phi an uong ben ngoai can theo doi vi tan suat nho nhung cong don nhanh.",
        "Utilities": "Khoan dich vu thiet yeu, nen theo doi xu huong tang bat thuong.",
        "Home Improvement": "Chi phi nha cua co the phat sinh theo du an, can tach khoan bat thuong.",
        "Mobile Phone": "Khoan dinh ky nen duoc kiem tra goi cuoc va muc su dung.",
        "Music": "Chi phi dang ky/giai tri nho, nen ra soat cac goi lap lai.",
        "Movies & Dvds": "Chi phi giai tri linh hoat, can dat gioi han neu muon tiet kiem.",
    }
    return comments.get(category, "Can theo doi dinh ky de toi uu ngan sach.")


def visualization_insights(category_totals: pd.DataFrame, monthly_totals: pd.DataFrame) -> list[str]:
    highest_category = category_totals.iloc[0]
    lowest_category = category_totals.iloc[-1]
    highest_month = monthly_totals.loc[monthly_totals["amount"].idxmax()]
    lowest_month = monthly_totals.loc[monthly_totals["amount"].idxmin()]

    return [
        (
            "Bar chart phu hop nhat de so sanh gia tri tuyet doi giua cac danh muc. "
            f"Trong bai nay, danh muc {highest_category['category']} cao nhat voi "
            f"{format_currency(highest_category['amount'])}, trong khi {lowest_category['category']} thap hon ro ret."
        ),
        (
            "Pie chart giup nhin nhanh co cau ngan sach theo ty trong phan tram. "
            f"Phan dien tich lon nhat thuoc ve {highest_category['category']}, cho thay day la nhom chi phi can uu tien kiem soat."
        ),
        (
            "Line chart the hien bien dong chi tieu theo thoi gian va ho tro nhan dien xu huong. "
            f"Du lieu cho thay muc cao nhat roi vao thang {highest_month['month']} voi {format_currency(highest_month['amount'])}, "
            f"trong khi muc thap nhat la {lowest_month['month']} voi {format_currency(lowest_month['amount'])}."
        ),
        "Xet ve tinh de doc, bar chart va line chart hieu qua hon pie chart khi can so sanh nhieu moc du lieu hoac nhieu danh muc cung luc.",
        "Han che cua truc quan hoa hien tai la chua co bieu do rieng cho giao dich bat thuong va chua tach phan so sanh theo quy, theo nam hoac theo phuong thuc thanh toan.",
    ]


def summarize_by_category(data_frame: pd.DataFrame) -> pd.DataFrame:
    # Groupby tren category cho biet nhom chi phi nao chiem ti trong lon nhat.
    category_totals = (
        data_frame.groupby("category", as_index=False)["amount"]
        .sum()
        .sort_values("amount", ascending=False)
        .reset_index(drop=True)
    )
    total_expense = float(data_frame["amount"].sum())
    category_totals["ratio"] = np.round(category_totals["amount"] / total_expense * 100, 2)
    return category_totals


def summarize_by_month(data_frame: pd.DataFrame) -> pd.DataFrame:
    # Gom tong chi tieu theo thang de phan tich xu huong thoi gian va phuc vu du bao.
    return (
        data_frame.groupby("month", as_index=False)["amount"]
        .sum()
        .sort_values("month")
        .reset_index(drop=True)
    )


def fit_linear_trend(monthly_totals: pd.DataFrame) -> tuple[float, float, pd.DataFrame]:
    # Chi lay 6 thang gan nhat vi bai toan can du bao ngan han, uu tien xu huong moi nhat.
    history_window = min(6, len(monthly_totals))
    recent_monthly_totals = monthly_totals.tail(history_window).reset_index(drop=True).copy()
    month_index = np.arange(history_window, dtype=float)
    monthly_amounts = recent_monthly_totals["amount"].to_numpy(dtype=float)

    # Neu du lieu qua it, dung muc trung binh de tranh loi khi hoi quy tuyen tinh.
    if history_window < 2:
        slope = 0.0
        intercept = float(monthly_amounts[0]) if history_window == 1 else 0.0
    else:
        slope, intercept = np.polyfit(month_index, monthly_amounts, 1)

    recent_monthly_totals["time_index"] = month_index
    return float(slope), float(intercept), recent_monthly_totals


def forecast_monthly_expenses(monthly_totals: pd.DataFrame, periods: int = 3) -> tuple[pd.DataFrame, float]:
    slope, intercept, training_window = fit_linear_trend(monthly_totals)
    history_window = len(training_window)
    future_index = np.arange(history_window, history_window + periods, dtype=float)
    future_amounts = np.maximum(intercept + slope * future_index, 0)

    last_month = pd.Period(monthly_totals["month"].iloc[-1], freq="M")
    future_months = [(last_month + offset).strftime("%Y-%m") for offset in range(1, periods + 1)]
    forecast_frame = pd.DataFrame({"month": future_months, "forecast_amount": np.round(future_amounts, 0)})
    return forecast_frame, float(slope)


def evaluate_forecast_model(monthly_totals: pd.DataFrame, test_periods: int = 3) -> dict[str, object]:
    total_months = len(monthly_totals)
    if total_months < 4:
        return {
            "train_months": [],
            "test_months": [],
            "predictions": pd.DataFrame(columns=["month", "actual_amount", "predicted_amount"]),
            "mae": None,
            "rmse": None,
            "r2": None,
            "test_periods": 0,
            "note": "Khong du so thang de tach tap huan luyen va tap kiem thu.",
        }

    test_periods = min(test_periods, max(1, total_months // 3))
    train_frame = monthly_totals.iloc[:-test_periods].reset_index(drop=True)
    test_frame = monthly_totals.iloc[-test_periods:].reset_index(drop=True)

    slope, intercept, training_window = fit_linear_trend(train_frame)
    start_index = len(training_window)
    prediction_index = np.arange(start_index, start_index + len(test_frame), dtype=float)
    predicted_amounts = np.maximum(intercept + slope * prediction_index, 0)
    actual_amounts = test_frame["amount"].to_numpy(dtype=float)
    errors = actual_amounts - predicted_amounts

    mae = float(np.mean(np.abs(errors)))
    rmse = float(np.sqrt(np.mean(np.square(errors))))
    total_variance = float(np.sum(np.square(actual_amounts - np.mean(actual_amounts))))
    residual_variance = float(np.sum(np.square(errors)))
    r2 = None if total_variance == 0 else float(1 - residual_variance / total_variance)

    prediction_frame = pd.DataFrame(
        {
            "month": test_frame["month"],
            "actual_amount": np.round(actual_amounts, 0),
            "predicted_amount": np.round(predicted_amounts, 0),
        }
    )
    return {
        "train_months": train_frame["month"].tolist(),
        "test_months": test_frame["month"].tolist(),
        "predictions": prediction_frame,
        "mae": mae,
        "rmse": rmse,
        "r2": r2,
        "test_periods": len(test_frame),
        "note": "",
    }


def top_expenses(data_frame: pd.DataFrame, limit: int = 5) -> pd.DataFrame:
    # Lay cac giao dich lon nhat de chi ra nhung khoan chi tieu dang chu y nhat trong du lieu.
    return data_frame.nlargest(limit, "amount")[["date", "category", "description", "amount"]]


def detect_unusual_expenses(data_frame: pd.DataFrame) -> pd.DataFrame:
    # Dung nguong mean + std de tim giao dich lon hon mat bang chi tieu thong thuong.
    threshold = float(np.mean(data_frame["amount"]) + np.std(data_frame["amount"]))
    unusual_expenses = data_frame[data_frame["amount"] >= threshold].copy()
    unusual_expenses["threshold"] = threshold
    return unusual_expenses.sort_values("amount", ascending=False)


def build_bar_chart_figure(category_totals: pd.DataFrame) -> plt.Figure:
    colors = ["#1f4e5f", "#2c7a7b", "#84a59d", "#f6bd60", "#f28482", "#6d597a", "#355070"]
    figure, axis = plt.subplots(figsize=(10, 6))
    # Bar chart la bieu do chinh vi phu hop nhat de so sanh tong chi tieu giua cac danh muc.
    bars = axis.bar(
        category_totals["category"],
        category_totals["amount"],
        color=colors[: len(category_totals)],
        edgecolor="white",
        linewidth=1.2,
    )
    axis.set_title("Bar chart: Tong chi tieu theo danh muc", fontsize=14, fontweight="bold")
    axis.set_xlabel("Danh muc")
    axis.set_ylabel("So tien (USD)")
    axis.grid(axis="y", linestyle="--", alpha=0.3)
    axis.set_axisbelow(True)

    max_amount = float(category_totals["amount"].max())
    axis.set_ylim(0, max_amount * 1.2)
    for bar, amount in zip(bars, category_totals["amount"]):
        axis.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + max_amount * 0.03,
            f"{amount:,.0f}",
            ha="center",
            va="bottom",
            fontsize=9,
        )

    figure.tight_layout()
    return figure


def build_pie_chart_figure(category_totals: pd.DataFrame) -> plt.Figure:
    colors = ["#355070", "#6D597A", "#B56576", "#E56B6F", "#EAAC8B", "#84A59D", "#52796F"]
    figure, axis = plt.subplots(figsize=(8, 8))
    # Pie chart dung de nhin nhanh ty trong tung nhom chi phi trong tong ngan sach.
    axis.pie(
        category_totals["amount"],
        labels=category_totals["category"],
        autopct="%1.1f%%",
        startangle=90,
        colors=colors[: len(category_totals)],
        wedgeprops={"edgecolor": "white", "linewidth": 1},
    )
    axis.set_title("Ty trong chi tieu theo danh muc")
    axis.axis("equal")
    figure.tight_layout()
    return figure


def build_line_chart_figure(monthly_totals: pd.DataFrame, forecast_totals: pd.DataFrame) -> plt.Figure:
    figure, axis = plt.subplots(figsize=(9, 5))
    # Duong lien la du lieu thuc te, giup theo doi muc chi tieu da phat sinh qua tung thang.
    axis.plot(
        monthly_totals["month"],
        monthly_totals["amount"],
        color="#d1495b",
        marker="o",
        linewidth=2.5,
        markersize=8,
        label="Thuc te",
    )
    # Duong dut doan bieu dien ket qua du bao 3 thang toi de tach biet voi du lieu da co.
    axis.plot(
        forecast_totals["month"],
        forecast_totals["forecast_amount"],
        color="#2c7a7b",
        marker="s",
        linewidth=2.5,
        markersize=7,
        linestyle="--",
        label="Du bao 3 thang toi",
    )
    axis.set_title("Line chart: Xu huong chi tieu theo thang", fontsize=14, fontweight="bold")
    axis.set_xlabel("Thang")
    axis.set_ylabel("So tien (USD)")
    axis.grid(True, linestyle="--", alpha=0.3)
    axis.legend()

    max_amount = float(monthly_totals["amount"].max())
    for month, amount in zip(monthly_totals["month"], monthly_totals["amount"]):
        axis.annotate(
            f"{amount:,.0f}",
            xy=(month, amount),
            xytext=(0, 8),
            textcoords="offset points",
            ha="center",
            fontsize=9,
        )

    for month, amount in zip(forecast_totals["month"], forecast_totals["forecast_amount"]):
        axis.annotate(
            f"{amount:,.0f}",
            xy=(month, amount),
            xytext=(0, 8),
            textcoords="offset points",
            ha="center",
            fontsize=9,
            color="#2c7a7b",
        )

    max_forecast = float(forecast_totals["forecast_amount"].max()) if not forecast_totals.empty else 0
    axis.set_ylim(0, max(max_amount, max_forecast) * 1.15)
    figure.tight_layout()
    return figure


def create_chart(figure: plt.Figure, output_file: Path) -> plt.Figure:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    # Luu anh ra file de dua vao bao cao ngay ca khi khong mo giao dien Tkinter.
    figure.savefig(output_file, dpi=200)
    return figure


def add_chart_tab(notebook: ttk.Notebook, title: str, description: str, figure: plt.Figure) -> None:
    frame = ttk.Frame(notebook)
    notebook.add(frame, text=title)

    description_label = ttk.Label(frame, text=description, font=("Arial", 10))
    description_label.pack(pady=(10, 0))

    canvas = FigureCanvasTkAgg(figure, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=12, pady=12)


def show_chart_window(bar_figure: plt.Figure, pie_figure: plt.Figure, line_figure: plt.Figure) -> None:
    window = tk.Tk()
    window.title("Truc quan hoa chi tieu ca nhan")
    window.geometry("1100x780")

    title_label = ttk.Label(
        window,
        text="Truc quan hoa chi tieu bang 3 bieu do",
        font=("Arial", 14, "bold"),
    )
    title_label.pack(pady=(12, 4))

    subtitle_label = ttk.Label(
        window,
        text="Dong cua so nay de ket thuc chuong trinh.",
        font=("Arial", 10),
    )
    subtitle_label.pack(pady=(0, 8))

    notebook = ttk.Notebook(window)
    notebook.pack(fill="both", expand=True, padx=12, pady=12)

    # Moi tab trinh bay mot goc nhin khac nhau cua cung bo du lieu.
    add_chart_tab(
        notebook,
        "Bar chart",
        "Bieu do quan trong nhat: so sanh tong chi tieu giua cac danh muc.",
        bar_figure,
    )
    add_chart_tab(
        notebook,
        "Pie chart",
        "The hien ty trong chi tieu cua tung danh muc trong tong ngan sach.",
        pie_figure,
    )
    add_chart_tab(
        notebook,
        "Line chart",
        "Theo doi xu huong bien dong chi tieu qua tung thang.",
        line_figure,
    )

    window.mainloop()
    plt.close(bar_figure)
    plt.close(pie_figure)
    plt.close(line_figure)


def build_report(data_frame: pd.DataFrame, cleaning_info: dict[str, int]) -> str:
    # Bao cao tong hop duoc ghep tu cac ket qua tinh toan de co the nop doc lap voi terminal.
    total_expense = float(data_frame["amount"].sum())
    average_expense = float(np.mean(data_frame["amount"]))
    median_expense = float(np.median(data_frame["amount"]))
    max_expense = float(np.max(data_frame["amount"]))
    min_expense = float(np.min(data_frame["amount"]))
    category_totals = summarize_by_category(data_frame)
    monthly_totals = summarize_by_month(data_frame)
    forecast_totals, forecast_slope = forecast_monthly_expenses(monthly_totals)
    evaluation_result = evaluate_forecast_model(monthly_totals)
    biggest = top_expenses(data_frame)
    unusual_expenses = detect_unusual_expenses(data_frame)
    chart_insights = visualization_insights(category_totals, monthly_totals)

    lines: list[str] = []
    lines.append("BAI TAP LON: PHAN TICH DU LIEU CHI TIEU CA NHAN")
    lines.append("=" * 55)
    lines.append("1. Quy trinh phan tich du lieu")
    lines.append("- Xac dinh muc tieu: tim nhom chi tieu lon, xu huong chi tieu va kha nang tang chi phi trong tuong lai.")
    lines.append("- Thu thap du lieu tu tep Excel/CSV va dua vao DataFrame de xu ly.")
    lines.append("- Lam sach du lieu, chuan hoa kieu du lieu va loai bo ban ghi khong hop le.")
    lines.append("- Tong hop, truc quan hoa va dua ra du bao chi tieu cho cac thang tiep theo.")
    lines.append("")
    lines.append("2. Mo ta de tai")
    lines.append("De tai su dung pandas va numpy de phan tich cac khoan chi tieu ca nhan,")
    lines.append("tu do nhan dien nhom chi phi lon, xu huong theo thoi gian va cac giao dich bat thuong.")
    lines.append("")
    lines.append("3. Doc va lam sach du lieu")
    lines.append(f"So dong ban dau: {cleaning_info['original_rows']}")
    lines.append(f"So dong trung lap da xoa: {cleaning_info['duplicate_rows_removed']}")
    lines.append(f"So dong loi du lieu da xoa: {cleaning_info['invalid_rows_removed']}")
    lines.append(f"So dong khong phai chi tieu debit da loai: {cleaning_info['non_expense_rows_removed']}")
    lines.append(f"So dong chuyen khoan/tra the da loai: {cleaning_info['transfer_rows_removed']}")
    lines.append(f"So dong gia tri am da xoa: {cleaning_info['negative_rows_removed']}")
    lines.append(f"So dong hop le sau lam sach: {cleaning_info['clean_rows']}")
    lines.append("- Chuyen cot date sang kieu datetime va amount sang so de san sang phan tich.")
    lines.append("")
    lines.append("4. Tong quan du lieu")
    lines.append(f"So giao dich: {len(data_frame)}")
    lines.append(f"Tong chi tieu: {format_currency(total_expense)}")
    lines.append(f"Chi tieu trung binh moi giao dich (NumPy mean): {format_currency(average_expense)}")
    lines.append(f"Trung vi chi tieu: {format_currency(median_expense)}")
    lines.append(f"Giao dich lon nhat (NumPy max): {format_currency(max_expense)}")
    lines.append(f"Giao dich nho nhat (NumPy min): {format_currency(min_expense)}")
    lines.append("")
    lines.append("5. Phan tich theo danh muc")
    for row in category_totals.itertuples(index=False):
        lines.append(
            f"- {row.category}: {format_currency(row.amount)} ({row.ratio:.2f}%). "
            f"{category_comment(row.category)}"
        )
    lines.append("")
    lines.append("6. Phan tich theo thang")
    for row in monthly_totals.itertuples(index=False):
        lines.append(f"- {row.month}: {format_currency(row.amount)}")
    lines.append("")
    lines.append("7. Truc quan hoa du lieu")
    lines.append(f"- Bar chart (quan trong nhat) da duoc tao tai: {BAR_CHART_FILE}")
    lines.append(f"- Pie chart da duoc tao tai: {PIE_CHART_FILE}")
    lines.append(f"- Line chart da duoc tao tai: {LINE_CHART_FILE}")
    lines.append("- Line chart da duoc bo sung duong du bao bang net dut cho 3 thang tiep theo.")
    lines.append("- Nhan xet cac khia canh truc quan hoa:")
    for insight in chart_insights:
        lines.append(f"  {insight}")
    lines.append("")
    lines.append("8. Cac giao dich co gia tri cao nhat")
    for row in biggest.itertuples(index=False):
        lines.append(
            f"- {row.date.strftime('%d/%m/%Y')} | {row.category} | {row.description} | "
            f"{format_currency(row.amount)}"
        )
    lines.append("")
    lines.append("9. Giao dich bat thuong")
    if unusual_expenses.empty:
        lines.append("- Khong phat hien giao dich nao vuot nguong bat thuong.")
    else:
        threshold = float(unusual_expenses["threshold"].iloc[0])
        lines.append(f"- Nguong bat thuong (trung binh + do lech chuan): {format_currency(threshold)}")
        for row in unusual_expenses.itertuples(index=False):
            lines.append(
                f"- {row.date.strftime('%d/%m/%Y')} | {row.category} | {row.description} | "
                f"{format_currency(row.amount)}"
            )
    lines.append("")
    lines.append("10. Phan tich du lieu du bao")
    # Ghi ro phuong phap du bao de khi doc bao cao co the giai thich cach uoc luong.
    lines.append("- Mo hinh su dung: hoi quy tuyen tinh (Linear Regression) bang numpy.polyfit.")
    lines.append("- Dau vao mo hinh: tong chi tieu theo thang sau khi da lam sach va tong hop du lieu.")
    if evaluation_result["train_months"]:
        train_months = evaluation_result["train_months"]
        test_months = evaluation_result["test_months"]
        lines.append(
            f"- Tap huan luyen: {len(train_months)} thang, tu {train_months[0]} den {train_months[-1]}."
        )
        lines.append(
            f"- Tap kiem thu: {len(test_months)} thang, tu {test_months[0]} den {test_months[-1]}."
        )
    else:
        lines.append("- Tap huan luyen va kiem thu chua duoc tach do chuoi du lieu theo thang qua ngan.")
    lines.append("- Phuong phap: uoc luong duong xu huong tuyen tinh tren 6 thang gan nhat de du bao 3 thang tiep theo.")
    lines.append(f"- Do doc xu huong uoc luong: {format_currency(forecast_slope)} moi thang.")
    for row in forecast_totals.itertuples(index=False):
        lines.append(f"- Du bao {row.month}: {format_currency(row.forecast_amount)}")
    if forecast_slope >= 0:
        lines.append("- Ket qua du bao cho thay chi tieu co xu huong tang dan, can lap ngan sach som cho cac khoan co dinh.")
    else:
        lines.append("- Ket qua du bao cho thay chi tieu co xu huong giam nhe, nhung van can theo doi cac khoan co dinh de tranh tang tro lai.")
    lines.append("- Danh gia mo hinh: mo hinh de hieu, de trien khai va phu hop voi bo du lieu nho.")
    if evaluation_result["predictions"].empty:
        lines.append(f"- Ket qua kiem thu: {evaluation_result['note']}")
    else:
        lines.append("- Ket qua du bao tren tap kiem thu:")
        for row in evaluation_result["predictions"].itertuples(index=False):
            lines.append(
                f"  {row.month}: thuc te {format_currency(row.actual_amount)}, "
                f"du bao {format_currency(row.predicted_amount)}"
            )
        lines.append(f"- MAE: {format_currency(evaluation_result['mae'])}")
        lines.append(f"- RMSE: {format_currency(evaluation_result['rmse'])}")
        if evaluation_result["r2"] is None:
            lines.append("- R2: khong tinh duoc vi tap kiem thu khong co do bien thien.")
        else:
            lines.append(f"- R2: {evaluation_result['r2']:.4f}")
    lines.append("- Han che mo hinh: mo hinh tuyen tinh co the bo sot tinh mua vu, bien dong dot bien va anh huong cua tung loai chi tieu rieng.")
    lines.append("")
    lines.append("11. Nhan xet")
    highest_category = category_totals.iloc[0]
    highest_month = monthly_totals.loc[monthly_totals["amount"].idxmax()]
    lines.append(
        f"- Danh muc chi tieu lon nhat la {highest_category['category']} voi "
        f"{format_currency(highest_category['amount'])}."
    )
    lines.append(
        f"- Thang chi tieu cao nhat la {highest_month['month']} voi "
        f"{format_currency(highest_month['amount'])}."
    )
    lines.append("- Nen dat ngan sach rieng cho nhom giai tri va mua sam de tranh vuot muc cho phep.")
    lines.append("- Nha o va an uong la cac khoan thiet yeu, can duoc uu tien truoc khi lap ke hoach tiet kiem.")
    lines.append("")
    lines.append("12. Ket luan")
    lines.append("- Bai toan da thuc hien du 4 yeu cau: quy trinh phan tich du lieu, xu ly du lieu, truc quan hoa va phan tich du bao.")
    lines.append("- Bar chart la bieu do hieu qua nhat de so sanh cac danh muc, trong khi pie chart va line chart bo tro ve co cau, xu huong va du bao.")
    lines.append("- Mo hinh hoi quy tuyen tinh hien tai phu hop de minh hoa xu huong ngan han, nhung can tach train/test va bo sung chi so danh gia neu muon bao cao theo huong hoc may day du.")
    lines.append("- Co the mo rong phan tich theo quy, theo phuong thuc thanh toan va ap dung mo hinh du bao nang cao hon.")
    return "\n".join(lines)


def print_dashboard(data_frame: pd.DataFrame, cleaning_info: dict[str, int]) -> None:
    # Dashboard terminal giup xem nhanh ket qua ma khong can mo file report.
    total_expense = float(data_frame["amount"].sum())
    average_expense = float(np.mean(data_frame["amount"]))
    max_expense = float(np.max(data_frame["amount"]))
    min_expense = float(np.min(data_frame["amount"]))
    category_totals = summarize_by_category(data_frame)
    monthly_totals = summarize_by_month(data_frame)
    forecast_totals, _ = forecast_monthly_expenses(monthly_totals)
    evaluation_result = evaluate_forecast_model(monthly_totals)
    unusual_expenses = detect_unusual_expenses(data_frame)

    print("PHAN TICH CHI TIEU CA NHAN BANG PANDAS VA NUMPY")
    print("=" * 48)
    print(f"So dong sau lam sach : {cleaning_info['clean_rows']}/{cleaning_info['original_rows']}")
    print(f"So giao dich         : {len(data_frame)}")
    print(f"Tong chi tieu        : {format_currency(total_expense)}")
    print(f"Chi tieu trung binh  : {format_currency(average_expense)}")
    print(f"Giao dich lon nhat   : {format_currency(max_expense)}")
    print(f"Giao dich nho nhat   : {format_currency(min_expense)}")
    print(
        f"Danh muc cao nhat    : {category_totals.iloc[0]['category']} "
        f"({format_currency(category_totals.iloc[0]['amount'])})"
    )
    print()
    print("Chi tieu theo danh muc")
    for row in category_totals.itertuples(index=False):
        print(f"- {row.category:<10} {format_currency(row.amount)} | {row.ratio:.2f}%")
    print()
    print("Chi tieu theo thang")
    for row in monthly_totals.itertuples(index=False):
        print(f"- {row.month}: {format_currency(row.amount)}")
    print()
    # In them phan du bao de nguoi xem thay bai khong chi mo ta qua khu ma con uoc luong tuong lai.
    print("Du bao 3 thang tiep theo")
    for row in forecast_totals.itertuples(index=False):
        print(f"- {row.month}: {format_currency(row.forecast_amount)}")
    print()
    if evaluation_result["predictions"].empty:
        print(f"Kiem thu mo hinh      : {evaluation_result['note']}")
    else:
        print(
            f"Tap huan luyen        : {len(evaluation_result['train_months'])} thang | "
            f"Tap kiem thu: {len(evaluation_result['test_months'])} thang"
        )
        print(f"MAE                   : {format_currency(evaluation_result['mae'])}")
        print(f"RMSE                  : {format_currency(evaluation_result['rmse'])}")
        if evaluation_result["r2"] is None:
            print("R2                    : Khong tinh duoc")
        else:
            print(f"R2                    : {evaluation_result['r2']:.4f}")
    print()
    print(f"So giao dich bat thuong: {len(unusual_expenses)}")
    print(f"Bar chart da duoc luu tai : {BAR_CHART_FILE}")
    print(f"Pie chart da duoc luu tai : {PIE_CHART_FILE}")
    print(f"Line chart da duoc luu tai: {LINE_CHART_FILE}")
    print(f"Bao cao chi tiet da duoc ghi vao: {REPORT_FILE}")


def export_four_sheet_workbook(source_file: Path, data_frame: pd.DataFrame, cleaning_info: dict[str, int]) -> Path:
    # Dung lai module xuat Excel de khi chay main.py cung tao file .xlsx 4 sheet.
    import sys

    sys.modules.setdefault("main", sys.modules[__name__])
    from scripts import export_analysis_xlsx

    raw_data = read_transaction_file(source_file)
    stats = export_analysis_xlsx.make_stats_frame(data_frame, cleaning_info)
    category_analysis = export_analysis_xlsx.make_category_analysis_frame(data_frame)
    export_analysis_xlsx.write_xlsx(
        EXCEL_REPORT_FILE,
        [
            ("dữ liệu gốc", export_analysis_xlsx.dataframe_rows(raw_data)),
            ("dữ liệu đã làm sạch", export_analysis_xlsx.dataframe_rows(data_frame)),
            ("thống kê", export_analysis_xlsx.dataframe_rows(stats)),
            ("phân tích danh mục", export_analysis_xlsx.dataframe_rows(category_analysis)),
        ],
    )
    return EXCEL_REPORT_FILE


def main() -> None:
    data_file = DATA_FILE if DATA_FILE.exists() else FALLBACK_DATA_FILE
    if not data_file.exists():
        raise FileNotFoundError(f"Khong tim thay tep du lieu: {DATA_FILE} hoac {FALLBACK_DATA_FILE}")

    # Luong xu ly chinh: doc du lieu -> tong hop -> ve bieu do -> ghi bao cao -> hien thi giao dien.
    data_frame, cleaning_info = load_expenses(data_file)
    category_totals = summarize_by_category(data_frame)
    monthly_totals = summarize_by_month(data_frame)
    forecast_totals, _ = forecast_monthly_expenses(monthly_totals)
    bar_chart_figure = create_chart(build_bar_chart_figure(category_totals), BAR_CHART_FILE)
    pie_chart_figure = create_chart(build_pie_chart_figure(category_totals), PIE_CHART_FILE)
    line_chart_figure = create_chart(build_line_chart_figure(monthly_totals, forecast_totals), LINE_CHART_FILE)
    report = build_report(data_frame, cleaning_info)
    REPORT_FILE.write_text(report, encoding="utf-8")
    excel_report_file = export_four_sheet_workbook(data_file, data_frame, cleaning_info)
    print_dashboard(data_frame, cleaning_info)
    print(f"File Excel 4 sheet da duoc xuat tai: {excel_report_file}")
    show_chart_window(bar_chart_figure, pie_chart_figure, line_chart_figure)


if __name__ == "__main__":
    main()
