from __future__ import annotations

from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_FILE = BASE_DIR / "data" / "expenses.csv"
TARGET_ROWS = 1000


CATEGORY_CONFIGS = [
    {
        "category": "Nha o",
        "descriptions": [
            "Tien tro hang thang",
            "Tien dien nuoc",
            "Phi internet",
            "Bao tri phong tro",
        ],
        "base_amount": 420000,
        "amount_step": 28000,
        "payment_methods": ["Chuyen khoan", "The"],
        "notes": [
            "Chi phi sinh hoat can thiet",
            "Thanh toan dinh ky",
            "Phat sinh trong thang",
        ],
    },
    {
        "category": "An uong",
        "descriptions": [
            "Mua thuc pham",
            "An trua van phong",
            "An toi cuoi tuan",
            "Dat do an online",
            "Ca phe buoi sang",
        ],
        "base_amount": 85000,
        "amount_step": 9000,
        "payment_methods": ["Tien mat", "The", "Vi dien tu"],
        "notes": [
            "Chi phi hang ngay",
            "Du tru bua an trong tuan",
            "Phuc vu lich hoc va lam",
        ],
    },
    {
        "category": "Di chuyen",
        "descriptions": [
            "Do xang xe",
            "Gui xe",
            "Di xe cong nghe",
            "Bao duong xe may",
        ],
        "base_amount": 70000,
        "amount_step": 11000,
        "payment_methods": ["Tien mat", "Vi dien tu", "Chuyen khoan"],
        "notes": [
            "Di hoc va di lam",
            "Phat sinh di lai trong thang",
            "Bao tri phuong tien",
        ],
    },
    {
        "category": "Hoc tap",
        "descriptions": [
            "Mua sach",
            "In tai lieu",
            "Hoc phi khoa ngan han",
            "Le phi thi",
            "Mua khoa hoc online",
        ],
        "base_amount": 120000,
        "amount_step": 35000,
        "payment_methods": ["Chuyen khoan", "The"],
        "notes": [
            "Dau tu cho ky nang",
            "Phuc vu bai tap va bao cao",
            "Nang cao kien thuc",
        ],
    },
    {
        "category": "Giai tri",
        "descriptions": [
            "Xem phim",
            "Ca phe voi ban be",
            "Dang ky nen tang giai tri",
            "Du lich ngan ngay",
            "Ve su kien",
        ],
        "base_amount": 95000,
        "amount_step": 20000,
        "payment_methods": ["The", "Vi dien tu", "Tien mat"],
        "notes": [
            "Thu gian cuoi tuan",
            "Chi phi giai tri theo thang",
            "Can theo doi de tranh vuot ngan sach",
        ],
    },
    {
        "category": "Suc khoe",
        "descriptions": [
            "Mua thuoc",
            "Kham tong quat",
            "Kham rang",
            "Mua vitamin",
        ],
        "base_amount": 100000,
        "amount_step": 26000,
        "payment_methods": ["Tien mat", "The", "Chuyen khoan"],
        "notes": [
            "Cham soc suc khoe dinh ky",
            "Du phong cho suc khoe",
            "Phat sinh khong deu",
        ],
    },
    {
        "category": "Mua sam",
        "descriptions": [
            "Mua do dung ca nhan",
            "Mua quan ao",
            "Mua phu kien hoc tap",
            "Mua do gia dung nho",
        ],
        "base_amount": 110000,
        "amount_step": 30000,
        "payment_methods": ["The", "Vi dien tu", "Tien mat"],
        "notes": [
            "Bo sung vat dung can thiet",
            "Can can nhac truoc khi mua",
            "Phuc vu sinh hoat ca nhan",
        ],
    },
]


def build_row(index: int) -> str:
    # Xoay vong qua cac cau hinh de tao du lieu lon nhung van giu phan bo category da dang.
    config = CATEGORY_CONFIGS[index % len(CATEGORY_CONFIGS)]
    month = (index % 12) + 1
    day = ((index * 3) % 28) + 1
    year = 2025 + (index // 365)
    description = config["descriptions"][index % len(config["descriptions"])]
    payment_method = config["payment_methods"][index % len(config["payment_methods"])]
    note = config["notes"][index % len(config["notes"])]
    amount = config["base_amount"] + (month * config["amount_step"]) + ((index // 7) % 5) * 17000
    return (
        f"{year}-{month:02d}-{day:02d},"
        f"{config['category']},"
        f"{description},"
        f"{amount},"
        f"{payment_method},"
        f"{note}"
    )


def main() -> None:
    rows = ["date,category,description,amount,payment_method,note"]
    # Sinh dong du lieu tu dong de co the tang quy mo dataset chi bang cach doi TARGET_ROWS.
    rows.extend(build_row(index) for index in range(TARGET_ROWS))
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    DATA_FILE.write_text("\n".join(rows) + "\n", encoding="utf-8")
    print(f"Wrote {TARGET_ROWS} rows to {DATA_FILE}")


if __name__ == "__main__":
    main()
