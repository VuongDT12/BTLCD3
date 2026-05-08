# Phan tich du lieu chi tieu ca nhan

Day la bai tap lon Python su dung `pandas` va `numpy` de phan tich du lieu chi tieu ca nhan tu tep Excel.

## Muc tieu

- Trinh bay quy trinh phan tich du lieu tu xac dinh muc tieu den ket luan.
- Doc du lieu giao dich tu tep Excel.
- Xu ly du lieu bang `DataFrame`.
- Loc giao dich `debit` va loai khoan tra the/chuyen khoan de phan tich dung chi tieu thuc.
- Tong hop chi phi theo danh muc va theo thang.
- Tim cac giao dich co gia tri lon nhat.
- Phat hien giao dich bat thuong bang cac chi so thong ke co ban.
- Du bao chi tieu 3 thang tiep theo dua tren xu huong tong chi tieu theo thang.
- Sinh bao cao tong hop de nop bai.

## Cau truc

- `main.py`: Chuong trinh chinh thuc hien phan tich.
- `data/personal_transactions_dashboard_ready (2).xlsx`: Du lieu giao dich ca nhan dung cho dashboard.
- `scripts/generate_expenses.py`: Script sinh bo du lieu mau so luong lon de phuc vu phan tich.
- `charts/expense_bar_chart.png`: Bar chart, bieu do quan trong nhat de so sanh tong chi tieu theo danh muc.
- `charts/expense_pie_chart.png`: Pie chart the hien ty trong chi tieu theo danh muc.
- `charts/expense_line_chart.png`: Line chart the hien xu huong chi tieu theo thang.
- `report.txt`: Tep bao cao duoc sinh ra sau khi chay chuong trinh.

## Thu vien su dung

- `pandas`: lam sach va tong hop du lieu.
- `zipfile` va `xml.etree.ElementTree`: doc file Excel don gian khi moi truong chua cai `openpyxl`.
- `numpy`: tinh trung binh, trung vi, do lech chuan va xac dinh nguong bat thuong.
- `matplotlib`: ve 3 bieu do bar, pie, line de truc quan hoa du lieu.
- `tkinter`: hien thi cua so gom 3 tab bieu do tren man hinh.
- `numpy.polyfit`: uoc luong xu huong de du bao chi tieu ngan han.

## Cach chay

```powershell
.venv\Scripts\python.exe main.py
```

Sau khi chay, chuong trinh se:

- in ket qua phan tich ra terminal
- tao cac file `charts/expense_bar_chart.png`, `charts/expense_pie_chart.png`, `charts/expense_line_chart.png`
- ghi them phan du bao 3 thang tiep theo vao `report.txt`
- mo cua so `tkinter` de hien thi 3 bieu do, trong do `Bar chart` la bieu do chinh

## Goi y trinh bay bao cao

Ban co the trinh bay bai tap lon theo bo cuc:

1. Ly do chon de tai.
2. Mo ta bo du lieu.
3. Quy trinh va cac buoc xu ly du lieu.
4. Ket qua truc quan hoa va phan tich.
5. Phan tich du lieu du bao.
6. Nhan xet va de xuat quan ly chi tieu.
