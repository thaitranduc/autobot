import re

import openpyxl

from vietlott import fetch_ketquadientoan_history, fetch_latest_vietlott_mega645

SHEET_NAME = "Lịch Sử Số Trúng (1200+ Kỳ)"
HEADER = [
    "Kỳ Quay",
    "Ngày Quay",
    "Thứ",
    "Số 1",
    "Số 2",
    "Số 3",
    "Số 4",
    "Số 5",
    "Số 6",
    "Jackpot (VNĐ)",
    "Số Vé Trúng JP",
]


def draw_id_to_number(draw_id):
    if draw_id is None:
        return 0
    match = re.search(r"(\d+)", str(draw_id))
    return int(match.group(1)) if match else 0


def rebuild_workbook(file_path="Vietlott_Mega_645_Full_Results.xlsx"):
    archive_rows = fetch_ketquadientoan_history("20-07-2016", "20-09-2026")
    latest_row = fetch_latest_vietlott_mega645()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADER)

    for row in archive_rows:
        ws.append([
            row["draw_id"],
            row["date"],
            row["weekday"],
            *row["nums"],
            row["jackpot"],
            row["winners"],
        ])

    last_draw_id = ws.cell(ws.max_row, 1).value if ws.max_row > 1 else None
    if last_draw_id != latest_row["draw_id"]:
        ws.append([
            latest_row["draw_id"],
            latest_row["date"],
            latest_row["weekday"],
            *latest_row["nums"],
            latest_row["jackpot"],
            latest_row["winners"],
        ])

    wb.save(file_path)
    print(f"Workbook rebuilt from archive page with {len(archive_rows)} archive rows + live row.")
    print(f"First draw: {ws.cell(2, 1).value} / {ws.cell(2, 2).value}")
    print(f"Last draw: {ws.cell(ws.max_row, 1).value} / {ws.cell(ws.max_row, 2).value}")


if __name__ == "__main__":
    rebuild_workbook()
