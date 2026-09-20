import datetime
import html as html_lib
import re
import sqlite3
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import requests


def parse_mega645_html(html_text):
    """Parse the latest Mega 6/45 result from the HTML page when the API is unavailable."""
    if not html_text:
        return None

    text = html_lib.unescape(html_text)
    text = text.replace("&nbsp;", " ").replace("\xa0", " ")
    text = re.sub(r"<[^>]+>", " ", text, flags=re.S)
    text = re.sub(r"\s+", " ", text).strip()

    row_match = re.search(
        r"(\d{2}/\d{2}/\d{4})\s*\|\s*0*(\d+)\s*\|\s*(?P<nums>(?:\d{1,2}\s+){5}\d{1,2})",
        text,
        re.I,
    )
    if row_match:
        draw_date_str = row_match.group(1)
        draw_id = f"#{int(row_match.group(2)):05d}"
        nums = [int(x) for x in row_match.group("nums").split()]
        dt = datetime.datetime.strptime(draw_date_str, "%d/%m/%Y")
        weekdays = [
            "Thứ 2",
            "Thứ 3",
            "Thứ 4",
            "Thứ 5",
            "Thứ 6",
            "Thứ 7",
            "Chủ Nhật",
        ]
        weekday = weekdays[dt.weekday()]
        return {
            "draw_id": draw_id,
            "date": draw_date_str,
            "weekday": weekday,
            "nums": sorted(nums),
            "jackpot": 0,
            "winners": 0,
        }

    direct_match = re.search(
        r"(\d{2}/\d{2}/\d{4})\s*#?0*(\d+)\s*(?P<nums>(?:\d{1,2}\s+){5}\d{1,2})",
        text,
        re.I,
    )
    if direct_match:
        draw_date_str = direct_match.group(1)
        draw_id = f"#{int(direct_match.group(2)):05d}"
        nums = [int(x) for x in direct_match.group("nums").split()]
        dt = datetime.datetime.strptime(draw_date_str, "%d/%m/%Y")
        weekdays = [
            "Thứ 2",
            "Thứ 3",
            "Thứ 4",
            "Thứ 5",
            "Thứ 6",
            "Thứ 7",
            "Chủ Nhật",
        ]
        weekday = weekdays[dt.weekday()]
        return {
            "draw_id": draw_id,
            "date": draw_date_str,
            "weekday": weekday,
            "nums": sorted(nums),
            "jackpot": 0,
            "winners": 0,
        }

    live_rows = re.finditer(
        r"<td[^>]*>\s*(\d{1,2}/\d{1,2}/\d{4})\s*</td>\s*"
        r"<td[^>]*>\s*(?:<a[^>]*>)?\s*#?0*(\d+)\s*(?:</a>)?\s*</td>\s*"
        r"<td[^>]*>(.*?)</td>",
        text,
        re.I | re.S,
    )
    for match in live_rows:
        draw_date_str = match.group(1)
        draw_id = f"#{int(match.group(2)):05d}"
        cell_html = match.group(3)
        nums = [int(x) for x in re.findall(r"<span[^>]*>\s*(\d{1,2})\s*</span>", cell_html, re.I)]
        if len(nums) == 6 and len(set(nums)) == 6:
            dt = datetime.datetime.strptime(draw_date_str, "%d/%m/%Y")
            weekdays = [
                "Thứ 2",
                "Thứ 3",
                "Thứ 4",
                "Thứ 5",
                "Thứ 6",
                "Thứ 7",
                "Chủ Nhật",
            ]
            weekday = weekdays[dt.weekday()]
            return {
                "draw_id": draw_id,
                "date": draw_date_str,
                "weekday": weekday,
                "nums": sorted(nums),
                "jackpot": 0,
                "winners": 0,
            }

    match_draw = re.search(r"Kết quả QSMT kỳ\s*#?(\d+)\s*ngày\s*(\d{2}/\d{2}/\d{4})", text, re.I)
    if not match_draw:
        return None

    draw_id = f"#{int(match_draw.group(1)):05d}"
    draw_date_str = match_draw.group(2)
    dt = datetime.datetime.strptime(draw_date_str, "%d/%m/%Y")
    weekdays = [
        "Thứ 2",
        "Thứ 3",
        "Thứ 4",
        "Thứ 5",
        "Thứ 6",
        "Thứ 7",
        "Chủ Nhật",
    ]
    weekday = weekdays[dt.weekday()]

    num_block = re.search(
        r"Kết quả QSMT kỳ.*?(?P<nums>(?:\d{2}\s+){5}\d{2})",
        text,
        re.I | re.S,
    )
    if not num_block:
        return None

    nums = [int(x) for x in num_block.group("nums").split()]

    jackpot_match = re.search(
        r"Jackpot\s+Mega\s+6/45\s+ước\s+tính.*?(?P<jackpot>[\d\.,]+)\s*VNĐ|(?P<jackpot2>[\d\.,]+)\s*VNĐ.*?Jackpot\s+Mega\s+6/45",
        text,
        re.I | re.S,
    )
    jackpot_val = 0
    if jackpot_match:
        jackpot_text = jackpot_match.group("jackpot") or jackpot_match.group("jackpot2")
        if jackpot_text:
            jackpot_val = int(jackpot_text.replace(".", "").replace(",", ""))

    return {
        "draw_id": draw_id,
        "date": draw_date_str,
        "weekday": weekday,
        "nums": sorted(nums),
        "jackpot": jackpot_val,
        "winners": 0,
    }


def parse_minhngoc_mega645_html(html_text):
    """Parse the Mega 6/45 block published on Minh Ngoc."""
    if not html_text:
        return None

    text = html_lib.unescape(html_text).replace("&nbsp;", " ").replace("\xa0", " ")
    draw_match = re.search(
        r"DT6X45_KY_VE[^>]*>\s*#?(\d+).*?"
        r"Ngày quay thưởng\s*(\d{1,2}/\d{1,2}/\d{4})",
        text,
        re.I | re.S,
    )
    if not draw_match:
        return None

    numbers_block = re.search(
        r'<ul[^>]*class=["\'][^"\']*result-number[^"\']*["\'][^>]*>'
        r"(.*?)(?:</ul>)",
        text,
        re.I | re.S,
    )
    if not numbers_block:
        return None

    nums = [
        int(value)
        for value in re.findall(
            r'class=["\'][^"\']*finnish[1-6][^"\']*["\'][^>]*>\s*(\d{1,2})\s*<',
            numbers_block.group(1),
            re.I,
        )
    ]
    if len(nums) != 6 or len(set(nums)) != 6 or not all(1 <= num <= 45 for num in nums):
        return None

    draw_date_str = draw_match.group(2)
    dt = datetime.datetime.strptime(draw_date_str, "%d/%m/%Y")
    weekdays = [
        "Thứ 2",
        "Thứ 3",
        "Thứ 4",
        "Thứ 5",
        "Thứ 6",
        "Thứ 7",
        "Chủ Nhật",
    ]
    jackpot_count_match = re.search(
        r'id=["\']DT6X45_S_JACKPOT["\'][^>]*>\s*([\d,.]+)',
        text,
        re.I,
    )
    jackpot_value_match = re.search(
        r'id=["\']DT6X45_G_JACKPOT["\'][^>]*>\s*([\d,.]+)',
        text,
        re.I,
    )

    def parse_amount(value):
        return int(value.replace(".", "").replace(",", "")) if value else 0

    return {
        "draw_id": f"#{int(draw_match.group(1)):05d}",
        "date": draw_date_str,
        "weekday": weekdays[dt.weekday()],
        "nums": sorted(nums),
        "jackpot": parse_amount(jackpot_value_match.group(1) if jackpot_value_match else ""),
        "winners": parse_amount(jackpot_count_match.group(1) if jackpot_count_match else ""),
    }


def add_manual_result(file_path, draw_id, date_str, nums, jackpot, winners=0):
    """Thêm dữ liệu kỳ quay theo cách thủ công (không cần API)."""
    # Lấy Thứ trong tuần
    dt = datetime.datetime.strptime(date_str, "%d/%m/%Y")
    weekdays = [
        "Thứ 2",
        "Thứ 3",
        "Thứ 4",
        "Thứ 5",
        "Thứ 6",
        "Thứ 7",
        "Chủ Nhật",
    ]
    weekday = weekdays[dt.weekday()]

    draw_data = {
        "draw_id": draw_id,
        "date": date_str,
        "weekday": weekday,
        "nums": sorted(nums),
        "jackpot": jackpot,
        "winners": winners,
    }
    
    append_to_excel(file_path, draw_data)


def fetch_latest_vietlott_mega645():
    """Lấy kết quả Mega 6/45 mới nhất. Hỗ trợ JSON cũ, HTML hiện tại và nhiều biến thể layout."""
    urls = [
        "https://www.minhngoc.net.vn/ket-qua-xo-so/dien-toan-vietlott.html",
        "https://vietlott.vn/api/front/get-result-mega645",
        "https://vietlott.vn/vi/trung-thuong/ket-qua-trung-thuong/645",
        "https://vietlott.vn/vi/trung-thuong/ket-qua-trung-thuong/mega-6-45",
        "https://vietlott.vn/vi/trung-thuong/ket-qua-trung-thuong/winning-number-645",
        "https://vietlott.vn/",
    ]
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
        "Referer": "https://vietlott.vn/",
    }

    for url in urls:
        try:
            print(f"Trying URL: {url}")
            res = requests.get(url, headers=headers, timeout=20)
            res.raise_for_status()
            content_type = res.headers.get("Content-Type", "")
            print(f"Status Code: {res.status_code}")
            print(f"Content-Type: {content_type}")

            if "application/json" in content_type:
                data = res.json()
                latest = data.get("result", {})
                if not latest:
                    print("⚠️ API JSON trả về nhưng không có trường result.")
                    continue

                draw_id = f"#{int(latest['drawId']):05d}"
                draw_date_str = latest["drawDate"]
                nums = [int(n) for n in latest["winningNumbers"]]
                jackpot_val = int(latest["jackpotAmount"])
                jackpot_winners = int(latest["jackpotWinners"])

                dt = datetime.datetime.strptime(draw_date_str, "%d/%m/%Y")
                weekdays = [
                    "Thứ 2",
                    "Thứ 3",
                    "Thứ 4",
                    "Thứ 5",
                    "Thứ 6",
                    "Thứ 7",
                    "Chủ Nhật",
                ]
                weekday = weekdays[dt.weekday()]

                return {
                    "draw_id": draw_id,
                    "date": draw_date_str,
                    "weekday": weekday,
                    "nums": sorted(nums),
                    "jackpot": jackpot_val,
                    "winners": jackpot_winners,
                }

            minhngoc_parsed = parse_minhngoc_mega645_html(res.text)
            if minhngoc_parsed:
                print(
                    "✅ Parsed Minh Ngoc Mega 6/45: "
                    f"{minhngoc_parsed['draw_id']} {minhngoc_parsed['date']}"
                )
                return minhngoc_parsed

            parsed = parse_mega645_html(res.text)
            if parsed:
                print(f"✅ Parse HTML thành công: {parsed['draw_id']} {parsed['date']}")
                return parsed

            print("⚠️ HTML response does not contain a Mega 6/45 result block on this URL.")
        except Exception as e:
            print(f"Lỗi khi lấy dữ liệu Vietlott từ {url}: {e}")

    error_message = (
        "Không thể lấy được dữ liệu Mega 6/45 từ mọi URL fallback. "
        "The upstream sources may be blocking GitHub Actions or have changed."
    )
    print(f"❌ {error_message}")
    raise RuntimeError(error_message)


def init_sqlite_db(db_path="vietlott.db"):
    """Create the SQLite database and table used for lightweight storage."""
    db_file = Path(db_path)
    db_file.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(db_file)
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS results (
            draw_id TEXT PRIMARY KEY,
            draw_date TEXT NOT NULL,
            weekday TEXT NOT NULL,
            n1 INTEGER NOT NULL,
            n2 INTEGER NOT NULL,
            n3 INTEGER NOT NULL,
            n4 INTEGER NOT NULL,
            n5 INTEGER NOT NULL,
            n6 INTEGER NOT NULL,
            jackpot INTEGER DEFAULT 0,
            winners INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    connection.commit()
    connection.close()
    return str(db_file)


def draw_id_to_number(draw_id):
    """Convert a draw id like '#01564' into its numeric value."""
    if not draw_id:
        return 0
    match = re.search(r"(\d+)", str(draw_id))
    return int(match.group(1)) if match else 0


def save_result_to_sqlite(db_path, draw_data):
    """Insert or update a Mega 6/45 result in SQLite. Returns True when saved."""
    db_file = Path(db_path)
    db_file.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(db_file)
    cursor = connection.cursor()

    incoming_num = draw_id_to_number(draw_data.get("draw_id"))
    if incoming_num <= 0:
        connection.close()
        return False

    cursor.execute(
        "DELETE FROM results WHERE CAST(substr(draw_id, 2) AS INTEGER) < ?",
        (incoming_num,),
    )

    existing_max = cursor.execute(
        "SELECT MAX(CAST(substr(draw_id, 2) AS INTEGER)) FROM results"
    ).fetchone()[0]
    existing_max = int(existing_max) if existing_max is not None else 0
    if incoming_num < existing_max:
        connection.close()
        return False

    nums = list(draw_data["nums"])[:6]
    while len(nums) < 6:
        nums.append(0)

    cursor.execute(
        """
        INSERT INTO results (
            draw_id, draw_date, weekday, n1, n2, n3, n4, n5, n6, jackpot, winners
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(draw_id) DO UPDATE SET
            draw_date = excluded.draw_date,
            weekday = excluded.weekday,
            n1 = excluded.n1,
            n2 = excluded.n2,
            n3 = excluded.n3,
            n4 = excluded.n4,
            n5 = excluded.n5,
            n6 = excluded.n6,
            jackpot = excluded.jackpot,
            winners = excluded.winners
        """,
        (
            draw_data["draw_id"],
            draw_data["date"],
            draw_data["weekday"],
            nums[0],
            nums[1],
            nums[2],
            nums[3],
            nums[4],
            nums[5],
            int(draw_data.get("jackpot", 0) or 0),
            int(draw_data.get("winners", 0) or 0),
        ),
    )
    connection.commit()
    connection.close()
    return True


def render_results_html(db_path="vietlott.db", limit=20):
    """Return a simple HTML table with the latest stored Vietlott results."""
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    rows = connection.execute(
        """
        SELECT draw_id, draw_date, weekday, n1, n2, n3, n4, n5, n6, jackpot, winners
        FROM results
        ORDER BY CAST(substr(draw_id, 2) AS INTEGER) DESC
        LIMIT ?
        """,
        (limit,),
    ).fetchall()
    connection.close()

    html_rows = "\n".join(
        """
        <tr>
            <td>{draw_id}</td>
            <td>{draw_date}</td>
            <td>{weekday}</td>
            <td>{n1}</td>
            <td>{n2}</td>
            <td>{n3}</td>
            <td>{n4}</td>
            <td>{n5}</td>
            <td>{n6}</td>
            <td>{jackpot}</td>
            <td>{winners}</td>
        </tr>
        """.format(
            draw_id=row["draw_id"],
            draw_date=row["draw_date"],
            weekday=row["weekday"],
            n1=row["n1"],
            n2=row["n2"],
            n3=row["n3"],
            n4=row["n4"],
            n5=row["n5"],
            n6=row["n6"],
            jackpot=row["jackpot"],
            winners=row["winners"],
        )
        for row in rows
    )

    return f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Vietlott Results</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 24px; }}
            table {{ border-collapse: collapse; width: 100%; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
            th {{ background: #f4f4f4; }}
        </style>
    </head>
    <body>
        <h1>Vietlott Results</h1>
        <table>
            <thead>
                <tr>
                    <th>Draw</th>
                    <th>Date</th>
                    <th>Weekday</th>
                    <th>N1</th>
                    <th>N2</th>
                    <th>N3</th>
                    <th>N4</th>
                    <th>N5</th>
                    <th>N6</th>
                    <th>Jackpot</th>
                    <th>Winners</th>
                </tr>
            </thead>
            <tbody>
                {html_rows}
            </tbody>
        </table>
    </body>
    </html>
    """


def append_to_excel(file_path, draw_data):
    """Ghi thêm dữ liệu kỳ mới vào file Excel."""
    wb = openpyxl.load_workbook(file_path)
    ws_data = wb["Lịch Sử Số Trúng (1200+ Kỳ)"]

    incoming_num = draw_id_to_number(draw_data.get("draw_id"))
    max_draw_num = 0
    for row_idx in range(1, ws_data.max_row + 1):
        cell_value = ws_data.cell(row=row_idx, column=1).value
        if cell_value:
            max_draw_num = max(max_draw_num, draw_id_to_number(cell_value))

    if incoming_num and max_draw_num and incoming_num < max_draw_num:
        print(
            f"Bỏ qua kỳ quay cũ {draw_data['draw_id']} vì Excel đã có kỳ mới hơn "
            f"({max_draw_num})."
        )
        return

    # Skip an exact duplicate, but repair stale rows where an old source
    # reused the same draw ID with a different date.
    last_row = ws_data.max_row
    last_draw_id = ws_data.cell(row=last_row, column=1).value

    if (
        last_draw_id == draw_data["draw_id"]
        and ws_data.cell(row=last_row, column=2).value == draw_data["date"]
    ):
        print(f"Kỳ quay {draw_data['draw_id']} đã tồn tại trong Excel!")
        return

    # Dùng thêm một lớp chống duplicate khi cùng draw_id xuất hiện ở nhiều nguồn hoặc chạy lại nhanh.
    for row_idx in range(1, last_row + 1):
        existing_draw_id = ws_data.cell(row=row_idx, column=1).value
        if existing_draw_id == draw_data["draw_id"]:
            existing_date = ws_data.cell(row=row_idx, column=2).value
            if existing_date != draw_data["date"]:
                print(
                    f"Kỳ quay {draw_data['draw_id']} có ngày cũ {existing_date}; "
                    f"cập nhật thành {draw_data['date']} ở hàng {row_idx}."
                )
                for col_idx, val in enumerate(
                    [
                        draw_data["draw_id"],
                        draw_data["date"],
                        draw_data["weekday"],
                        *draw_data["nums"],
                        draw_data["jackpot"],
                        draw_data["winners"],
                    ],
                    start=1,
                ):
                    ws_data.cell(row=row_idx, column=col_idx, value=val)
                wb.save(file_path)
                return
            print(f"Kỳ quay {draw_data['draw_id']} đã tồn tại ở hàng {row_idx}; bỏ qua ghi đè.")
            return

    # Chuẩn bị dòng mới
    new_row_idx = last_row + 1
    new_data = [
        draw_data["draw_id"],
        draw_data["date"],
        draw_data["weekday"],
        *draw_data["nums"],
        draw_data["jackpot"],
        draw_data["winners"],
    ]

    # Styles
    font_regular = Font(name="Calibri", size=11)
    font_num = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    fill_gold = PatternFill(
        start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    # Ghi dữ liệu & định dạng
    for col_idx, val in enumerate(new_data, start=1):
        cell = ws_data.cell(row=new_row_idx, column=col_idx, value=val)
        cell.border = thin_border
        cell.font = font_regular

        if col_idx in [1, 2, 3]:
            cell.alignment = Alignment(horizontal="center")
        elif 4 <= col_idx <= 9:
            cell.alignment = Alignment(horizontal="center")
            cell.number_format = "00"
            cell.font = font_num
        elif col_idx == 10:
            cell.alignment = Alignment(horizontal="right")
            cell.number_format = "#,##0"
        elif col_idx == 11:
            cell.alignment = Alignment(horizontal="center")
            if val > 0:
                cell.fill = fill_gold
                cell.font = Font(
                    name="Calibri", size=11, bold=True, color="B45309"
                )

    wb.save(file_path)
    print(f"✅ Đã thêm thành công kỳ {draw_data['draw_id']} vào Excel!")


# Chạy bot
if __name__ == "__main__":
    result = fetch_latest_vietlott_mega645()
    append_to_excel("Vietlott_Mega_645_Full_Results.xlsx", result)
    init_sqlite_db("vietlott.db")
    save_result_to_sqlite("vietlott.db", result)
    with open("results.html", "w", encoding="utf-8") as file:
        file.write(render_results_html("vietlott.db", limit=20))