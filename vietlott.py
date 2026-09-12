import datetime
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import requests


def parse_mega645_html(html_text):
    """Parse the Vietlott Mega 6/45 page when the API endpoint no longer returns JSON."""
    if not html_text:
        return None

    match_draw = re.search(r"Kết quả QSMT kỳ\s*#?(\d+)\s*ngày\s*(\d{2}/\d{2}/\d{4})", html_text, re.I)
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

    numbers_match = re.search(r"(?:\b|>)(\d{2})\s+(\d{2})\s+(\d{2})\s+(\d{2})\s+(\d{2})\s+(\d{2})(?:\b|<)", html_text)
    if not numbers_match:
        return None

    nums = [int(x) for x in numbers_match.groups()]

    jackpot_match = re.search(r"Jackpot\s+Mega\s+6/45\s+ước\s+tính.*?([\d\.]+)\s*VNĐ|([\d\.]+)\s*VNĐ.*?Jackpot\s+Mega\s+6/45", html_text, re.I | re.S)
    jackpot_val = 0
    if jackpot_match:
        num_str = jackpot_match.group(1) or jackpot_match.group(2)
        if num_str:
            jackpot_val = int(num_str.replace('.', '').replace(',', ''))

    return {
        "draw_id": draw_id,
        "date": draw_date_str,
        "weekday": weekday,
        "nums": sorted(nums),
        "jackpot": jackpot_val,
        "winners": 0,
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
    """Lấy kết quả Mega 6/45 mới nhất. Hỗ trợ cả JSON cũ lẫn HTML mới của Vietlott."""
    url = "https://vietlott.vn/vi/trung-thuong/ket-qua-trung-thuong/645"
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
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
                return None

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

        print("⚠️ API không trả về JSON; đang thử parse HTML trang kết quả Mega 6/45.")
        parsed = parse_mega645_html(res.text)
        if parsed:
            print(f"✅ Parse HTML thành công: {parsed['draw_id']} {parsed['date']}")
            return parsed

        print("❌ Không thể parse dữ liệu Mega 6/45 từ phản hồi hiện tại.")
        print(res.text[:1000])
        return None
    except Exception as e:
        print(f"Lỗi khi lấy dữ liệu Vietlott: {e}")
        return None


def append_to_excel(file_path, draw_data):
    """Ghi thêm dữ liệu kỳ mới vào file Excel."""
    wb = openpyxl.load_workbook(file_path)
    ws_data = wb["Lịch Sử Số Trúng (1200+ Kỳ)"]

    # Kiểm tra xem kỳ quay đã tồn tại chưa
    last_row = ws_data.max_row
    last_draw_id = ws_data.cell(row=last_row, column=1).value

    if last_draw_id == draw_data["draw_id"]:
        print(f"Kỳ quay {draw_data['draw_id']} đã tồn tại trong Excel!")
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
    if result:
        append_to_excel("Vietlott_Mega_645_Full_Results.xlsx", result)
    else:
        print("\n💡 Để thêm dữ liệu theo cách thủ công, chạy:")
        print('python -c "from vietlott import append_to_excel; ')
        print('data = {')
        print('    \"draw_id\": \"#01234\",')
        print('    \"date\": \"15/08/2026\",')
        print('    \"weekday\": \"Thứ 5\",')
        print('    \"nums\": [2, 4, 6, 23, 31, 39],')
        print('    \"jackpot\": 50000000,')
        print('    \"winners\": 0')
        print('};')
        print('append_to_excel(\'Vietlott_Mega_645_Full_Results.xlsx\', data)"')