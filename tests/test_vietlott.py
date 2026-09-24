import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import vietlott
from vietlott import (
    append_to_excel,
    parse_mega645_html,
    parse_minhngoc_mega645_html,
    save_result_to_sqlite,
    render_results_html,
)
import repair_workbook


def test_parse_mega645_html():
    html = """
    <html>
        <body>
            <div>Kết quả QSMT kỳ #01561 ngày 11/09/2026</div>
            <div>14 18 20 21 26 27</div>
            <div>Jackpot Mega 6/45 ước tính</div>
            <div>61.850.545.000 VNĐ</div>
        </body>
    </html>
    """

    result = parse_mega645_html(html)

    assert result is not None
    assert result["draw_id"] == "#01561"
    assert result["date"] == "11/09/2026"
    assert result["weekday"] == "Thứ 6"
    assert result["nums"] == [14, 18, 20, 21, 26, 27]
    assert result["jackpot"] == 61850545000
    assert result["winners"] == 0


def test_parse_mega645_history_table():
    html = """
    <html><body>
    | 11/09/2026 | 01561 | 14 18 20 21 26 27 |
    | 09/09/2026 | 01560 | 12 17 20 21 36 43 |
    </body></html>
    """

    result = parse_mega645_html(html)

    assert result is not None
    assert result["draw_id"] == "#01561"
    assert result["date"] == "11/09/2026"
    assert result["weekday"] == "Thứ 6"
    assert result["nums"] == [14, 18, 20, 21, 26, 27]
    assert result["jackpot"] == 0
    assert result["winners"] == 0


def test_parse_live_vietlott_result_table():
    html = """
    <table>
      <tbody>
        <tr>
          <td>18/09/2026</td>
          <td><a href="/vi/trung-thuong/ket-qua-trung-thuong/645?id=01564&nocatche=1">01564</a></td>
          <td>
            <div class="day_so_ket_qua_v2">
              <span class="bong_tron ">07</span>
              <span class="bong_tron ">12</span>
              <span class="bong_tron ">26</span>
              <span class="bong_tron ">27</span>
              <span class="bong_tron ">41</span>
              <span class="bong_tron no-margin-right ">43</span>
            </div>
          </td>
        </tr>
      </tbody>
    </table>
    """

    result = parse_mega645_html(html)

    assert result is not None
    assert result["draw_id"] == "#01564"
    assert result["date"] == "18/09/2026"
    assert result["weekday"] == "Thứ 6"
    assert result["nums"] == [7, 12, 26, 27, 41, 43]
    assert result["jackpot"] == 0
    assert result["winners"] == 0


def test_parse_minhngoc_mega645_html():
    html = """
    <div class="boxkqxsdientoan">
      <td align="center">Kỳ vé: <span id="DT6X45_KY_VE">#01562</span>
      | Ngày quay thưởng 13/09/2026</td>
      <ul class="result-number">
        <li><div class="finnish1 bool">04</div></li>
        <li><div class="finnish2 bool">12</div></li>
        <li><div class="finnish3 bool">31</div></li>
        <li><div class="finnish4 bool">34</div></li>
        <li><div class="finnish5 bool">38</div></li>
        <li><div class="finnish6 bool">41</div></li>
      </ul>
      <td id="DT6X45_S_JACKPOT">0</td>
      <b id="DT6X45_G_JACKPOT">69,161,109,500<sup>đ</sup></b>
    </div>
    """

    result = parse_minhngoc_mega645_html(html)

    assert result == {
        "draw_id": "#01562",
        "date": "13/09/2026",
        "weekday": "Chủ Nhật",
        "nums": [4, 12, 31, 34, 38, 41],
        "jackpot": 69161109500,
        "winners": 0,
    }


def test_append_repairs_stale_draw_id(tmp_path):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Lịch Sử Số Trúng (1200+ Kỳ)"
    worksheet.append(["Draw", "Date", "Weekday", "N1", "N2", "N3", "N4", "N5", "N6", "Jackpot", "Winners"])
    worksheet.append(["#01563", "12/07/2026", "Chủ Nhật", 1, 2, 19, 37, 44, 45, 100, 0])
    path = tmp_path / "results.xlsx"
    workbook.save(path)

    append_to_excel(
        path,
        {
            "draw_id": "#01563",
            "date": "16/09/2026",
            "weekday": "Thứ 4",
            "nums": [4, 12, 31, 34, 38, 41],
            "jackpot": 69161109500,
            "winners": 0,
        },
    )

    updated = openpyxl.load_workbook(path, data_only=True).active
    assert updated.cell(2, 2).value == "16/09/2026"
    assert [updated.cell(2, col).value for col in range(4, 10)] == [4, 12, 31, 34, 38, 41]


def test_append_formats_jackpot_column_as_vnd(tmp_path):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Lịch Sử Số Trúng (1200+ Kỳ)"
    worksheet.append(["Draw", "Date", "Weekday", "N1", "N2", "N3", "N4", "N5", "N6", "Jackpot", "Winners"])
    path = tmp_path / "results.xlsx"
    workbook.save(path)

    append_to_excel(
        path,
        {
            "draw_id": "#01564",
            "date": "18/09/2026",
            "weekday": "Thứ 6",
            "nums": [7, 12, 26, 27, 41, 43],
            "jackpot": 69161109500,
            "winners": 0,
        },
    )

    updated = openpyxl.load_workbook(path).active
    assert updated["J2"].number_format == '#,##0 "VND"'
    assert updated["J2"].value == 69161109500


def test_fetch_latest_vietlott_mega645_prefers_live_official_page(monkeypatch):
    responses = {
        "https://www.minhngoc.net.vn/ket-qua-xo-so/dien-toan-vietlott.html": """
        <div class="boxkqxsdientoan">
          <td align="center">Kỳ vé: <span id="DT6X45_KY_VE">#01400</span>
          | Ngày quay thưởng 19/09/2026</td>
          <ul class="result-number">
            <li><div class="finnish1 bool">04</div></li>
            <li><div class="finnish2 bool">07</div></li>
            <li><div class="finnish3 bool">11</div></li>
            <li><div class="finnish4 bool">18</div></li>
            <li><div class="finnish5 bool">22</div></li>
            <li><div class="finnish6 bool">25</div></li>
          </ul>
          <td id="DT6X45_S_JACKPOT">0</td>
          <b id="DT6X45_G_JACKPOT">86,148,921,500<sup>đ</sup></b>
        </div>
        """,
        "https://vietlott.vn/vi/trung-thuong/ket-qua-trung-thuong/winning-number-645": """
        <html><body>
          <h5>Kỳ quay thưởng <b>#01564</b> ngày <b>18/09/2026</b></h5>
          <div class="day_so_ket_qua_v2">
            <span class="bong_tron">07</span>
            <span class="bong_tron">12</span>
            <span class="bong_tron">26</span>
            <span class="bong_tron">27</span>
            <span class="bong_tron">41</span>
            <span class="bong_tron no-margin-right">43</span>
          </div>
        </body></html>
        """,
    }

    class FakeResponse:
        def __init__(self, url):
            self.url = url
            self.text = responses[url]
            self.headers = {"Content-Type": "text/html; charset=UTF-8"}

        def raise_for_status(self):
            return None

    def fake_get(url, headers=None, timeout=None):
        return FakeResponse(url)

    monkeypatch.setattr(vietlott.requests, "get", fake_get)

    result = vietlott.fetch_latest_vietlott_mega645()
    assert result["draw_id"] == "#01564"
    assert result["date"] == "18/09/2026"


def test_save_result_to_sqlite_and_render_html(tmp_path):
    db_path = tmp_path / "vietlott.db"
    result = {
        "draw_id": "#01564",
        "date": "20/09/2026",
        "weekday": "Thứ 7",
        "nums": [6, 12, 18, 25, 33, 42],
        "jackpot": 500000000,
        "winners": 1,
    }

    saved = save_result_to_sqlite(db_path, result)
    assert saved is True

    html = render_results_html(db_path, limit=10)
    assert "#01564" in html
    assert "20/09/2026" in html
    assert "6" in html
    assert "500000000" in html


def test_save_result_to_sqlite_keeps_historical_rows(tmp_path):
    db_path = tmp_path / "vietlott.db"
    save_result_to_sqlite(db_path, {
        "draw_id": "#01564",
        "date": "18/09/2026",
        "weekday": "Thứ 6",
        "nums": [7, 12, 26, 27, 41, 43],
        "jackpot": 0,
        "winners": 0,
    })
    save_result_to_sqlite(db_path, {
        "draw_id": "#01400",
        "date": "19/09/2026",
        "weekday": "Thứ 7",
        "nums": [4, 7, 11, 18, 22, 25],
        "jackpot": 0,
        "winners": 0,
    })
    assert "#01400" in str(render_results_html(db_path, limit=10))

    save_result_to_sqlite(db_path, {
        "draw_id": "#01564",
        "date": "18/09/2026",
        "weekday": "Thứ 6",
        "nums": [7, 12, 26, 27, 41, 43],
        "jackpot": 0,
        "winners": 0,
    })

    html = render_results_html(db_path, limit=10)
    assert "#01564" in html
    assert "#01400" in html


def test_rebuild_workbook_keeps_latest_live_result(tmp_path, monkeypatch):
    archive_rows = [
        {
            "draw_id": "#00001",
            "date": "24/07/2016",
            "weekday": "Chủ Nhật",
            "nums": [1, 10, 16, 18, 23, 38],
            "jackpot": 0,
            "winners": 0,
        },
        {
            "draw_id": "#00002",
            "date": "25/07/2016",
            "weekday": "Thứ 2",
            "nums": [2, 11, 17, 19, 24, 39],
            "jackpot": 0,
            "winners": 0,
        },
    ]
    live_result = {
        "draw_id": "#01564",
        "date": "18/09/2026",
        "weekday": "Thứ 6",
        "nums": [7, 12, 26, 27, 41, 43],
        "jackpot": 0,
        "winners": 0,
    }

    monkeypatch.setattr(repair_workbook, "fetch_ketquadientoan_history", lambda *args, **kwargs: archive_rows)
    monkeypatch.setattr(repair_workbook, "fetch_latest_vietlott_mega645", lambda: live_result)

    file_path = tmp_path / "rebuild.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Lịch Sử Số Trúng (1200+ Kỳ)"
    sheet.append(["Kỳ Quay", "Ngày Quay", "Thứ", "Số 1", "Số 2", "Số 3", "Số 4", "Số 5", "Số 6", "Jackpot (VNĐ)", "Số Vé Trúng JP"])
    workbook.save(file_path)

    repair_workbook.rebuild_workbook(file_path)

    saved = openpyxl.load_workbook(file_path, data_only=True)
    ws = saved["Lịch Sử Số Trúng (1200+ Kỳ)"]
    assert ws.max_row == 4
    assert ws.cell(2, 1).value == "#00001"
    assert ws.cell(4, 1).value == "#01564"
    assert ws.cell(4, 2).value == "18/09/2026"
