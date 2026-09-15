import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from vietlott import parse_mega645_html, parse_minhngoc_mega645_html


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
    assert result["weekday"] == "Thứ 4"
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
    assert result["weekday"] == "Thứ 4"
    assert result["nums"] == [14, 18, 20, 21, 26, 27]
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
