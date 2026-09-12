import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from vietlott import parse_mega645_html


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
