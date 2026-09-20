import sqlite3

with sqlite3.connect("vietlott.db") as connection:
    print(connection.execute(
        "SELECT COUNT(*), MIN(draw_id), MAX(draw_id) FROM results"
    ).fetchone())
