import openpyxl

wb = openpyxl.load_workbook(r'd:\Git\autobot\Vietlott_Mega_645_Full_Results.xlsx', data_only=True)
ws = wb['Lịch Sử Số Trúng (1200+ Kỳ)']
used = set()

for row in ws.iter_rows(min_row=2, min_col=4, max_col=9):
    for cell in row:
        if cell.value is not None:
            used.add(int(cell.value))

missing = [n for n in range(1, 46) if n not in used]
print('missing_count=', len(missing))
print('first_missing=', missing[:12])
print('suggested=', missing[:6])
