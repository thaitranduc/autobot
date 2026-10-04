import openpyxl
from collections import Counter

wb = openpyxl.load_workbook(r'd:\Git\autobot\Vietlott_Mega_645_Full_Results.xlsx', data_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
print('worksheet rows:', len(rows))
print('header:', rows[0])

nums = []
last_draws = []
for r in rows[1:]:
    if not r or len(r) < 10:
        continue
    n = [int(r[i]) for i in range(3, 9) if isinstance(r[i], (int, float))]
    if len(n) == 6:
        nums.extend(n)
        last_draws.append(sorted(n))

print('draw count:', len(last_draws))
print('most recent draw:', last_draws[-1])
print('all-time top 15:', Counter(nums).most_common(15))
recent = []
for d in last_draws[-100:]:
    recent.extend(d)
print('last 100 top 15:', Counter(recent).most_common(15))
print('last 10 draws:', last_draws[-10:])
