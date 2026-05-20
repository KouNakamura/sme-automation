import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import openpyxl
from datetime import datetime

CF_PATH = r'\\YOUR_NAS_IP\keiri\01.キャッシュフロー\キャッシュフロー改定(2026.5.7).xlsx'
wb = openpyxl.load_workbook(CF_PATH, data_only=True)
ws = wb['キャッシュ表_予定【融資有】']

print('=== キャッシュ表 先頭100行×15列 ===')
for r in range(1, 101):
    row_data = []
    for c in range(1, 16):
        v = ws.cell(r, c).value
        if v is not None:
            row_data.append(f'C{c}:{str(v)[:15]}')
    if row_data:
        print(f'  行{r:3d}: {" | ".join(row_data)}')
