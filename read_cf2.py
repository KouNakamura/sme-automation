import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import openpyxl
from datetime import datetime, date

CF_PATH = r'\\YOUR_NAS_IP\YOUR_FOLDER\01.cashflow\cashflow_YYYYMMDD.xlsx'
wb = openpyxl.load_workbook(CF_PATH, data_only=True)
ws = wb['キャッシュ表_予定【融資有】']

# 2026-04以降の行を探す
print('=== キャッシュ表 2026年4月以降のエントリ ===')
current_month = None
in_range = False
for r in range(1, ws.max_row + 1):
    c1 = ws.cell(r, 1).value
    c2 = ws.cell(r, 2).value
    c4 = ws.cell(r, 4).value
    c5 = ws.cell(r, 5).value
    c6 = ws.cell(r, 6).value
    c7 = ws.cell(r, 7).value
    c8 = ws.cell(r, 8).value

    # 年月ヘッダー行を検出
    if isinstance(c1, (datetime, date)) or (isinstance(c1, str) and '2026' in str(c1)):
        current_month = c1
        if hasattr(c1, 'year') and c1.year >= 2026 and c1.month >= 4:
            in_range = True
        elif isinstance(c1, str) and '2026' in c1:
            in_range = True
        else:
            if in_range:
                break

    if in_range and (c2 or c4 or c5 or c6):
        c5s = f'{int(c5):,}' if isinstance(c5, (int, float)) and c5 else ''
        c6s = f'{int(c6):,}' if isinstance(c6, (int, float)) and c6 else ''
        c7s = f'{int(c7):,}' if isinstance(c7, (int, float)) and c7 else ''
        c8s = str(c8)[:30] if c8 else ''
        c1s = str(c1)[:12] if c1 else ''
        c4s = str(c4)[:30] if c4 else ''
        print(f'  行{r:5d} [{c1s:12}] 日:{str(c2):4} {c4s:30} 収入:{c5s:>12} 支出:{c6s:>12} 残高:{c7s:>14} {c8s}')
