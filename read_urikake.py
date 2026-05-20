import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import xlrd
from datetime import date

wb = xlrd.open_workbook(r'\\YOUR_NAS_IP\eigyo\01_請求関連\☆売上\売掛一覧.xls')
ws = wb.sheet_by_name('Sheet1')

def xl_date(val):
    try:
        if val and val > 40000:
            return xlrd.xldate_as_datetime(val, wb.datemode).date()
    except:
        pass
    return None

apr_start = date(2026, 4, 1)
apr_end   = date(2026, 4, 30)

results = []
for r in range(1, ws.nrows):
    try:
        bangou  = str(ws.cell_value(r, 0))
        col1    = str(ws.cell_value(r, 1))
        col2    = str(ws.cell_value(r, 2))
        hakkouv = ws.cell_value(r, 3)
        kingaku = ws.cell_value(r, 4)
        nyukin_v= ws.cell_value(r, 5)
        nyukin_c= ws.cell_value(r, 7)

        hakkou    = xl_date(hakkouv)
        nyukin    = xl_date(nyukin_v)
        confirmed = xl_date(nyukin_c) if nyukin_c else None

        if hakkou and apr_start <= hakkou <= apr_end:
            results.append({
                'bangou': bangou, 'col1': col1[:35], 'col2': col2[:25],
                'hakkou': hakkou, 'kingaku': kingaku,
                'nyukin': nyukin, 'confirmed': confirmed
            })
    except:
        pass

print(f'=== 2026年4月度 発行分 ({len(results)}件) ===')
total = 0
nyukin_done = 0
nyukin_done_amt = 0
for d in results:
    nyukin_str = str(d['nyukin']) if d['nyukin'] else '未定'
    if d['confirmed']:
        conf_str = f'入金済 {d["confirmed"]}'
        nyukin_done += 1
        nyukin_done_amt += d['kingaku'] if d['kingaku'] else 0
    else:
        conf_str = '未入金'
    print(f'  {d["bangou"]:12}  {d["col2"]:25}  {d["col1"][:30]}')
    print(f'    発行:{d["hakkou"]}  金額:{d["kingaku"]:>12,.0f}円  入金予定:{nyukin_str}  {conf_str}')
    total += d['kingaku'] if d['kingaku'] else 0

print()
print(f'  請求合計 : {total:>12,.0f}円  ({len(results)}件)')
print(f'  入金済   : {nyukin_done_amt:>12,.0f}円  ({nyukin_done}件)')
print(f'  未入金   : {total - nyukin_done_amt:>12,.0f}円  ({len(results) - nyukin_done}件)')
