"""
Cybozuトラブル管理(did=YOUR_DB_ID)のオープン件をチェック
usage: python -X utf8 cybozu_trouble_check.py
"""
import sys, io, os, time, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)
sys.path.insert(0, os.path.dirname(__file__))

from check_travel import _cybozu_login_driver, CYBOZU_BASE
from bs4 import BeautifulSoup

BASE = CYBOZU_BASE
DID  = YOUR_DB_ID
OPEN_STATUS  = {'対応中', '確認中', '調査中', '保留中', 'オープン', '未対応', '新規'}
CLOSED_STATUS = {'対応完了', '完了', '解決済み', 'クローズ', 'Closed', '終了'}

def parse_record_fast(driver, rid):
    """レコード詳細ページから主要フィールドを取得"""
    driver.get(f'{BASE}/o/ag.cgi?page=DBRecord&did={DID}&rid={rid}')
    time.sleep(0.8)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    data = {'rid': rid}
    seen = set()
    target_keys = {'レコード番号', '登録日時', '更新日時', '対象者', '部長',
                   'トラブル区分', '内容', '状況'}
    for row in soup.find_all('tr'):
        cells = row.find_all(['th', 'td'])
        if len(cells) >= 2:
            key = cells[0].get_text(strip=True)
            val = cells[1].get_text(separator=' ', strip=True)
            if key and key not in seen and key in target_keys:
                data[key] = val
                seen.add(key)
    return data

def get_all_rids(driver):
    """一覧から全ridを収集（ページング対応）"""
    all_rids = []
    url = f'{BASE}/o/ag.cgi?page=DBView&did={DID}'
    visited = set()
    while url and url not in visited:
        visited.add(url)
        driver.get(url)
        time.sleep(1.5)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        rids = list(dict.fromkeys(
            m.group(1)
            for a in soup.find_all('a', href=True)
            for m in [re.search(r'rid=(\d+)', a.get('href', ''))]
            if m and 'DBRecord' in a.get('href', '')
        ))
        all_rids.extend(r for r in rids if r not in all_rids)
        nxt = soup.find('a', string=re.compile(r'次のページ|次へ'))
        if nxt and nxt.get('href'):
            h = nxt['href']
            url = h if h.startswith('http') else f'{BASE}/o/{h.lstrip("/")}'
        else:
            url = None
    return all_rids

def run(verbose=True):
    driver = _cybozu_login_driver()
    if not driver:
        print('サイボウズログイン失敗'); return []

    try:
        rids = get_all_rids(driver)
        if verbose:
            print(f'  トラブル管理: 全{len(rids)}件を確認中...')

        open_records = []
        for rid in rids:
            rec = parse_record_fast(driver, rid)
            status = rec.get('状況', '')
            if status in OPEN_STATUS:
                open_records.append(rec)
            elif not status or status in CLOSED_STATUS:
                pass  # 完了済みはスキップ

        return open_records

    finally:
        try: driver.quit()
        except: pass

def print_report(records):
    if not records:
        print('  オープントラブルなし ✓')
        return

    print(f'  ★ オープントラブル: {len(records)}件')
    for r in records:
        rid      = r.get('rid', '?')
        person   = r.get('対象者', '不明')
        category = r.get('トラブル区分', '')
        status   = r.get('状況', '')
        reg_date = r.get('登録日時', '')[:10]
        upd_date = r.get('更新日時', '')[:10]
        content  = r.get('内容', '')[:100].replace('\n', ' ')
        url      = f'{BASE}/o/ag.cgi?page=DBRecord&did={DID}&rid={rid}'
        print(f'  ─────────────────────────────')
        print(f'  [{rid}] {person}  [{category}]  状況: {status}')
        print(f'        登録: {reg_date}  最終更新: {upd_date}')
        print(f'        {content}...')
        print(f'        → {url}')

if __name__ == '__main__':
    print('=== Cybozuトラブル管理チェック ===')
    records = run(verbose=True)
    print_report(records)
