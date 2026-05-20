"""
週次営業レポート — 毎週月曜朝にブレイン営業chへ投稿
  - 直近7日間の引き合い技術カテゴリ（前週比）
  - 地域別（関東/関西）件数（前週比）
  - Cybozuトラブル管理 件数（前週比）
usage: python -X utf8 chatwork_weekly_report.py
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

import ctypes, ctypes.wintypes
import urllib.request, urllib.parse, json, os, time, re
from datetime import datetime, timedelta
from collections import Counter

# ─── 設定 ────────────────────────────────────────────
ROOM_ID    = YOUR_SALES_ROOM_ID   # ブレイン営業ch
CACHE_FILE = os.path.join(os.path.dirname(__file__), 'weekly_report_cache.json')
BASE_CW    = 'https://api.chatwork.com/v2'
BASE_CY    = 'https://YOUR_COMPANY.cybozu.com'
DID        = YOUR_DB_ID

CATEGORIES = {
    'Java系':    ['Java', 'Spring', 'Struts', 'J2EE'],
    '.NET系':    ['.NET', 'C#', 'ASP.NET', 'VB.NET'],
    'Python系':  ['Python', 'Django', 'Flask', 'FastAPI'],
    'Web系':     ['JavaScript', 'TypeScript', 'React', 'Vue', 'Angular', 'Node.js', 'フロント'],
    'COBOL系':   ['COBOL', '汎用機', 'メインフレーム', 'AS/400'],
    'AWS':       ['AWS', 'EC2', 'Lambda', 'ECS', 'EKS', 'CloudFormation'],
    'Azure':     ['Azure', 'Intune', 'Entra'],
    'インフラ基盤': ['インフラ', 'Linux', 'Windows Server', 'VMware', '仮想化'],
    'クラウド移行': ['マイグレーション', 'クラウド移行', 'リフト', '移行'],
    'テスト/QA': ['テスト', 'QA', 'テスター'],
    'データ/AI': ['データ分析', 'AI', '機械学習', 'BI', 'DWH', 'ETL'],
    'PMO/PM':    ['PMO', 'PM ', 'プロジェクトマネ'],
    'セキュリティ': ['セキュリティ', 'SOC', '脆弱性'],
    'SAP系':     ['SAP', 'ABAP', 'S/4HANA'],
}

KANSAI_WORDS = ['大阪', '梅田', '難波', '心斎橋', '天王寺', '京橋', '北浜',
                '神戸', '三宮', '兵庫', '京都', '奈良', '滋賀', '和歌山', '関西', '近畿']
KANTO_WORDS  = ['東京', '渋谷', '新宿', '品川', '丸の内', '大手町', '赤坂', '六本木',
                '秋葉原', '池袋', '銀座', '横浜', '川崎', '埼玉', '千葉',
                '関東', '首都圏', 'リモート', '在宅', 'フルリモート']

OPEN_STATUS = {'対応中', '確認中', '調査中', '保留中', 'オープン', '未対応', '新規'}


# ─── 認証 ─────────────────────────────────────────────
def _load_cw_token():
    class FILETIME(ctypes.Structure):
        _fields_ = [('Low', ctypes.wintypes.DWORD), ('High', ctypes.wintypes.DWORD)]
    class CREDENTIAL(ctypes.Structure):
        _fields_ = [
            ('Flags', ctypes.wintypes.DWORD), ('Type', ctypes.wintypes.DWORD),
            ('TargetName', ctypes.c_wchar_p), ('Comment', ctypes.c_wchar_p),
            ('LastWritten', FILETIME), ('CredentialBlobSize', ctypes.wintypes.DWORD),
            ('CredentialBlob', ctypes.POINTER(ctypes.c_ubyte)),
            ('Persist', ctypes.wintypes.DWORD), ('AttributeCount', ctypes.wintypes.DWORD),
            ('Attributes', ctypes.c_void_p), ('TargetAlias', ctypes.c_wchar_p),
            ('UserName', ctypes.c_wchar_p),
        ]
    advapi32 = ctypes.windll.advapi32
    count = ctypes.wintypes.DWORD()
    creds = ctypes.POINTER(ctypes.POINTER(CREDENTIAL))()
    if advapi32.CredEnumerateW(None, 0, ctypes.byref(count), ctypes.byref(creds)):
        for i in range(count.value):
            c = creds[i].contents
            if (c.TargetName or '').lower() == 'chatwork_api_token':
                size = c.CredentialBlobSize
                if size > 0:
                    blob = bytes(c.CredentialBlob[j] for j in range(size))
                    try:    return blob.decode('utf-16-le').rstrip('\x00')
                    except: return blob.decode('utf-8', errors='replace')
        advapi32.CredFree(creds)
    return 'b0c820c2a79087252f191c16b9b9d4ef'


# ─── メール分析 ───────────────────────────────────────
def analyze_emails(days=7):
    import win32com.client
    since_dt  = datetime.now() - timedelta(days=days)
    since_str = since_dt.strftime('%m/%d/%Y %H:%M %p')

    outlook = win32com.client.Dispatch('Outlook.Application')
    ns      = outlook.GetNamespace('MAPI')

    seen_keys = set()
    texts = []

    for acc in ns.Accounts:
        addr = acc.SmtpAddress or ''
        if 'brain-d.jp' not in addr:
            continue
        try:
            store    = acc.DeliveryStore
            inbox    = store.GetDefaultFolder(6)
            items    = inbox.Items
            filtered = items.Restrict("[ReceivedTime] >= '" + since_str + "'")
            item = filtered.GetFirst()
            while item is not None:
                try:
                    subject     = item.Subject or ''
                    sender_addr = (item.SenderEmailAddress or '').lower()
                    recv_time   = item.ReceivedTime
                    body_prev   = (item.Body or '')[:500]

                    if 'brain-d.jp' in sender_addr:
                        item = filtered.GetNext()
                        continue

                    try:    date_str = recv_time.strftime('%Y%m%d')
                    except: date_str = '00000000'
                    key = (sender_addr, subject.strip(), date_str)

                    if key not in seen_keys:
                        seen_keys.add(key)
                        txt = subject + ' ' + body_prev
                        texts.append(txt)
                except Exception:
                    pass
                try:    item = filtered.GetNext()
                except: break
        except Exception:
            pass

    cat_counts = {}
    for cat, keywords in CATEGORIES.items():
        cat_counts[cat] = sum(1 for t in texts if any(kw in t for kw in keywords))

    kanto  = sum(1 for t in texts if any(w in t for w in KANTO_WORDS))
    kansai = sum(1 for t in texts if any(w in t for w in KANSAI_WORDS))

    return {
        'total':      len(texts),
        'categories': cat_counts,
        'kanto':      kanto,
        'kansai':     kansai,
        'date':       datetime.now().strftime('%Y-%m-%d'),
    }


# ─── Cybozuトラブル件数取得 ───────────────────────────
def get_trouble_count():
    try:
        sys.path.insert(0, os.path.dirname(__file__))
        from check_travel import _cybozu_login_driver
        from bs4 import BeautifulSoup

        driver = _cybozu_login_driver()
        if not driver:
            return None, []

        try:
            all_rids = []
            url = f'{BASE_CY}/o/ag.cgi?page=DBView&did={DID}'
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
                    url = h if h.startswith('http') else f'{BASE_CY}/o/{h.lstrip("/")}'
                else:
                    url = None

            open_records = []
            for rid in all_rids:
                driver.get(f'{BASE_CY}/o/ag.cgi?page=DBRecord&did={DID}&rid={rid}')
                time.sleep(0.8)
                soup = BeautifulSoup(driver.page_source, 'html.parser')
                data = {}
                for row in soup.find_all('tr'):
                    cells = row.find_all(['th', 'td'])
                    if len(cells) >= 2:
                        key = cells[0].get_text(strip=True)
                        val = cells[1].get_text(separator=' ', strip=True)
                        if key in {'状況', '対象者', 'トラブル区分'}:
                            data[key] = val
                if data.get('状況', '') in OPEN_STATUS:
                    open_records.append({
                        'rid':    rid,
                        'person': data.get('対象者', '不明'),
                        'type':   data.get('トラブル区分', ''),
                        'status': data.get('状況', ''),
                        'url':    f'{BASE_CY}/o/ag.cgi?page=DBRecord&did={DID}&rid={rid}',
                    })
            return len(open_records), open_records
        finally:
            try: driver.quit()
            except: pass
    except Exception as e:
        print(f'Cybozu取得失敗: {e}')
        return None, []


# ─── 差分フォーマット ─────────────────────────────────
def diff(curr, prev, key):
    c = curr.get(key, 0)
    p = prev.get(key, 0) if prev else None
    if p is None:
        return f'{c}件'
    d = c - p
    if d > 0:  return f'{c}件（▲{d}）'
    if d < 0:  return f'{c}件（▽{abs(d)}）'
    return f'{c}件（→）'


# ─── 投稿 ─────────────────────────────────────────────
def post_to_cw(token, msg):
    data = urllib.parse.urlencode({'body': msg}).encode()
    req  = urllib.request.Request(
        f'{BASE_CW}/rooms/{ROOM_ID}/messages',
        data=data,
        headers={'X-ChatWorkToken': token},
        method='POST'
    )
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode()).get('message_id')


# ─── メイン ───────────────────────────────────────────
def main():
    today = datetime.now()
    week_label = today.strftime('%Y/%m/%d')

    print('メール分析中...')
    curr = analyze_emails(days=7)

    # キャッシュ読み込み
    prev = None
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                prev = json.load(f)
        except Exception:
            pass

    print('Cybozuトラブル確認中...')
    trouble_count, trouble_records = get_trouble_count()
    prev_trouble = prev.get('trouble_count') if prev else None

    # 差分サマリー
    def tdiff(c, p):
        if p is None: return f'{c}件'
        d = c - p
        if d > 0:  return f'{c}件（▲{d} 増）'
        if d < 0:  return f'{c}件（▽{abs(d)} 減）'
        return f'{c}件（変化なし）'

    # メッセージ構成
    lines = [f'【週次営業レポート {week_label}】（直近7日間）']

    # トラブル管理
    tc_str = tdiff(trouble_count, prev_trouble) if trouble_count is not None else '取得失敗'
    lines.append(f'\n■ トラブル管理 オープン件数: {tc_str}')
    for r in trouble_records:
        lines.append(f'  ・{r["person"]}（{r["type"]}）{r["status"]}')
        lines.append(f'    {r["url"]}')

    # 地域別
    k_str  = tdiff(curr['kanto'],  prev.get('kanto')  if prev else None)
    ks_str = tdiff(curr['kansai'], prev.get('kansai') if prev else None)
    lines.append(f'\n■ 地域別 引き合い件数')
    lines.append(f'  関東（リモート含む）: {k_str}')
    lines.append(f'  関西:                 {ks_str}')

    # 技術カテゴリ
    lines.append(f'\n■ 技術カテゴリ別 引き合い（前週比）')
    prev_cats = prev.get('categories', {}) if prev else {}
    sorted_cats = sorted(curr['categories'].items(), key=lambda x: -x[1])
    for cat, cnt in sorted_cats:
        if cnt == 0:
            continue
        p = prev_cats.get(cat)
        if p is None:
            diff_str = ''
        else:
            d = cnt - p
            diff_str = f'（▲{d}）' if d > 0 else (f'（▽{abs(d)}）' if d < 0 else '（→）')
        lines.append(f'  {cat}: {cnt}件{diff_str}')

    prev_date = prev.get('date', '初回') if prev else '初回'
    lines.append(f'\n比較基準: {prev_date}週')
    lines.append(f'分析対象: 重複排除後 {curr["total"]}件（社外メール）')

    msg = '\n'.join(lines)
    print(msg)
    print()

    token = _load_cw_token()
    mid = post_to_cw(token, msg)
    print(f'投稿成功: message_id={mid}')

    # キャッシュ保存（次回比較用）
    cache = {
        'date':          curr['date'],
        'total':         curr['total'],
        'categories':    curr['categories'],
        'kanto':         curr['kanto'],
        'kansai':        curr['kansai'],
        'trouble_count': trouble_count,
    }
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    print(f'キャッシュ保存: {CACHE_FILE}')


if __name__ == '__main__':
    main()
