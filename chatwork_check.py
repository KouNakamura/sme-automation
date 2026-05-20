"""
Chatwork 前営業日〜現在の動向チェック
usage: python -X utf8 chatwork_check.py
"""
import sys, io, ctypes, ctypes.wintypes
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import urllib.request, json, re
from datetime import datetime, timedelta, date

BASE  = 'https://api.chatwork.com/v2'
CW_BASE = 'https://www.chatwork.com/#!rid'
TOKEN = None  # 起動時に資格情報マネージャーから取得

# ブレイン営業チャンネル（メインの営業共有ch）
SALES_ROOM_ID = YOUR_SALES_ROOM_ID

# 重要キーワード（これを含むメッセージは強調）
ALERT_WORDS = [
    'クレーム', '苦情', 'トラブル', '問題', '遅延', '遅れ', '障害',
    '退職', '休職', '体調', '病気', '緊急', '至急', '重要',
    '請求', '未払', '支払', 'エラー', 'ミス', '事故',
    'お詫び', '謝罪', '申し訳', '確認お願い', 'ご確認',
]

# 営業チャンネルでのトラブル検出キーワード
TROUBLE_WORDS = [
    'クレーム', '苦情', 'トラブル', '問題', '遅延', '遅れ', '障害',
    '退職', '休職', '体調不良', '病気', '緊急', '至急',
    '未払', '督促', '警告', 'ミス', '事故', 'やばい', 'まずい',
    '対応中', '調査中', '確認中', '保留',
]

def _load_token():
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
    return 'b0c820c2a79087252f191c16b9b9d4ef'  # fallback

def cw_get(path, token):
    req = urllib.request.Request(f'{BASE}{path}',
                                 headers={'X-ChatWorkToken': token})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode())

def prev_business_day(today=None):
    """前営業日（土→金、月→金、日→金）の朝9時のタイムスタンプを返す"""
    d = today or date.today()
    delta = 1
    if d.weekday() == 0:   delta = 3   # 月→金
    elif d.weekday() == 6: delta = 2   # 日→金
    prev = d - timedelta(days=delta)
    return int(datetime(prev.year, prev.month, prev.day, 9, 0).timestamp())

def is_alert(text):
    return any(w in text for w in ALERT_WORDS)

def is_trouble(text):
    return any(w in text for w in TROUBLE_WORDS)

def room_url(room_id):
    return f'{CW_BASE}{room_id}'

def check_sales_channel(token, since_ts):
    """営業チャンネルのトラブル状況を専用チェック"""
    rid = SALES_ROOM_ID
    url = room_url(rid)
    print(f'\n■ 営業チャンネル専用チェック')
    print(f'  URL: {url}')

    try:
        msgs = cw_get(f'/rooms/{rid}/messages?force=1', token)
        recent = [m for m in msgs if m.get('send_time', 0) >= since_ts]
    except Exception as e:
        print(f'  取得失敗: {e}')
        return

    if not recent:
        print(f'  前営業日以降のメッセージなし')
        return

    print(f'  前営業日以降: {len(recent)}件のメッセージ')

    trouble_msgs = [m for m in recent if is_trouble(m.get('body', ''))]

    if not trouble_msgs:
        print(f'  トラブル関連メッセージなし ✓')
    else:
        print(f'\n  ⚠ トラブル関連メッセージ ({len(trouble_msgs)}件):')
        for m in trouble_msgs:
            sender = m.get('account', {}).get('name', '')[:15]
            dt_str = datetime.fromtimestamp(m.get('send_time', 0)).strftime('%m/%d %H:%M')
            body   = m.get('body', '').replace('\n', ' ')[:120]
            print(f'    [{dt_str}] {sender}')
            print(f'    → {body}')
            print()

    # 全メッセージのサマリー（最新5件）
    print(f'  最新メッセージ ({min(5, len(recent))}件):')
    for m in recent[-5:]:
        sender = m.get('account', {}).get('name', '')[:15]
        dt_str = datetime.fromtimestamp(m.get('send_time', 0)).strftime('%m/%d %H:%M')
        body   = m.get('body', '').replace('\n', ' ')[:80]
        mark = '⚠ ' if is_trouble(m.get('body', '')) else '  '
        print(f'  {mark}{dt_str} {sender}: {body}')


def run_check():
    token = _load_token()
    since_ts = prev_business_day()
    since_dt = datetime.fromtimestamp(since_ts).strftime('%m/%d(%a) %H:%M')

    print(f'=== Chatwork チェック（{since_dt}〜現在） ===')

    # サマリー
    status = cw_get('/my/status', token)
    print(f'  未読: {status["unread_room_num"]}ルーム  '
          f'メンション: {status["mention_num"]}件  '
          f'自分のタスク: {status["mytask_num"]}件')

    # 自分の未完了タスク
    tasks = cw_get('/my/tasks?status=open', token)
    if tasks:
        print(f'\n■ 自分の未完了タスク ({len(tasks)}件)')
        for t in tasks:
            room = t.get('room', {}).get('name', '')[:18]
            body = t.get('body', '')[:50]
            due  = t.get('due_date', '') or 'なし'
            print(f'  [{room}] {body}  期限:{due}')

    # 営業チャンネル専用チェック（常に実行）
    check_sales_channel(token, since_ts)

    # 前営業日以降にアクティブなルーム
    rooms = cw_get('/rooms', token)
    active = sorted(
        [r for r in rooms if r.get('last_update_time', 0) >= since_ts],
        key=lambda x: x['last_update_time'], reverse=True
    )

    if not active:
        print(f'\n前営業日以降の動きなし')
        return

    print(f'\n■ 前営業日からの動き ({len(active)}ルーム)')

    alert_items = []
    mention_items = []
    normal_items = []

    for room in active:
        rid      = room['room_id']
        rname    = room['name']
        unread   = room.get('unread_num', 0)
        mention  = room.get('mention_num', 0)
        last_ts  = room['last_update_time']
        last_dt  = datetime.fromtimestamp(last_ts).strftime('%m/%d %H:%M')
        r_url    = room_url(rid)

        # メッセージ取得
        try:
            msgs = cw_get(f'/rooms/{rid}/messages?force=1', token)
            recent = [m for m in msgs if m.get('send_time', 0) >= since_ts]
        except:
            recent = []

        if not recent:
            continue

        # 重要度判定
        has_alert   = any(is_alert(m.get('body', '')) for m in recent)
        has_mention = mention > 0

        entry = {
            'room': rname, 'last_dt': last_dt,
            'msgs': recent, 'mention': mention, 'unread': unread,
            'alert': has_alert, 'url': r_url,
        }
        if has_alert:
            alert_items.append(entry)
        elif has_mention:
            mention_items.append(entry)
        else:
            normal_items.append(entry)

    def print_room(entry, show_msgs=True):
        flag = '🔴' if entry['alert'] else ('💬' if entry['mention'] else '  ')
        ment = f'【言及{entry["mention"]}】' if entry['mention'] else ''
        print(f'  {flag} {entry["last_dt"]} {entry["room"][:30]}  {ment}')
        print(f'       URL: {entry["url"]}')
        if show_msgs:
            for m in entry['msgs'][-3:]:
                sender = m.get('account', {}).get('name', '')[:10]
                body   = m.get('body', '').replace('\n', ' ')[:70]
                mts    = datetime.fromtimestamp(m.get('send_time', 0)).strftime('%H:%M')
                alert_mark = '⚠ ' if is_alert(body) else ''
                print(f'    {mts} {sender}: {alert_mark}{body}')

    if alert_items:
        print(f'\n  ⚠ 要注意ルーム ({len(alert_items)}件)')
        for e in alert_items:
            print_room(e, show_msgs=True)

    if mention_items:
        print(f'\n  💬 メンションあり ({len(mention_items)}件)')
        for e in mention_items:
            print_room(e, show_msgs=True)

    if normal_items:
        print(f'\n  動きあり ({len(normal_items)}件)')
        for e in normal_items:
            print_room(e, show_msgs=False)

if __name__ == '__main__':
    run_check()
