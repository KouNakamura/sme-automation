"""
Outlookローカルメール 前営業日〜現在のチェック
usage: python -X utf8 outlook_check.py
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

import win32com.client
from datetime import datetime, timedelta, date, timezone
import re

ALERT_WORDS = [
    'クレーム', '苦情', 'トラブル', '問題', '障害', '不具合',
    '遅延', '遅れ', '納期',
    '退職', '休職', '体調', '病気', '緊急', '至急', '重要',
    '請求', '未払', '支払', 'エラー', 'ミス', '事故',
    'お詫び', '謝罪', '申し訳', '確認お願い', 'ご確認',
    '督促', '警告', 'アラート', '失敗', '拒否', '却下',
]

IGNORE_SENDERS = {
    'noreply', 'no-reply', 'donotreply', 'do-not-reply',
    'notification', 'newsletter', 'marketing', 'mailer-daemon',
}

TARGET_ACCOUNTS = ['your_email@example.com', 'your_email2@example.com']


def prev_business_day(today=None):
    d = today or date.today()
    delta = 1
    if d.weekday() == 0:   delta = 3   # 月→金
    elif d.weekday() == 6: delta = 2   # 日→金
    prev = d - timedelta(days=delta)
    return datetime(prev.year, prev.month, prev.day, 9, 0)


def is_alert(text):
    return any(w in text for w in ALERT_WORDS)


def is_ignorable_sender(addr):
    addr_lower = (addr or '').lower()
    return any(kw in addr_lower for kw in IGNORE_SENDERS)


def get_outlook_mails(since_dt):
    outlook = win32com.client.Dispatch('Outlook.Application')
    ns = outlook.GetNamespace('MAPI')

    results = []

    for acc in ns.Accounts:
        addr = acc.SmtpAddress or ''
        if TARGET_ACCOUNTS and addr.lower() not in [a.lower() for a in TARGET_ACCOUNTS]:
            continue

        try:
            store = acc.DeliveryStore
            inbox = store.GetDefaultFolder(6)  # olFolderInbox = 6
        except Exception as e:
            print(f'  [{addr}] 受信トレイ取得失敗: {e}')
            continue

        # フィルタ文字列（Outlook DASL）
        since_str = since_dt.strftime('%m/%d/%Y %H:%M %p')
        filter_str = f"[ReceivedTime] >= '{since_str}'"

        try:
            items = inbox.Items
            items.Sort('[ReceivedTime]', True)  # 降順
            filtered = items.Restrict(filter_str)
            count = filtered.Count
        except Exception as e:
            print(f'  [{addr}] フィルタ失敗: {e}')
            continue

        mails = []
        try:
            item = filtered.GetFirst()
            while item is not None:
                try:
                    subject  = item.Subject or ''
                    sender   = item.SenderName or ''
                    sender_addr = item.SenderEmailAddress or ''
                    recv_time = item.ReceivedTime
                    # COMのDateTimeはローカル時刻のdatetime
                    if hasattr(recv_time, 'year'):
                        dt = recv_time.replace(tzinfo=None)
                    else:
                        dt = datetime.now()

                    body_preview = (item.Body or '')[:200].replace('\r\n', ' ').replace('\n', ' ').strip()
                    unread = item.UnRead

                    mails.append({
                        'account': addr,
                        'subject': subject,
                        'sender': sender,
                        'sender_addr': sender_addr,
                        'dt': dt,
                        'body': body_preview,
                        'unread': unread,
                        'alert': is_alert(subject + ' ' + body_preview),
                        'ignorable': is_ignorable_sender(sender_addr),
                    })
                except Exception:
                    pass
                try:
                    item = filtered.GetNext()
                except Exception:
                    break
        except Exception as e:
            print(f'  [{addr}] メール取得エラー: {e}')

        # 未読総数を取得
        try:
            unread_total = inbox.UnReadItemCount
        except Exception:
            unread_total = '?'

        results.append({
            'account': addr,
            'mails': mails,
            'unread_total': unread_total,
        })

    return results


def print_report(account_results, since_dt):
    since_str = since_dt.strftime('%m/%d(%a) %H:%M')
    print(f'=== Outlookメールチェック（{since_str}〜現在） ===')

    for acc_data in account_results:
        addr = acc_data['account']
        mails = acc_data['mails']
        unread_total = acc_data['unread_total']

        print(f'\n【{addr}】  未読合計: {unread_total}件  期間内: {len(mails)}件')

        if not mails:
            print('  期間内のメールなし')
            continue

        alert_mails   = [m for m in mails if m['alert'] and not m['ignorable']]
        unread_mails  = [m for m in mails if m['unread'] and not m['ignorable'] and not m['alert']]
        normal_mails  = [m for m in mails if not m['alert'] and not m['unread'] and not m['ignorable']]
        ignored_mails = [m for m in mails if m['ignorable']]

        def fmt_mail(m, prefix='  '):
            dt_str = m['dt'].strftime('%m/%d %H:%M')
            flag = '⚠ ' if m['alert'] else ('● ' if m['unread'] else '  ')
            print(f'{prefix}{flag}{dt_str}  {m["sender"][:20]}')
            print(f'{prefix}      件名: {m["subject"][:60]}')
            if m['alert'] or m['unread']:
                print(f'{prefix}      {m["body"][:80]}')

        if alert_mails:
            print(f'\n  ⚠ 要注意メール ({len(alert_mails)}件)')
            for m in alert_mails[:10]:
                fmt_mail(m)

        if unread_mails:
            print(f'\n  ● 未読メール ({len(unread_mails)}件)')
            for m in unread_mails[:10]:
                fmt_mail(m)

        if normal_mails:
            print(f'\n  既読メール ({len(normal_mails)}件)')
            for m in normal_mails[:5]:
                print(f'    {m["dt"].strftime("%m/%d %H:%M")}  {m["sender"][:20]}  {m["subject"][:50]}')

        if ignored_mails:
            print(f'\n  （自動送信等 {len(ignored_mails)}件 省略）')


def run_check():
    since_dt = prev_business_day()
    try:
        account_results = get_outlook_mails(since_dt)
    except Exception as e:
        print(f'Outlook接続失敗: {e}')
        return

    print_report(account_results, since_dt)


if __name__ == '__main__':
    run_check()
