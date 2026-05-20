"""
直近1ヶ月の営業メール傾向分析
usage: python -X utf8 outlook_sales_analysis.py
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

import win32com.client
from datetime import datetime, timedelta
from collections import defaultdict, Counter

SALES_KEYWORDS = [
    '案件', '提案', '商談', '見積', '契約', 'ご紹介', 'SES', 'BP', '協力会社',
    'リプレイス', '受注', '失注', '常駐', '稼働', 'アサイン', '要員',
    'Java', 'Python', '.NET', 'COBOL', 'インフラ', 'AWS', 'Azure', 'SAP', 'PMO',
    '募集', '紹介', '人材',
]

HR_KEYWORDS = [
    '入社', '退社', '採用', '内定', '退職', '休職', '評価', '昇給', '昇格', '給与',
    '始業', '終業', '勤怠',
]

TROUBLE_KEYWORDS = [
    'クレーム', '苦情', 'トラブル', '緊急', '至急', 'ミス', '事故', '謝罪', 'お詫び',
]

IGNORE_DOMAINS = {
    'gmail.com', 'yahoo.co.jp', 'hotmail.com', 'outlook.com',
    'amazon.co.jp', 'amazon.com', 'rakuten.co.jp', 'noreply.github.com',
    'americanexpress.com', 'adobe.com', 'microsoft.com',
}


def run_analysis():
    since_dt = datetime.now() - timedelta(days=30)
    since_str = since_dt.strftime('%m/%d/%Y %H:%M %p')

    outlook = win32com.client.Dispatch('Outlook.Application')
    ns = outlook.GetNamespace('MAPI')

    all_mails = []

    for acc in ns.Accounts:
        addr = acc.SmtpAddress or ''
        if 'brain-d.jp' not in addr:
            continue
        try:
            store = acc.DeliveryStore
            inbox = store.GetDefaultFolder(6)
            items = inbox.Items
            items.Sort('[ReceivedTime]', True)
            filtered = items.Restrict("[ReceivedTime] >= '" + since_str + "'")

            item = filtered.GetFirst()
            while item is not None:
                try:
                    subject     = item.Subject or ''
                    sender      = item.SenderName or ''
                    sender_addr = (item.SenderEmailAddress or '').lower()
                    recv_time   = item.ReceivedTime
                    body_prev   = (item.Body or '')[:300]
                    unread      = item.UnRead

                    all_mails.append({
                        'account':     addr,
                        'subject':     subject,
                        'sender':      sender,
                        'sender_addr': sender_addr,
                        'dt':          recv_time,
                        'body':        body_prev,
                        'unread':      unread,
                    })
                except Exception:
                    pass
                try:
                    item = filtered.GetNext()
                except Exception:
                    break
        except Exception as e:
            print(f'  [{addr}] 取得失敗: {e}')

    print(f'=== 直近1ヶ月の営業メール傾向分析（{since_dt.strftime("%m/%d")}〜本日）===')
    print(f'取得メール総数: {len(all_mails)}件')
    print()

    def contains_any(text, keywords):
        return any(kw in text for kw in keywords)

    sales_mails   = [m for m in all_mails if contains_any(m['subject'] + m['body'], SALES_KEYWORDS)]
    hr_mails      = [m for m in all_mails if contains_any(m['subject'] + m['body'], HR_KEYWORDS)]
    trouble_mails = [m for m in all_mails if contains_any(m['subject'] + m['body'], TROUBLE_KEYWORDS)]

    print('■ カテゴリ別件数')
    print(f'  営業・案件関連: {len(sales_mails)}件')
    print(f'  人事・勤怠関連: {len(hr_mails)}件')
    print(f'  要注意（クレーム・緊急等）: {len(trouble_mails)}件')
    print()

    # 送信者ドメイン別集計（外部）
    domain_counter = Counter()
    for m in all_mails:
        addr = m['sender_addr']
        if '@' not in addr:
            continue
        domain = addr.split('@')[1]
        if domain not in IGNORE_DOMAINS and 'brain-d.jp' not in domain:
            domain_counter[domain] += 1

    print('■ 主要取引先ドメイン（上位15件）')
    for domain, cnt in domain_counter.most_common(15):
        print(f'  {domain}: {cnt}件')
    print()

    # 週別トレンド
    week_counts = defaultdict(int)
    for m in all_mails:
        try:
            dt = m['dt']
            week = dt.strftime('%Y-W%U')
            week_counts[week] += 1
        except Exception:
            pass

    print('■ 週別メール量（受信トレンド）')
    for wk in sorted(week_counts.keys()):
        bar = '█' * (week_counts[wk] // 10)
        print(f'  {wk}: {week_counts[wk]:3d}件 {bar}')
    print()

    # 要注意メール
    if trouble_mails:
        print(f'■ 要注意メール一覧（{len(trouble_mails)}件）')
        for m in sorted(trouble_mails, key=lambda x: x['dt'], reverse=True)[:15]:
            try:
                dt_str = m['dt'].strftime('%m/%d %H:%M')
            except Exception:
                dt_str = '??'
            print(f'  {dt_str}  {m["sender"][:20]}  /  {m["subject"][:55]}')
        print()

    # 営業案件キーワード頻出分析
    kw_counter = Counter()
    for m in sales_mails:
        txt = m['subject'] + ' ' + m['body']
        for kw in SALES_KEYWORDS:
            if kw in txt:
                kw_counter[kw] += 1

    print('■ 案件キーワード頻出（営業メール内）')
    for kw, cnt in kw_counter.most_common(15):
        print(f'  {kw}: {cnt}回')
    print()

    # 主要送信者（社外）上位
    sender_counter = Counter()
    for m in all_mails:
        addr = m['sender_addr']
        if 'brain-d.jp' not in addr and addr:
            sender_counter[m['sender']] += 1

    print('■ 社外からのメール送信者（上位10名）')
    for sender, cnt in sender_counter.most_common(10):
        print(f'  {sender}: {cnt}件')
    print()

    # 勤怠連絡のある社員（始業・終業・体調）
    kinmu_counter = Counter()
    for m in all_mails:
        subject = m['subject']
        if any(kw in subject for kw in ['始業', '終業', '勤怠', '在宅']):
            kinmu_counter[m['sender']] += 1

    if kinmu_counter:
        print('■ 勤怠連絡メール（社員別）')
        for sender, cnt in kinmu_counter.most_common(20):
            print(f'  {sender}: {cnt}件')
        print()

    # 結論サマリー
    print('■ 傾向サマリー')
    total = len(all_mails)
    if total > 0:
        sales_pct   = len(sales_mails) * 100 // total
        trouble_pct = len(trouble_mails) * 100 // total
        print(f'  受信メール全体のうち案件関連: {sales_pct}%（{len(sales_mails)}件）')
        print(f'  要注意メールの割合: {trouble_pct}%（{len(trouble_mails)}件）')

    top_domain = domain_counter.most_common(1)
    if top_domain:
        print(f'  最も頻繁な取引先ドメイン: {top_domain[0][0]}（{top_domain[0][1]}件）')

    top_kw = kw_counter.most_common(1)
    if top_kw:
        print(f'  最頻出案件キーワード: 「{top_kw[0][0]}」（{top_kw[0][1]}回）')


if __name__ == '__main__':
    run_analysis()
