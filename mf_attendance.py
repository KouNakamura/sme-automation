"""
MFクラウド勤怠 自動CSV取得・集計スクリプト

機能:
  - メール+パスワードでログイン（資格情報マネージャー使用）
  - 打刻漏れ・未申請チェック（毎日）
  - 残業時間集計・アラート
  - 月次勤怠サマリーレポート
  - CSVをNASに保存

usage: python -X utf8 mf_attendance.py [--mode daily|monthly|check]
"""
import sys, io, os, time, re, csv, ctypes, ctypes.wintypes
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

import urllib.request, urllib.parse, json
from datetime import datetime, date, timedelta
from pathlib import Path

# ─── 設定 ────────────────────────────────────────────
MF_BASE      = 'https://attendance.moneyforward.com'
NAS_SAVE_DIR = r'\\YOUR_NAS_IP\keiri\99.claude連携\reports\mf_attendance'
CW_ROOM_ID   = YOUR_SALES_ROOM_ID
CW_BASE      = 'https://api.chatwork.com/v2'

# 残業アラートの閾値（月間）
OVERTIME_WARN_H  = 30   # 注意
OVERTIME_ALERT_H = 45   # アラート（法定上限警告）


def _load_credential(target_name):
    """Windowsの資格情報マネージャーからID/PWを取得"""
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
            if (c.TargetName or '').lower() == target_name.lower():
                username = c.UserName or ''
                size = c.CredentialBlobSize
                if size > 0:
                    blob = bytes(c.CredentialBlob[j] for j in range(size))
                    try:    pw = blob.decode('utf-16-le').rstrip('\x00')
                    except: pw = blob.decode('utf-8', errors='replace')
                    return username, pw
        advapi32.CredFree(creds)
    return None, None


def _load_cw_token():
    _, token = _load_credential('chatwork_api_token')
    return token or 'b0c820c2a79087252f191c16b9b9d4ef'


def _get_otp_from_outlook(since_dt, max_wait=60):
    """Outlookで受信してMF IDのOTPコードを取得する（最大max_wait秒リトライ）"""
    import win32com.client
    since_str = since_dt.strftime('%m/%d/%Y %H:%M %p')

    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        ns = outlook.GetNamespace('MAPI')
    except Exception as e:
        print(f'    Outlook接続失敗: {e}')
        return None

    def check_inbox():
        for acc in ns.Accounts:
            try:
                store    = acc.DeliveryStore
                inbox    = store.GetDefaultFolder(6)
                filtered = inbox.Items.Restrict("[ReceivedTime] >= '" + since_str + "'")
                item = filtered.GetFirst()
                while item is not None:
                    subject = item.Subject or ''
                    sender  = item.SenderEmailAddress or ''
                    body    = item.Body or ''
                    if 'moneyforward' in sender.lower() or 'マネーフォワード' in subject:
                        codes = re.findall(r'\b(\d{6})\b', body)
                        if codes:
                            return codes[0]
                    try:    item = filtered.GetNext()
                    except: break
            except Exception:
                pass
        return None

    elapsed = 0
    while elapsed < max_wait:
        # 受信を強制トリガー
        try:
            ns.SendAndReceive(False)
        except Exception:
            pass
        time.sleep(8)
        elapsed += 8

        otp = check_inbox()
        if otp:
            print(f'    OTP発見: {otp} ({elapsed}秒後)')
            return otp
        print(f'    OTP待ち中... ({elapsed}秒経過)')

    return None


def _mf_login_driver():
    """MFクラウド勤怠にSeleniumでログイン"""
    import tempfile
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    email, password = _load_credential('MFCloud_Attendance')
    if not email or not password:
        print('⚠ 資格情報が見つかりません。以下を実行してください:')
        print('  cmdkey /generic:MFCloud_Attendance /user:"メールアドレス" /pass:"パスワード"')
        return None

    driver_path = os.path.join(tempfile.gettempdir(), 'edgedriver_mf', 'msedgedriver.exe')
    opts = Options()
    opts.add_argument('--headless')
    opts.add_argument('--no-sandbox')
    opts.add_argument('--disable-dev-shm-usage')
    opts.add_argument('--disable-gpu')
    opts.add_experimental_option('prefs', {
        'download.default_directory': str(Path(tempfile.gettempdir()) / 'mf_dl'),
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
    })

    svc = Service(driver_path)
    driver = webdriver.Edge(service=svc, options=opts)

    try:
        from datetime import datetime as _dt
        login_start = _dt.now()
        print(f'MFクラウド勤怠 ログイン中... ({login_start.strftime("%H:%M:%S")})')

        # STEP1: 勤怠ログインページ → MF IDボタンをクリック
        driver.get(f'{MF_BASE}/employee_session/new')
        time.sleep(2)
        for a in driver.find_elements(By.TAG_NAME, 'a'):
            if 'マネーフォワード ID' in a.text:
                a.click()
                break
        time.sleep(2)

        # STEP2: MF ID ログイン（メール → パスワード → OTP）
        wait = WebDriverWait(driver, 10)
        from selenium.webdriver.common.keys import Keys
        try:
            # メール入力 → Enter
            email_field = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[name='mfid_user[email]']")
            ))
            email_field.send_keys(email)
            email_field.send_keys(Keys.RETURN)
            time.sleep(2)

            # パスワード入力 → Enter
            pw_field = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[type='password']")
            ))
            pw_field.send_keys(password)
            pw_field.send_keys(Keys.RETURN)
            time.sleep(3)

            # OTP画面チェック（/email_otp）
            if 'email_otp' in driver.current_url or 'otp' in driver.current_url.lower():
                print('  メールOTP認証が必要です。Outlookからコードを取得中...')
                otp = _get_otp_from_outlook(login_start)
                if otp:
                    otp_field = wait.until(EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "input[name='email_otp']")
                    ))
                    otp_field.send_keys(otp)
                    otp_field.send_keys(Keys.RETURN)
                    time.sleep(3)
                    print(f'  OTP入力完了: {otp}')
                else:
                    print('  ⚠ OTPをOutlookから取得できませんでした')
                    return None

            # パスキー登録プロモーションページをスキップ
            if 'passkey_promotion' in driver.current_url:
                for a in driver.find_elements(By.TAG_NAME, 'a'):
                    if any(kw in a.text for kw in ['登録せず', 'スキップ', '後で', '次へ']):
                        print(f'  パスキー登録スキップ: {a.text.strip()!r}')
                        a.click()
                        time.sleep(3)
                        break

        except Exception as e:
            print(f'  ログイン入力失敗: {e}')
            print(f'  現在URL: {driver.current_url}')
            return None

        current = driver.current_url
        print(f'  ログイン後URL: {current}')
        # id.moneyforward.com 上に留まっていたらログイン失敗
        if not current.startswith('https://attendance.moneyforward.com'):
            print('  ⚠ ログイン失敗（勤怠サイトへのリダイレクト未完了）')
            return None
        if 'session' in current or 'login' in current:
            print('  ⚠ ログイン失敗（セッション未確立）')
            return None
        print('  ✓ ログイン成功')
        return driver

    except Exception as e:
        print(f'ログインエラー: {e}')
        try: driver.quit()
        except: pass
        return None


def _parse_time_hhmm(val):
    """'HH:MM' 形式の時間文字列を分に変換（例: '88:50' → 5330）"""
    if not val or not val.strip():
        return 0
    m = re.match(r'(\d+):(\d+)', val.strip())
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return 0


def download_csv(driver, year=None, month=None):
    """指定年月の月別勤怠CSVをエクスポートAPIで取得してテキストを返す（CP932）"""
    from selenium.webdriver.common.by import By
    import tempfile

    if year is None:
        year  = date.today().year
    if month is None:
        month = date.today().month

    dl_dir = Path(tempfile.gettempdir()) / 'mf_dl'
    dl_dir.mkdir(exist_ok=True)
    for f in dl_dir.glob('*.csv'):
        try: f.unlink()
        except: pass

    print(f'  CSV取得: {year}年{month}月')
    EXPORTER_TYPE = 'CsvExporter::MonthlyAttendanceItemCsvExporter'
    FORM_PATH     = 'monthly_attendance_item_csv_exporters'
    FORM_PREFIX   = 'admin_settings_exporters_monthly_attendance_item_csv_exporter_form'

    def _get_export_entry():
        """export_histories.json から今月の完了済みエントリを返す"""
        driver.get(f'{MF_BASE}/admin/export_histories.json')
        time.sleep(1)
        try:
            data = json.loads(driver.find_element(By.TAG_NAME, 'body').text)
        except Exception:
            return None
        for entry in data.get('exports', []):
            opts = entry.get('options', {})
            if (entry.get('model_type') == EXPORTER_TYPE
                    and opts.get('year') == year
                    and opts.get('month') == month
                    and entry.get('status') == 'succeeded'):
                return entry
        return None

    # 既存エントリがあればそのまま使う
    entry = _get_export_entry()

    if not entry:
        print('  エクスポートジョブを作成中...')
        driver.get(f'{MF_BASE}/admin/settings/exporters/{FORM_PATH}/new')
        time.sleep(2)
        driver.execute_script(f"""
            var mf = document.querySelector("input[name='{FORM_PREFIX}[month]']");
            if (mf) mf.value = '{month}';
            var yf = document.querySelector("input[name='{FORM_PREFIX}[year]']");
            if (yf) yf.value = '{year}';
        """)
        submit_el = driver.find_element(By.CSS_SELECTOR, "input[name='commit']")
        driver.execute_script("arguments[0].click();", submit_el)

        # ジョブ完了を最大60秒ポーリング
        for i in range(10):
            time.sleep(6)
            entry = _get_export_entry()
            if entry:
                print(f'  エクスポート完了 ({6*(i+1)}秒後)')
                break
            print(f'  ジョブ待ち... ({6*(i+1)}秒)')
        if not entry:
            print('  ⚠ エクスポートジョブがタイムアウト')
            return None

    # ダウンロード実行
    dl_url = f'{MF_BASE}{entry["download_url"]}?file_format=utf8_csv'
    print(f'  ダウンロード: {entry["filename"]}')
    driver.get(dl_url)
    time.sleep(4)

    csv_files = list(dl_dir.glob('*.csv'))
    if not csv_files:
        print('  ⚠ CSVファイルが落ちてきませんでした')
        return None

    print(f'  ✓ ダウンロード成功: {csv_files[0].name}')
    return csv_files[0].read_bytes().decode('cp932', errors='replace')


def parse_csv(csv_bytes_or_text):
    """CSVを解析して勤怠データを返す"""
    if not csv_bytes_or_text:
        return []
    if isinstance(csv_bytes_or_text, bytes):
        text = csv_bytes_or_text.decode('cp932', errors='replace')
    else:
        text = csv_bytes_or_text
    rows = []
    reader = csv.DictReader(io.StringIO(text))
    for row in reader:
        rows.append(row)
    return rows


def analyze_attendance(rows):
    """月別CSVから残業超過・欠勤・遅刻を集計して返す"""
    overtime = {}   # 氏名 → (所定外分, 法定外分)
    absences = {}   # 氏名 → 欠勤日数
    lates    = {}   # 氏名 → 遅刻回数

    for row in rows:
        name = row.get('氏名', '').strip()
        if not name:
            continue

        # 残業時間（所定外 = 所定外時間（平日）、法定外 = 法定外時間（平日））
        ot_soc = _parse_time_hhmm(row.get('所定外時間（平日）', ''))
        ot_leg = _parse_time_hhmm(row.get('法定外時間（平日）', ''))
        overtime[name] = (ot_soc, ot_leg)

        # 欠勤
        try:
            absence_days = float(row.get('欠勤日数（平日）', '0') or '0')
        except ValueError:
            absence_days = 0
        if absence_days > 0:
            absences[name] = absence_days

        # 遅刻
        try:
            late_count = int(row.get('遅刻回数（平日）', '0') or '0')
        except ValueError:
            late_count = 0
        if late_count > 0:
            lates[name] = late_count

    ot_alerts = []
    for name, (ot_soc, ot_leg) in overtime.items():
        hours = ot_leg / 60  # 法定外時間で判定
        if hours >= OVERTIME_ALERT_H:
            ot_alerts.append({'name': name, 'hours': ot_soc/60, 'legal_hours': hours, 'level': 'アラート'})
        elif hours >= OVERTIME_WARN_H:
            ot_alerts.append({'name': name, 'hours': ot_soc/60, 'legal_hours': hours, 'level': '注意'})

    return ot_alerts, absences, lates, overtime


def post_to_cw(token, msg):
    data = urllib.parse.urlencode({'body': msg}).encode()
    req  = urllib.request.Request(
        f'{CW_BASE}/rooms/{CW_ROOM_ID}/messages',
        data=data,
        headers={'X-ChatWorkToken': token},
        method='POST'
    )
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode()).get('message_id')


# ─── メイン ───────────────────────────────────────────
def main():
    mode  = sys.argv[1] if len(sys.argv) > 1 else 'check'
    today = date.today()

    # 月指定オプション: python mf_attendance.py monthly 2026/04
    target_year  = today.year
    target_month = today.month
    if len(sys.argv) > 2:
        try:
            td = datetime.strptime(sys.argv[2], "%Y/%m")
            target_year  = td.year
            target_month = td.month
            print(f'対象月指定: {target_year}/{target_month:02d}')
        except ValueError:
            print(f'月指定形式エラー: {sys.argv[2]} (例: 2026/04)')

    driver = _mf_login_driver()
    if not driver:
        sys.exit(1)

    try:
        csv_text = download_csv(driver, target_year, target_month)
        if not csv_text:
            print('CSVを取得できませんでした')
            return

        # NASに保存
        try:
            nas_dir = Path(NAS_SAVE_DIR)
            nas_dir.mkdir(parents=True, exist_ok=True)
            fname = nas_dir / f'attendance_{target_year}{target_month:02d}.csv'
            fname.write_text(csv_text, encoding='utf-8-sig')
            print(f'NAS保存: {fname}')
        except Exception as e:
            print(f'NAS保存失敗（スキップ）: {e}')

        rows = parse_csv(csv_text)
        print(f'取得レコード数: {len(rows)}行')

        ot_alerts, absences, lates, overtime = analyze_attendance(rows)

        if mode in ('daily', 'check'):
            lines = [f'[info][title]■ 勤怠チェック（{today.strftime("%Y/%m/%d")} {target_year}年{target_month}月分）[/title]']

            if ot_alerts:
                lines.append(f'【残業アラート】')
                for a in ot_alerts:
                    lines.append(f'・{a["name"]} 所定外:{a["hours"]:.1f}h 法定外:{a["legal_hours"]:.1f}h ─ {a["level"]}')
            else:
                lines.append('【残業アラート】なし ✓')

            if absences:
                lines.append(f'\n【欠勤】')
                for name, days in sorted(absences.items(), key=lambda x: -x[1])[:10]:
                    lines.append(f'・{name}: {days:.0f}日')
            else:
                lines.append('\n【欠勤】なし ✓')

            if lates:
                lines.append(f'\n【遅刻】')
                for name, count in sorted(lates.items(), key=lambda x: -x[1])[:10]:
                    lines.append(f'・{name}: {count}回')
            else:
                lines.append('\n【遅刻】なし ✓')

            lines.append('[/info]')
            msg = '\n'.join(lines)
            print(msg)
            token = _load_cw_token()
            mid = post_to_cw(token, msg)
            print(f'投稿成功: message_id={mid}')

        elif mode == 'monthly':
            lines = [f'[info][title]■ 月次勤怠サマリー（{target_year}年{target_month}月）[/title]']
            lines.append(f'従業員数: {len(rows)}名\n')

            # 残業ランキング（所定外時間で並び替え）
            lines.append('【残業時間ランキング（上位10名）】')
            sorted_ot = sorted(overtime.items(), key=lambda x: -x[1][0])  # 所定外時間で降順
            for name, (ot_soc, ot_leg) in sorted_ot[:10]:
                h_soc = ot_soc // 60;  m_soc = ot_soc % 60
                h_leg = ot_leg // 60;  m_leg = ot_leg % 60
                flag = ' ⚠' if ot_leg/60 >= OVERTIME_ALERT_H else (' △' if ot_leg/60 >= OVERTIME_WARN_H else '')
                lines.append(f'・{name}: 所定外 {h_soc}h{m_soc:02d}m / 法定外 {h_leg}h{m_leg:02d}m{flag}')

            if absences:
                lines.append(f'\n【欠勤あり】{len(absences)}名')
                for name, days in sorted(absences.items(), key=lambda x: -x[1])[:10]:
                    lines.append(f'・{name}: {days:.0f}日')

            if lates:
                lines.append(f'\n【遅刻あり】{len(lates)}名')
                for name, count in sorted(lates.items(), key=lambda x: -x[1])[:10]:
                    lines.append(f'・{name}: {count}回')

            if ot_alerts:
                lines.append(f'\n【残業アラート】{len(ot_alerts)}名')
                for a in ot_alerts:
                    lines.append(f'・{a["name"]} 法定外:{a["legal_hours"]:.1f}h ─ {a["level"]}')

            lines.append('[/info]')
            msg = '\n'.join(lines)
            print(msg)
            token = _load_cw_token()
            mid = post_to_cw(token, msg)
            print(f'投稿成功: message_id={mid}')

    finally:
        try: driver.quit()
        except: pass


if __name__ == '__main__':
    main()
