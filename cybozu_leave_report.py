"""
サイボウズ有休・休暇申請 × MFクラウド勤怠 クロス集計レポート
Usage: python -X utf8 cybozu_leave_report.py [YYYY/MM]
デフォルト: 前月
"""
import sys, io, os, time, re, json, csv, ctypes
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

import calendar
from pathlib import Path
from datetime import date, datetime

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

BASE         = "https://YOUR_COMPANY.cybozu.com"
EDGE_DRIVER  = os.path.join(os.environ["TEMP"], "edgedriver_mf", "msedgedriver.exe")
EDGE_PROFILE = os.path.join(os.environ["USERPROFILE"], ".config", "cybozu_profile")
NAS_REPORT   = r"\\YOUR_NAS_IP\YOUR_FOLDER\99.automation\reports\attendance"

# サイボウズのフォームID
FID_LEAVE    = 153    # 休暇申請/遅早届
FID_HOLIDAY  = 5233  # 休日出勤申請

# --- 対象月 ---
if len(sys.argv) > 1:
    try:
        target = datetime.strptime(sys.argv[1], "%Y/%m")
    except ValueError:
        print(f"日付形式エラー: {sys.argv[1]} (例: 2026/04)"); sys.exit(1)
else:
    today = date.today()
    m = today.month - 1 or 12
    y = today.year if today.month > 1 else today.year - 1
    target = datetime(y, m, 1)

year  = target.year
month = target.month
print(f"対象月: {year}/{month:02d}")

# ─── Credential Manager ───────────────────────────────────────────
def _read_cred(target_name: str):
    CRED_TYPE_GENERIC = 1
    class _CREDENTIAL(ctypes.Structure):
        _fields_ = [
            ("Flags",              ctypes.c_ulong),
            ("Type",               ctypes.c_ulong),
            ("TargetName",         ctypes.c_wchar_p),
            ("Comment",            ctypes.c_wchar_p),
            ("LastWritten",        ctypes.c_ulonglong),
            ("CredentialBlobSize", ctypes.c_ulong),
            ("CredentialBlob",     ctypes.c_char_p),
            ("Persist",            ctypes.c_ulong),
            ("AttributeCount",     ctypes.c_ulong),
            ("Attributes",         ctypes.c_void_p),
            ("TargetAlias",        ctypes.c_wchar_p),
            ("UserName",           ctypes.c_wchar_p),
        ]
    advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
    ptr = ctypes.POINTER(_CREDENTIAL)()
    ok  = advapi32.CredReadW(target_name, CRED_TYPE_GENERIC, 0, ctypes.byref(ptr))
    if not ok:
        return None, None
    cred = ptr.contents
    blob = cred.CredentialBlob[:cred.CredentialBlobSize].decode("utf-16-le", errors="replace")
    user = cred.UserName or ""
    advapi32.CredFree(ptr)
    return user, blob

# ─── Edge ドライバー ──────────────────────────────────────────────
def make_driver(headless=True):
    Path(EDGE_PROFILE).mkdir(parents=True, exist_ok=True)
    opts = Options()
    opts.add_argument(f"--user-data-dir={EDGE_PROFILE}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-extensions")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1280,900")
    svc = Service(executable_path=EDGE_DRIVER)
    return webdriver.Edge(service=svc, options=opts)

def is_logged_in(driver):
    url = driver.current_url
    return BASE in url and '/login' not in url and 'cloudLogin' not in url.lower()

def login_cybozu(driver):
    email, password = _read_cred("Cybozu_Login")
    if not email:
        print("  認証情報 Cybozu_Login が見つかりません"); return False
    driver.get(BASE + "/o/ag.cgi")
    time.sleep(4)
    if is_logged_in(driver):
        return True
    try:
        email_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))
        )
        email_input.clear()
        email_input.send_keys(email)
        email_input.submit()
        time.sleep(3)
        pw_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))
        )
        pw_input.clear()
        pw_input.send_keys(password)
        pw_input.submit()
        time.sleep(5)
        for _ in range(3):
            try:
                driver.find_element(By.ID, "idBtn_Back").click(); time.sleep(2)
            except: pass
        driver.get(BASE + "/o/ag.cgi")
        time.sleep(3)
    except Exception as e:
        print(f"  ログイン例外: {e}")
    return is_logged_in(driver)

# ─── テーブル解析（ヘッダー行を探して正しいテーブルを選択） ────────
def _parse_result_table(driver, seen_wids=None):
    """
    WorkFlowSearch/WorkFlowManageWhole の結果テーブルを解析。
    「番号」「申請者」「状況」「申請日」をヘッダーに持つテーブルを探す。
    """
    if seen_wids is None:
        seen_wids = set()
    rows = []

    for table in driver.find_elements(By.TAG_NAME, 'table'):
        trs = table.find_elements(By.TAG_NAME, 'tr')
        if len(trs) < 3:
            continue

        # ヘッダー行を探す
        header_idx = None
        col_map    = {}
        for i, tr in enumerate(trs):
            cells = tr.find_elements(By.TAG_NAME, 'th') or tr.find_elements(By.TAG_NAME, 'td')
            texts = [c.text.strip() for c in cells]
            if '番号' in texts and '申請者' in texts and '状況' in texts:
                header_idx = i
                col_map = {t: j for j, t in enumerate(texts)}
                break

        if header_idx is None:
            continue

        # データ行を解析
        for tr in trs[header_idx + 1:]:
            tds = tr.find_elements(By.TAG_NAME, 'td')
            if len(tds) < 4:
                continue
            texts = [td.text.strip() for td in tds]

            # 番号列が数値かチェック
            serial = texts[0] if texts else ''
            if not serial.isdigit():
                continue
            if serial in seen_wids:
                continue
            seen_wids.add(serial)

            # リンクから wid を取得
            links = tds[1].find_elements(By.TAG_NAME, 'a') if len(tds) > 1 else []
            href  = links[0].get_attribute('href') if links else ''
            wid_m = re.search(r'wid=(\d+)', href)
            wid   = int(wid_m.group(1)) if wid_m else 0

            def _get(idx, fallback=''):
                try: return texts[idx]
                except IndexError: return fallback

            # インデックスを col_map から安全に取得（デフォルト: 位置順）
            fi = col_map.get('申請フォーム名（標題）', 1)
            si = col_map.get('状況',   2)
            ai = col_map.get('申請者', 3)
            di = col_map.get('申請日', 4)
            form_name = _get(fi if fi < len(texts) else 1)
            status    = _get(si if si < len(texts) else 2)
            applicant = _get(ai if ai < len(texts) else 3)
            app_date  = _get(di if di < len(texts) else 4)

            rows.append({
                'serial':    serial,
                'wid':       wid,
                'form_name': form_name,
                'status':    status,
                'applicant': applicant,
                'date':      app_date,
                'href':      href,
            })

        if rows:
            break  # 見つかったテーブルで終了

    return rows

# ─── 日付ユーティリティ ───────────────────────────────────────────
def _parse_cybozu_date(date_str):
    """2026/4/29 や 2026/04/29 を (year, month, day) に変換"""
    m = re.match(r'(\d{4})/(\d{1,2})/(\d{1,2})', (date_str or '').strip())
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    return None, None, None

def _in_month(date_str, year, month):
    y, mo, _ = _parse_cybozu_date(date_str)
    return y == year and mo == month

def _before_month(date_str, year, month):
    y, mo, _ = _parse_cybozu_date(date_str)
    if y is None:
        return False
    return (y, mo) < (year, month)

# ─── WorkFlowSearch で休暇申請を検索（全ページ） ─────────────────
def search_leave_cybozu(driver, year, month):
    """
    WorkFlowSearch で休暇申請/遅早届を検索し、対象月の申請一覧を全件返す。
    受信申請として中村さんが承認者になっている全社員分が取れる。
    """
    print(f"  サイボウズ休暇申請検索: {year}/{month:02d}")

    results   = []
    seen_wids = set()

    driver.get(f"{BASE}/o/ag.cgi?page=WorkFlowSearch&cp=ww")
    time.sleep(2)

    driver.execute_script("""
        var f = document.querySelector('form input[name="Text0"]');
        if (!f) return;
        var form = f.closest('form');

        var rowsSel = form.querySelector('select[name="Rows"]');
        if (rowsSel) rowsSel.value = '100';

        var iid = form.querySelector('select[name="IID0"]');
        if (iid) iid.value = 'Name';

        var expr = form.querySelector('select[name="Expr0"]');
        if (expr) expr.value = '1';

        f.value = '休暇申請';

        var submit = form.querySelector('input[name="Submit"]');
        if (submit) submit.click();
    """)
    time.sleep(3)

    page_num   = 0
    found_old  = False
    while True:
        page_num += 1
        page_rows = _parse_result_table(driver, seen_wids)
        print(f"    ページ{page_num}: {len(page_rows)}件")

        in_month = [r for r in page_rows if _in_month(r.get('date', ''), year, month)]
        before   = [r for r in page_rows if _before_month(r.get('date', ''), year, month)]

        results.extend(in_month)

        # 対象月より古いデータが出始めたら次ページはない
        if before:
            found_old = True
        if found_old and not in_month:
            print(f"    対象月より古いデータのみ → 終了")
            break
        if not page_rows:
            break

        next_links = [a for a in driver.find_elements(By.TAG_NAME, 'a')
                      if '次の' in (a.text or '') and '件へ' in (a.text or '')]
        if not next_links or page_num >= 30:
            break
        next_links[0].click()
        time.sleep(2)

    return results

# ─── WorkFlowView で申請詳細取得 ─────────────────────────────────
def get_leave_detail(driver, wid):
    """申請詳細ページから休暇日付・種別を取得"""
    driver.get(f"{BASE}/o/ag.cgi?page=WorkFlowManageView&wid={wid}&cp=wwma&cid=0&fid={FID_LEAVE}")
    time.sleep(1)
    body  = driver.find_element(By.TAG_NAME, 'body').text

    detail = {}

    # 休暇種別を探す
    leave_types = ['年次有給休暇', '有給休暇', '特別休暇', '慶弔休暇', '育児休業', '産前産後休業',
                   '病気休暇', '半日有給', '時間有給', '振替休日', '代休', '欠勤', '遅刻', '早退',
                   '介護', '公用外出', 'AM半休', 'PM半休']
    found_types = [lt for lt in leave_types if lt in body]
    detail['leave_types'] = found_types

    # 日付パターンを探す（申請内容中の休暇日付）
    # フォームの入力値テーブルを解析
    date_cells = []
    for table in driver.find_elements(By.TAG_NAME, 'table'):
        trs = table.find_elements(By.TAG_NAME, 'tr')
        for tr in trs:
            tds = tr.find_elements(By.TAG_NAME, 'td')
            for td in tds:
                text = td.text.strip()
                m = re.search(r'20\d{2}/\d{1,2}/\d{1,2}', text)
                if m:
                    date_cells.append(m.group(0))

    detail['leave_dates'] = sorted(set(date_cells))
    detail['body_excerpt'] = body[200:600] if len(body) > 200 else body

    return detail

# ─── NASからMF Cloud CSVを読み込む ──────────────────────────────
def load_mf_csv(year, month):
    """NASに保存済みのMFクラウド勤怠CSVを読み込む"""
    nas_path = Path(NAS_REPORT)
    pattern  = f"attendance_{year}{month:02d}.csv"
    csv_path = nas_path / pattern

    if not csv_path.exists():
        print(f"  MF CSV不在: {csv_path}")
        print(f"  → python -X utf8 mf_attendance.py monthly で {year}/{month:02d} のCSVを取得してください")
        return []

    print(f"  MF CSV読み込み: {csv_path}")
    raw_bytes = csv_path.read_bytes()
    # mf_attendance.py は utf-8-sig で保存、旧来はcp932の場合あり
    if raw_bytes[:3] == b'\xef\xbb\xbf':
        raw = raw_bytes.decode('utf-8-sig', errors='replace')
    else:
        raw = raw_bytes.decode('cp932', errors='replace')
    reader = csv.DictReader(raw.splitlines())
    rows   = list(reader)
    print(f"  MF CSV: {len(rows)}名分")
    return rows

# ─── MF Cloud 欠勤者抽出 ─────────────────────────────────────────
def _parse_hhmm(s):
    s = (s or '').strip()
    if not s or s == '-':
        return 0
    m = re.match(r'^(\d+):(\d{2})$', s)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    try:
        return int(float(s) * 60)
    except:
        return 0

def extract_absences(mf_rows):
    absences = {}
    for row in mf_rows:
        name = (row.get('氏名') or '').strip()
        if not name:
            continue
        try:
            days = float(row.get('欠勤日数（平日）') or '0')
        except:
            days = 0
        if days > 0:
            absences[name] = days
    return absences

# ─── メイン ──────────────────────────────────────────────────────
driver = None
try:
    driver = make_driver(headless=True)

    # ログイン確認
    driver.get(BASE + "/o/ag.cgi")
    time.sleep(3)
    if not is_logged_in(driver):
        print("セッション切れ → ログイン中...")
        if not login_cybozu(driver):
            print("ログイン失敗"); sys.exit(1)
    print("サイボウズ: ログイン済み")

    # ─── Step 1: WorkFlowSearch で休暇申請を取得 ──────────────
    print(f"\n=== Cybozu 休暇申請取得 ({year}/{month:02d}) ===")
    leave_requests = search_leave_cybozu(driver, year, month)

    print(f"\n  対象月の総休暇申請: {len(leave_requests)}件")

    # 承認済み（完了）のみ
    approved = [r for r in leave_requests if '完了' in r.get('status', '')]
    pending  = [r for r in leave_requests if '進行中' in r.get('status', '')]
    print(f"  完了（承認済み）: {len(approved)}件")
    print(f"  進行中（未完了）: {len(pending)}件")

    # ─── Step 2: 申請者ごとに集計 ─────────────────────────────
    from collections import defaultdict
    person_approved = defaultdict(list)
    person_pending  = defaultdict(list)
    for r in approved:
        person_approved[r['applicant']].append(r)
    for r in pending:
        person_pending[r['applicant']].append(r)

    print(f"\n--- 承認済み申請者一覧 ---")
    for name, reqs in sorted(person_approved.items()):
        print(f"  {name}: {len(reqs)}件")
        for r in reqs:
            print(f"    #{r['serial']} {r['form_name'][:35]} ({r['date']})")

    if pending:
        print(f"\n--- 未承認（進行中）申請者一覧 ---")
        for name, reqs in sorted(person_pending.items()):
            print(f"  {name}: {len(reqs)}件")
            for r in reqs:
                print(f"    #{r['serial']} {r['form_name'][:35]} ({r['date']})")

    # ─── Step 3: 承認済み申請の詳細（休暇日付・種別）を取得 ─────
    print(f"\n=== 承認済み申請の詳細取得（最大20件） ===")
    detailed = []
    for req in approved[:20]:
        if not req.get('wid'):
            continue
        try:
            detail = get_leave_detail(driver, req['wid'])
            req.update(detail)
            dates  = detail.get('leave_dates', [])
            types  = detail.get('leave_types', [])
            print(f"  #{req['serial']} {req['applicant']}: {types} | 日付: {dates[:5]}")
        except Exception as e:
            print(f"  #{req['serial']} 詳細取得失敗: {e}")
        detailed.append(req)
        time.sleep(0.5)

    # ─── Step 4: MF Cloud 欠勤データ ──────────────────────────
    print(f"\n=== MF Cloud 欠勤データ ({year}/{month:02d}) ===")
    mf_rows  = load_mf_csv(year, month)
    absences = extract_absences(mf_rows)

    if absences:
        print(f"  欠勤あり: {len(absences)}名")
        for name, days in sorted(absences.items(), key=lambda x: -x[1]):
            print(f"  {name}: {days}日")
    else:
        print("  欠勤者なし（またはCSV未取得）")

    # ─── Step 5: クロス集計レポート ─────────────────────────────
    print(f"\n{'='*60}")
    print(f"  クロス集計レポート: {year}/{month:02d}")
    print(f"{'='*60}")

    approved_names = set(r['applicant'] for r in approved)

    print(f"\n【Cybozu 承認済み休暇申請者: {len(approved_names)}名】")
    for name in sorted(approved_names):
        reqs  = person_approved[name]
        types_set = set()
        dates_set = set()
        for r in reqs:
            types_set.update(r.get('leave_types', []))
            dates_set.update(r.get('leave_dates', []))
        type_str = '・'.join(types_set) if types_set else '（種別未取得）'
        print(f"  {name}: {len(reqs)}件 [{type_str}]")

    if mf_rows:
        print(f"\n【欠勤 × 休暇申請 突き合わせ】")
        print(f"{'氏名':<14}{'MF欠勤':>8}{'申請状況':<30}{'判定'}")
        print("-" * 65)

        for name, days in sorted(absences.items(), key=lambda x: -x[1]):
            # 名前部分一致（姓一致でもOK）
            matched_reqs = [r for r in approved
                           if r['applicant'] == name
                           or name in r['applicant']
                           or r['applicant'] in name
                           or (name.split() and r['applicant'].split()
                               and name.split()[0] == r['applicant'].split()[0])]

            if matched_reqs:
                types  = set()
                for r in matched_reqs:
                    types.update(r.get('leave_types', []))
                type_str = '・'.join(types) if types else '（種別不明）'
                verdict  = f"✓ 有休等承認済み [{type_str}]"
            else:
                verdict = "⚠ 要確認（Cybozu申請なし）"

            req_info = f"{len(matched_reqs)}件承認" if matched_reqs else "申請なし"
            print(f"  {name:<12}{days:>6.1f}日  {req_info:<25}  {verdict}")

        # サマリ
        matched = sum(1 for name in absences
                      if any(r['applicant'] == name or name in r['applicant'] or r['applicant'] in name
                             for r in approved))
        unmatched = len(absences) - matched

        print(f"\n【サマリ】")
        print(f"  MF Cloud 欠勤者:     {len(absences)}名")
        print(f"  うち Cybozu承認あり: {matched}名（有休・特別休暇等）")
        print(f"  うち 申請なし:       {unmatched}名（要確認）")
        print(f"  Cybozu 承認済み休暇: {len(approved_names)}名（MF欠勤未計上含む）")

    else:
        # MFデータなし → Cybozu申請一覧のみ表示
        print(f"\n【MF Cloudデータなし - Cybozu申請一覧のみ表示】")
        print(f"  MF Cloudの{year}/{month:02d}データを取得するには:")
        print(f"  python -X utf8 mf_attendance.py monthly")
        print(f"  （mf_attendance.py に月指定引数を追加する必要があります）")

    # ─── Step 6: 全申請一覧（参考） ────────────────────────────
    print(f"\n【全休暇申請一覧（{year}/{month:02d}）】")
    print(f"{'番号':>6} {'状況':^6} {'申請者':<12} {'申請日':<12} {'種別'}")
    print("-" * 65)
    for r in sorted(leave_requests, key=lambda x: x.get('serial', '0'), reverse=True):
        types = '・'.join(r.get('leave_types', [])) or '—'
        print(f"  {r['serial']:>5} [{r['status'][:4]:^4}] {r['applicant']:<12} {r['date']:<12} {types}")

finally:
    if driver:
        try: driver.quit()
        except: pass
