"""
サイボウズOffice 総合チェックスクリプト (MS365 SSO対応)
usage: python cybozu_schedule.py [days=7]
"""
import sys, io, os, time, ctypes, ctypes.wintypes, pickle, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

from datetime import date, timedelta
from pathlib import Path

EDGE_DRIVER    = os.path.join(os.environ["TEMP"], "edgedriver_mf", "msedgedriver.exe")
COOKIE_FILE    = os.path.join(os.environ["TEMP"], "cybozu_cookies.pkl")  # 旧方式（互換性のため残置）
EDGE_PROFILE   = os.path.join(os.environ["USERPROFILE"], ".config", "cybozu_profile")
BASE           = "https://YOUR_COMPANY.cybozu.com"
HEADLESS       = "--headless" in sys.argv or "--bg" in sys.argv
NOTIFY         = "--notify" in sys.argv   # 未提出者へメッセージ送信
FULL           = "--full" in sys.argv     # ファイル管理・カスタムアプリも含むフルスキャン
_day_args      = [a for a in sys.argv[1:] if not a.startswith('-')]
DAYS           = int(_day_args[0]) if _day_args else 7

# チェック対象ユーザー（名前, UID）— UIDが判明している場合は直接指定
WATCH_USERS  = [("麻野", "349")]

class CREDENTIAL(ctypes.Structure):
    _fields_ = [
        ('Flags', ctypes.wintypes.DWORD), ('Type', ctypes.wintypes.DWORD),
        ('TargetName', ctypes.wintypes.LPWSTR), ('Comment', ctypes.wintypes.LPWSTR),
        ('LastWritten', ctypes.wintypes.FILETIME), ('CredentialBlobSize', ctypes.wintypes.DWORD),
        ('CredentialBlob', ctypes.POINTER(ctypes.c_byte)), ('Persist', ctypes.wintypes.DWORD),
        ('AttributeCount', ctypes.wintypes.DWORD), ('Attributes', ctypes.c_void_p),
        ('TargetAlias', ctypes.wintypes.LPWSTR), ('UserName', ctypes.wintypes.LPWSTR),
    ]

def read_cred(target):
    advapi32 = ctypes.windll.advapi32
    for t in (1, 2):
        cp = ctypes.POINTER(CREDENTIAL)()
        if advapi32.CredReadW(target, t, 0, ctypes.byref(cp)):
            c = cp.contents
            pw = bytes(c.CredentialBlob[:c.CredentialBlobSize]).decode('utf-16-le') if c.CredentialBlobSize > 0 else ''
            advapi32.CredFree(cp)
            return c.UserName or '', pw
    return '', ''

def make_driver(headless=False):
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.edge.options import Options
    Path(EDGE_PROFILE).mkdir(parents=True, exist_ok=True)
    opts = Options()
    # ヘッドレス・対話ともに同じ永続プロファイルを使用（OAuthトークンで自動セッション維持）
    opts.add_argument(f"--user-data-dir={EDGE_PROFILE}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    if headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1280,900")
        opts.add_argument("--disable-extensions")
    else:
        opts.add_argument("--start-maximized")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    svc = Service(executable_path=EDGE_DRIVER)
    return webdriver.Edge(service=svc, options=opts)

def is_logged_in(driver):
    url = driver.current_url
    return BASE in url and '/login' not in url

def login_microsoft_sso(driver, email, password):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    wait = WebDriverWait(driver, 20)

    try:
        print("  Microsoftメールアドレス入力中...")
        email_input = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "input[type='email'],input[name='loginfmt'],#i0116")))
        email_input.clear()
        email_input.send_keys(email)
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR,
            "input[type='submit'][value='Next'],input#idSIButton9,button#idSIButton9").click()
        time.sleep(3)
    except Exception as e:
        print(f"  メール入力エラー: {e}")

    try:
        print("  パスワード入力中...")
        pw_input = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "input[type='password'],input[name='passwd'],#i0118")))
        pw_input.clear()
        pw_input.send_keys(password)
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR,
            "input[type='submit'],input#idSIButton9,button#idSIButton9").click()
        time.sleep(3)
    except Exception as e:
        print(f"  パスワード入力エラー: {e}")

    print("  認証フロー処理中...")
    for _ in range(60):
        url = driver.current_url
        if BASE in url:
            break
        # KMSI「サインインを保持しますか？」→「はい」のみクリック（「次へ」は誤検知のためスキップ）
        try:
            btn = driver.find_element(By.ID, "idSIButton9")
            val = btn.get_attribute('value') or btn.text or ''
            # 「次へ」「Next」はメール入力ステップのボタンなのでKMSIではない
            if val.strip() in ('次へ', 'Next', ''):
                pass
            elif any(w in val for w in ('はい', 'Yes', 'サインイン')):
                print(f"  KMSI画面検出 ('{val}') → クリック")
                btn.click()
                time.sleep(2)
                continue
        except:
            pass
        if 'login.microsoftonline.com' in url or 'microsoft.com' in url:
            if HEADLESS:
                # ヘッドレス時はMFA待ちできないため終了
                print("  ヘッドレスモードでMFA要求 → ブラウザあり(対話)モードで先にログインしてください")
                return False
            print(f"\n  MFA等の追加認証が必要です")
            print("  ブラウザで認証を完了してEnterを押してください: ", end='', flush=True)
            input()
            time.sleep(2)
            continue
        time.sleep(1)

    handle_service_selection(driver)
    return is_logged_in(driver)

def handle_service_selection(driver):
    from selenium.webdriver.common.by import By
    url = driver.current_url
    if BASE not in url or '/o/' in url:
        return
    print(f"  サービス選択画面を検出")
    for sel in ["a[href*='/o/']", "a[href*='cybozu.com/o']"]:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                href = el.get_attribute('href') or ''
                if '/o/' in href and 'cybozu' in href:
                    print(f"  サイボウズOfficeリンク → クリック")
                    el.click()
                    time.sleep(3)
                    return
        except:
            pass

def login(driver, email, password):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    driver.get(BASE + "/")
    time.sleep(4)
    url = driver.current_url
    print(f"  アクセス後URL: {url}")

    if 'login.microsoftonline.com' in url or 'microsoft.com/wsfed' in url:
        print("  Microsoft SSO フロー検出")
        return login_microsoft_sso(driver, email, password)

    handle_service_selection(driver)
    if is_logged_in(driver):
        return True

    try:
        wait = WebDriverWait(driver, 10)
        inp = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "input[name='_Account'],input[type='email'],#username")))
        inp.clear(); inp.send_keys(email)
        pw = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        pw.clear(); pw.send_keys(password)
        from selenium.webdriver.common.keys import Keys
        pw.send_keys(Keys.RETURN)
        time.sleep(4)
        return is_logged_in(driver)
    except Exception as e:
        print(f"  ログインフォームエラー: {e}")
        return False

def try_cookie_login(driver):
    if not os.path.exists(COOKIE_FILE):
        return False
    try:
        driver.get(BASE + "/")
        time.sleep(2)
        for ck in pickle.load(open(COOKIE_FILE, "rb")):
            try: driver.add_cookie(ck)
            except: pass
        driver.refresh()
        time.sleep(3)
        return is_logged_in(driver)
    except:
        return False

# ──────────────────────────────────────────────
# スケジュール取得
# ──────────────────────────────────────────────

def get_schedule_links(driver):
    """ScheduleIndexページからUID・GIDを取得する"""
    from selenium.webdriver.common.by import By
    driver.get(f"{BASE}/o/ag.cgi?page=ScheduleIndex")
    time.sleep(3)
    links = driver.find_elements(By.TAG_NAME, "a")
    uid = gid = None
    for link in links:
        href = link.get_attribute("href") or ""
        if "ag.cgi" not in href:
            continue
        if 'page=ScheduleUserDay' in href:
            m_uid = re.search(r'UID=(\d+)', href)
            m_gid = re.search(r'GID=(\d+)', href)
            if m_uid:
                uid = m_uid.group(1)
                gid = m_gid.group(1) if m_gid else '0'
                break
    if not uid:
        src = driver.page_source
        uids = re.findall(r'UID=(\d+)', src)
        if uids: uid = uids[0]
    return uid, gid

def fetch_day_events(driver, uid, gid, target_date):
    from selenium.webdriver.common.by import By
    cy_date = f"da.{target_date.year}.{target_date.month}.{target_date.day}"
    url = f"{BASE}/o/ag.cgi?page=ScheduleUserDay&UID={uid}&GID={gid}&Date={cy_date}"
    driver.get(url)
    time.sleep(2)
    events = []
    seen = set()
    try:
        for link in driver.find_elements(By.TAG_NAME, "a"):
            href = link.get_attribute("href") or ""
            txt  = link.text.strip()
            if not txt or len(txt) <= 1:
                continue
            if not any(k in href for k in ["ScheduleDetail", "ScheduleView", "EventDetail"]):
                continue
            if txt in seen:
                continue
            seen.add(txt)
            time_str = ''
            try:
                parent_txt = link.find_element(By.XPATH, "..").text
                m = re.search(r'\d{1,2}:\d{2}', parent_txt)
                if m: time_str = m.group()
            except:
                pass
            events.append({'date': target_date.strftime('%Y-%m-%d'), 'title': txt[:60], 'time': time_str})
    except:
        pass
    return events

def get_schedule(driver, uid, gid, label="自分", days=None):
    if days is None:
        days = DAYS
    today = date.today()
    end_date = today + timedelta(days=days)
    all_events = []
    print(f"\n[{label}] 予定取得中: {today} ～ {end_date}")
    current = today
    while current <= end_date:
        evs = fetch_day_events(driver, uid, gid, current)
        all_events.extend(evs)
        current += timedelta(days=1)

    seen = set()
    unique = []
    for ev in all_events:
        k = (ev['date'], ev['title'])
        if k not in seen:
            seen.add(k); unique.append(ev)

    def sort_key(ev):
        t = ev['time']
        if t:
            try:
                h, m = t.split(':')
                return (ev['date'], int(h)*60+int(m))
            except: pass
        return (ev['date'], -1)

    return sorted(unique, key=sort_key)

# ──────────────────────────────────────────────
# ユーザーUID検索
# ──────────────────────────────────────────────

def find_user_uid(driver, search_name, own_gid):
    """名前でユーザーを検索してUIDを返す"""
    from selenium.webdriver.common.by import By

    # 方法1: グループスケジュール画面からユーザー名＋UIDを抽出
    try:
        today = date.today()
        cy_date = f"da.{today.year}.{today.month}.{today.day}"
        driver.get(f"{BASE}/o/ag.cgi?page=ScheduleGroupDay&GID={own_gid}&Date={cy_date}")
        time.sleep(3)
        src = driver.page_source
        # 「麻野」の前後にある UID= を探す
        # パターン: UID=XXX ... 麻野 or 麻野 ... UID=XXX
        for link in driver.find_elements(By.TAG_NAME, "a"):
            href = link.get_attribute("href") or ""
            txt  = link.text.strip()
            if search_name in txt and 'UID=' in href:
                m = re.search(r'UID=(\d+)', href)
                if m:
                    uid = m.group(1)
                    print(f"  {search_name} UID発見: {uid} (グループスケジュール)")
                    return uid
    except Exception as e:
        print(f"  グループスケジュール検索エラー: {e}")

    # 方法2: ユーザー一覧ページ
    for page in ["AddressBook", "UserList", "UserSearch"]:
        try:
            driver.get(f"{BASE}/o/ag.cgi?page={page}")
            time.sleep(2)
            for link in driver.find_elements(By.TAG_NAME, "a"):
                href = link.get_attribute("href") or ""
                txt  = link.text.strip()
                if search_name in txt and 'UID=' in href:
                    m = re.search(r'UID=(\d+)', href)
                    if m:
                        uid = m.group(1)
                        print(f"  {search_name} UID発見: {uid} ({page})")
                        return uid
        except:
            continue

    # 方法3: ページソース全体から名前近傍のUIDを探す
    try:
        driver.get(f"{BASE}/o/ag.cgi?page=ScheduleIndex")
        time.sleep(2)
        src = driver.page_source
        idx = src.find(search_name)
        if idx >= 0:
            fragment = src[max(0, idx-200):idx+200]
            uids = re.findall(r'UID=(\d+)', fragment)
            if uids:
                uid = uids[0]
                print(f"  {search_name} UID発見: {uid} (ページソース近傍)")
                return uid
    except:
        pass

    print(f"  {search_name} のUIDが見つかりませんでした")
    return None

# ──────────────────────────────────────────────
# ナビゲーションからモジュールURLを発見
# ──────────────────────────────────────────────

def discover_module_urls(driver):
    """Cybozuトップページのナビゲーションから各モジュールの実URLを収集する"""
    from selenium.webdriver.common.by import By

    driver.get(f"{BASE}/o/ag.cgi")
    time.sleep(3)

    # キーワードとモジュール名のマッピング
    keyword_map = {
        'workflow':    ('ワークフロー', []),
        'Workflow':    ('ワークフロー', []),
        'bulletin':    ('掲示板',       []),
        'Bulletin':    ('掲示板',       []),
        'cabinet':     ('ファイル管理', []),
        'Cabinet':     ('ファイル管理', []),
        'report':      ('レポート',     []),
        'Report':      ('レポート',     []),
        'memo':        ('メモ',         []),
        'Memo':        ('メモ',         []),
        'ToDo':        ('ToDo',         []),
        'todo':        ('ToDo',         []),
    }

    module_urls = {}  # モジュール名 → URL

    all_links = driver.find_elements(By.TAG_NAME, "a")
    for link in all_links:
        href = link.get_attribute("href") or ""
        txt  = link.text.strip()
        if not href or BASE not in href:
            continue
        for kw, (mod_name, _) in keyword_map.items():
            if kw in href and mod_name not in module_urls:
                module_urls[mod_name] = href
                print(f"  モジュール発見: {mod_name} → {href.replace(BASE,'')}")
                break

    # ページに「ワークフロー」等のテキストリンクがあれば追加収集
    nav_keywords = {
        'ワークフロー': 'ワークフロー',
        '掲示板':       '掲示板',
        'ファイル管理': 'ファイル管理',
        'レポート':     'レポート',
        'ToDo':         'ToDo',
    }
    for link in all_links:
        href = link.get_attribute("href") or ""
        txt  = link.text.strip()
        if txt in nav_keywords and nav_keywords[txt] not in module_urls and BASE in href:
            module_urls[nav_keywords[txt]] = href
            print(f"  モジュール発見(テキスト): {txt} → {href.replace(BASE,'')}")

    return module_urls

def _is_valid_page(src):
    ng = ["このページは表示できません", "不正なURL", "Not Found", "404"]
    return not any(x in src for x in ng)

def _collect_links(driver, detail_keywords, limit=15):
    """ページ内から詳細リンクを収集する汎用関数"""
    from selenium.webdriver.common.by import By
    items = []
    seen  = set()
    for link in driver.find_elements(By.TAG_NAME, "a"):
        href = link.get_attribute("href") or ""
        txt  = link.text.strip()
        if not txt or len(txt) <= 1 or txt in seen:
            continue
        if any(k in href for k in detail_keywords):
            seen.add(txt)
            items.append(txt[:70])
            if len(items) >= limit:
                break
    return items

# ──────────────────────────────────────────────
# ワークフロー確認
# ──────────────────────────────────────────────

def get_workflow_items(driver, module_urls):
    from selenium.webdriver.common.by import By
    items = []

    # 発見したURLを優先、なければ候補パターンを試す
    wf_url = module_urls.get('ワークフロー')
    candidates = [wf_url] if wf_url else []
    candidates += [
        f"{BASE}/o/ag.cgi?page=WorkflowIndex",
        f"{BASE}/o/ag.cgi?page=WorkflowMyTask",
        f"{BASE}/o/ag.cgi?page=WorkflowTop",
    ]

    for url in candidates:
        if not url:
            continue
        try:
            driver.get(url)
            time.sleep(2)
            src = driver.page_source
            if not _is_valid_page(src):
                print(f"  ワークフロー: ページ無効 ({url.replace(BASE,'')})")
                continue
            print(f"  ワークフロー: アクセス成功 ({url.replace(BASE,'')})")
            found = _collect_links(driver,
                ["WorkflowDetail", "WorkflowView", "workflow_id=", "wid=", "Request"])
            items.extend(found)
            # サブページのリンクも辿る
            for link in driver.find_elements(By.TAG_NAME, "a"):
                href = link.get_attribute("href") or ""
                txt  = link.text.strip()
                if BASE not in href:
                    continue
                if any(k in href for k in ["MyTask", "Pending", "Request", "受信", "申請"]):
                    try:
                        driver.get(href)
                        time.sleep(2)
                        if _is_valid_page(driver.page_source):
                            sub = _collect_links(driver,
                                ["WorkflowDetail", "WorkflowView", "workflow_id=", "wid="])
                            items.extend(f"[{txt}] {x}" for x in sub)
                    except:
                        pass
            break
        except Exception as e:
            print(f"  ワークフローエラー: {e}")
            continue

    # 重複除去
    seen = set(); result = []
    for it in items:
        if it not in seen:
            seen.add(it); result.append(it)
    return result

# ──────────────────────────────────────────────
# 掲示板確認
# ──────────────────────────────────────────────

def get_bulletin_items(driver, module_urls):
    bb_url = module_urls.get('掲示板')
    candidates = [bb_url] if bb_url else []
    candidates += [
        f"{BASE}/o/ag.cgi?page=BulletinIndex",
        f"{BASE}/o/ag.cgi?page=BulletinTop",
    ]
    for url in candidates:
        if not url:
            continue
        try:
            driver.get(url)
            time.sleep(2)
            src = driver.page_source
            if not _is_valid_page(src):
                print(f"  掲示板: ページ無効 ({url.replace(BASE,'')})")
                continue
            print(f"  掲示板: アクセス成功 ({url.replace(BASE,'')})")
            return _collect_links(driver, ["BulletinDetail", "BulletinView", "bid=", "topic_id="])
        except Exception as e:
            print(f"  掲示板エラー: {e}")
    return []

# ──────────────────────────────────────────────
# ファイル管理確認
# ──────────────────────────────────────────────

_UI_SKIP = {
    "ファイルを追加する", "フォルダを追加する", "フォルダ/アプリを順番変更する",
    "詳細検索", "アプリを追加する", "カスタムアプリ（ルートフォルダ）",
    "ファイル管理", "カスタムアプリ", "採用管理", "トップページへ",
}

def _get_content_links(driver, detail_keys, skip_keys=None):
    """ページ内のフォルダ・ファイル・アプリリンクを収集（UIボタン除外）"""
    from selenium.webdriver.common.by import By
    if skip_keys is None:
        skip_keys = _UI_SKIP
    result = []
    seen_href = set()
    for link in driver.find_elements(By.TAG_NAME, "a"):
        href = link.get_attribute("href") or ""
        txt  = link.text.strip()
        if not txt or len(txt) <= 1:
            continue
        if txt in skip_keys:
            continue
        if href in seen_href:
            continue
        if BASE not in href:
            continue
        if any(k in href for k in detail_keys):
            seen_href.add(href)
            result.append((txt[:60], href))
    return result

def get_cabinet_items(driver, module_urls):
    """ファイル管理: フォルダ一覧 → 各フォルダ内の実ファイルのみ取得"""
    from selenium.webdriver.common.by import By

    print(f"  ファイル管理: トップ取得中...")
    driver.get(f"{BASE}/o/ag.cgi?page=FileIndex")
    time.sleep(2)
    if not _is_valid_page(driver.page_source):
        print("  ファイル管理: ページ無効")
        return {}

    # トップのフォルダ一覧
    # Cybozuのファイル管理フォルダURLは page=FileIndex& (追加パラメータあり)
    all_links = driver.find_elements(By.TAG_NAME, "a")
    top_folder_names = set()
    folder_links = []
    seen_href = set()
    for link in all_links:
        href = link.get_attribute("href") or ""
        txt  = link.text.strip()
        if not txt or txt in _UI_SKIP or BASE not in href:
            continue
        if href in seen_href:
            continue
        # FileIndex に追加パラメータがある = フォルダ（page=FileIndex& または FileIndex?xxx=yyy 等）
        if "FileIndex" in href and "page=FileIndex&" in href:
            seen_href.add(href)
            top_folder_names.add(txt)
            folder_links.append((txt, href))

    # page=FileIndex& がない場合は & 付きの FileIndex を広めに拾う
    if not folder_links:
        seen_href.clear()
        for link in all_links:
            href = link.get_attribute("href") or ""
            txt  = link.text.strip()
            if not txt or txt in _UI_SKIP or BASE not in href:
                continue
            if href in seen_href:
                continue
            # FileIndex かつ何らかのパラメータが付いているリンク（純粋なトップを除外）
            pure_top = f"{BASE}/o/ag.cgi?page=FileIndex"
            if "FileIndex" in href and href.rstrip("/") != pure_top:
                seen_href.add(href)
                top_folder_names.add(txt)
                folder_links.append((txt, href))

    print(f"  ファイル管理: フォルダ {len(folder_links)}件検出")
    structure = {}

    for folder_name, folder_href in folder_links[:15]:
        try:
            driver.get(folder_href)
            time.sleep(1.5)
            if not _is_valid_page(driver.page_source):
                continue

            files = []
            sub_folders = []
            for link in driver.find_elements(By.TAG_NAME, "a"):
                href = link.get_attribute("href") or ""
                txt  = link.text.strip()
                if not txt or txt in _UI_SKIP or BASE not in href:
                    continue
                # 実ファイル（FileDetail/FileView）
                if "FileDetail" in href or "FileView" in href:
                    files.append(txt[:60])
                # サブフォルダ（FileIndex系でトップフォルダ名と異なるもの）
                elif "FileIndex" in href and txt not in top_folder_names and "page=FileIndex" in href:
                    sub_folders.append(f"[フォルダ] {txt}")

            all_items = sub_folders + files
            structure[folder_name] = all_items
            print(f"    [{folder_name}] ファイル{len(files)}件 / サブフォルダ{len(sub_folders)}件")
        except Exception as e:
            print(f"    [{folder_name}] エラー: {e}")
            structure[folder_name] = []
    return structure

def get_weekly_report_status(driver):
    """週報・稼働時間報告書の提出状況を確認する"""
    from selenium.webdriver.common.by import By

    today = date.today()
    week_start = today - timedelta(days=today.weekday())

    print(f"  週報状況確認中 (対象週: {week_start})...")

    weekly_app_url = None
    # 幅広く名前をマッチ（アプリ名が何であれ稼働/週報/報告を含めばヒット）
    search_names = ["週報", "稼働時間", "稼働報告", "勤怠報告", "月次報告", "週次報告", "報告書"]

    def _links_from_page(driver):
        result = []
        for link in driver.find_elements(By.TAG_NAME, "a"):
            href = link.get_attribute("href") or ""
            txt  = link.text.strip()
            if txt and BASE in href:
                result.append((txt, href))
        return result

    def _find_report_app(links):
        for txt, href in links:
            if any(n in txt for n in search_names):
                return txt, href
        return None, None

    # ── ステップ1: DBIndex トップで直接探す ──
    driver.get(f"{BASE}/o/ag.cgi?page=DBIndex")
    time.sleep(2)
    if not _is_valid_page(driver.page_source):
        print("  週報: DBIndexにアクセスできません")
        return None

    top_links = _links_from_page(driver)
    found_name, weekly_app_url = _find_report_app(top_links)
    if weekly_app_url:
        print(f"  週報アプリ発見(トップ): {found_name}")

    # ── ステップ2: 各フォルダを1階層掘る ──
    if not weekly_app_url:
        folder_links = [(t, h) for t, h in top_links
                        if "DBIndex" in h and ("did=" in h or "page=DBIndex&" in h)
                        and t not in _UI_SKIP]
        print(f"  週報: フォルダ {len(folder_links)}件を検索中...")
        for folder_name, folder_href in folder_links[:25]:
            try:
                driver.get(folder_href)
                time.sleep(1.5)
                if not _is_valid_page(driver.page_source):
                    continue
                sub_links = _links_from_page(driver)
                found_name, weekly_app_url = _find_report_app(sub_links)
                if weekly_app_url:
                    print(f"  週報アプリ発見: [{folder_name}] {found_name}")
                    break
                # 何が入っているかデバッグ出力（最初の5件）
                app_names = [t for t, h in sub_links
                             if ("DBView" in h or "aid=" in h) and t not in _UI_SKIP]
                if app_names:
                    print(f"    [{folder_name}] アプリ: {', '.join(app_names[:5])}")
            except:
                continue

    if not weekly_app_url:
        print("  週報アプリが見つかりませんでした（アプリ名を確認してください）")
        return None

    # アプリにアクセスしてレコード取得
    driver.get(weekly_app_url)
    time.sleep(2)
    if not _is_valid_page(driver.page_source):
        print("  週報: アプリページ無効")
        return None

    src = driver.page_source

    # 件数を取得
    cnt_m = re.search(r'(\d+)\s*件', src)
    total = cnt_m.group(1) if cnt_m else "?"

    # レコード一覧のテーブル行から提出者情報を取得
    # Cybozuカスタムアプリは「閲覧する」リンク + 隣接セルに名前/日付
    submitted = []
    seen_rid = set()

    # まずテーブル行から取得を試みる
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "table tr")
        for row in rows:
            row_text = row.text.strip()
            if not row_text or len(row_text) <= 2:
                continue
            # rid= リンクがある行のみ対象
            row_links = row.find_elements(By.TAG_NAME, "a")
            rid = None
            for rl in row_links:
                href = rl.get_attribute("href") or ""
                rid_m = re.search(r'(?:rid|record_id)=(\d+)', href)
                if rid_m:
                    rid = rid_m.group(1)
                    break
            if not rid or rid in seen_rid:
                continue
            seen_rid.add(rid)
            # 行テキストから「閲覧する」「編集する」等の操作ボタン文言を除去
            clean = re.sub(r'閲覧する|編集する|複写する|削除する', '', row_text).strip()
            if clean:
                submitted.append(clean[:80])
            if len(submitted) >= 30:
                break
    except Exception as e:
        print(f"  週報: テーブル行取得エラー: {e}")

    # テーブルで取得できなかった場合、リンクのみで補完
    if not submitted:
        for link in driver.find_elements(By.TAG_NAME, "a"):
            href = link.get_attribute("href") or ""
            txt  = link.text.strip()
            rid_m = re.search(r'(?:rid|record_id)=(\d+)', href)
            if not rid_m:
                continue
            rid = rid_m.group(1)
            if rid in seen_rid:
                continue
            seen_rid.add(rid)
            submitted.append(f"[rid={rid}] {txt}")
            if len(submitted) >= 30:
                break

    return {
        "url": weekly_app_url,
        "week_start": week_start,
        "submitted": submitted,
        "total": total,
    }


def get_all_employees(driver, gid):
    """複数の方法で全社員のUID・名前を取得する"""
    from selenium.webdriver.common.by import By

    employees = {}
    print("  社員一覧: 取得中...")

    today = date.today()
    cy_date = f"da.{today.year}.{today.month}.{today.day}"

    _NAME_SKIP_PAT = re.compile(r'[【】（）()「」\[\]]|^\d+$|^[a-zA-Z\s]{15,}$')

    def _extract_from_page(driver):
        """グループスケジュールページからユーザー名リンク（ScheduleUser*）のみ抽出"""
        result = {}
        # ScheduleUserDay / ScheduleUserWeek へのリンク = ユーザー名リンク
        user_link_pat = re.compile(r'ScheduleUser(?:Day|Week|Month)')
        for link in driver.find_elements(By.TAG_NAME, "a"):
            href = link.get_attribute("href") or ""
            txt  = link.text.strip()
            if not txt or len(txt) < 2 or txt in _UI_SKIP:
                continue
            if _NAME_SKIP_PAT.search(txt):
                continue  # 【会議】等のイベントタイトルを除外
            if user_link_pat.search(href) and "UID=" in href:
                m = re.search(r'UID=(\d+)', href)
                if m and txt not in result:
                    result[txt] = m.group(1)
        return result

    def _extract_from_source(src):
        """ページソース全体から UID=xxx の前後にある日本語名前を広く抽出"""
        result = {}
        for m in re.finditer(r'UID=(\d+)', src):
            uid = m.group(1)
            fragment = src[max(0, m.start()-300):m.end()+300]
            # >名前< の形式（2〜8文字の日本語テキスト）
            for nm in re.finditer(r'>([^<>\s]{2,10})\s*([^<>\s]{1,8})?<', fragment):
                cand = (nm.group(1) + (' ' + nm.group(2) if nm.group(2) else '')).strip()
                if re.search(r'[一-鿿぀-ヿ]', cand) and cand not in _UI_SKIP:
                    if cand not in result:
                        result[cand] = uid
                    break
        return result

    def _try_group(driver, target_gid, label=""):
        """指定GIDのScheduleGroupDayからユーザーリストを取得"""
        result = {}
        try:
            driver.get(f"{BASE}/o/ag.cgi?page=ScheduleGroupDay&GID={target_gid}&Date={cy_date}")
            time.sleep(3)
            if _is_valid_page(driver.page_source):
                found = _extract_from_page(driver)
                if found:
                    print(f"    ScheduleGroupDay({label or target_gid}): {len(found)}名")
                    result.update(found)
        except Exception as e:
            print(f"    ScheduleGroupDay({target_gid}) エラー: {e}")
        return result

    # スケジュールナビゲーション系の除外ワード
    _SCHED_NAV = {'月予定', '週予定', '当日予定', '日予定', '年予定',
                  '今日', '前月', '次月', '前週', '次週', '今週', '今月',
                  '当日', '翌日', '前日', 'スケジュール'}

    # 日本語名前らしいか判定（漢字/ひらがな/カタカナを含み2〜10文字）
    _name_re = re.compile(r'[一-鿿々ぁ-んァ-ン]')
    def _is_name_like(txt):
        cleaned = txt.replace('　', ' ').strip()
        return _name_re.search(cleaned) and 2 <= len(cleaned.replace(' ', '')) <= 10

    # ── 方法1: UserListIndex&GID=0（全社員一覧） ──
    try:
        driver.get(f"{BASE}/o/ag.cgi?page=UserListIndex&GID=0")
        time.sleep(3)
        if _is_valid_page(driver.page_source):
            # 各UID に紐づくリンクテキストを全収集 → 名前らしいものを選択
            uid_texts = {}  # uid -> [txt, ...]
            for link in driver.find_elements(By.TAG_NAME, "a"):
                href = link.get_attribute("href") or ""
                txt  = link.text.strip()
                if not txt or txt in _UI_SKIP or txt in _SCHED_NAV:
                    continue
                if "UID=" in href:
                    m = re.search(r'UID=(\d+)', href)
                    if m:
                        uid_texts.setdefault(m.group(1), []).append(txt)
            for uid, texts in uid_texts.items():
                for txt in texts:
                    if _is_name_like(txt) and txt not in employees:
                        employees[txt] = uid
                        break
            print(f"    UserListIndex: {len(employees)}名")
    except Exception as e:
        print(f"    UserListIndex エラー: {e}")

    # ── 方法2（補完）: ScheduleGroupDay でUIDを補完 ──
    if len(employees) < 3:
        employees.update(_try_group(driver, gid))

    # 重複チェック: UID重複を持つ名前エントリを除去（短い名前より長い名前を優先）
    uid_to_names = {}
    for name, uid in employees.items():
        if uid not in uid_to_names:
            uid_to_names[uid] = name
        elif len(name) > len(uid_to_names[uid]):
            uid_to_names[uid] = name
    employees = {name: uid for uid, name in uid_to_names.items()}

    print(f"  社員一覧: 合計{len(employees)}名取得")
    return employees


def extract_submitter_names(submitted_records):
    """週報レコードの行テキストから提出者名セットを抽出"""
    names = set()
    for record_text in submitted_records:
        # "登録者 金井 亮輔" パターンで抽出
        m = re.search(r'登録者\s+(.+?)(?:\n|更新|$)', record_text)
        if m:
            name = re.split(r'\s{2,}|更新', m.group(1))[0].strip()
            if name and len(name) >= 2:
                names.add(name)
    return names


def send_cybozu_message(driver, to_uid_list, subject, body):
    """サイボウズのメッセージ送信（MessageSendページ使用）"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    if not to_uid_list:
        print("  送信対象なし")
        return True

    # UIDs をクエリパラメータで渡してフォームを開く（Cybozu Office 仕様）
    uid_param = "&".join(f"UID={uid}" for uid in to_uid_list)
    driver.get(f"{BASE}/o/ag.cgi?page=MessageSend&{uid_param}")
    time.sleep(2)
    if not _is_valid_page(driver.page_source):
        driver.get(f"{BASE}/o/ag.cgi?page=MessageSend")
        time.sleep(2)
    if not _is_valid_page(driver.page_source):
        print("  メッセージ送信: ページ無効")
        return False

    try:
        wait = WebDriverWait(driver, 10)

        # 件名入力
        for sel in ["input[name='Subject']", "input[name='title']", "input[name='subject']"]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                el.clear(); el.send_keys(subject); break
            except: pass

        # 本文入力
        for sel in ["textarea[name='Body']", "textarea[name='body']", "textarea[name='message']"]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                el.clear(); el.send_keys(body); break
            except: pass

        # 送信ボタンクリック
        for sel in ["input[value='送信']", "input[value*='送']", "input[type='submit']",
                    "button[type='submit']"]:
            try:
                btn = driver.find_element(By.CSS_SELECTOR, sel)
                btn.click()
                time.sleep(3)
                print(f"  メッセージ送信完了: {len(to_uid_list)}名")
                return True
            except: pass

        print(f"  フォーム入力済み・送信ボタン未検出（手動送信が必要な場合あり）")
        return False

    except Exception as e:
        print(f"  メッセージ送信エラー: {e}")
        return False


def get_custom_app_items(driver):
    """カスタムアプリ: フォルダ→アプリ→レコード件数まで取得"""
    from selenium.webdriver.common.by import By

    print(f"  カスタムアプリ: トップ取得中...")
    driver.get(f"{BASE}/o/ag.cgi?page=DBIndex")
    time.sleep(2)
    if not _is_valid_page(driver.page_source):
        print("  カスタムアプリ: ページ無効")
        return {}

    # トップのフォルダ・アプリ一覧
    top_links = _get_content_links(driver, ["DBIndex", "DBView", "did=", "aid="])
    print(f"  カスタムアプリ: {len(top_links)}件検出")

    structure = {}
    for item_name, item_href in top_links[:15]:
        try:
            driver.get(item_href)
            time.sleep(1.5)
            if not _is_valid_page(driver.page_source):
                continue
            src = driver.page_source

            # フォルダの場合: 中のアプリ一覧を収集
            if "DBIndex" in item_href:
                sub_links = _get_content_links(driver, ["DBIndex", "DBView", "did=", "aid="])
                apps = []
                for app_name, app_href in sub_links[:10]:
                    try:
                        driver.get(app_href)
                        time.sleep(1.5)
                        if not _is_valid_page(driver.page_source):
                            continue
                        # レコード一覧の件数を探す
                        cnt_match = re.search(r'(\d+)\s*件', driver.page_source)
                        count = cnt_match.group(0) if cnt_match else ""
                        # 最新レコードのタイトルリンクを取得
                        records = _get_content_links(driver,
                            ["DBDetail", "DBView", "record_id=", "rid="])
                        rec_names = [n for n, _ in records[:3]]
                        apps.append((app_name, count, rec_names))
                        driver.back()
                        time.sleep(1)
                    except:
                        apps.append((app_name, "", []))
                structure[item_name] = apps
                print(f"    [{item_name}] アプリ{len(apps)}件")
            # アプリの場合: レコード一覧を直接収集
            else:
                cnt_match = re.search(r'(\d+)\s*件', src)
                count = cnt_match.group(0) if cnt_match else ""
                records = _get_content_links(driver,
                    ["DBDetail", "DBView", "record_id=", "rid="])
                rec_names = [n for n, _ in records[:5]]
                structure[item_name] = [(item_name, count, rec_names)]
                print(f"    [{item_name}] {count}")
        except Exception as e:
            print(f"    [{item_name}] エラー: {e}")
            structure[item_name] = []
    return structure

# ──────────────────────────────────────────────
# フォーマット・出力
# ──────────────────────────────────────────────

def format_schedule(events, label=""):
    if not events:
        return "  予定はありません。"
    lines = []
    current_date = None
    weekdays = ['月', '火', '水', '木', '金', '土', '日']
    for ev in events:
        d = date.fromisoformat(ev['date'])
        if d != current_date:
            wd = weekdays[d.weekday()]
            lines.append(f"\n  【{d.strftime('%m/%d')}（{wd}）】")
            current_date = d
        pfx = f"    {ev['time']} " if ev['time'] else "    "
        lines.append(f"{pfx}{ev['title']}")
    return '\n'.join(lines)

# ──────────────────────────────────────────────
# main
# ──────────────────────────────────────────────

def main():
    email, password = read_cred('Cybozu_Login')
    if not email:
        print("エラー: 資格情報 'Cybozu_Login' が見つかりません")
        sys.exit(1)

    print(f"=== サイボウズOffice 総合チェック ===")
    print(f"ユーザー: {email}  期間: {DAYS}日間")

    driver = make_driver(headless=HEADLESS)
    if HEADLESS:
        print("（ヘッドレスモード：ブラウザ非表示）")
    try:
        # プロファイルのOAuthトークンで自動ログイン試行、失敗時は資格情報でMS SSO
        print("セッション確認中...")
        driver.get(BASE + "/o/ag.cgi")
        time.sleep(3)
        logged_in = is_logged_in(driver)
        if logged_in:
            print("プロファイルセッション有効!")
        else:
            print("セッション切れ → MS365 SSOでログイン中...")
            logged_in = login(driver, email, password)

        if not logged_in:
            print(f"ログイン失敗: {driver.current_url}")
            driver.save_screenshot(os.path.join(os.environ["TEMP"], "cybozu_fail.png"))
            sys.exit(1)

        print("ログイン成功!")

        # 自分のUID・GID取得
        my_uid, my_gid = get_schedule_links(driver)
        print(f"自分: UID={my_uid}, GID={my_gid}")

        # 自分のスケジュール
        my_events = get_schedule(driver, my_uid, my_gid, "中村")

        # 麻野さんほか対象ユーザーのスケジュール
        other_schedules = {}
        for name, known_uid in WATCH_USERS:
            if known_uid:
                uid = known_uid
                print(f"\n[{name}] UID={uid}（指定値）")
            else:
                uid = find_user_uid(driver, name, my_gid)
            if uid:
                evs = get_schedule(driver, uid, my_gid, name)
                other_schedules[name] = evs
            else:
                other_schedules[name] = None

        # ワークフロー確認（URLは実績値を直接使用）
        print("ワークフロー確認中...")
        wf_items = get_workflow_items(driver, {'ワークフロー': f"{BASE}/o/ag.cgi?page=WorkFlowIndex"})

        # 週報・稼働時間報告書の提出状況確認
        weekly_status = get_weekly_report_status(driver)

        # ファイル管理・カスタムアプリ（--full 時のみ）
        cab_items = {}
        db_items  = {}
        if FULL:
            print("ファイル管理確認中...")
            cab_items = get_cabinet_items(driver, {})
            print("カスタムアプリ確認中...")
            db_items = get_custom_app_items(driver)

        # 全社員リスト取得 & 未提出者特定
        all_employees = {}
        submitted_names = set()
        non_submitters_names = set()
        non_submitters_uids  = []
        if weekly_status:
            all_employees   = get_all_employees(driver, my_gid)
            submitted_names = extract_submitter_names(weekly_status['submitted'])
            if all_employees:
                non_submitters_names = set(all_employees.keys()) - submitted_names
                non_submitters_uids  = [all_employees[n] for n in non_submitters_names
                                        if n in all_employees]
                print(f"  未提出者: {len(non_submitters_names)}名 / 全{len(all_employees)}名")

        # 未提出者へ通知（--notify 指定時のみ）
        if NOTIFY and non_submitters_uids:
            ws_str = weekly_status['week_start'].strftime('%Y/%m/%d')
            subject = f"【週報提出のお願い】{ws_str}週分"
            body = (
                f"お疲れ様です。\n\n"
                f"{ws_str}週分の週報・稼働時間報告書がまだ提出されていません。\n"
                f"お手数ですが、サイボウズOfficeの「全社」→「週報・稼働時間報告書」より\n"
                f"ご提出いただけますようお願いします。\n\n"
                f"※このメッセージは自動送信されました。\n"
                f"　ご不明な点は管理本部（麻野）までお問い合わせください。"
            )
            print("未提出者へ通知中...")
            send_cybozu_message(driver, non_submitters_uids, subject, body)

        # ──── レポート出力 ────
        sep = "=" * 55
        print(f"\n{sep}")
        print(f"サイボウズOffice チェックレポート ({date.today()})")
        print(sep)

        print(f"\n■ 中村さんの予定 ({DAYS}日間)")
        print(format_schedule(my_events))
        print(f"  合計: {len(my_events)}件")

        for name, evs in other_schedules.items():
            print(f"\n■ {name}さんの予定 ({DAYS}日間)")
            if evs is None:
                print(f"  UIDが特定できませんでした")
            else:
                print(format_schedule(evs))
                print(f"  合計: {len(evs)}件")

        print(f"\n■ ワークフロー ({len(wf_items)}件)")
        if wf_items:
            for it in wf_items[:15]:
                print(f"  ・{it}")
        else:
            print("  未処理・申請中なし")

        if FULL:
            print(f"\n■ ファイル管理")
            if cab_items:
                for folder, files in cab_items.items():
                    print(f"  [{folder}] ({len(files)}件)")
                    for f in files[:5]:
                        print(f"    ・{f}")
                    if len(files) > 5:
                        print(f"    ・… 他{len(files)-5}件")
            else:
                print("  取得できませんでした")

            print(f"\n■ カスタムアプリ")
            if db_items:
                for folder, apps in db_items.items():
                    if isinstance(apps, list) and apps and isinstance(apps[0], tuple):
                        print(f"  [{folder}]")
                        for app_name, count, records in apps:
                            rec_str = "、".join(records) if records else ""
                            print(f"    ・{app_name} {count}" + (f" → {rec_str}" if rec_str else ""))
                    else:
                        print(f"  [{folder}]")
            else:
                print("  取得できませんでした")

        print(f"\n■ 週報・稼働時間報告書")
        if weekly_status:
            ws      = weekly_status['week_start']
            submitted = weekly_status['submitted']
            total   = weekly_status['total']
            print(f"  対象週: {ws.strftime('%Y/%m/%d')}（月）〜")
            total_members = len(all_employees) if all_employees else "?"
            print(f"  全社員: {total_members}名  提出済み: {len(submitted_names)}名  "
                  f"未提出: {len(non_submitters_names)}名")

            # 提出済み（名前のみ）
            if submitted_names:
                slist = sorted(submitted_names)
                print(f"  【提出済み】{' / '.join(slist)}")
            else:
                print("  【注意】今週の提出が見当たりません")

            # 未提出者
            if non_submitters_names:
                nlist = sorted(non_submitters_names)
                print(f"\n  【★未提出者★】")
                for n in nlist:
                    print(f"    ・{n}")
                if NOTIFY:
                    print(f"  → 上記 {len(nlist)}名 にサイボウズメッセージを送信しました")
                else:
                    print(f"  ※通知するには --notify フラグを付けて実行してください")
            else:
                if all_employees:
                    print("  全員提出済みです ✓")
        else:
            print("  週報アプリが見つかりませんでした（アプリ名を確認してください）")

        print(f"\n{sep}")

    finally:
        # driver.quit() がヘッドレスでハングする場合があるため、タイムアウト付きで実行
        import threading
        def _quit():
            try: driver.quit()
            except: pass
        t = threading.Thread(target=_quit, daemon=True)
        t.start()
        t.join(timeout=8)

if __name__ == "__main__":
    main()
