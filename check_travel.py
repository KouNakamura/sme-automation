"""
フライト・ホテル予約とサイボウズ予定の整合性チェック
usage: python check_travel.py [--headless]
"""
import sys, io, os, time, re, json, pickle, ctypes, ctypes.wintypes
from datetime import date, datetime, timedelta
# PYTHONIOENCODING=utf-8 で起動するか、-X utf8 フラグを使用すること

EDGE_DRIVER  = os.path.join(os.environ["TEMP"], "edgedriver_mf", "msedgedriver.exe")
COOKIE_FILE  = os.path.join(os.environ["TEMP"], "cybozu_cookies.pkl")
HEADLESS     = "--no-headless" not in sys.argv  # デフォルトはヘッドレス
TMPDIR       = os.environ["TEMP"]
CYBOZU_BASE  = "https://YOUR_COMPANY.cybozu.com"

# フライト前後にサイボウズ予定があれば警告する余裕時間（分）
BUFFER_MIN = 180  # 3時間

# ──────────────────────────────────
# 共通ユーティリティ
# ──────────────────────────────────

class CREDENTIAL(ctypes.Structure):
    _fields_ = [('Flags', ctypes.wintypes.DWORD), ('Type', ctypes.wintypes.DWORD),
                ('TargetName', ctypes.wintypes.LPWSTR), ('Comment', ctypes.wintypes.LPWSTR),
                ('LastWritten', ctypes.wintypes.FILETIME), ('CredentialBlobSize', ctypes.wintypes.DWORD),
                ('CredentialBlob', ctypes.POINTER(ctypes.c_byte)), ('Persist', ctypes.wintypes.DWORD),
                ('AttributeCount', ctypes.wintypes.DWORD), ('Attributes', ctypes.c_void_p),
                ('TargetAlias', ctypes.wintypes.LPWSTR), ('UserName', ctypes.wintypes.LPWSTR)]

def read_cred(target):
    a = ctypes.windll.advapi32
    for t in (1, 2):
        cp = ctypes.POINTER(CREDENTIAL)()
        if a.CredReadW(target, t, 0, ctypes.byref(cp)):
            c = cp.contents
            pw = bytes(c.CredentialBlob[:c.CredentialBlobSize]).decode('utf-16-le') if c.CredentialBlobSize > 0 else ''
            a.CredFree(cp); return c.UserName or '', pw
    return '', ''

def make_driver(headless=None):
    """headless=None → グローバルHEADLESS変数に従う。headless=False → 非ヘッドレス強制"""
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.edge.options import Options
    opts = Options()
    use_headless = HEADLESS if headless is None else headless
    if use_headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1440,900"); opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage"); opts.add_argument("--disable-extensions")
    opts.add_argument("--remote-debugging-port=0")
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--lang=ja")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--no-first-run")
    opts.add_argument("--disable-features=IsolateOrigins,site-per-process,VizDisplayCompositor")
    opts.add_argument("--disable-site-isolation-trials")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--disable-background-timer-throttling")
    opts.add_argument("--disable-renderer-backgrounding")
    opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "autofill.profile_enabled": False,
        "profile.default_content_setting_values.automatic_downloads": 1,
    })
    driver = webdriver.Edge(service=Service(EDGE_DRIVER), options=opts)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": (
            "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"
            "Object.defineProperty(navigator,'plugins',{get:()=>[1,2,3,4,5]});"
            "Object.defineProperty(navigator,'languages',{get:()=>['ja-JP','ja','en-US','en']});"
            "window.chrome={runtime:{}};"
        )
    })
    return driver

def ss(driver, name):
    path = os.path.join(TMPDIR, f"travel_{name}.png")
    driver.save_screenshot(path)

def js_fill_input(driver, selector, value):
    """React対応フォーム入力（send_keysが効かないフォーム用）"""
    return driver.execute_script("""
        var el = document.querySelector(arguments[0]);
        if (!el) return false;
        try {
            var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            setter.call(el, arguments[1]);
        } catch(e) { el.value = arguments[1]; }
        ['input','change','keyup','keydown'].forEach(function(t){
            el.dispatchEvent(new Event(t, {bubbles:true}));
        });
        return true;
    """, selector, value)

# ──────────────────────────────────
# ANA 予約取得
# ──────────────────────────────────

def fetch_ana_bookings(member_no, password):
    """ANAにログインして予約一覧を返す（非ヘッドレスでbot検知回避）"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys

    print("\n【ANA】予約情報取得中...")
    driver = make_driver()
    bookings = []
    try:
        wait = WebDriverWait(driver, 20)

        # ── ANAトップページからログイン（aswbeサービシングポータルはシステム移行期間中に閉鎖されるため使用しない）──
        driver.get("https://www.ana.co.jp/ja/jp/")
        time.sleep(5)
        try:
            ss(driver, "ana_top")
        except Exception:
            pass
        cur = driver.current_url
        print(f"  ANAトップURL: {cur}")

        # ── ログインボタンをクリック ──
        login_clicked = driver.execute_script("""
            for (var el of document.querySelectorAll('a,button')) {
                var t = el.textContent.trim();
                if (t === 'ログイン' || t === 'ログイン／新規入会') {
                    var r = el.getBoundingClientRect();
                    if (r.width > 0 && r.height > 0) { el.click(); return 'clicked:' + t; }
                }
            }
            return 'none';
        """)
        print(f"  ログインボタン: {login_clicked}")
        time.sleep(4)
        try:
            ss(driver, "ana_login_page")
        except Exception:
            pass

        # ログインが必要かチェック（ログインページへリダイレクトされるはず）
        needs_login = 'login' in cur.lower() or 'account' in cur.lower() or \
                      bool(driver.find_elements(By.CSS_SELECTOR, "input[type='password']")) or \
                      'booking-list' not in cur
        print(f"  ログイン要否: {'要' if needs_login else '不要（既にログイン済み）'}")

        # ── ログインフォームへの入力 ──
        if needs_login:
            ss(driver, "ana_login_modal")

        if not needs_login:
            print("  ログイン不要（セッション有効）")
        else:
            # ── 会員番号入力 ──
            member_filled = False
            for sel in ["input[placeholder*='123456']", "input[placeholder*='1234567']",
                        "input[name='loginId']", "input[name='customerId']", "input[name='amcNo']",
                        "input[id*='loginId']", "input[id*='customer']"]:
                try:
                    el = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, sel)))
                    if el.is_displayed():
                        el.click(); time.sleep(0.2)
                        el.send_keys(Keys.CONTROL + 'a'); el.send_keys(Keys.DELETE); time.sleep(0.1)
                        el.send_keys(member_no)
                        print(f"  会員番号: {sel} len={len(el.get_attribute('value') or '')}")
                        member_filled = True; break
                except: pass
            if not member_filled:
                try:
                    for inp in driver.find_elements(By.CSS_SELECTOR,
                            "[role='dialog'] input, dialog input, .modal input, "
                            "input[type='text'], input[type='tel'], input[type='number']"):
                        if inp.is_displayed():
                            inp.click(); time.sleep(0.2)
                            inp.send_keys(Keys.CONTROL + 'a'); inp.send_keys(Keys.DELETE); time.sleep(0.1)
                            inp.send_keys(member_no)
                            print(f"  会員番号（FB）: len={len(inp.get_attribute('value') or '')}")
                            member_filled = True; break
                except Exception as fb_e:
                    print(f"  会員番号FB失敗: {fb_e}")

            # ── 「次へ」ボタン（2段階ログイン対応） ──
            time.sleep(1)
            next_res = driver.execute_script("""
                for (var b of document.querySelectorAll('button')) {
                    var t = b.textContent.trim();
                    if ((t==='次へ'||t==='次') && b.offsetParent!==null) { b.click(); return 'clicked:'+t; }
                }
                return 'none';
            """)
            if next_res and next_res.startswith('clicked'):
                print(f"  「次へ」: {next_res}")
                time.sleep(4)
            ss(driver, "ana_before_pw")

            # ── パスワード入力（JS native setter でReact状態を更新） ──
            from selenium.webdriver.common.action_chains import ActionChains
            pw_el = None
            pw_filled = False
            # モーダル内のパスワードフィールドを優先して取得（ページ外の非表示フィールドを除外）
            for sel in ["[role='dialog'] input[type='password']",
                        "dialog input[type='password']",
                        ".modal input[type='password']",
                        "input[type='password']"]:
                try:
                    el = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, sel)))
                    if not el.is_displayed(): continue
                    # JS native setter + 複数イベント発火
                    diag = driver.execute_script("""
                        var el=arguments[0], pw=arguments[1];
                        el.focus();
                        var nativeSetter=Object.getOwnPropertyDescriptor(
                            window.HTMLInputElement.prototype,'value').set;
                        nativeSetter.call(el,'');
                        el.dispatchEvent(new InputEvent('input',{bubbles:true,inputType:'deleteContentBackward'}));
                        nativeSetter.call(el,pw);
                        ['input','change','keyup','blur','focus'].forEach(function(t){
                            el.dispatchEvent(new Event(t,{bubbles:true}));
                        });
                        el.dispatchEvent(new InputEvent('input',{bubbles:true,inputType:'insertText',data:pw}));
                        return el.value.length;
                    """, el, password)
                    print(f"  PW JS: len={diag}")
                    if diag > 0:
                        pw_el = el; pw_filled = True; break
                except Exception as e:
                    print(f"  PW例外({sel}): {e}")
            if not pw_filled:
                print("  ⚠ PW設定失敗")
            time.sleep(1)
            ss(driver, "ana_form_filled")

            # ── ログインボタン：JS でモーダル内ボタンを直接クリック ──
            try:
                click_res = driver.execute_script("""
                    // 1. モーダル内のボタン優先
                    var dialogs = document.querySelectorAll('[role="dialog"], dialog, .modal, [class*="modal"], [class*="dialog"]');
                    for (var d of dialogs) {
                        var btns = d.querySelectorAll('button');
                        for (var b of btns) {
                            var t = b.textContent.trim();
                            var r = b.getBoundingClientRect();
                            if (t.includes('ログイン') && r.width > 0 && r.height > 0) {
                                b.click();
                                return 'modal-btn: ' + t + ' x=' + Math.round(r.x);
                            }
                        }
                    }
                    // 2. ページ全体のボタン（offsetParent非nullで表示確認）
                    var allBtns = Array.from(document.querySelectorAll('button'));
                    var cands = allBtns.filter(function(b) {
                        return b.textContent.trim().includes('ログイン') && b.offsetParent !== null;
                    });
                    // x座標最大（右端）を選ぶ
                    if (cands.length > 0) {
                        cands.sort(function(a,b){ return b.getBoundingClientRect().x - a.getBoundingClientRect().x; });
                        cands[0].click();
                        return 'page-btn: ' + cands[0].textContent.trim() + ' x=' + Math.round(cands[0].getBoundingClientRect().x);
                    }
                    // 3. type=submit ボタン
                    var subBtns = document.querySelectorAll('button[type="submit"], input[type="submit"]');
                    for (var s of subBtns) {
                        if (s.offsetParent !== null) { s.click(); return 'submit-btn'; }
                    }
                    return 'not-found';
                """)
                print(f"  ログインボタン(JS): {click_res}")
                if 'not-found' in str(click_res):
                    print("  ⚠ ログインボタン未検出 → Enter送信")
                    if pw_el:
                        pw_el.send_keys(Keys.RETURN)
            except Exception as e:
                print(f"  ログインボタン例外: {e}")
                if pw_el:
                    try: pw_el.send_keys(Keys.RETURN); print("  Enter送信(fallback)")
                    except: pass
            time.sleep(3)

            # ── ログイン完了待ち（モーダルの消滅 = パスワードフィールドが全て不可視になる） ──
            url_before = driver.current_url
            for i in range(25):
                time.sleep(1)
                # 表示中のパスワードフィールドが残っているか確認
                gone = driver.execute_script("""
                    var pws = document.querySelectorAll('input[type="password"]');
                    for (var pw of pws) {
                        var r = pw.getBoundingClientRect();
                        if (r.width > 0 && r.height > 0) return false;  // まだ見えている
                    }
                    return true;  // 全て非表示
                """)
                cur_url = driver.current_url
                if gone or cur_url != url_before:
                    print(f"  ログイン完了 ({i+1}秒, url_changed={cur_url != url_before})")
                    break
            else:
                print("  ※ログイン完了タイムアウト（モーダルが閉じず）")

            ss(driver, "ana_after_login")
            print(f"  ログイン後URL: {driver.current_url}")

        # ── 予約確認ページへ（通常ANAサイト経由）──
        time.sleep(2)
        cur = driver.current_url
        print(f"  ログイン後URL: {cur}")
        try:
            ss(driver, "ana_after_login_page")
        except Exception:
            pass

        # 国内線予約確認タブ or マイANAの予約リンクを探す
        clicked = False
        nav_result = driver.execute_script("""
            var keywords = ['国内線予約確認', '予約確認', '予約一覧', 'reservation', 'booking'];
            for (var kw of keywords) {
                for (var el of document.querySelectorAll('a, button, [role="tab"]')) {
                    var t = (el.textContent||'').trim();
                    var h = (el.getAttribute('href')||'');
                    if (t.includes(kw) || h.includes(kw)) {
                        var r = el.getBoundingClientRect();
                        if (r.width > 0 && r.height > 0) {
                            el.click(); return 'clicked:' + t + '|' + h.substring(0,80);
                        }
                    }
                }
            }
            return 'none';
        """)
        print(f"  予約リンク: {nav_result}")
        if nav_result != 'none':
            time.sleep(5); clicked = True

        # フォールバック: マイANA予約確認ページへ直接遷移（複数URL試行）
        if not clicked:
            for try_url in [
                "https://www.ana.co.jp/ja/jp/my/",
                "https://www.ana.co.jp/ja/jp/my/reservation/itinerary-list/",
            ]:
                print(f"  → 直接遷移: {try_url}")
                driver.get(try_url)
                time.sleep(5)
                cur2 = driver.current_url
                # ログイン状態確認（ログインページに飛ばされたら失敗）
                if 'login' in cur2.lower() or 'signin' in cur2.lower():
                    print(f"    ログインページへリダイレクト: {cur2}")
                    continue
                body2 = driver.find_element(By.TAG_NAME, 'body').text
                # 予約らしき情報があるか確認
                if any(kw in body2 for kw in ['NH', 'ANA', '予約番号', '便名', 'フライト']):
                    print(f"    予約情報あり → このページを使用")
                    clicked = True
                    break
                # マイANAトップで予約リンクを探す
                nav2 = driver.execute_script("""
                    var kws=['予約確認','旅程','itinerary','booking-list','reservation/list'];
                    for(var kw of kws){
                        for(var el of document.querySelectorAll('a')){
                            var h=(el.getAttribute('href')||'');
                            var t=(el.textContent||'').trim();
                            if(h.includes(kw)||t.includes(kw)){
                                var r=el.getBoundingClientRect();
                                if(r.width>0&&r.height>0){el.click();return'clicked:'+t+'|'+h;}
                            }
                        }
                    }
                    return 'none';
                """)
                print(f"    マイANA予約リンク: {nav2}")
                if nav2 != 'none':
                    time.sleep(5); clicked = True; break

        cur = driver.current_url
        print(f"  予約ページURL: {cur}")
        try:
            ss(driver, "ana_reservations")
        except Exception:
            pass
        body_text = driver.find_element(By.TAG_NAME, 'body').text
        print(f"  ページテキスト冒頭: {body_text[:500].replace(chr(10),' ')}")

        # ── 予約情報パース（aswbe.ana.co.jp 形式） ──
        # フォーマット例:
        #   予約番号 FM4TKD
        #   NH022 大阪(伊丹) 東京(羽田)
        #   2026年5月23日(土) 11:00 → 12:15
        bookings = _parse_ana_booking_page(driver, body_text)

        if not bookings:
            # フォールバック：ページソースから直接解析
            src = driver.page_source
            bookings = _parse_ana_booking_src(src)

        if not bookings:
            print("  ※ 自動解析できませんでした。スクリーンショットを確認してください。")
            print(f"     → {os.path.join(TMPDIR, 'travel_ana_reservations.png')}")

    except Exception as e:
        print(f"  ANA取得エラー: {e}")
        import traceback; traceback.print_exc()
        try: ss(driver, "ana_error")
        except: pass
    finally:
        try:
            import threading
            threading.Thread(target=driver.quit, daemon=True).start()
        except Exception:
            pass

    return bookings


def _parse_ana_booking_page(driver, body_text):
    """aswbe.ana.co.jp/webapps/servicing/booking-list のページをパース

    フォーマット例（body_text を改行で分割）:
        FM4TKD
        5月23日（土）
        11:00   大阪(伊丹)
        12:15   東京(羽田)
        NH022
        ...
        5月24日（日）
        10:00   東京(羽田)
        ...
        NH019
    """
    from selenium.webdriver.common.by import By
    bookings = []
    seen = set()

    lines = [l.strip() for l in body_text.splitlines() if l.strip()]

    # 予約番号パターン: 大文字英数字6文字（行単独）
    conf_pattern = re.compile(r'^([A-Z0-9]{6})$')
    # 日付パターン: N月N日（曜日）または YYYY年N月N日
    date_pattern = re.compile(
        r'(?:(\d{4})年)?(\d{1,2})月(\d{1,2})日')
    # 時刻パターン
    time_pattern = re.compile(r'^(\d{1,2}):(\d{2})$')
    # フライト番号
    flt_pattern = re.compile(r'^(NH|ANA)(\d{3,4})$')

    cur_year = date.today().year
    cur_conf = None
    cur_date = None
    cur_times = []
    cur_flt = None

    def flush():
        nonlocal cur_conf, cur_date, cur_times, cur_flt
        if cur_conf and cur_date and cur_flt and cur_times:
            dep_time = f"{int(cur_times[0][0]):02d}:{cur_times[0][1]}"
            arr_time = f"{int(cur_times[1][0]):02d}:{cur_times[1][1]}" if len(cur_times) >= 2 else ''
            key = (cur_flt, cur_date)
            if key not in seen:
                seen.add(key)
                e = {'type':'flight','flight_no':cur_flt,'date':cur_date,
                     'dep_time':dep_time,'arr_time':arr_time,'confirmation':cur_conf,
                     'raw':f'{cur_flt} {cur_date} {dep_time}〜{arr_time}'}
                bookings.append(e)
                print(f"  予約確定: {cur_flt} {cur_date} {dep_time}〜{arr_time} (予約番号:{cur_conf})")
        cur_date = None; cur_times = []; cur_flt = None

    for line in lines:
        # 予約番号検出 → 新しい予約ブロック開始
        m = conf_pattern.match(line)
        if m:
            flush()
            cur_conf = m.group(1)
            continue

        # 日付検出 → フライトブロック開始（前のフライトを確定）
        m = date_pattern.search(line)
        if m and cur_conf:
            flush()
            yr = int(m.group(1)) if m.group(1) else cur_year
            mo, dy = int(m.group(2)), int(m.group(3))
            cur_date = date(yr, mo, dy)
            cur_times = []
            cur_flt = None
            continue

        # 時刻検出
        m = time_pattern.match(line)
        if m and cur_date is not None:
            cur_times.append((m.group(1), m.group(2)))
            continue

        # フライト番号検出
        m = flt_pattern.match(line)
        if m and cur_date is not None:
            cur_flt = m.group(1) + m.group(2)
            continue

    flush()  # 最後のブロック

    if not bookings:
        # フォールバック: 全文を平坦化してパターンマッチ
        flat = re.sub(r'\s+', ' ', body_text)
        # 予約番号 → 日付 → 時刻 → 便名 の順に検索
        for conf_m in re.finditer(r'\b([A-Z0-9]{6})\b', flat):
            conf = conf_m.group(1)
            ctx = flat[conf_m.start():conf_m.start()+800]
            for date_m in re.finditer(r'(\d{1,2})月(\d{1,2})日', ctx):
                mo, dy = int(date_m.group(1)), int(date_m.group(2))
                sub = ctx[date_m.start():date_m.start()+200]
                times = re.findall(r'(\d{1,2}):(\d{2})', sub)
                flt_m = re.search(r'(NH|ANA)(\d{3,4})', sub)
                if flt_m and times:
                    flt_no = flt_m.group(1)+flt_m.group(2)
                    flt_date = date(cur_year, mo, dy)
                    dep_time = f"{int(times[0][0]):02d}:{times[0][1]}"
                    key = (flt_no, flt_date)
                    if key not in seen:
                        seen.add(key)
                        bookings.append({'type':'flight','flight_no':flt_no,
                                         'date':flt_date,'dep_time':dep_time,
                                         'confirmation':conf,'raw':sub[:80]})
                        print(f"  予約確定(FB): {flt_no} {flt_date} {dep_time}")

    return bookings


def _parse_ana_booking_src(src):
    """ページソースからフライト情報を抽出（フォールバック）"""
    bookings = []
    seen = set()

    # NH番号 + 日付 + 時刻の組み合わせを探す
    # 日本語日付形式
    date_blocks = list(re.finditer(r'(\d{4})年(\d{1,2})月(\d{1,2})日', src))
    for dm in date_blocks:
        yr, mo, dy = int(dm.group(1)), int(dm.group(2)), int(dm.group(3))
        ctx = src[max(0, dm.start()-300):dm.end()+200]
        flt_m = re.search(r'(NH|ANA)\s*(\d{3,4})', ctx)
        if not flt_m: continue
        flt_no = flt_m.group(1) + flt_m.group(2)
        times = re.findall(r'(\d{1,2}):(\d{2})', ctx)
        dep_time = f"{int(times[0][0]):02d}:{times[0][1]}" if times else '00:00'
        key = (flt_no, (yr, mo, dy))
        if key in seen: continue
        seen.add(key)
        bookings.append({
            'type': 'flight',
            'flight_no': flt_no,
            'date': date(yr, mo, dy),
            'dep_time': dep_time,
            'raw': re.sub(r'\s+', ' ', ctx[:100]),
        })
        print(f"  予約確定(FB): {flt_no} {yr}-{mo:02d}-{dy:02d} {dep_time}")

    return bookings

# ──────────────────────────────────
# Hilton 予約取得
# ──────────────────────────────────

def fetch_hilton_bookings(honors_no, password):
    """Hilton Honorsにログインして予約一覧を返す"""
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    print("\n【Hilton】予約情報取得中...")
    driver = make_driver()
    bookings = []
    try:
        # Hiltonサインインページへ直接アクセス
        driver.get("https://www.hilton.com/en/hilton-honors/login/")
        time.sleep(6)
        ss(driver, "hilton_top")
        print(f"  サインインページURL: {driver.current_url}")

        wait = WebDriverWait(driver, 20)

        # Cookie同意バナーがあれば「Accept」クリック
        try:
            accept_btn = driver.find_element(By.XPATH,
                "//button[normalize-space(text())='Accept' or normalize-space(text())='ACCEPT']")
            if accept_btn.is_displayed():
                driver.execute_script("arguments[0].click()", accept_btn)
                print("  Cookie同意: Accept クリック")
                time.sleep(2)
        except: pass

        # Honors番号入力（JS React対応）
        username_filled = False
        for sel in ["input[id='username']", "input[name='username']",
                    "input[autocomplete='username']", "input[type='email']", "input[type='text']"]:
            try:
                el = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, sel)))
                if el.is_displayed():
                    js_fill_input(driver, sel, honors_no)
                    time.sleep(0.3)
                    el.clear(); el.send_keys(honors_no)  # JS後に send_keys でも確実に入力（clear必須）
                    val = el.get_attribute('value') or ''
                    print(f"  オナーズ番号入力: {sel} → value={val!r}")
                    username_filled = True; break
            except: pass

        if not username_filled:
            print("  ⚠ オナーズ番号フィールド未発見")

        time.sleep(0.5)

        # パスワード入力（JS React対応）
        from selenium.webdriver.common.keys import Keys
        pw_el = None
        for sel in ["input[id='password']", "input[name='password']", "input[type='password']"]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                if el.is_displayed():
                    js_fill_input(driver, sel, password)
                    time.sleep(0.3)
                    el.clear(); el.send_keys(password)  # clear必須（二重入力防止）
                    print(f"  パスワード入力: {sel} → len={len(el.get_attribute('value') or '')}")
                    pw_el = el; break
            except: pass

        time.sleep(1)
        ss(driver, "hilton_signin")

        # サインインボタン（テキストで特定してクリック + Enterキーのバックアップ）
        signin_submitted = False
        try:
            # "Sign In"テキストを持つボタンを優先
            btns = driver.find_elements(By.XPATH,
                "//button[contains(normalize-space(text()),'Sign In') or "
                "contains(normalize-space(text()),'SIGN IN') or "
                "contains(normalize-space(text()),'sign in')]")
            for btn in btns:
                if btn.is_displayed() and btn.is_enabled():
                    # スクロールしてクリック
                    driver.execute_script("arguments[0].scrollIntoView(true)", btn)
                    time.sleep(0.3)
                    btn.click()
                    print(f"  サインインボタンクリック: {btn.text!r}")
                    signin_submitted = True; break
        except: pass

        # フォールバック: submit型ボタン
        if not signin_submitted:
            for sel in ["button[data-testid*='signin']", "button[data-testid*='submit']",
                        "input[type='submit']"]:
                try:
                    btn = driver.find_element(By.CSS_SELECTOR, sel)
                    if btn.is_displayed() and btn.is_enabled():
                        driver.execute_script("arguments[0].click()", btn)
                        print(f"  サインインボタンクリック(FB): {sel}")
                        signin_submitted = True; break
                except: pass

        # 最終フォールバック: パスワードフィールドでEnterキー
        if not signin_submitted and pw_el:
            try:
                pw_el.send_keys(Keys.RETURN)
                print("  Enterキーでフォーム送信")
                signin_submitted = True
            except: pass

        # ログイン完了を待つ（URLが変わるまで最大30秒）
        for _ in range(30):
            time.sleep(1)
            cur = driver.current_url
            if 'login' not in cur and 'signin' not in cur.lower():
                print(f"  ログイン完了確認"); break
        else:
            # エラーメッセージ確認
            try:
                err_el = driver.find_element(By.CSS_SELECTOR, "[class*='error'], [class*='alert'], [role='alert']")
                print(f"  ログインエラーメッセージ: {err_el.text[:100]}")
            except: pass

        ss(driver, "hilton_after_login")
        print(f"  ログイン後URL: {driver.current_url}")

        # 予約アクティビティページへ（正しいURL）
        driver.get("https://www.hilton.com/en/hilton-honors/guest/activity/")
        time.sleep(8)
        print(f"  予約ページURL: {driver.current_url}")
        ss(driver, "hilton_reservations")

        # ページテキストをデバッグ出力
        body_text = driver.find_element(By.TAG_NAME, 'body').text
        print(f"  ページテキスト冒頭: {body_text[:300].replace(chr(10),' ')}")

        # 予約情報抽出（Hilton Activityページ対応）
        MONTHS = {'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,
                  'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12}
        DAYS_OF_WEEK = {'SAT','SUN','MON','TUE','WED','THU','FRI'}

        def parse_hilton_date(txt):
            """'23 MAY SAT' や 'May 23, 2026' などを date に変換"""
            yr = date.today().year
            m = re.search(r'(\d{1,2})\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)', txt.upper())
            if m:
                return date(yr, MONTHS[m.group(2)], int(m.group(1)))
            for fmt in ['%B %d, %Y', '%b %d, %Y', '%B %d %Y', '%b %d %Y']:
                try: return datetime.strptime(re.sub(r'\s+', ' ', txt.strip()), fmt).date()
                except: pass
            m2 = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', txt)
            if m2: return date(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
            return None

        # body_text全体を空白で平坦化（改行区切りの日付に対応）
        body_flat = re.sub(r'\s+', ' ', body_text)
        print(f"  body_flat長: {len(body_flat)}文字")

        # Confirmation番号を全て検索し、前後コンテキストから予約情報を抽出
        for conf_match in re.finditer(r'Confirmation\s*#?\s*(\d{6,})', body_flat, re.IGNORECASE):
            conf_no = conf_match.group(1)
            ctx_start = max(0, conf_match.start() - 350)
            ctx_end = min(len(body_flat), conf_match.end() + 100)
            ctx = body_flat[ctx_start:ctx_end]
            print(f"  Confirmation #{conf_no} コンテキスト: {ctx[:300]}")

            date_chunks = re.findall(
                r'\d{1,2}\s+(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)'
                r'(?:\s+(?:SAT|SUN|MON|TUE|WED|THU|FRI))?',
                ctx.upper())
            hotel_name_m = re.search(
                r'(?:Conrad|DoubleTree by Hilton|Waldorf Astoria|Curio Collection|LXR Hotels|'
                r'Tapestry Collection|Embassy Suites|Hampton Inn|Home2 Suites|Tru by Hilton)\s+'
                r'[\w\s]+?(?=\s+Confirmation|\s+Points|\s+\d{1,2}\s+(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)|$)|'
                r'Hilton\s+(?:Tokyo|Osaka|Kyoto|Nagoya|Fukuoka|Garden|Inn|Shinjuku|Odaiba|Bay|'
                r'Roppongi|Shibuya|Ginza|Marunouchi|Airport|[\w]+)',
                ctx, re.IGNORECASE)
            hotel_name = hotel_name_m.group(0).strip() if hotel_name_m else 'Hilton'
            # 末尾の余分なスペースや不要語を除去
            hotel_name = re.sub(r'\s+(Confirmation|Points|Pending|View).*$', '', hotel_name, flags=re.IGNORECASE).strip()

            if len(date_chunks) >= 2:
                ci = parse_hilton_date(date_chunks[0])
                co = parse_hilton_date(date_chunks[1])
                if ci and co and ci < co:
                    bookings.append({'type':'hotel','hotel':hotel_name,'checkin':ci,'checkout':co,
                                     'raw':ctx[:100],'confirmation':conf_no})
                    print(f"  → 予約確定: {hotel_name} {ci}〜{co} (Conf#{conf_no})")
                    continue

            # フォールバック: 従来形式
            date_strs = re.findall(r'\w+ \d{1,2},? \d{4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}', ctx)
            if len(date_strs) >= 2:
                ci = parse_hilton_date(date_strs[0])
                co = parse_hilton_date(date_strs[1])
                if ci and co and ci < co:
                    bookings.append({'type':'hotel','hotel':hotel_name,'checkin':ci,'checkout':co,
                                     'raw':ctx[:100],'confirmation':conf_no})
                    print(f"  → 予約確定(FB): {hotel_name} {ci}〜{co} (Conf#{conf_no})")

        if not bookings:
            # さらにフォールバック: ページ全体から日付ペアを探す
            rows = driver.find_elements(By.CSS_SELECTOR,
                "[class*='reservation'], [class*='booking'], article, "
                "[class*='activity'], [class*='stay-card']")
            print(f"  ホテル予約行候補(FB): {len(rows)}件")
            for row in rows[:30]:
                txt = row.text.strip()
                if not txt or len(txt) < 15: continue
                if not any(kw in txt for kw in ['Confirmation', 'Pending Stay', 'Completed Stay']): continue
                date_chunks = re.findall(
                    r'\d{1,2}\s+(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)'
                    r'(?:\s+(?:SAT|SUN|MON|TUE|WED|THU|FRI))?',
                    re.sub(r'\s+', ' ', txt).upper())
                hotel_name_m = re.search(r'Hilton[\w\s]+', txt)
                hotel_name = hotel_name_m.group(0).strip() if hotel_name_m else 'Hilton'
                if len(date_chunks) >= 2:
                    ci = parse_hilton_date(date_chunks[0])
                    co = parse_hilton_date(date_chunks[1])
                    if ci and co and ci < co:
                        bookings.append({'type':'hotel','hotel':hotel_name,'checkin':ci,'checkout':co,'raw':txt[:100]})
                        print(f"  → 予約確定(DOM): {hotel_name} {ci}〜{co}")

        if not bookings:
            print("  ※ 自動解析できませんでした。スクリーンショットを確認してください。")
            print(f"     → {os.path.join(TMPDIR, 'travel_hilton_reservations.png')}")

    except Exception as e:
        print(f"  Hilton取得エラー: {e}")
        try:
            ss(driver, "hilton_error")
        except Exception:
            pass
    finally:
        try:
            import threading
            threading.Thread(target=driver.quit, daemon=True).start()
        except Exception:
            pass

    return bookings

# ──────────────────────────────────
# サイボウズ スケジュール取得
# ──────────────────────────────────

def fetch_cybozu_events(start_dt, end_dt):
    """指定期間のサイボウズ予定を取得する"""
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    print(f"\n【サイボウズ】予定取得中 ({start_dt} ～ {end_dt})...")
    cb_email, cb_pw = read_cred('Cybozu_Login')
    driver = make_driver()
    events = []

    try:
        # Cookie認証を試みる
        logged_in = False
        if os.path.exists(COOKIE_FILE):
            driver.get(CYBOZU_BASE + "/"); time.sleep(2)
            for ck in pickle.load(open(COOKIE_FILE, "rb")):
                try: driver.add_cookie(ck)
                except: pass
            driver.refresh(); time.sleep(3)
            logged_in = CYBOZU_BASE in driver.current_url and '/login' not in driver.current_url

        if not logged_in:
            # MS365 SSO ログイン
            driver.get(CYBOZU_BASE + "/"); time.sleep(4)
            url = driver.current_url
            if 'microsoftonline.com' in url or 'microsoft.com' in url:
                wait = WebDriverWait(driver, 20)
                try:
                    ei = wait.until(EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "input[type='email'],#i0116")))
                    ei.clear(); ei.send_keys(cb_email); time.sleep(1)
                    driver.find_element(By.CSS_SELECTOR,
                        "input[type='submit'][value='Next'],input#idSIButton9").click()
                    time.sleep(3)
                except: pass
                try:
                    pi = wait.until(EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "input[type='password'],#i0118")))
                    pi.clear(); pi.send_keys(cb_pw); time.sleep(1)
                    driver.find_element(By.CSS_SELECTOR,
                        "input[type='submit'],input#idSIButton9").click()
                    time.sleep(3)
                except: pass
                for _ in range(30):
                    if CYBOZU_BASE in driver.current_url: break
                    try:
                        btn = driver.find_element(By.ID, "idSIButton9")
                        val = (btn.get_attribute('value') or btn.text or '').strip()
                        if any(w in val for w in ('はい', 'Yes', 'サインイン')) and val not in ('次へ', 'Next'):
                            btn.click(); time.sleep(2); continue
                    except: pass
                    time.sleep(1)
            # サービス選択
            if CYBOZU_BASE in driver.current_url and '/o/' not in driver.current_url:
                for el in driver.find_elements(By.CSS_SELECTOR, "a[href*='/o/']"):
                    href = el.get_attribute('href') or ''
                    if '/o/' in href:
                        el.click(); time.sleep(3); break
            logged_in = CYBOZU_BASE in driver.current_url and '/login' not in driver.current_url
            if logged_in:
                pickle.dump(driver.get_cookies(), open(COOKIE_FILE, "wb"))

        if not logged_in:
            print("  サイボウズログイン失敗"); return []

        # UID取得
        driver.get(f"{CYBOZU_BASE}/o/ag.cgi?page=ScheduleIndex"); time.sleep(3)
        uid = gid = None
        for link in driver.find_elements(By.TAG_NAME, "a"):
            href = link.get_attribute("href") or ""
            if 'page=ScheduleUserDay' in href and 'UID=' in href:
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
        print(f"  UID={uid}, GID={gid}")

        # 日付ごとに予定を取得
        current = start_dt
        while current <= end_dt:
            cy_date = f"da.{current.year}.{current.month}.{current.day}"
            url = f"{CYBOZU_BASE}/o/ag.cgi?page=ScheduleUserDay&UID={uid}&GID={gid}&Date={cy_date}"
            driver.get(url); time.sleep(1.5)
            seen = set()
            for link in driver.find_elements(By.TAG_NAME, "a"):
                href = link.get_attribute("href") or ""
                txt = link.text.strip()
                if not txt or len(txt) <= 1: continue
                if not any(k in href for k in ["ScheduleDetail", "ScheduleView", "EventDetail"]): continue
                if txt in seen: continue
                seen.add(txt)
                time_str = ''
                try:
                    parent_txt = link.find_element(By.XPATH, "..").text
                    m = re.search(r'(\d{1,2}:\d{2})', parent_txt)
                    if m: time_str = m.group(1)
                except: pass
                events.append({
                    'date': current,
                    'date_str': current.strftime('%Y-%m-%d'),
                    'title': txt[:60],
                    'time': time_str,
                })
            current += timedelta(days=1)

        print(f"  取得完了: {len(events)}件")

    except Exception as e:
        print(f"  サイボウズ取得エラー: {e}")
    finally:
        import threading
        threading.Thread(target=driver.quit, daemon=True).start()

    return events

# ──────────────────────────────────
# サイボウズ スケジュール同期
# ──────────────────────────────────

def _cybozu_login_driver():
    """サイボウズにログイン済みのdriverを返す（Cookie優先）"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    cb_email, cb_pw = read_cred('Cybozu_Login')
    driver = make_driver(headless=True)
    logged_in = False

    if os.path.exists(COOKIE_FILE):
        driver.get(CYBOZU_BASE + "/"); time.sleep(2)
        for ck in pickle.load(open(COOKIE_FILE, "rb")):
            try: driver.add_cookie(ck)
            except: pass
        driver.refresh(); time.sleep(3)
        logged_in = CYBOZU_BASE in driver.current_url and '/login' not in driver.current_url

    if not logged_in:
        driver.get(CYBOZU_BASE + "/"); time.sleep(4)
        url = driver.current_url
        if 'microsoftonline.com' in url or 'microsoft.com' in url:
            wait = WebDriverWait(driver, 20)
            try:
                ei = wait.until(EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "input[type='email'],#i0116")))
                ei.clear(); ei.send_keys(cb_email); time.sleep(1)
                driver.find_element(By.CSS_SELECTOR,
                    "input[type='submit'][value='Next'],input#idSIButton9").click()
                time.sleep(3)
            except: pass
            try:
                pi = wait.until(EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "input[type='password'],#i0118")))
                pi.clear(); pi.send_keys(cb_pw); time.sleep(1)
                driver.find_element(By.CSS_SELECTOR,
                    "input[type='submit'],input#idSIButton9").click()
                time.sleep(3)
            except: pass
            for _ in range(30):
                if CYBOZU_BASE in driver.current_url: break
                try:
                    btn = driver.find_element(By.ID, "idSIButton9")
                    val = (btn.get_attribute('value') or btn.text or '').strip()
                    if any(w in val for w in ('はい', 'Yes', 'サインイン')) and val not in ('次へ', 'Next'):
                        btn.click(); time.sleep(2); continue
                except: pass
                time.sleep(1)
        if CYBOZU_BASE in driver.current_url and '/o/' not in driver.current_url:
            for el in driver.find_elements(By.CSS_SELECTOR, "a[href*='/o/']"):
                href = el.get_attribute('href') or ''
                if '/o/' in href:
                    el.click(); time.sleep(3); break
        logged_in = CYBOZU_BASE in driver.current_url and '/login' not in driver.current_url
        if logged_in:
            pickle.dump(driver.get_cookies(), open(COOKIE_FILE, "wb"))

    if not logged_in:
        driver.quit(); return None
    return driver


def _cybozu_event_exists(driver, date_obj, keyword):
    """指定日にキーワードを含む予定が既にあるか確認"""
    from selenium.webdriver.common.by import By
    cy_date = f"da.{date_obj.year}.{date_obj.month}.{date_obj.day}"
    url = f"{CYBOZU_BASE}/o/ag.cgi?page=ScheduleUserDay&UID=170&Date={cy_date}"
    driver.get(url); time.sleep(1.5)
    return keyword in driver.page_source


def _cybozu_find_eids_by_keyword(driver, date_obj, keyword):
    """指定日にキーワードを含む予定のEIDリストを返す"""
    from selenium.webdriver.common.by import By
    cy_date = f"da.{date_obj.year}.{date_obj.month}.{date_obj.day}"
    url = f"{CYBOZU_BASE}/o/ag.cgi?page=ScheduleUserDay&UID=170&GID=154&Date={cy_date}"
    driver.get(url); time.sleep(1.5)
    eids = []
    for a in driver.find_elements(By.TAG_NAME, "a"):
        txt = a.text.strip()
        href = a.get_attribute("href") or ""
        if keyword in txt and 'sEID=' in href:
            m = re.search(r'sEID=(\d+)', href)
            if m and m.group(1) not in eids:
                eids.append(m.group(1))
    return eids


def _cybozu_delete_by_keyword(driver, date_obj, keyword):
    """指定日のキーワードを含む予定を削除してEIDリストを返す"""
    from selenium.webdriver.common.by import By
    eids = _cybozu_find_eids_by_keyword(driver, date_obj, keyword)
    deleted = []
    cy_date = f"da.{date_obj.year}.{date_obj.month}.{date_obj.day}"
    for eid in eids:
        view_url = f"{CYBOZU_BASE}/o/ag.cgi?page=ScheduleView&date={cy_date}&uid=170&gid=154&sEID={eid}"
        driver.get(view_url); time.sleep(2)
        for el in driver.find_elements(By.TAG_NAME, 'a') + driver.find_elements(By.CSS_SELECTOR, "input[type='submit']"):
            txt = el.text or el.get_attribute('value') or ''
            if '削除' in txt:
                driver.execute_script('arguments[0].click()', el); time.sleep(2)
                for btn2 in driver.find_elements(By.CSS_SELECTOR, "input[type='submit']"):
                    val2 = btn2.get_attribute('value') or btn2.text or ''
                    if '削除' in val2:
                        driver.execute_script('arguments[0].click()', btn2); time.sleep(2); break
                print(f"    削除: EID={eid} ({keyword})")
                deleted.append(eid); break
    return deleted


def _cybozu_set_gyoji_mark(driver):
    """scheduleEventValue hidden input に行事マーク(s3)をJSでセット"""
    try:
        driver.execute_script("""
            var val = document.getElementById('scheduleEventValue');
            if (val) val.value = 's3,【行事】';
            var title = document.getElementById('scheduleEventTitle');
            if (title) title.innerText = '【行事】';
            var color = document.getElementById('scheduleEventColor');
            if (color) {
                color.className = ' scheduleMarkEventMenu scheduleMarkEventMenu3';
                color.style.display = '';
            }
        """)
        print("    行事マーク: 設定")
        return True
    except Exception as ex:
        print(f"    行事マーク設定失敗: {ex}")
        return False


def _cybozu_set_time_selects(driver, start_time, end_time):
    """SetTime.Hour/Minute と EndTime.Hour/Minute を設定"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import Select
    h_s, m_s = start_time.split(':')
    h_e, m_e = end_time.split(':')
    try:
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='SetTime.Hour']")).select_by_value(str(int(h_s)))
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='SetTime.Minute']")).select_by_value(str(int(m_s)))
        print(f"    開始時刻: {start_time}")
    except Exception as ex:
        print(f"    開始時刻設定失敗: {ex}"); return False
    try:
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='EndTime.Hour']")).select_by_value(str(int(h_e)))
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='EndTime.Minute']")).select_by_value(str(int(m_e)))
        print(f"    終了時刻: {end_time}")
    except Exception as ex:
        print(f"    終了時刻設定失敗: {ex}"); return False
    return True


def _cybozu_add_timed_event(driver, date_obj, title, start_time, end_time, memo=''):
    """サイボウズに時刻指定の行事予定を追加"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import Select
    cy_date = f"da.{date_obj.year}.{date_obj.month}.{date_obj.day}"

    # ScheduleUserDay → 「予定を登録する」クリック
    ud_url = f"{CYBOZU_BASE}/o/ag.cgi?page=ScheduleUserDay&UID=170&GID=154&Date={cy_date}"
    driver.get(ud_url); time.sleep(3)

    add_clicked = False
    for el in driver.find_elements(By.XPATH, "//a[contains(., '予定を登録')]"):
        if el.is_displayed():
            driver.execute_script("arguments[0].click()", el)
            time.sleep(3); add_clicked = True; break

    if not add_clicked:
        print("    「予定を登録する」リンクが見つかりません")
        return False

    # タイトル入力
    for sel in ["input[name='Detail']", "input[name='Description']", "input[name='Title']"]:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            if el.is_displayed():
                el.clear(); el.send_keys(title); break
        except: pass

    # 行事マークをJSで設定
    _cybozu_set_gyoji_mark(driver)

    # 時刻設定（確定済みSELECT名）
    _cybozu_set_time_selects(driver, start_time, end_time)

    # メモ
    if memo:
        for sel in ["textarea[name='Memo']", "textarea[name='memo']", "textarea"]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                if el.is_displayed():
                    el.clear(); el.send_keys(memo); break
            except: pass

    # 登録する
    for sel in ["input[type='submit'][name='Entry']", "input[type='submit'][value*='登録する']",
                "input[type='submit']"]:
        try:
            btn = driver.find_element(By.CSS_SELECTOR, sel)
            if btn.is_displayed() and btn.get_attribute('name') != 'AddressSearch':
                driver.execute_script("arguments[0].click()", btn)
                time.sleep(3); return True
        except: pass
    return False


def _cybozu_add_allday_event(driver, date_obj, title, memo=''):
    """サイボウズに終日予定を追加（外出扱い）"""
    from selenium.webdriver.common.by import By
    cy_date = f"da.{date_obj.year}.{date_obj.month}.{date_obj.day}"
    add_url = f"{CYBOZU_BASE}/o/ag.cgi?page=ScheduleAddShow&Date={cy_date}"
    driver.get(add_url); time.sleep(3)

    # 初回は診断用スクリーンショット保存
    ss_path = os.path.join(TMPDIR, "cybozu_add_form.png")
    if not os.path.exists(ss_path):
        driver.save_screenshot(ss_path)

    # タイトル入力
    title_filled = False
    for sel in ["input[name='Description']", "input[name='Title']",
                "input[name='schedule_title']", "input[id*='title']",
                "input[id*='subject']", "input[id*='Title']"]:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            if el.is_displayed():
                el.clear(); el.send_keys(title); title_filled = True; break
        except: pass

    if not title_filled:
        # ページ内の最初のテキスト入力欄
        inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
        for inp in inputs:
            if inp.is_displayed():
                inp.clear(); inp.send_keys(title); title_filled = True; break

    # 終日チェック
    for sel in ["input[name='Allday']", "input[name='allday']",
                "input[id*='allday']", "input[value='1'][name*='day']"]:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            if not el.is_selected():
                driver.execute_script("arguments[0].click()", el)
            break
        except: pass

    # 外出ラジオボタン（EventType/Category）
    for sel in ["input[name='Category'][value='outside']",
                "input[name='EventType'][value='2']",
                "input[value='outside']", "input[value='外出']"]:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            driver.execute_script("arguments[0].click()", el); break
        except: pass

    # メモ入力
    if memo:
        for sel in ["textarea[name='Memo']", "textarea[name='memo']",
                    "textarea[name='Description']", "textarea"]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                if el.is_displayed():
                    el.clear(); el.send_keys(memo); break
            except: pass

    # 保存ボタン
    for sel in ["input[type='submit'][name*='Add']", "input[type='submit'][value*='追加']",
                "input[type='submit'][value*='保存']", "input[type='submit']",
                "button[type='submit']"]:
        try:
            btn = driver.find_element(By.CSS_SELECTOR, sel)
            if btn.is_displayed():
                driver.execute_script("arguments[0].click()", btn)
                time.sleep(3); return True
        except: pass

    return False


def _cybozu_add_period_event(driver, start_date, end_date, title, memo=''):
    """サイボウズ「期間予定」タブ(ScheduleBannerEntry)でホテル等の複数日予定を追加"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import Select
    cy_start = f"da.{start_date.year}.{start_date.month}.{start_date.day}"

    # ScheduleUserDay → 「予定を登録する」クリック
    ud_url = f"{CYBOZU_BASE}/o/ag.cgi?page=ScheduleUserDay&UID=170&GID=154&Date={cy_start}"
    driver.get(ud_url); time.sleep(3)

    add_clicked = False
    for el in driver.find_elements(By.XPATH, "//a[contains(., '予定を登録')]"):
        if el.is_displayed():
            driver.execute_script("arguments[0].click()", el)
            time.sleep(3); add_clicked = True; break

    if not add_clicked:
        print("    「予定を登録する」リンクが見つかりません")
        return False

    # 「期間予定」タブをクリック (ScheduleBannerEntry)
    tab_clicked = False
    for el in driver.find_elements(By.XPATH, "//a[contains(., '期間予定')]"):
        if el.is_displayed():
            driver.execute_script("arguments[0].click()", el)
            time.sleep(3); tab_clicked = True; break
    if not tab_clicked:
        print("    「期間予定」タブが見つかりません")

    # タイトル入力
    for sel in ["input[name='Detail']", "input[name='Description']", "input[name='Title']"]:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            if el.is_displayed():
                el.clear(); el.send_keys(title); break
        except: pass

    # 終了日設定（期間予定の終了日）
    end_str = f"{end_date.year}/{end_date.month:02d}/{end_date.day:02d}"
    end_filled = False
    # ScheduleBannerEntry の終了日セレクトを試みる
    try:
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='EndDate.Year']")).select_by_value(str(end_date.year))
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='EndDate.Month']")).select_by_value(str(end_date.month))
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='EndDate.Day']")).select_by_value(str(end_date.day))
        print(f"    終了日(SELECT): {end_str}"); end_filled = True
    except: pass
    if not end_filled:
        for sel in ["input[name='EndDate']", "input[name='end_date']", "input[name='EndDay']"]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                if el.is_displayed():
                    el.clear(); el.send_keys(end_str)
                    print(f"    終了日(INPUT): {end_str}"); end_filled = True; break
            except: pass
    if not end_filled:
        print(f"    終了日フィールド見つからず")

    # メモ
    if memo:
        for sel in ["textarea[name='Memo']", "textarea[name='memo']", "textarea"]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                if el.is_displayed():
                    el.clear(); el.send_keys(memo); break
            except: pass

    # 登録する（AddressSearchボタンを誤クリックしないよう確認）
    for sel in ["input[type='submit'][name='Entry']", "input[type='submit'][value*='登録する']",
                "input[type='submit'][value*='保存する']"]:
        try:
            btn = driver.find_element(By.CSS_SELECTOR, sel)
            if btn.is_displayed() and btn.get_attribute('name') != 'AddressSearch':
                driver.execute_script("arguments[0].click()", btn)
                time.sleep(3); return True
        except: pass
    return False


def sync_cybozu_schedule(flights, hotels):
    """フライト・ホテル予約をサイボウズに終日・外出予定として同期する"""
    if not flights and not hotels:
        print("\n【サイボウズ同期】同期対象なし（予約情報なし）")
        return [], []

    print("\n【サイボウズ同期】スケジュール更新中...")
    driver = _cybozu_login_driver()
    if not driver:
        print("  サイボウズログイン失敗 → 同期スキップ")
        return [], []

    synced, skipped = [], []
    try:
        # フライト予定を登録（行事・時刻指定）
        for f in flights:
            fn = f.get('flight_no', 'フライト')
            dep = f.get('dep_time', '00:00')
            arr = f.get('arr_time', '')
            title = f"移動【{fn}】"
            # 旧形式（✈ NH022 ...）が残っていれば削除
            old_eids = _cybozu_delete_by_keyword(driver, f['date'], f"✈ {fn}")
            if old_eids:
                print(f"  旧イベント削除: {f['date']} ✈{fn} EID={old_eids}")
            # 新形式も削除して再登録（時刻を正しく入れるため）
            new_eids = _cybozu_delete_by_keyword(driver, f['date'], title)
            if new_eids:
                print(f"  再登録のため削除: {f['date']} {title} EID={new_eids}")
            if arr:
                ok = _cybozu_add_timed_event(driver, f['date'], title, dep, arr,
                                             memo=f.get('raw', ''))
            else:
                # 到着時刻不明の場合は終日
                ok = _cybozu_add_period_event(driver, f['date'], f['date'], title,
                                              memo=f.get('raw', ''))
            if ok:
                t_range = f"{dep}〜{arr}" if arr else dep
                print(f"  登録: {f['date']} {t_range} {title}")
                synced.append(title)
            else:
                print(f"  登録失敗: {f['date']} {title}")

        # ホテル滞在を「期間予定」で登録（チェックイン〜チェックアウト前日）
        for h in hotels:
            hotel_name = h['hotel']
            conf_no = h.get('confirmation', '')
            title = "【出張】"
            memo = hotel_name + (f" Conf#{conf_no}" if conf_no else "")
            # 既存の「【出張】」を削除して期間予定で再登録
            del_eids = _cybozu_delete_by_keyword(driver, h['checkin'], "【出張】")
            if del_eids:
                print(f"  再登録のため削除: {h['checkin']} 【出張】 EID={del_eids}")
            ok = _cybozu_add_period_event(driver, h['checkin'], h['checkout'],
                                          title, memo=memo)
            if ok:
                print(f"  登録: {h['checkin']}〜{h['checkout']} {title}（{memo}）")
                synced.append(f"{h['checkin']}-{h['checkout']} {title}")
            else:
                print(f"  登録失敗: {h['checkin']}〜{h['checkout']} {title}")

    except Exception as e:
        print(f"  同期エラー: {e}")
    finally:
        import threading
        threading.Thread(target=driver.quit, daemon=True).start()

    print(f"  完了: 登録 {len(synced)}件 / スキップ {len(skipped)}件")
    return synced, skipped


# ──────────────────────────────────
# 整合性チェック・レポート
# ──────────────────────────────────

def time_to_min(t_str):
    """'HH:MM' → 分"""
    try:
        h, m = t_str.split(':'); return int(h)*60 + int(m)
    except: return -1

def check_and_report(flights, hotels, events):
    print("\n" + "="*60)
    print("  旅程・サイボウズ整合性レポート")
    print("="*60)

    if not flights and not hotels:
        print("\n予約情報が取得できませんでした。")
        print("スクリーンショットを確認してください:")
        for name in ['ana_reservations', 'hilton_reservations']:
            print(f"  → {os.path.join(TMPDIR, f'travel_{name}.png')}")
        return

    # ── フライト一覧 ──
    if flights:
        print(f"\n■ フライト予約 ({len(flights)}件)")
        for f in sorted(flights, key=lambda x: x['date']):
            print(f"  {f['date'].strftime('%Y/%m/%d')} {f['dep_time']} {f.get('flight_no','?')}  {f.get('raw','')[:50]}")
    else:
        print("\n■ フライト予約: 取得なし")

    # ── ホテル一覧 ──
    if hotels:
        print(f"\n■ ホテル予約 ({len(hotels)}件)")
        for h in sorted(hotels, key=lambda x: x['checkin']):
            nights = (h['checkout'] - h['checkin']).days
            print(f"  チェックイン {h['checkin']} → チェックアウト {h['checkout']} ({nights}泊) {h['hotel']}")
    else:
        print("\n■ ホテル予約: 取得なし")

    # ── 出張期間のサイボウズ予定 ──
    travel_dates = set()
    for f in flights:
        travel_dates.add(f['date'])
    for h in hotels:
        d = h['checkin']
        while d <= h['checkout']:
            travel_dates.add(d); d += timedelta(days=1)

    if travel_dates and events:
        travel_events = [e for e in events if e['date'] in travel_dates]
        print(f"\n■ 出張期間中のサイボウズ予定 ({len(travel_events)}件)")
        for e in sorted(travel_events, key=lambda x: (x['date_str'], x['time'] or '99:99')):
            print(f"  {e['date_str']} {e['time']:5s}  {e['title']}")

    # ── 競合チェック ──
    print(f"\n■ 競合・注意事項")
    issues = []

    # フライト ↔ サイボウズ
    for f in flights:
        dep_min = time_to_min(f['dep_time'])
        if dep_min < 0: continue
        for e in events:
            if e['date'] != f['date']: continue
            ev_min = time_to_min(e['time'])
            if ev_min < 0: continue
            diff = dep_min - ev_min
            if 0 < diff <= BUFFER_MIN:
                issues.append(f"⚠ {f['date']} フライト{f.get('flight_no','')} {f['dep_time']}出発の"
                              f"{BUFFER_MIN//60}時間以内にサイボウズ予定あり → 「{e['title']}」({e['time']})")
            elif -60 <= diff < 0:
                issues.append(f"🚨 {f['date']} フライト{f.get('flight_no','')} {f['dep_time']}出発と"
                              f"サイボウズ予定「{e['title']}」({e['time']})が重複の可能性")

    # ホテル ↔ フライト
    for h in hotels:
        arr_flights = [f for f in flights if f['date'] == h['checkin']]
        dep_flights = [f for f in flights if f['date'] == h['checkout']]
        if arr_flights:
            for f in arr_flights:
                issues.append(f"ℹ チェックイン日({h['checkin']})にフライト{f.get('flight_no','')} {f['dep_time']}あり → ホテル到着時間を確認")
        if dep_flights:
            for f in dep_flights:
                issues.append(f"ℹ チェックアウト日({h['checkout']})にフライト{f.get('flight_no','')} {f['dep_time']}あり → チェックアウト時刻({h['hotel']})を確認")

    # ホテル滞在期間のオフィス系予定
    for h in hotels:
        office_events = [e for e in events
                         if h['checkin'] <= e['date'] <= h['checkout']
                         and any(kw in e['title'] for kw in ['会議','ミーティング','打合','面談','オフィス','来社'])]
        for e in office_events:
            issues.append(f"⚠ ホテル滞在中({h['checkin']}〜{h['checkout']})にオフィス系予定あり → 「{e['title']}」({e['date_str']})")

    if issues:
        for iss in issues:
            print(f"  {iss}")
    else:
        print("  問題は検出されませんでした ✓")

    print("\n" + "="*60)

# ──────────────────────────────────
# メイン
# ──────────────────────────────────

if __name__ == '__main__':
    # 実行対象の選択: --ana / --hilton / (なし=両方)
    ana_only    = '--ana'    in sys.argv
    hilton_only = '--hilton' in sys.argv
    run_ana     = ana_only or (not hilton_only)
    run_hilton  = hilton_only or (not ana_only)

    target = 'ANA のみ' if ana_only else 'Hilton のみ' if hilton_only else 'ANA + Hilton'
    print(f"=== 旅程・サイボウズ整合性チェック [{target}] ===")
    print(f"実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    # 認証情報読み込み
    ana_no, ana_pw = read_cred('ANA_Login')    if run_ana    else ('', '')
    hlt_no, hlt_pw = read_cred('Hilton_Login') if run_hilton else ('', '')

    if run_ana and not ana_no:
        print("⚠ ANA認証情報が見つかりません (cmdkey /add:ANA_Login /user:会員番号 /pass:パスワード)")
    elif run_ana and not ana_pw:
        print("⚠ ANAパスワードが空です (cmdkey /delete:ANA_Login → /add:ANA_Login /user:... /pass:...)")
    if run_hilton and not hlt_no:
        print("⚠ Hilton認証情報が見つかりません (cmdkey /add:Hilton_Login /user:HonorsNo /pass:パスワード)")
    elif run_hilton and not hlt_pw:
        print("⚠ Hiltonパスワードが空です (cmdkey /delete:Hilton_Login → /add:Hilton_Login /user:... /pass:...)")

    # --ana-direct: ANAログイン不要で予約を直接指定（ログイン失敗時の回避策）
    # 使い方: python check_travel.py --ana-direct "NH219,2026-05-23,HND,KIX" "NH220,2026-05-24,KIX,HND"
    # フォーマット: "便名,日付,出発地,到着地"
    if '--ana-direct' in sys.argv:
        flights = []
        idx = sys.argv.index('--ana-direct')
        for arg in sys.argv[idx+1:]:
            if arg.startswith('--'):
                break
            try:
                parts = arg.split(',')
                if len(parts) >= 4:
                    flt_no, flt_date, orig, dest = parts[0], parts[1], parts[2], parts[3]
                    dep_time = parts[4] if len(parts) >= 5 else '00:00'
                    arr_time = parts[5] if len(parts) >= 6 else ''
                    dt = datetime.strptime(flt_date, '%Y-%m-%d').date()
                    flights.append({'type':'flight','flight_no':flt_no,'date':dt,'origin':orig,
                                    'destination':dest,'dep_time':dep_time,'arr_time':arr_time,
                                    'raw':f'{flt_no} {orig}→{dest} {dep_time}〜{arr_time}'})
                    print(f"  【直接指定】{flt_no} {flt_date} {orig}→{dest} {dep_time}〜{arr_time}")
            except Exception as e:
                print(f"  ⚠ --ana-direct引数エラー: {arg} ({e})")
        print(f"  ANA直接指定: {len(flights)}件")
    else:
        # 各サービスから予約取得
        flights = fetch_ana_bookings(ana_no, ana_pw) if (run_ana and ana_no and ana_pw) else []

    # --hilton-direct: スクリーンショット確認済みの予約を直接指定（bot検知回避）
    if '--hilton-direct' in sys.argv:
        hotels = [{'type':'hotel','hotel':'Hilton Tokyo',
                   'checkin':date(2026,5,23),'checkout':date(2026,5,24),
                   'confirmation':'3451900389','raw':'直接指定'}]
        print("  【直接指定】Hilton Tokyo 2026-05-23〜24 (Conf#3451900389)")
    else:
        hotels  = fetch_hilton_bookings(hlt_no, hlt_pw)  if (run_hilton and hlt_no and hlt_pw)  else []

    # 旅程期間を計算（前後30日も含めて広めに取得）
    all_dates = [f['date'] for f in flights] + \
                [h['checkin'] for h in hotels] + [h['checkout'] for h in hotels]

    if all_dates:
        start_dt = min(all_dates) - timedelta(days=1)
        end_dt   = max(all_dates) + timedelta(days=1)
    else:
        start_dt = date.today()
        end_dt   = date.today() + timedelta(days=60)

    # サイボウズ取得: --no-sync か --no-cybozu があればスキップ
    # （スクレイピングテスト時に不要な128件取得を防ぐ）
    skip_cybozu = '--no-sync' in sys.argv or '--no-cybozu' in sys.argv
    events = [] if skip_cybozu else fetch_cybozu_events(start_dt, end_dt)
    if skip_cybozu:
        print("\n【サイボウズ】スキップ（--no-sync / --no-cybozu）")

    # 整合性チェック・レポート出力
    check_and_report(flights, hotels, events)

    # サイボウズにフライト・出張予定を同期（--no-sync 引数で無効化可能）
    if '--no-sync' not in sys.argv:
        synced, skipped = sync_cybozu_schedule(flights, hotels)
    else:
        synced, skipped = [], []

    # JSON保存
    out = {
        'checked_at': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'flights': [{**f, 'date': f['date'].isoformat()} for f in flights],
        'hotels': [{**h, 'checkin': h['checkin'].isoformat(), 'checkout': h['checkout'].isoformat()} for h in hotels],
        'events': [{**e, 'date': e['date'].isoformat()} for e in events],
        'synced': synced,
        'skipped': skipped,
    }
    out_path = os.path.join(TMPDIR, 'travel_check_result.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\n結果JSON保存 → {out_path}")
