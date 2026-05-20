"""
2025年以前のオープントラブルをクローズ（状況→対応完了）
対象: rid=113（久保田/2018）, rid=114（中田/2018）, rid=115（澤/2021）
"""
import sys, io, os, time, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)
sys.path.insert(0, os.path.dirname(__file__))

from check_travel import _cybozu_login_driver, CYBOZU_BASE
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup

BASE = CYBOZU_BASE
DID  = YOUR_DB_ID

TARGETS = [
    {'rid': 115, 'person': '澤 良樹',    'year': 2021},
    {'rid': 114, 'person': '中田 知佐',  'year': 2018},
    {'rid': 113, 'person': '久保田 英幸','year': 2018},
]

def close_record(driver, rid, person):
    # 詳細ページから「編集する」リンクを経由してロックを正しく取得
    driver.get(f'{BASE}/o/ag.cgi?page=DBRecord&did={DID}&rid={rid}')
    time.sleep(2)

    try:
        edit_link = driver.find_element(By.XPATH,
            "//a[contains(@href,'DBForm') and (contains(text(),'編集') or contains(@href,'rw='))]")
        edit_link.click()
        time.sleep(2)
    except Exception as e:
        print(f'    「編集する」リンクが見つからず直接アクセス: {e}')
        driver.get(f'{BASE}/o/ag.cgi?page=DBForm&did={DID}&rid={rid}&rw=in')
        time.sleep(2)

    # 状況フィールドを「対応完了」に変更
    try:
        sel_el = driver.find_element(By.CSS_SELECTOR, "select[name='17']")
        before = Select(sel_el).first_selected_option.text
        Select(sel_el).select_by_visible_text('対応完了')
        print(f'    状況: "{before}" → "対応完了"')
    except Exception as e:
        print(f'    ⚠ 状況フィールド操作失敗: {e}')
        return False

    # 本文フィールドがedit modeの場合は編集ラジオを選択しない（デフォルトのまま）
    # EditMode20のradio[0]（編集しない）を選択してテキストを保持
    try:
        radio_no_edit = driver.find_element(By.CSS_SELECTOR, "input[name='EditMode20'][value='0']")
        if not radio_no_edit.is_selected():
            radio_no_edit.click()
    except:
        pass

    # 送信
    try:
        btn = driver.find_element(By.CSS_SELECTOR, "input[name='Submit']")
        btn.click()
        time.sleep(3)
    except Exception as e:
        print(f'    ⚠ 送信失敗: {e}')
        return False

    # 結果確認
    url_after = driver.current_url
    src = driver.page_source
    soup = BeautifulSoup(src, 'html.parser')

    # エラーメッセージを探す
    err_texts = []
    for el in soup.find_all(class_=re.compile(r'error|alert|caution', re.I)):
        t = el.get_text(strip=True)
        if t: err_texts.append(t[:80])

    if 'DBRecord' in url_after or 'DBView' in url_after:
        print(f'    ✓ 保存成功: {person}')
        return True
    elif err_texts:
        print(f'    ⚠ エラー: {err_texts}')
    else:
        # DBFormのまま→ページ内容を確認
        body_text = soup.get_text(separator=' ', strip=True)
        # エラー文言を探す
        for kw in ['エラー', '必須', 'error', '入力してください']:
            idx = body_text.lower().find(kw.lower())
            if idx >= 0:
                snippet = body_text[max(0,idx-20):idx+60]
                print(f'    ⚠ "{kw}" 検出: ...{snippet}...')

        print(f'    URL変化なし: {url_after}')
        print(f'    → JavaScriptで強制送信を試みます')

        # JSで直接フォームをsubmit
        try:
            driver.execute_script("""
                var frms = document.querySelectorAll('form');
                for(var f of frms){
                    var s = f.querySelector("select[name='17']");
                    if(s){
                        for(var o of s.options){ if(o.text=='対応完了') o.selected=true; }
                        var btn = f.querySelector("input[name='Submit']");
                        if(btn) btn.click();
                        break;
                    }
                }
            """)
            time.sleep(3)
            url_after2 = driver.current_url
            if 'DBRecord' in url_after2 or 'DBView' in url_after2:
                print(f'    ✓ JS送信成功: {person}')
                return True
            else:
                print(f'    ✗ JS送信後もURL変化なし: {url_after2}')
                return False
        except Exception as ex:
            print(f'    ✗ JS送信失敗: {ex}')
            return False

    return False

def main():
    print('=== 2025年以前のトラブルクローズ処理 ===')
    for t in TARGETS:
        print(f'  rid={t["rid"]}  {t["person"]}  ({t["year"]}年登録)')
    print()

    driver = _cybozu_login_driver()
    if not driver:
        print('ログイン失敗'); return

    success, failed = [], []
    try:
        for t in TARGETS:
            print(f'処理中: {t["person"]} (rid={t["rid"]}, {t["year"]}年登録)')
            ok = close_record(driver, t['rid'], t['person'])
            (success if ok else failed).append(t)
            print()
    finally:
        try: driver.quit()
        except: pass

    print(f'=== 結果 ===')
    print(f'  成功: {len(success)}件 → {[t["person"] for t in success]}')
    if failed:
        print(f'  失敗: {len(failed)}件 → {[t["person"] for t in failed]}')

if __name__ == '__main__':
    main()
