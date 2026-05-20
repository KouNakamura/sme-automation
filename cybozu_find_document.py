"""
サイボウズOffice ファイル管理から指定キーワードのファイルを検索・表示するスクリプト
usage: python -X utf8 cybozu_find_document.py "退職金規程"
"""
import sys, io, os, time, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

import ctypes, ctypes.wintypes
from pathlib import Path

EDGE_DRIVER  = os.path.join(os.environ["TEMP"], "edgedriver_mf", "msedgedriver.exe")
EDGE_PROFILE = os.path.join(os.environ["USERPROFILE"], ".config", "cybozu_profile")
BASE         = "https://YOUR_COMPANY.cybozu.com"

KEYWORD = sys.argv[1] if len(sys.argv) > 1 else "退職金規程"


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

def make_driver():
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.edge.options import Options
    Path(EDGE_PROFILE).mkdir(parents=True, exist_ok=True)
    opts = Options()
    opts.add_argument(f"--user-data-dir={EDGE_PROFILE}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1280,900")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    return webdriver.Edge(service=Service(executable_path=EDGE_DRIVER), options=opts)

def main():
    from selenium.webdriver.common.by import By

    driver = make_driver()
    try:
        print(f'サイボウズ接続中... キーワード: 「{KEYWORD}」')
        driver.get(BASE + "/")
        time.sleep(4)
        url = driver.current_url
        print(f'  URL: {url[:80]}')

        if BASE not in url:
            print('  ⚠ ログインが必要です（cybozu_schedule.py をブラウザ付きで先に実行してください）')
            return

        # ファイル管理を検索（サイボウズOfficeのファイル管理ページ）
        file_mgmt_urls = [
            BASE + "/o/ag.cgi?page=FilesFolderTopPc",
            BASE + "/o/ag.cgi?page=FilesTop",
        ]

        found_links = []
        for fm_url in file_mgmt_urls:
            driver.get(fm_url)
            time.sleep(3)
            body_text = driver.find_element(By.TAG_NAME, 'body').text
            if KEYWORD in body_text:
                print(f'  ✓ 「{KEYWORD}」がファイル管理に見つかりました')
                # リンクを収集
                for a in driver.find_elements(By.TAG_NAME, 'a'):
                    txt  = a.text.strip()
                    href = a.get_attribute('href') or ''
                    if KEYWORD in txt or KEYWORD in href:
                        found_links.append((txt, href))
                break
            else:
                print(f'  {fm_url[-40:]} → キーワードなし')

        # 検索機能を試みる（サイボウズOfficeの全文検索）
        if not found_links:
            print('\n全文検索を試みます...')
            search_urls = [
                BASE + f"/o/ag.cgi?page=Search&Keywords={KEYWORD}",
                BASE + f"/o/search?q={KEYWORD}",
            ]
            for su in search_urls:
                driver.get(su)
                time.sleep(3)
                body = driver.find_element(By.TAG_NAME, 'body').text
                if KEYWORD in body:
                    print(f'  ✓ 検索結果に「{KEYWORD}」が見つかりました')
                    # 結果リンクを収集
                    for a in driver.find_elements(By.TAG_NAME, 'a'):
                        txt = a.text.strip()
                        href = a.get_attribute('href') or ''
                        if txt and (KEYWORD in txt or KEYWORD in href or 'Files' in href):
                            found_links.append((txt[:60], href[:120]))
                    break
                else:
                    print(f'  {su[-60:]} → キーワードなし')

        # フォルダ一覧を全走査して探す
        if not found_links:
            print('\nフォルダ一覧を走査中...')
            driver.get(BASE + "/o/ag.cgi?page=FilesFolderTopPc")
            time.sleep(3)
            # 全フォルダリンクを取得
            folder_links = []
            for a in driver.find_elements(By.TAG_NAME, 'a'):
                href = a.get_attribute('href') or ''
                if 'FilesFolder' in href or 'Files' in href:
                    folder_links.append(href)

            print(f'  フォルダ数: {len(folder_links)}')
            for fl in folder_links[:20]:
                driver.get(fl)
                time.sleep(2)
                body = driver.find_element(By.TAG_NAME, 'body').text
                if KEYWORD in body:
                    print(f'  ✓ 発見: {fl[-80:]}')
                    for a in driver.find_elements(By.TAG_NAME, 'a'):
                        txt = a.text.strip()
                        href = a.get_attribute('href') or ''
                        if KEYWORD in txt:
                            found_links.append((txt, href))

        if found_links:
            print(f'\n=== 発見したリンク ===')
            for txt, href in found_links[:10]:
                print(f'  [{txt}]')
                print(f'  → {href[:120]}')

            # 最初のファイルを開いてテキスト抽出
            if found_links:
                print(f'\n=== ファイル内容確認 ===')
                driver.get(found_links[0][1])
                time.sleep(3)
                content = driver.find_element(By.TAG_NAME, 'body').text
                # 退職金関連部分を抽出
                lines = content.split('\n')
                in_relevant = False
                for line in lines:
                    if any(kw in line for kw in ['退職', '勤続', '支給', '算定', '基準', '月数', '倍率']):
                        print(f'  {line.strip()}')
        else:
            print(f'\n「{KEYWORD}」はサイボウズのファイル管理で見つかりませんでした')
            print('  掲示板・カスタムアプリも確認します...')
            # 掲示板検索
            driver.get(BASE + f"/o/ag.cgi?page=BulletinBoardSearch&Keywords={KEYWORD}")
            time.sleep(3)
            body = driver.find_element(By.TAG_NAME, 'body').text
            if KEYWORD in body:
                print(f'  掲示板に見つかりました')
                for line in body.split('\n'):
                    if KEYWORD in line: print(f'  {line.strip()[:100]}')
            else:
                print('  掲示板にも見つかりませんでした')
                # カスタムアプリ（DB）も確認
                driver.get(BASE + f"/o/ag.cgi?page=DBRecordSearch&DB=&Keywords={KEYWORD}")
                time.sleep(3)
                body = driver.find_element(By.TAG_NAME, 'body').text
                if KEYWORD in body:
                    print(f'  カスタムアプリに見つかりました')

    finally:
        driver.quit()

if __name__ == '__main__':
    main()
