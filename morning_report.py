# -*- coding: utf-8 -*-
"""
毎朝7時自動実行：メール＋Chatworkの動向をマイチャットに投稿
"""
import sys, io, subprocess, datetime, ctypes, urllib.request, urllib.parse

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

SCRIPTS_DIR  = r'C:\Users\nakko\Scripts'
CW_ROOM_ID   = 'YOUR_MYCHAT_ROOM_ID'   # マイチャット（中村 光）
PYTHON       = sys.executable

def get_cred(target):
    class _CRED(ctypes.Structure):
        _fields_ = [("Flags",ctypes.c_ulong),("Type",ctypes.c_ulong),
                    ("TargetName",ctypes.c_wchar_p),("Comment",ctypes.c_wchar_p),
                    ("LastWritten",ctypes.c_ulonglong),("CredentialBlobSize",ctypes.c_ulong),
                    ("CredentialBlob",ctypes.c_void_p),("Persist",ctypes.c_ulong),
                    ("AttributeCount",ctypes.c_ulong),("Attributes",ctypes.c_void_p),
                    ("TargetAlias",ctypes.c_wchar_p),("UserName",ctypes.c_wchar_p)]
    adv = ctypes.WinDLL("advapi32", use_last_error=True)
    ptr = ctypes.POINTER(_CRED)()
    if not adv.CredReadW(target, 1, 0, ctypes.byref(ptr)):
        return None
    raw = ctypes.string_at(ptr.contents.CredentialBlob, ptr.contents.CredentialBlobSize)
    token = raw.decode("utf-16-le").strip()
    adv.CredFree(ptr)
    return token

def post_chatwork(token, room_id, message):
    url  = f'https://api.chatwork.com/v2/rooms/{room_id}/messages'
    data = urllib.parse.urlencode({'body': message}).encode('utf-8')
    req  = urllib.request.Request(url, data=data, headers={'X-ChatWorkToken': token})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.status == 200
    except Exception as e:
        print(f"Chatwork投稿エラー: {e}")
        return False

def run_script(script_name):
    """スクリプトを実行して出力を返す。エラーの場合はエラーメッセージを返す"""
    try:
        result = subprocess.run(
            [PYTHON, '-X', 'utf8', f'{SCRIPTS_DIR}\\{script_name}'],
            capture_output=True, text=True, encoding='utf-8',
            timeout=120, errors='replace'
        )
        out = result.stdout.strip()
        if result.returncode != 0 and result.stderr:
            out += f"\n[エラー] {result.stderr[:200]}"
        return out or '（出力なし）'
    except subprocess.TimeoutExpired:
        return '（タイムアウト）'
    except Exception as e:
        return f'（実行失敗: {e}）'

def has_activity(text):
    """実質的な動きがあるか判定（空・エラーのみは除く）"""
    ignore = ['出力なし', 'タイムアウト', '実行失敗', '期間内: 0件', '未読合計: 0件']
    return any(line.strip() and not any(ig in line for ig in ignore)
               for line in text.splitlines() if line.strip())

def main():
    today = datetime.date.today()
    now   = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    token = get_cred('Chatwork_API_Token')
    if not token:
        print('Chatwork APIトークン取得失敗')
        return

    print(f'[{now}] 朝の自動レポート開始')

    # ── メールチェック ──
    print('Outlookメールチェック中...')
    mail_result = run_script('outlook_check.py')

    # ── Chatworkチェック ──
    print('Chatworkチェック中...')
    cw_result = run_script('chatwork_check.py')

    # ── 動きがあるかチェック ──
    mail_active = has_activity(mail_result)
    cw_active   = has_activity(cw_result)

    if not mail_active and not cw_active:
        print('メール・Chatworkともに動きなし → 投稿スキップ')
        return

    # ── Chatworkメッセージ作成 ──
    parts = [f'[info][title]📋 朝の自動レポート {today.strftime("%m/%d(%a)")}[/title]']

    if mail_active:
        # 要注意メールと未読の概要だけ抜粋（全文は長すぎるため）
        mail_lines = mail_result.splitlines()
        summary_lines = []
        for line in mail_lines:
            if any(kw in line for kw in ['⚠', '未読合計', '期間内', '要注意', '件']):
                summary_lines.append(line.rstrip())
            if len(summary_lines) >= 15:
                break
        parts.append('[メール]\n' + '\n'.join(summary_lines) if summary_lines else '[メール]\n' + mail_result[:300])
    else:
        parts.append('[メール]\n動きなし')

    if cw_active:
        cw_lines = cw_result.splitlines()
        summary_lines = []
        for line in cw_lines:
            if any(kw in line for kw in ['⚠', '未読', '件', '要注意', '●', '【']):
                summary_lines.append(line.rstrip())
            if len(summary_lines) >= 15:
                break
        parts.append('[Chatwork]\n' + '\n'.join(summary_lines) if summary_lines else '[Chatwork]\n' + cw_result[:300])
    else:
        parts.append('[Chatwork]\n動きなし')

    parts.append('[/info]')
    message = '\n\n'.join(parts)

    # ── 投稿 ──
    ok = post_chatwork(token, CW_ROOM_ID, message)
    print(f'Chatwork投稿: {"成功" if ok else "失敗"}')

if __name__ == '__main__':
    main()
