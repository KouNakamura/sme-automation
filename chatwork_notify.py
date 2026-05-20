"""Chatwork へメッセージを投稿するヘルパー
Usage: python -X utf8 chatwork_notify.py <message>
"""
import sys, io, ctypes, urllib.request, urllib.parse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

CW_ROOM_ID = 'YOUR_MYCHAT_ROOM_ID'  # マイチャット（中村 光）

def _get_cred(target_name):
    class _CREDENTIAL(ctypes.Structure):
        _fields_ = [
            ("Flags",              ctypes.c_ulong),
            ("Type",               ctypes.c_ulong),
            ("TargetName",         ctypes.c_wchar_p),
            ("Comment",            ctypes.c_wchar_p),
            ("LastWritten",        ctypes.c_ulonglong),
            ("CredentialBlobSize", ctypes.c_ulong),
            ("CredentialBlob",     ctypes.c_void_p),  # c_char_pだとNULLで打ち切られるためc_void_p
            ("Persist",            ctypes.c_ulong),
            ("AttributeCount",     ctypes.c_ulong),
            ("Attributes",         ctypes.c_void_p),
            ("TargetAlias",        ctypes.c_wchar_p),
            ("UserName",           ctypes.c_wchar_p),
        ]
    advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
    ptr = ctypes.POINTER(_CREDENTIAL)()
    if not advapi32.CredReadW(target_name, 1, 0, ctypes.byref(ptr)):
        return None
    cred = ptr.contents
    raw = ctypes.string_at(cred.CredentialBlob, cred.CredentialBlobSize)
    blob = raw.decode("utf-16-le").strip()
    advapi32.CredFree(ptr)
    return blob

def post_chatwork(token, room_id, message):
    url  = f'https://api.chatwork.com/v2/rooms/{room_id}/messages'
    data = urllib.parse.urlencode({'body': message}).encode('utf-8')
    req  = urllib.request.Request(url, data=data, headers={'X-ChatWorkToken': token})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status

if __name__ == '__main__':
    message = sys.argv[1] if len(sys.argv) > 1 else '通知'
    token   = _get_cred('Chatwork_API_Token')
    if not token:
        print('ERROR: Chatwork_API_Token が資格情報マネージャーに見つかりません')
        sys.exit(1)
    try:
        status = post_chatwork(token, CW_ROOM_ID, message)
        print(f'Chatwork投稿完了 (HTTP {status})')
    except Exception as e:
        print(f'Chatwork投稿失敗: {e}')
        sys.exit(1)
