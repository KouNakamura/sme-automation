"""資格情報をCredential Managerに安全に保存するスクリプト"""
import ctypes, ctypes.wintypes, getpass, sys

target = sys.argv[1] if len(sys.argv) > 1 else input("Target name: ")
user   = sys.argv[2] if len(sys.argv) > 2 else input("Username: ")
pw     = getpass.getpass(f"Password for {user}: ")

class CREDENTIAL(ctypes.Structure):
    _fields_ = [
        ("Flags",              ctypes.wintypes.DWORD),
        ("Type",               ctypes.wintypes.DWORD),
        ("TargetName",         ctypes.wintypes.LPWSTR),
        ("Comment",            ctypes.wintypes.LPWSTR),
        ("LastWritten",        ctypes.wintypes.FILETIME),
        ("CredentialBlobSize", ctypes.wintypes.DWORD),
        ("CredentialBlob",     ctypes.POINTER(ctypes.c_byte)),
        ("Persist",            ctypes.wintypes.DWORD),
        ("AttributeCount",     ctypes.wintypes.DWORD),
        ("Attributes",         ctypes.c_void_p),
        ("TargetAlias",        ctypes.wintypes.LPWSTR),
        ("UserName",           ctypes.wintypes.LPWSTR),
    ]

b = pw.encode("utf-16-le")
buf = (ctypes.c_byte * len(b))(*b)
cred = CREDENTIAL(Flags=0, Type=1, TargetName=target, CredentialBlobSize=len(b),
                  CredentialBlob=buf, Persist=2, UserName=user)

ok = ctypes.windll.advapi32.CredWriteW(ctypes.byref(cred), 0)
print("保存完了!" if ok else f"保存失敗 (error={ctypes.GetLastError()})")
