# -*- coding: utf-8 -*-
"""从 Windows 凭据管理器读取 git:https://github.com 的凭据（CredRead API）。

安全约定：token 只写入 .gh_token 与内存，屏幕仅显示脱敏片段。
"""
import ctypes
import os
import sys
from ctypes import wintypes

TARGET = "git:https://github.com"
CRED_TYPE_GENERIC = 1


class CREDENTIAL(ctypes.Structure):
    _fields_ = [
        ("Flags", wintypes.DWORD),
        ("Type", wintypes.DWORD),
        ("TargetName", wintypes.LPWSTR),
        ("Comment", wintypes.LPWSTR),
        ("LastWritten", wintypes.FILETIME),
        ("CredentialBlobSize", wintypes.DWORD),
        ("CredentialBlob", ctypes.POINTER(ctypes.c_byte)),
        ("Persist", wintypes.DWORD),
        ("AttributeCount", wintypes.DWORD),
        ("Attributes", ctypes.c_void_p),
        ("TargetAlias", wintypes.LPWSTR),
        ("UserName", wintypes.LPWSTR),
    ]


def read_cred(target):
    advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
    advapi32.CredReadW.argtypes = [
        wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD), ctypes.POINTER(ctypes.POINTER(CREDENTIAL))]
    advapi32.CredReadW.restype = wintypes.BOOL

    pcb = wintypes.DWORD()
    pcred = ctypes.POINTER(CREDENTIAL)()
    if not advapi32.CredReadW(target, CRED_TYPE_GENERIC, 0,
                              ctypes.byref(pcb), ctypes.byref(pcred)):
        err = ctypes.get_last_error()
        return None, "CredRead 失败 (WinError %d)" % err

    c = pcred.contents
    size = c.CredentialBlobSize
    blob = bytes(ctypes.string_at(c.CredentialBlob, size)) if size else b""
    user = c.UserName or ""
    ctypes.windll.advapi32.CredFree(pcred)
    return (user, blob), None


def mask(s, head=8, tail=4):
    if len(s) <= head + tail:
        return s[:2] + "****"
    return s[:head] + "****" + s[-tail:] + " (长度%d)" % len(s)


def main():
    r, err = read_cred(TARGET)
    if err:
        print("[失败] %s" % err)
        sys.exit(1)

    user, blob = r
    print("[凭据条目] %s" % TARGET)
    print("[UserName] %s" % (user or "(空)"))
    print("[Blob大小] %d 字节" % len(blob))

    # GCM 的 blob 可能是：纯 token / user:token / JSON
    text = None
    for enc in ("utf-8", "utf-16-le"):
        try:
            text = blob.decode(enc)
            if text and text.isprintable():
                break
            text = None
        except Exception:
            continue
    if text is None:
        print("[Blob] 无法解码为文本（可能是加密块），需要 GCM 解密")
        sys.exit(2)

    print("[Blob内容(脱敏)] %s" % mask(text.strip()))

    # 提取 token：优先找 ghp_/github_pat_/gho_ 前缀，否则按 user:pass 或 JSON 拆
    token = ""
    import re
    m = re.search(r"(ghp_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}|gho_[A-Za-z0-9]{30,})", text)
    if m:
        token = m.group(1)
    elif ":" in text:
        token = text.split(":", 1)[1].strip()
    elif text.strip().startswith("{"):
        try:
            import json
            d = json.loads(text)
            token = d.get("password") or d.get("token") or d.get("accessToken") or ""
        except Exception:
            pass
    else:
        token = text.strip()

    if not token:
        print("[失败] 未能从凭据中提取 token")
        sys.exit(3)

    print("[提取token(脱敏)] %s" % mask(token))

    # 写入 .gh_token（项目根，已被 .gitignore）
    root = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(root, ".gh_token")
    with open(p, "w", encoding="utf-8") as f:
        f.write(token + "\n")
    os.system('attrib +h "%s"' % p)
    print("[已写入] %s（隐藏属性，已列入 .gitignore）" % p)
    print("\n下一步：python push_via_api.py")


if __name__ == "__main__":
    main()
