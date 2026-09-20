#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GitHub Data API 链式推送（用于 github.com:443 不通、但 api.github.com 可达时）。

背景：本项目部署在国内网络，github.com（git push 走这个域名）经常连不上，
     但 api.github.com 通常可达。本脚本用 REST API 把本地提交精确复刻到远端，
      生成的 commit SHA 与本地一致，效果等价于 git push。

用法：
    python push_via_api.py <TOKEN>
    python push_via_api.py            # 自动从 .gh_token 文件或 GH_TOKEN 环境变量读取

    # 从指定基线推送（默认自动探测远端当前 main）
    python push_via_api.py <TOKEN> --base <SHA>

    # 只预览不推送
    python push_via_api.py <TOKEN> --dry-run

TOKEN 获取：GitHub → Settings → Developer settings → Personal access tokens
           → Tokens (classic) → 勾选 repo 权限。仅保存在本机，可用 .gh_token 文件。

安全：token 只在本机内存与可选 .gh_token 中使用，绝不要提交入库（已在 .gitignore）。
"""
import argparse
import base64
import os
import subprocess
import sys
import time

try:
    import requests
except ImportError:
    print("[错误] 需要 requests：pip install requests")
    sys.exit(1)

REPO = "AssistingtAI/st-assist"
BRANCH = "main"
API = "https://api.github.com"
ROOT = os.path.dirname(os.path.abspath(__file__))


def build_headers(token):
    return {
        "Authorization": "token %s" % token,
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


class Api:
    def __init__(self, token, dry=False):
        self.h = build_headers(token)
        self.dry = dry
        # 直连 api.github.com：忽略环境代理变量（HTTP_PROXY 等指向的本机代理
        # 对 api.github.com 时通时断，直连实测稳定）。api.github.com 与
        # github.com 是不同域名，国内网络下后者常被阻断、前者通常可达。
        self.s = requests.Session()
        self.s.trust_env = False
        self.s.headers.update(self.h)

    def req(self, method, path, **kw):
        url = API + path
        if self.dry and method != "GET":
            print("   [dry-run] %s %s" % (method, path))
            # dry-run 时返回可用的假响应，让后续逻辑能走通
            if path.endswith("/git/blobs"):
                return {"sha": "0" * 40}
            if path.endswith("/git/trees"):
                return {"sha": "0" * 40}
            if path.endswith("/git/commits"):
                return {"sha": "0" * 40}
            return {"object": {"sha": "0" * 40}}
        for attempt in range(3):
            try:
                r = self.s.request(method, url, timeout=60, **kw)
            except requests.RequestException as e:
                if attempt == 2:
                    raise SystemExit("[网络错误] %s %s -> %s" % (method, path, e))
                print("   网络抖动，重试 %d/3…" % (attempt + 1))
                time.sleep(2)
                continue
            if r.status_code >= 300:
                msg = r.text[:500]
                if r.status_code == 401:
                    raise SystemExit("[认证失败 401] token 无效或已过期。\n%s" % msg)
                if r.status_code == 403:
                    raise SystemExit("[权限不足 403] 请确认 token 勾选了 repo 权限。\n%s" % msg)
                raise SystemExit("[API 错误] %s %s -> %d: %s" % (method, path, r.status_code, msg))
            return r.json()
        raise SystemExit("重试耗尽")


def git(*args, binary=False):
    r = subprocess.run(["git"] + list(args), cwd=ROOT, capture_output=True)
    if r.returncode != 0:
        raise SystemExit("git %s 失败: %s" % (" ".join(args),
                                             r.stderr.decode("utf-8", "replace")[:300]))
    return r.stdout if binary else r.stdout.decode("utf-8", "replace")


def resolve_token(cli_token):
    if cli_token:
        return cli_token.strip()
    env = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if env:
        return env.strip()
    p = os.path.join(ROOT, ".gh_token")
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return f.read().strip()
    print("[错误] 未提供 token。三种方式任选：")
    print("  1) python push_via_api.py <TOKEN>")
    print("  2) 设置环境变量 GH_TOKEN")
    print("  3) 把 token 写入 %s （一行纯文本，已在 .gitignore）" % p)
    sys.exit(1)


def main():
    ap = argparse.ArgumentParser(description="GitHub Data API 链式推送")
    ap.add_argument("token", nargs="?", help="GitHub PAT（可省略，改用环境变量或 .gh_token）")
    ap.add_argument("--repo", default=REPO, help="仓库，默认 %s" % REPO)
    ap.add_argument("--branch", default=BRANCH, help="分支，默认 %s" % BRANCH)
    ap.add_argument("--base", help="远端基线 SHA，默认自动探测")
    ap.add_argument("--dry-run", action="store_true", help="只预览不实际推送")
    args = ap.parse_args()

    token = resolve_token(args.token)
    api = Api(token, dry=args.dry_run)
    repo, branch = args.repo, args.branch

    # ---- 1) 探测认证与远端基线 ----
    me = api.req("GET", "/user")
    print("认证成功: %s" % me.get("login"))
    ref = api.req("GET", "/repos/%s/git/ref/heads/%s" % (repo, branch))
    remote_sha = ref["object"]["sha"]
    remote_tree = api.req("GET", "/repos/%s/git/commits/%s" % (repo, remote_sha))["tree"]["sha"]
    print("远端 %s/%s = %s" % (repo, branch, remote_sha[:8]))

    base = args.base or remote_sha

    # ---- 2) 收集待推送提交（旧→新）----
    # 注意：分叉场景下 merge-base --is-ancestor 会返回非 0，这属正常（远端有本地没有的提交），
    #       不能用 git() 包装（它会当成失败抛错），这里单独判断。
    anc = subprocess.run(["git", "merge-base", "--is-ancestor", base, "HEAD"],
                         cwd=ROOT, capture_output=True)
    fork = None
    if anc.returncode != 0:
        # 分叉：找到共同祖先，从那里开始复刻，保证 parent 链与本地一致（SHA 不变）
        fork = git("merge-base", "HEAD", base).strip()
        print("检测到分叉：")
        print("  远端 main = %s" % remote_sha[:8])
        print("  共同祖先  = %s" % fork[:8])
        print("  → 从共同祖先开始复刻本地提交，最终远端历史将与本地完全一致")
        base = fork
        remote_sha = fork

    rng = "%s..HEAD" % base
    commits = [c for c in git("rev-list", "--reverse", rng).split("\n") if c.strip()]
    if not commits:
        print("\n远端已是最新，无需推送。")
        return 0

    # 检查本地是否有 base 之后的分叉（远端领先本地）
    behind = [c for c in git("rev-list", "--reverse", "HEAD..%s" % base).split("\n") if c.strip()]
    print("待推送 %d 个提交%s" % (len(commits),
                              "（远端另有 %d 个提交，将被覆盖）" % len(behind) if behind else ""))

    # ---- 3) 链式复刻 ----
    # 确保起点对象在远端存在：分叉点（fork）可能只在本地有，远端没有对应 commit。
    # 若 GET 失败，就用本地对象在远端"补建"该 commit（parent 指向它自己的 parent）。
    cur_remote = base
    try:
        cur_tree = api.req("GET", "/repos/%s/git/commits/%s" % (repo, base))["tree"]["sha"]
    except SystemExit:
        print("起点 %s 在远端不存在，先补建…" % base[:8])
        p = git("rev-parse", "%s^" % base).strip()
        pt = git("rev-parse", "%s^{tree}" % base).strip()
        # 递归补建父链（最多回溯到远端已知点）
        chain = []
        cur = base
        while True:
            try:
                api.req("GET", "/repos/%s/git/commits/%s" % (repo, cur))
                break
            except SystemExit:
                chain.append(cur)
                cur = git("rev-parse", "%s^" % cur).strip()
        print("需补建 %d 个对象" % len(chain))
        for sha in reversed(chain):
            pm = git("log", "-1", "--format=%an%x00%ae%x00%aI%x00%cn%x00%ce%x00%cI%x00%B", sha).split("\x00")
            pl = git("rev-parse", "%s^" % sha).strip()
            ptree = git("rev-parse", "%s^{tree}" % sha).strip()
            c = api.req("POST", "/repos/%s/git/commits" % repo, json={
                "message": pm[6].rstrip("\n"), "tree": ptree, "parents": [pl],
                "author": {"name": pm[0], "email": pm[1], "date": pm[2]},
                "committer": {"name": pm[3], "email": pm[4], "date": pm[5]},
            })
            print("  补建 %s -> %s" % (sha[:8], c["sha"][:8]))
        cur_tree = git("rev-parse", "%s^{tree}" % base).strip()

    for i, lc in enumerate(commits, 1):
        full = git("rev-parse", lc).strip()
        m = git("log", "-1", "--format=%an%x00%ae%x00%aI%x00%cn%x00%ce%x00%cI%x00%B", full).split("\x00")
        an, ae, ad, cn, ce, cd, msg = m[0], m[1], m[2], m[3], m[4], m[5], m[6].rstrip("\n")
        lparent = git("rev-parse", "%s^" % full).strip()
        ltree = git("rev-parse", "%s^{tree}" % full).strip()

        raw = git("diff", "--name-status", "--no-renames", "-z", lparent, full, binary=True)
        toks = [t for t in raw.decode("utf-8").split("\0") if t]
        files = []
        k = 0
        while k + 1 < len(toks):
            files.append((toks[k].strip(), toks[k + 1]))
            k += 2
        title = msg.splitlines()[0][:48] if msg.splitlines() else ""
        print("\n[%d/%d] %s %s (%d 文件)" % (i, len(commits), full[:8], title, len(files)))

        entries = []
        for st, f in files:
            if st.startswith("D"):
                # 删除型变更：路径在目标树已不存在，读 blob 必然失败。
                # GitHub 建树 API 的语义：sha=null 的条目 = 从 base_tree 中删除该路径。
                entries.append({"path": f, "mode": "100644", "type": "blob", "sha": None})
                print("   删  %s" % f)
                continue
            content = git("cat-file", "blob", "%s:%s" % (full, f), binary=True)
            blob = api.req("POST", "/repos/%s/git/blobs" % repo,
                           json={"content": base64.b64encode(content).decode(), "encoding": "base64"})
            lsha = git("rev-parse", "%s:%s" % (full, f)).strip()
            if blob["sha"] != lsha and not args.dry_run:
                print("   ! blob 不一致 %s: local=%s api=%s" % (f, lsha[:8], blob["sha"][:8]))
            entries.append({"path": f, "mode": "100644", "type": "blob", "sha": blob["sha"]})

        t = api.req("POST", "/repos/%s/git/trees" % repo,
                    json={"base_tree": cur_tree, "tree": entries})
        if not args.dry_run:
            print("   tree  %s  与本地%s" % (t["sha"][:8], "一致" if t["sha"] == ltree else "不一致"))

        c = api.req("POST", "/repos/%s/git/commits" % repo, json={
            "message": msg, "tree": t["sha"], "parents": [cur_remote],
            "author": {"name": an, "email": ae, "date": ad},
            "committer": {"name": cn, "email": ce, "date": cd},
        })
        if not args.dry_run:
            print("   commit %s  与本地%s" % (c["sha"][:8], "一致" if c["sha"] == full else "不一致"))
        cur_remote, cur_tree = c["sha"], t["sha"]

    # ---- 4) 更新分支引用 ----
    if args.dry_run:
        print("\n[dry-run] 未实际更新分支。")
        return 0

    out = api.req("PATCH", "/repos/%s/git/refs/heads/%s" % (repo, branch),
                  json={"sha": cur_remote, "force": True})
    print("\n分支 %s 已更新 -> %s" % (branch, out["object"]["sha"][:8]))
    print("DONE")
    print("\n本地 HEAD: %s" % git("rev-parse", "HEAD").strip()[:8])
    if out["object"]["sha"] == git("rev-parse", "HEAD").strip():
        print("远端与本地完全一致，推送成功。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
