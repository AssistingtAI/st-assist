# -*- coding: utf-8 -*-
"""网络恢复后自动推送（GitHub 访问不稳定时使用）。

用法：双击本文件，或在 CMD 中运行：
    python push_retry.py
    python push_retry.py --interval 15 --max 120     # 自定义间隔与次数

逻辑：
  1. 先用 git ls-remote 探测 GitHub 是否可达（快，不触发认证）
  2. 可达后执行 git push origin main --force-with-lease
  3. 未达通则每隔 N 秒重试，直到成功或达到次数上限

说明：--force-with-lease 会在推送前校验远端未被他人改动，比 --force 安全。
"""
import argparse
import subprocess
import sys
import time

REPO = r"C:\Users\xiaot\WorkBuddy\2026-05-16-task-2"


def sh(cmd, timeout=120):
    """执行命令，返回 (returncode, stdout+stderr)。"""
    try:
        r = subprocess.run(
            cmd, cwd=REPO, shell=True, capture_output=True,
            timeout=timeout,
        )
        out = (r.stdout or b"") + (r.stderr or b"")
        try:
            text = out.decode("utf-8")
        except UnicodeDecodeError:
            text = out.decode("gbk", "replace")
        return r.returncode, text
    except subprocess.TimeoutExpired:
        return -1, "(超时)"


def online():
    """探测 GitHub 是否可达。"""
    rc, out = sh("git ls-remote origin refs/heads/main", timeout=40)
    return rc == 0 and "refs/heads/main" in out, out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--interval", type=int, default=20, help="重试间隔秒数，默认 20")
    ap.add_argument("--max", type=int, default=90, help="最大尝试次数，默认 90")
    args = ap.parse_args()

    print("=" * 56)
    print("  GitHub 自动重试推送")
    print("  仓库: %s" % REPO)
    print("  策略: 每 %d 秒探测一次，最多 %d 次" % (args.interval, args.max))
    print("  按 Ctrl+C 可随时中止")
    print("=" * 56)

    for i in range(1, args.max + 1):
        stamp = time.strftime("%H:%M:%S")
        ok, out = online()
        if not ok:
            reason = out.strip().splitlines()
            reason = reason[-1][:70] if reason else "未知"
            print("[%s] 第 %d/%d 次: 网络未就绪 - %s"
                  % (stamp, i, args.max, reason))
            time.sleep(args.interval)
            continue

        print("[%s] 第 %d/%d 次: 网络已通，开始推送…" % (stamp, i, args.max))
        rc, out2 = sh("git push origin main --force-with-lease", timeout=300)
        if rc == 0:
            print()
            print("=" * 56)
            print("  推送成功！")
            print("=" * 56)
            print(out2.strip()[:1500])
            print()
            print("验证线上页面（1-2 分钟后生效）：")
            print("  https://assistingtai.github.io/st-assist/apply.html")
            return 0
        else:
            print("[%s] 推送失败（可能是认证问题）：" % stamp)
            print(out2.strip()[:800])
            print("  将在 %d 秒后重试。若是认证窗口，请手动登录 GitHub。" % args.interval)
            time.sleep(args.interval)

    print()
    print("已达最大尝试次数仍未成功。建议：")
    print("  1. 检查是否开启了代理软件（GitHub 在国内访问不稳定）")
    print("  2. 稍后重新运行本脚本")
    return 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n已手动中止。")
