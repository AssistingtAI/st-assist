#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Git Data API 链式推送：把本地一串 commit 精确复刻到远端（内容取自 git 对象，SHA 应与本地一致）。

用法: python _push_chain_via_api.py <TOKEN> <REPO> <REMOTE_BASE_SHA> <c1> <c2> ...
"""
import base64
import subprocess
import sys

import requests

token, repo, remote_base, commits = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4:]
API = 'https://api.github.com'
H = {'Authorization': f'token {token}', 'Accept': 'application/vnd.github+json'}


def req(method, path, **kw):
    r = requests.request(method, API + path, headers=H, timeout=180, **kw)
    if r.status_code >= 300:
        raise SystemExit(f'!! {method} {path} -> {r.status_code}: {r.text[:400]}')
    return r.json()


def git(*args, binary=False):
    r = subprocess.run(['git'] + list(args), capture_output=True)
    if r.returncode != 0:
        raise SystemExit(f'git {" ".join(args)} failed: {r.stderr[:300]}')
    return r.stdout if binary else r.stdout.decode('utf-8', 'replace')


cur_remote = req('GET', f'/repos/{repo}/git/commits/{remote_base}')['sha']
cur_tree = req('GET', f'/repos/{repo}/git/commits/{remote_base}')['tree']['sha']
print(f'remote base {cur_remote[:8]} tree {cur_tree[:8]}')

for lc in commits:
    full = git('rev-parse', lc).strip()
    m = git('log', '-1', '--format=%an%x00%ae%x00%aI%x00%cn%x00%ce%x00%cI%x00%B', full).split('\x00')
    an, ae, ad, cn, ce, cd, msg = m[0], m[1], m[2], m[3], m[4], m[5], m[6].rstrip('\n')
    lparent = git('rev-parse', f'{full}^').strip()
    ltree = git('rev-parse', f'{full}^{{tree}}').strip()
    files = [p for p in git('diff', '--name-only', '-z', lparent, full, binary=True).decode('utf-8').split('\0') if p]
    print(f'\n== {full[:8]} {msg.splitlines()[0][:50]} ({len(files)} files)')

    entries = []
    for f in files:
        content = git('cat-file', 'blob', f'{full}:{f}', binary=True)
        blob = req('POST', f'/repos/{repo}/git/blobs',
                   json={'content': base64.b64encode(content).decode(), 'encoding': 'base64'})
        lsha = git('rev-parse', f'{full}:{f}').strip()
        if blob['sha'] != lsha:
            print(f'   X blob mismatch {f}: local={lsha[:8]} api={blob["sha"][:8]}')
        entries.append({'path': f, 'mode': '100644', 'type': 'blob', 'sha': blob['sha']})

    t = req('POST', f'/repos/{repo}/git/trees', json={'base_tree': cur_tree, 'tree': entries})
    tree_ok = t['sha'] == ltree
    print(f'   tree {t["sha"][:8]} local={ltree[:8]} match={tree_ok}')

    c = req('POST', f'/repos/{repo}/git/commits', json={
        'message': msg, 'tree': t['sha'], 'parents': [cur_remote],
        'author': {'name': an, 'email': ae, 'date': ad},
        'committer': {'name': cn, 'email': ce, 'date': cd},
    })
    print(f'   commit {c["sha"][:8]} local={full[:8]} match={c["sha"]==full}')
    cur_remote, cur_tree = c['sha'], t['sha']

ref = req('PATCH', f'/repos/{repo}/git/refs/heads/main', json={'sha': cur_remote, 'force': True})
print('\nref ->', ref['object']['sha'])
print('DONE')
