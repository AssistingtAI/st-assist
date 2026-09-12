# -*- coding: utf-8 -*-
"""
convert_reports_html.py — 报告库 docx → 只读HTML研究档案（合规版）
- 输入: reports_baokeng_v2/*.docx（2026-09-01 十三维历史档案）
- 输出: reports/<code>.html × N + reports_index.json
- 合规处理:
  1) 每页固定"历史研究档案"横幅（口径差异说明 + 不构成投资建议 + noindex）
  2) 安全词归一: 摸鱼榜→壳市值观察 / 摸鱼指数→壳市值指数 / 摸鱼池→观察池 /
     V2保壳分→V2评分 / 保壳分→V2评分（品牌名"ST摸鱼风云"保留，属报告体系官方名）
  3) 过滤模板残留段（目录域刷新提示等）
  4) 页脚免责声明
- 分数/日期等数字一律原样保留，不篡改历史档案
"""
import os, re, json, glob, html as H
import docx
from docx.document import Document as _Doc
from docx.table import Table
from docx.text.paragraph import Paragraph

SRC_DIR = 'reports_baokeng_v2'
OUT_DIR = 'reports'
INDEX_FILE = 'reports_index.json'

PAGE_TPL = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>{title} · 历史研究档案</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;background:#f4f7f2;color:#2c3e2d;line-height:1.85}}
.banner{{background:#fff8e6;border-bottom:1px solid #f0ddb0;padding:14px 18px;font-size:12.5px;color:#8a6d1f}}
.banner b{{color:#7a5c10}}
.wrap{{max-width:820px;margin:0 auto;padding:28px 20px 60px}}
.doc-card{{background:#fff;border:1px solid #e3ece0;border-radius:12px;padding:34px 36px;box-shadow:0 2px 10px rgba(59,109,17,.06)}}
h1.doc-title{{font-size:22px;color:#3B6D11;margin-bottom:4px}}
.doc-sub{{font-size:13px;color:#888;margin-bottom:22px;padding-bottom:14px;border-bottom:2px solid #eef4ea}}
h1{{font-size:19px;color:#3B6D11;margin:26px 0 10px}}
h2{{font-size:16px;color:#4d7d22;margin:20px 0 8px}}
p{{margin:8px 0;font-size:14.5px}}
table{{border-collapse:collapse;width:100%;margin:12px 0;font-size:13.5px}}
th{{background:#3B6D11;color:#fff;padding:8px 10px;text-align:left;font-weight:600}}
td{{padding:7px 10px;border-bottom:1px solid #eef2ea}}
tr:nth-child(even) td{{background:#f7faf4}}
.footer{{max-width:820px;margin:0 auto;padding:0 20px 40px;font-size:12px;color:#8a9a85}}
.footer .box{{background:#fff;border:1px dashed #cfdcc8;border-radius:10px;padding:14px 16px}}
.back{{display:inline-block;margin-bottom:16px;font-size:13px;color:#3B6D11;text-decoration:none}}
.back:hover{{text-decoration:underline}}
</style>
</head>
<body>
<div class="banner"><b>📜 历史研究档案（2026-09-01 生成，十三维旧口径）</b>：本档案为历史版本留存，评分口径与当前榜单的<b>十二维评分</b>存在差异，分数仅供档案参考，<b>最新评分请以榜单页为准</b>。内容基于当时公开数据，不构成任何投资建议。</div>
<div class="wrap">
<a class="back" href="../index.html">← 返回榜单</a>
<div class="doc-card">
{body}
</div>
</div>
<div class="footer"><div class="box">⚠️ 免责声明：本页面为独立数据研究工具输出的历史档案，不提供证券投资咨询服务，不推荐任何证券，不构成任何投资建议。请以交易所及上市公司官方披露为准。据此操作，风险自担。</div></div>
</body>
</html>"""

SKIP_PATTERNS = ('目录域：', 'Ctrl+A', '打开文档后', '九章式深度结构')

WORD_MAP = [
    ('ST摸鱼风云-V2 · 深度跟读报告', '小调AI-ST摸鱼报告V2 · 深度跟读报告'),
    ('ST摸鱼风云-V2', 'ST摸鱼-V2'),
    ('摸鱼风云', '摸鱼'),
    ('V168+G 九章式深度结构 × V2 评分刻度：评分唯一口径为 ST保壳评分系统V2（十三维100分制），未使用 V168 六/八维评分，两套刻度语义不同，请勿与 V168 报告交叉比较。',
     '小调AI-ST摸鱼报告V2（深度报告结构）：评分唯一口径为 ST保壳评分系统V2（十三维100分制）。'),
    ('摸鱼指数', '壳市值指数'),
    ('摸鱼榜', '壳市值观察'),
    ('摸鱼池', '观察池'),
    ('V2保壳分', 'V2评分'),
    ('保壳分', 'V2评分'),
]

def norm(text):
    for a, b in WORD_MAP:
        text = text.replace(a, b)
    return text

def esc(t):
    return H.escape(norm(t)).replace('\n', ' · ')

def esc_body(t):
    return H.escape(norm(t))

def iter_block_items(doc):
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, doc)
        elif child.tag.endswith('}tbl'):
            yield Table(child, doc)

def render_doc(path):
    d = docx.Document(path)
    parts = []
    # 标题：取前两个非空段落（简称+代码 / 报告名）
    title = ''
    first_paras = []
    for b in iter_block_items(d):
        if isinstance(b, Paragraph):
            t = b.text.strip()
            if t:
                first_paras.append((b, t))
            if len(first_paras) >= 2 and b.style.name == 'Heading 1':
                break
        else:
            break
    # 正文顺序渲染
    title_done = 0
    for b in iter_block_items(d):
        if isinstance(b, Paragraph):
            t = b.text.strip()
            if not t:
                continue
            if any(s in t for s in SKIP_PATTERNS):
                continue
            st = b.style.name
            if title_done < 2 and st == 'Heading 1':
                if title_done == 0:
                    parts.append(f'<h1 class="doc-title">{esc(t)}</h1>')
                else:
                    parts.append(f'<div class="doc-sub">{esc(t)}</div>')
                title_done += 1
                continue
            if title_done < 2:
                # 前两段若是普通文本也并入标题区
                if title_done == 0 and len(t) < 30:
                    parts.append(f'<h1 class="doc-title">{esc(t)}</h1>')
                    title_done += 1
                    continue
                if title_done == 1 and len(t) < 60:
                    parts.append(f'<div class="doc-sub">{esc(t)}</div>')
                    title_done += 1
                    continue
            if st == 'Heading 1':
                parts.append(f'<h1>{esc(t)}</h1>')
            elif st == 'Heading 2':
                parts.append(f'<h2>{esc(t)}</h2>')
            else:
                parts.append(f'<p>{esc(t)}</p>')
        else:  # Table
            rows = []
            for ri, row in enumerate(b.rows):
                cells = ''.join(
                    f'<th>{esc(c.text.strip())}</th>' if ri == 0 else f'<td>{esc(c.text.strip())}</td>'
                    for c in row.cells)
                rows.append(f'<tr>{cells}</tr>')
            if rows:
                parts.append(f'<table>{"".join(rows)}</table>')
    return parts, first_paras

def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    files = sorted(glob.glob(os.path.join(SRC_DIR, '*.docx')))
    index = {}
    ok = fail = 0
    for f in files:
        base = os.path.basename(f)
        m = re.match(r'ST摸鱼风云-V2跟读报告_(\d+)_(\d{6})_(.+)_(\d{8})\.docx', base)
        if not m:
            print('SKIP(命名不符):', base)
            fail += 1
            continue
        seq, code, name, date = m.groups()
        try:
            parts, _ = render_doc(f)
            body = '\n'.join(parts)
            page = PAGE_TPL.format(title=f'{name}({code})', body=body)
            out = os.path.join(OUT_DIR, f'{code}.html')
            with open(out, 'w', encoding='utf-8') as fh:
                fh.write(page)
            index[code] = {'name': name, 'date': date, 'file': f'reports/{code}.html'}
            ok += 1
        except Exception as e:
            print('FAIL:', base, repr(e))
            fail += 1
    with open(INDEX_FILE, 'w', encoding='utf-8') as fh:
        json.dump(index, fh, ensure_ascii=False, indent=1)
    print(f'DONE ok={ok} fail={fail} -> {OUT_DIR}/ + {INDEX_FILE}')

if __name__ == '__main__':
    main()
