#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""前十大股东采集管道 — 东财数据中心 RPT_F10_EH_HOLDERS（2026-09-02 新增）
供 S1 评分用：取每只 ST/*ST 最新报告期前十大股东，
输出 st_top_holders.json（每 code：report_date + top_holders + top1 + top3_sum）。
用途：①S1前置(第一大股东=实控?+持股>20%) ②民企前3大>50%上探10 ③实控人持股比例。
"""
import json
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import defaultdict

import requests

BASE = os.path.dirname(os.path.abspath(__file__))
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    'Referer': 'https://emweb.securities.eastmoney.com/',
}
URL = 'https://datacenter.eastmoney.com/securities/api/data/v1/get'


def secu_code(code: str) -> str:
    c3 = code[:3]
    if c3 in ('600', '601', '603', '605', '688', '689', '900'):
        return code + '.SH'
    if c3 == '920' or code.startswith('4') or code.startswith('8'):
        return code + '.BJ'
    return code + '.SZ'


def fetch_one(code: str):
    sc = secu_code(code)
    for attempt in range(4):
        try:
            r = requests.get(URL, params={
                'reportName': 'RPT_F10_EH_HOLDERS', 'columns': 'ALL',
                'filter': f'(SECUCODE="{sc}")',
                'pageNumber': 1, 'pageSize': 500, 'source': 'HSF10', 'client': 'PC',
            }, headers=HEADERS, timeout=15)
            data = ((r.json() or {}).get('result') or {}).get('data') or []
            if not data:
                return code, None  # 接口通但无股东数据（北交所/新股可能）
            # 按报告期分组，取最新一期
            by_date = defaultdict(list)
            for d in data:
                ed = (d.get('END_DATE') or '')[:10]
                if ed:
                    by_date[ed].append(d)
            if not by_date:
                return code, None
            latest = sorted(by_date.keys())[-1]
            rows = sorted(by_date[latest], key=lambda x: x.get('HOLDER_RANK') or 99)
            holders = []
            for d in rows[:10]:
                holders.append({
                    'rank': d.get('HOLDER_RANK'),
                    'name': (d.get('HOLDER_NAME') or '').strip(),
                    'ratio': d.get('HOLD_NUM_RATIO'),  # 可能 None/缺失
                })
            # 前3大合计（ratio 有效值求和）
            ratios = [h['ratio'] for h in holders[:3] if isinstance(h.get('ratio'), (int, float))]
            info = {
                'secucode': sc,
                'report_date': latest,
                'top_holders': holders,
                'top1_name': holders[0]['name'] if holders else None,
                'top1_ratio': holders[0]['ratio'] if holders else None,
                'top3_sum': round(sum(ratios), 2) if len(ratios) == len(holders[:3]) else (sum(ratios) if ratios else None),
            }
            return code, info
        except Exception:
            time.sleep(1 + attempt)
    return code, None


def main():
    with open(os.path.join(BASE, 'st_names.json'), encoding='utf-8') as f:
        name_map = json.load(f)
    codes = list(name_map.keys())
    print(f'[TOP10] 开始批量采集前十大股东：{len(codes)}家')

    out, fails = {}, []
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = {ex.submit(fetch_one, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            code, info = fut.result()
            done += 1
            if info:
                out[code] = info
            else:
                fails.append(code)
            if done % 50 == 0:
                print(f'  进度 {done}/{len(codes)}，失败 {len(fails)}')

    # 需补全 code 级名（评分处用 code 索引，name 从 st_names 现取）
    payload = {
        'meta': {
            'source': 'datacenter.eastmoney.com RPT_F10_EH_HOLDERS(最新报告期前十大)',
            'fetched_at': time.strftime('%Y-%m-%d %H:%M:%S'),
            'total': len(codes), 'ok': len(out), 'fail': len(fails),
            'rule': 'S1前置:第一大股东=实控+持股>20%满分;民企前3大>50%上探10;质押>60%减半;实控冻结限高折减;涉造假封顶4',
        },
        'data': out,
    }
    with open(os.path.join(BASE, 'st_top_holders.json'), 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)

    miss = [c for c in codes if c not in out]
    print(f'[TOP10] 完成：成功 {len(out)}，失败 {len(fails)}')
    if fails:
        print('  失败清单:', fails)
    # 冀凯样例核验
    if '002691' in out:
        h = out['002691']
        print('  冀凯样例:', h['report_date'], '| top1:', h['top1_name'], h['top1_ratio'],
              '| top3_sum:', h['top3_sum'])


if __name__ == '__main__':
    main()
