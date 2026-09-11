# 保壳风云榜 项目记忆

工作目录：C:\Users\xiaot\WorkBuddy\2026-05-16-task-2（GitHub Pages: **assistingtai.github.io/baokeng-ranking**；2026-09-11 GitHub账号 Zsheng007→**AssistingtAI**，仓库名仍为 baokeng-ranking；旧 zsheng007.github.io 已404不重定向）

## 评分体系（当前权威口径）
**ST保壳评分系统V2 · 十二维100分制**（2026-08-28老Z定稿切换，2026-09-02 13→12重构 config v2.3；**十三维/V1已废弃**，历史见每日日志+git）。分数越高=保壳越容易；评级 A(>70)/B(51-70)/C(31-50)/D(≤30)。score_v2.py 是唯一评分实现源，榜单分数=个股报告分数零偏差。完整维度档位/规则见 skill `~/.workbuddy/skills/baokeng-rank/SKILL.md`。

**12 维满分**：C1面值6 / C2壳价值10(反转,越便宜分越高,基准33.81亿) / S1实控人12 / S2质押6 / A1净资产12 / A2扣非主营12 / A3扣非6 / D1现金流4 / B1立案10 / B2审计12 / F1重组纾困6(原F2更名) / H1司法4 = 100。

**S1(2026-09-02新版连乘)**：性质base(央企/省级/市县国资12、院所/国资未分层10、民企/无实控8)→ 折减(质押>60%×0.5、top1持股≤20%×0.8、实控冻结×0.5/限高×0.6、实控刑事固定3)→ 上探(民企前3大>50%→10、主体>500亿满12未启用)→ 涉造假封顶4，下界1。判持股只用 top1_ratio>20%，**勿用股东名=实控名**(间接持股会误杀)。

**通道封顶(闸门，JS+引擎同步)**：C1=0或B2=0→总分封顶50；B1=0→封顶30。

**2026-09-02 13→12重构**：取消原F1趋势4分→A1+2(→12)/C2+2(→10)；F2更名F1。config v2.3。分布 A20/B86/C60/D41，均分51.7，覆盖207家。冀凯 70/B/rank21(13维为67/rank30)。

## 数据与管道
唯一口径链：st_names/st_market_data/st_financials + st_risk_flags(巨潮定向) + st_controllers/st_top_holders/st_pledges/st_deduct_income → build_baokeng_v2.py → st_scores_v2.json → generate_html.py → baokeng-rank.html+index.html。st_trends 不再消费(原F1维度取消)。verify_html.js 生成后必跑(RAW结构+V2求和=total+等级+JS语法四重)。

## 关键铁律/坑
- **generate_html.py 模板是唯一权威，HTML是产物**：改页面文案必落模板，直接改HTML会被下次生成覆盖；模板JS内`${...}`需转义`${{...}}`；镜像 _check.js 由verify自动同步。
- 页面分数=数据源total；JS calcScore2 索引与V2RAW逐位一致(12维idx4-15,total20)。
- 巨潮采集须**公司定向查询**(orgId取自szse_stock.json,pageSize钳30)，大窗口关键词搜索翻不到深排公告；按(code,bucket,title,date)去重。
- 北交所(920)巨潮未覆盖维度走bj中性降级。
- 壳基准 SHELL_BASE=33.81亿(own口径,动态读 shell-fee-base/shell_transactions.json)；壳费中位22.02亿。

## 报告体系（刻度不得混用）
- **ST摸鱼风云-V5**(score_v2十二维,榜单同刻度)：模板 `~/.workbuddy/skills/baokeng-rank/scripts/st_moyu_fengyun_v5_report_template.py`(九章式,A4可用宽16cm+fixed防撑破)；冀凯实例 gen_v5_jikai_full.py。
- **V168-G**(st-risk-analyzer)：内部尽调独立刻度，报告不并列双分数。
- 产品结论：**页面=纯免费榜单**(报告功能2026-09-01下线,交付走线下docx直发)；reports_baokeng_v2/ 保留报告库可人工直发。

## 自动化 & 环境
- automation-1782198499699 每周五9:00 全自动更新(13步链)；automation-1787916836110 每周五11:00名单核验。
- weekly_update_friday.py 主脚本；fetch_financials/fetch_risk_flags 链外刷新。
- venv Python ~/.workbuddy/binaries/python/envs/default/Scripts/python.exe(docx/requests)；Node ~/.workbuddy/binaries/node/versions/22.12.0/node.exe。
- GitHub被限流用 Git Data API 模拟推送(blobs→trees→commits→refs)。
