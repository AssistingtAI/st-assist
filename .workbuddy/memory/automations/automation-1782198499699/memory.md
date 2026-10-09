# automation-1782198499699 - 保壳风云榜每周五更新

## 2026-09-11 10:07 执行（首次记录）

**结果：数据链+验证全通过，GitHub push 阻断，未上线**

- weekly_update_friday.py 13步：✅ 成功（207家，V2 A20/B88/C60/D39，均分51.9；新增002743/301117，移除000711/002586）
- verify_html.js：✅ 全通过（V2 sum!=total 0，JS syntax OK，莫高63.6/B）
- git commit：✅ 本地成功（commit 5139e03）
- git push origin HEAD：❌ 失败 → **github.com:443 TCP不可达**（3次重试），且Windows凭据管理器无github.com条目
- 对照：api.github.com 达(200) / 百度达(200) / gitee凭据在 → 仅github.com:443被封
- 按用户指令"任一步失败即停止"，未强行发布

**下次执行要点**
1. 若仍失败，先测 `curl -o /dev/null -w "%{http_code}" --max-time 15 https://api.github.com` 与 `/dev/tcp/github.com/443`，快速定位是网络封禁还是凭据缺失
2. 恢复通道：代理直连 push；或配 GitHub PAT 后走 Git Data API（api.github.com 本机可达）
3. **不要重跑数据链**——本地 commit 已在，网络恢复后直接 `git push origin HEAD`
4. `git credential fill` 会挂起等待交互（无存储凭据），务必加 `GIT_TERMINAL_PROMPT=0`

## 2026-09-18 09:05 执行

**结果：数据链+验证+commit 全通过，GitHub push 再次阻断，未上线**

- weekly_update_friday.py 13步：✅ 成功（207家，V2 A20/B87/C60/D40，均分51.9；新增0，移除0）
- verify_html.js：✅ 全通过（V2 sum!=total 0，JS syntax OK，莫高63.6/B）
- git commit：✅ 本地成功（commit d5b344d）
- git push origin HEAD：❌ 失败 → GCM 追踪确认 Windows 凭据管理器中 **没有** `https://github.com` 的已存凭据；GIT_TERMINAL_PROMPT=0 下无法弹窗或终端输入，报错 `Cannot prompt because terminal prompts have been disabled`
- 按用户指令“任一步失败即停止”，未强行发布

**后续恢复**
- 在 Windows 凭据管理器中为 `git:https://github.com` 添加 GitHub Personal Access Token（classic 或 fine-grained，需 repo 权限）
- 配置后在本仓库执行 `git push origin HEAD` 即可把已就绪的 d5b344d 推送到远程
- 或配置 GitHub PAT 后走 `push_via_api.py` 经 api.github.com 推送

## 2026-10-02 16:23 执行

**结果：全链路成功，线上已更新（连续3周push阻断后首次完整上线）**

- weekly_update_friday.py 13步：✅ 成功（204家，V2 A19/B84/C63/D38；新增600363 ST联光，移除300527 ST应急；行情/市值204/204）
- verify_html.js：✅ 全通过（RAW/V2各204行、row len 21、sum!=total 0、missing 0、JS语法OK；莫高63.6/B）
- git commit：✅ 本地成功（8fcec8d，14文件）
- 推送：✅ **github.com:443 本日恢复连通**（常规push到达远端但因SHA分叉被non-fast-forward拒绝），改走 push_via_api.py 成功，远端 main → be4805d9，复刻2个提交（4768590上周+8fcec8d今日）
- 线上验证：✅ assistingtai.github.io/st-assist/baokeng-rank.html 已含 ST联光 + "覆盖公司：204 家"，index.html 200
- 注意：600363 新简称 **ST联光**（名单源叫联创光电）；自动化提示词中"十三维"与旧URL已过时，实际=十二维+st-assist
- 下次可选：github.com已通，可 git pull --rebase 对齐分叉历史后恢复常规push（等价补丁自动跳过）

## 2026-09-25 09:06 执行

**结果：数据链+验证+commit 全通过，GitHub push 再次阻断，未上线**

- weekly_update_friday.py 13步：✅ 成功（204家，V2 A18/B86/C61/D39，均分51.8；新增0，移除300096/300211/301139）
- verify_html.js：✅ 全通过（RAW rows 204, V2 row len 21, V2 sum!=total 0, V2 missing codes 0, JS syntax OK；莫高V2 sum=63.6/total=63.6/level=B）
- git commit：✅ 本地成功（commit 4768590，14 files changed, 6988 insertions, 7291 deletions）
- git push origin HEAD：❌ 失败 → git-remote-https 进程挂起近5分钟无响应，手动终止；Windows 凭据管理器中仍无 `git:https://github.com` 条目；api.github.com 可达（200），确认是 HTTPS 凭据缺失导致 push 挂起
- 按用户指令“任一步失败即停止”，未强行发布

**下次执行要点**
1. 必须在 Windows 凭据管理器中添加 `git:https://github.com` 凭据（GitHub PAT classic/fine-grained，需 repo 权限），或配置 Git 使用 `push_via_api.py` 经 api.github.com 推送
2. 当前本地未推送的 commit 已累积 3 个（d5b344d、4768590 等），恢复通道后 `git push origin HEAD` 会一次性推送
3. 数据链已就绪，网络恢复后无需重跑
