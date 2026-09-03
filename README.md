# UESTC Job Monitor

电子科技大学就业网招聘信息自动监控项目。程序低频访问公开页面，监控“现场招聘”和首页“需求信息”（站内实际类型为“网上招聘”），解析公司、岗位、时间、地点、学历、专业、截止时间、投递方式及真实详情链接，并把历史数据持久化在 Git 仓库中。

目标网站：<https://jiuye.uestc.edu.cn/career/>

## 网站数据方式

当前网站是 Nuxt 应用。项目依据网站实际前端代码使用以下公开接口，而不是猜测接口：

- `GET api/home/calendar`：按月读取现场招聘日历；
- `POST api/home/recruitmentList`：分页读取近期网上招聘；
- 招聘详情页：`recruitment/{栏目}/{站点编码ID}`。

详情路由中的 ID 按网站前端自身使用的 base-36 算法编码。程序会实际访问每个详情页，并从页面的 `__NUXT_DATA__` 服务端渲染数据中解析详情；只有列表公司名称或标题与详情相符时，链接才标记为 `verified`。无法确认的链接标为 `unverified`，访问失败标为 `failed`，不会虚构 URL 或字段。

## 安装和手动运行

```bash
python -m pip install -r requirements.txt
python scripts/update_jobs.py
```

运行测试：

```bash
python -m pytest -q
```

默认时区是 `Asia/Shanghai`。抓取窗口、请求间隔和重试次数位于 `config/config.json`。首次运行只建立近期基线，不抓取网站全部历史；默认读取最近 14 天发布的需求信息，以及最近 14 天至未来 30 天的现场招聘。窗口内所有尚未入库的真实公告都会打开详情页，不因数量上限而漏掉；为控制日常站点负载，已有公告每个栏目每次最多复查 80 个。候选排序保持确定性，重复执行不会把同一记录再次写入。

## 数据文件

- `data/jobs.csv`：机器可读的权威历史数据，编码为 UTF-8-SIG；
- `data/UESTC招聘信息.xlsx`：使用 `openpyxl` 维护的人类可读 Excel；
- `reports/YYYY-MM-DD.md`：仅当本次存在新增或更新时生成/追加当日日报。

Excel 包含“招聘记录”和“更新记录”两个工作表。招聘记录冻结表头、启用筛选，详情和投递 URL 可点击；更新记录只记录新增或有意义的字段变化。没有新增或更新且 Excel 正常时，程序不会重写 Excel，因此不会产生无意义的 Git 二进制差异。

## 去重和更新规则

程序启动后先读取 CSV 和 Excel，历史状态不依赖 Codex 对话上下文。文本先做 Unicode 与空白规范化，URL 去除末尾斜杠、片段和明显跟踪参数。

- 有详情 URL 和明确岗位：`信息类型 + 规范化详情URL + 岗位`；
- 有详情 URL、无明确岗位：`信息类型 + 规范化详情URL`；
- 无详情 URL：`信息类型 + 公司 + 招聘标题 + 发布时间/招聘日期 + 岗位`。

上述材料计算 SHA-256 作为唯一 ID。同一公告的多个明确岗位分别保存。相同 ID 不新增；已有字段从空变为明确值、内容确实变化或链接状态升级时更新原行。网络瞬时失败不会清空历史、覆盖已有非空字段或把 `verified` 降级。

CSV 和 Excel 都先写入同目录临时文件，校验记录数和唯一 ID 后再原子替换。网站整体不可访问时，程序输出“本次抓取失败，历史数据未修改”。

## Codex Automation 建议

建议在 Codex Cloud 中以 `Asia/Shanghai` 每天上午执行一次，任务指令可写为：

> 在已检出的 UESTC Job Monitor 仓库中先运行 `python -m pytest -q`，测试通过后运行 `python scripts/update_jobs.py`。检查 `git status`，仅在代码、数据或日报确有变化时提交本项目文件；数据更新提交信息使用 `Update UESTC jobs YYYY-MM-DD`，代码修复使用 `Improve UESTC job monitor`。当前分支有远程写权限时正常 push，禁止 force push；无变化不提交，push 失败如实报告。

GitHub 仓库本身是长期存储。自动任务每次必须从已检出的仓库读取 `data/jobs.csv`，执行后正常 commit/push；不要把 Token、Cookie、密码或其他 Secret 写入配置、日志、数据或提交。

## 常见故障

- `403` / `429` / `5xx`：程序只做有限重试；不要缩短请求间隔或增加攻击性并发。
- API/HTML 结构改变：对应栏目会明确报错；历史文件保持不变，先根据网站真实前端修复解析器。
- 单个详情失败：其他详情继续处理，失败链接不会导致整批退出。
- Excel 损坏但 CSV 正常：程序从 CSV 重建 Excel；CSV 仍是机器可读历史源。
- CSV 损坏且无法安全读取：程序停止，避免用空数据覆盖历史。
- 第二次运行仍有新增：先确认网站在两次运行之间是否发布了新公告，再检查唯一 ID 和 URL 规范化测试。
- Git push 失败：保留本地 commit，不绕过分支保护或 force push。
