# Weekly Report Aggregator (周报汇总自动化)

自动化收集学生周报并合并发送的工具。从本地文件夹和 QQ 邮箱两个来源读取学生周报 PDF，按 `students.txt` 顺序合并成一个带书签的 PDF，并在指定时间通过 QQ 邮箱发送到指定收件人。

---

## 功能特性

- **双数据源**：同时从本地目录（`FilePath`）和 QQ 邮箱（最近 N 天）读取周报
- **邮件抓取重试**：单封邮件读取超时会自动重试并重连 IMAP，失败后只跳过该邮件，不影响后续邮件
- **智能匹配**：通过文件名、PDF 内容、邮件发件人三种信号匹配学生
- **拼音变体**：基于 `students.txt` 中手动提供的拼音，生成 `zhangsan`/`zhang_san`/`zhang-san`/`sanzhang` 等多种匹配形式，避免中文多音字导致的错误匹配
- **冲突解决**：本地优先，同源取最新；同名学生在已交/未交中只出现一次
- **PDF 合并 + 书签**：每个学生一个顶级书签（中文名）
- **大附件分卷**：超过 20MB 自动按学生边界拆分成多封邮件
- **定时执行**：到 `TargetTime`（北京时间）才触发发送
- **日志 + 匹配报告**：每次运行生成日志和详细匹配明细
- **Dry-run 模式**：预览匹配/合并结果但不发邮件
- **错误兜底**：单个 PDF 损坏不影响整体流程

---

## 目录结构

```
weekly_report_aggregator/
├── main.py                     # 程序入口
├── requirements.txt
├── config.json                 # 配置模板（含 IMAP/SMTP 凭据）
├── README.md
├── students.txt                # 学生名单示例
├── TargetEmail.txt             # 收件人示例
├── src/
│   ├── __init__.py
│   ├── config.py               # 配置 / students.txt / TargetEmail.txt 加载
│   ├── pinyin_utils.py         # 基于手动拼音生成匹配变体
│   ├── pdf_utils.py            # PDF 读取 / 合并 / 书签 / 分卷
│   ├── matcher.py              # PDF→学生 匹配器
│   ├── email_fetcher.py        # IMAP 抓取邮件附件，支持单封邮件 fetch 重试与断线重连
│   ├── email_sender.py         # SMTP 发送
│   ├── scheduler.py            # TargetTime 等待
│   └── logger.py               # 日志配置
├── logs/                       # 运行时自动生成
└── output/                     # 合并 PDF + 匹配报告
```

---

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 准备配置文件

复制示例配置并填入你的信息

#### `config.json` 字段说明

```json
{
    "imap": {
        "host": "imap.qq.com",
        "port": 993,
        "user": "xxx@qq.com",
        "password": "QQ邮箱授权码（不是登录密码）"
    },
    "smtp": {
        "host": "smtp.qq.com",
        "port": 465,
        "user": "xxx@qq.com",
        "password": "QQ邮箱授权码"
    },
    "FilePath": "/Users/xxx/Documents/weekly_reports",
    "TargetTime": "2026_05_08_23_59",
    "lookback_days": 3,
    "max_attachment_size_mb": 20,
    "sender_display_name": "李伟"
}
```

> **关于 QQ 邮箱授权码**
> QQ 邮箱不允许直接用登录密码进行 IMAP/SMTP，必须先开启「IMAP/SMTP 服务」并生成「授权码」。
> 路径：QQ 邮箱网页版 → 设置 → 账户 → 「POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务」→ 开启 IMAP/SMTP → 生成授权码。

#### `students.txt` 格式

每行一个学生：中文名 + 手动拼音 + 一个或多个邮箱。

推荐使用中文逗号 `，`、英文逗号 `,`、中文分号 `；` 或英文分号 `;` 分隔字段。拼音字段内部可以包含空格，例如 `li zhuoran`。

```
张三 zhangsan@qq.com
李四，lisi@gmail.com；lisi2@qq.com
王五 wangwu@example.com,wangwu@163.com
赵六 zhaoliu@qq.com
```

> 顺序非常重要——合并 PDF 中各学生的章节顺序、邮件正文中"已交/未交"的学生顺序，都按本文件的顺序。

> 注意：当前版本不再自动把中文名转换为拼音。每个学生都必须手动提供拼音，尤其适合处理中文姓名中的多音字。

#### `TargetEmail.txt` 格式

每行一个收件人邮箱（也可用上面的分隔符放多个邮箱在一行）：

```
advisor1@example.com
advisor2@example.com
```

#### 邮件抓取稳定性说明

邮件附件通过 IMAP 从 QQ 邮箱读取。为避免一封邮件超时影响整个流程，程序对单封邮件 fetch 设置了重试机制(这些参数目前是写死在 `email_fetcher.py` 里的)：

```
_IMAP_TIMEOUT_SECONDS = 60
_FETCH_MAX_ATTEMPTS = 3
_FETCH_RETRY_SLEEP_SECONDS = 2
```

```
IMAP socket timeout: 60 秒
单封邮件 fetch 最大重试次数: 3 次
重试间隔: 2 秒
```
每次 fetch 超时后，程序会关闭当前 IMAP 连接并重新连接；如果同一封邮件多次失败，则跳过该邮件并继续处理下一封。

### 3. 运行

```bash
# 等到 config.json 中的 TargetTime（北京时间）才执行
python main.py

# 立即执行（忽略 TargetTime 等待，但仍用其命名输出）
python main.py --no-wait

# 演练模式（不发邮件，只生成 PDF + 匹配报告供检查）
python main.py --dry-run --no-wait

# 自定义路径
python main.py \
    --config myconfig.json \
    --students mystudents.txt \
    --target-email myrecipients.txt \
    --output-dir output_2026_05_08
```

### 4. 命令行参数

| 参数 | 说明 |
|------|------|
| `--config` | 配置文件路径，默认 `config.json` |
| `--students` | 学生名单路径，默认 `students.txt` |
| `--target-email` | 收件人列表路径，默认 `TargetEmail.txt` |
| `--dry-run` | 完整运行但不发邮件，输出全部保留 |
| `--no-wait` | 不等待 TargetTime，立即执行 |
| `--output-dir` | 合并 PDF + 匹配报告输出目录，默认 `output/` |
| `--keep-temp` | 保留邮件附件临时目录（默认运行结束后删除） |

---

## 工作流程

```
启动
 │
 ├─→ 加载 config.json / students.txt / TargetEmail.txt
 │
 ├─→ 等待至 TargetTime（北京时间） ──[--no-wait 跳过]
 │
 ├─→ 扫描 FilePath 下所有 .pdf  →  本地候选列表
 │
 ├─→ 通过 IMAP 搜索 [TargetTime - lookback_days, TargetTime] 间的邮件 UID
 │     对每封邮件执行 fetch：
 │       • fetch 超时 / 连接中断时，自动关闭旧连接并重新连接
 │       • 单封邮件最多重试若干次
 │       • 多次失败后只跳过该邮件，不影响后续邮件
 │     提取所有 PDF 附件保存到临时目录 → 邮件候选列表
 │
 ├─→ 匹配器：为每个候选 PDF 算出最佳学生
 │     • 文件名包含中文名     +60
 │     • 文件名包含拼音变体   +50
 │     • PDF 内容含中文名     +40
 │     • PDF 内容含拼音变体   +25
 │     • 发件人邮箱匹配学生   +20
 │     得分 ≥ 40 才算匹配；并列时不匹配
 │
 ├─→ 每个学生选一份 PDF
 │     本地优先 > 邮件
 │     同源取 mtime 最新
 │
 ├─→ 按 students.txt 顺序，合并 PDF + 添加书签（中文名）
 │     如果 > max_attachment_size_mb，按学生边界分卷
 │
 ├─→ 写匹配报告：output/match_report_<TargetTime>.txt
 │
 └─→ 通过 SMTP 发送 → TargetEmail 中的所有收件人
       多卷时分别发送，subject 加 (i/n)
       Dry-run 模式只构造邮件不发送
```

---

## 输出说明

### 合并的 PDF

- 单个文件：`output/WeeklyReport_2026_05_08_23_59.pdf`
- 多卷（超过 20MB）：`output/WeeklyReport_2026_05_08_23_59_part1of3.pdf`、`_part2of3.pdf`、`_part3of3.pdf`

每个文件都有书签，PDF 阅读器左侧大纲可直接跳转到每个学生的章节。

### 匹配报告

`output/match_report_<TargetTime>.txt` —— 详细列出：

- 每个候选 PDF 匹配到了谁、得分多少、匹配原因
- 每个学生最终选了哪份 PDF（本地 vs 邮件）
- 未匹配的 PDF 列表（便于排查命名或拼音问题）
- 如果邮件附件没有成功保存，则该附件不会出现在匹配报告中，需要查看日志中的 IMAP / fetch 记录

### 日志

`logs/run_<时间戳>.log` —— 每次运行一份，含 IMAP/SMTP 操作、邮件 fetch 重试与重连记录、附件保存记录、匹配过程、错误等。

如果某封邮件读取失败，日志会显示类似：

```
Failed to fetch message b'41' (attempt 1/3): The read operation timed out
Failed to fetch message b'41' (attempt 2/3): ...
Skipped message b'41' after 3 failed fetch attempt(s): ...
```
这表示只跳过该邮件，程序会继续处理后续邮件。

### 邮件

主题：`周报汇总 - 2026年5月8日`（从 TargetTime 解析）

正文：

```
已交周报：张三，李四，王五
未交周报：赵六，孙七
```

---

## 匹配规则细节

### 拼音变体生成

对中文名 `李伟`（`li`/`zhuo`/`ran`），生成以下变体：

```
姓在前： xxx, li_zhuoran, li-zhuoran, li_zhuo_ran, li-zhuo-ran
名在前： zhuoranli, zhuoran_li, zhuoran-li, zhuo_ran_li, zhuo-ran-li
```

匹配时统一转小写后做子串匹配。

### 来源去重规则

| 学生有... | 选择 |
|-----------|------|
| 本地 1 份 + 邮件 0 份 | 本地那份 |
| 本地 0 份 + 邮件 1 份 | 邮件那份 |
| 本地 N 份 + 邮件 0 份 | 本地 mtime 最新的 |
| 本地 0 份 + 邮件 N 份 | 邮件 mtime 最新的（按邮件 Date 头） |
| 本地 ≥1 份 + 邮件 ≥1 份 | **本地** mtime 最新的 |
| 本地 0 份 + 邮件 0 份 | 进入"未交"名单 |

### 已交/未交 约束

- `students.txt` 中的每个学生**必且仅**出现在一个名单（已交 或 未交）
- 多个邮箱不会导致重复——按中文名去重

---

## 错误处理

| 异常情况 | 处理方式 |
|----------|----------|
| 单个 PDF 损坏 / 无法读取 | 跳过该文件，记录到日志和匹配报告，不影响其他 PDF |
| 邮箱认证失败 | 立即报错并退出（无法继续） |
| IMAP 搜索邮件失败 / 首次连接失败 | 记录错误，邮件来源不可用时继续使用本地 PDF |
| 单封邮件 fetch 超时 / 连接中断 | 自动重试并重连 IMAP；多次失败后只跳过该邮件，不影响后续邮件 |
| 单封邮件解析失败 | 跳过该邮件，记录到日志，不影响后续邮件 |
| 邮件附件保存失败 | 跳过该附件，记录到日志，不影响其他附件 |
| FilePath 不存在或非目录 | 记录错误，仅用邮件来源（如果有的话） |
| SMTP 发送失败 | 记录错误，返回非零退出码 |
| 没有任何学生交了周报 | 退出码 2，不发邮件 |
| 同名 / 同分匹配并列 | 该 PDF 留作"未匹配"，写入报告供人工处理 |

---

## 常见问题

**Q：QQ 邮箱授权码在哪里申请？**
A：网页登录 QQ 邮箱 → 设置 → 账户 → 找到「POP3/IMAP/SMTP...」服务 → 开启 IMAP/SMTP，按提示发送验证短信，得到 16 位授权码。

**Q：程序需要一直运行直到 TargetTime 吗？**
A：是的，按需求采用方案 A（程序常驻直到指定时间）。如果想用系统级的 cron / 计划任务，可以加 `--no-wait` 参数让程序立即执行。

**Q：邮件附件中带毒/过大怎么办？**
A：超过 `max_attachment_size_mb` 的单个 PDF 仍会被处理（按学生边界单独成卷发送），并在日志中提示。

**Q：能不能让程序运行多次？**
A：每次执行是独立的（无状态）。修改 `config.json` 的 `TargetTime` 重新启动即可。也可以多个实例并行（比如不同班级）但要注意 QQ 邮箱的 IMAP 并发限制。

**Q：如何调试匹配失败？**
A：先用 `--dry-run --no-wait` 跑一遍，查看 `output/match_report_*.txt`，里面有每个 PDF 的得分明细。

**Q：为什么我明明在 QQ 邮箱里收到了学生邮件，结果仍显示未交？**

A：请先查看 `logs/run_<时间戳>.log` 和 `output/match_report_*.txt`。

如果日志中 `Total PDFs saved from email` 为 0，说明程序没有成功从邮件中保存 PDF 附件。当前版本会对单封邮件 fetch 超时自动重试并重连；如果某封邮件多次失败，只会跳过该邮件，不会影响后续邮件。

可以重点检查日志中是否有以下信息：

```
Saved email attachment: ...
Failed to fetch message ... (attempt 1/3)
Skipped message ... after 3 failed fetch attempt(s)
```
如果 `Saved email attachment` 出现了，但匹配报告仍显示该学生未交，再检查 `match_report_*.txt` 中该 PDF 的匹配得分和匹配原因。

**Q：为什么邮件抓取使用 UID 而不是普通 message id？**

A：因为普通 `message sequence number` 在 IMAP 重连后可能变化，而 UID 更稳定。程序在 fetch 超时后会关闭旧连接并重新连接，因此使用 UID 可以更安全地继续处理后续邮件。

---

## 依赖

- Python 3.9+（用到 `zoneinfo`）
- pypdf（PDF 合并 + 书签）
- pdfplumber（更准的 PDF 文本提取）

详见 `requirements.txt`。

---

## 安全提醒

- `config.json` 含有 QQ 邮箱授权码，**不要**提交到 git。建议加入 `.gitignore`：
  ```
  config.json
  logs/
  output/
  ```
- 授权码泄露后，到 QQ 邮箱后台重新生成一次即可作废旧的。
