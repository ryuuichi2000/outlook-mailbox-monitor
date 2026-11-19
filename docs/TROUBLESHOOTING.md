# 常见问题

## 配置阶段

**`--check-config` 报 `invalid_config` / 占位符未替换**
逐个修正 `errors` 里的字段。`config.example.json` 中所有 `REPLACE_WITH_*`、
`example.com`、`YOUR_*` 都必须换成你自己的真实值；`quota_bytes` 必须填正整数
字节数。

**不知道 `quota_bytes` 填多少**
从 Exchange / OWA / 管理员后台的邮箱容量页面读取配额，再换算成字节：
`GB × 1073741824`。例如 5 GB 填 `5368709120`（此处仅为通用举例）。
工具本身不做配额猜测。

**`storage.data_directory` 被拒绝**
必须是绝对路径、本地固定卷、非联接/符号链接、且仅当前用户（及 SYSTEM /
管理员）可写。权限过宽或路径是网络映射都会被拒绝。

## Outlook 阶段

**`probe.ok = false` / `reason = outlook_not_ready`**
确认：使用的是**经典版桌面 Outlook**（不是网页版/新版 Outlook）；运行时 Outlook
已打开；目标邮箱已配置在**同一个 Windows 用户**的账户列表里；SMTP 地址拼写与
配置完全一致。

**错误 `0x800401E3 MK_E_UNAVAILABLE`**
通常表示发起调用的进程与 Outlook 不在同一完整性级别 / 会话。最常见于以
**管理员或提升令牌**运行的自动化宿主。两种办法：

- 让脚本以打开 Outlook 的普通交互用户身份运行（推荐）；
- 或通过 `runtime.launcher_command` 接入你自己的降权 / 包装工具，见
  [SETUP.md 第 5 节](SETUP.md)。

本仓库不内置、也不依赖任何特定降权程序。

**连接模式不是 700 / 800**
例如 `400`（断连）、`500`/`600`（仅下载邮件头等）。工具会当轮告警且不归档。
请在 Outlook 中恢复连接后重试。完整缓存模式 `700` 需
`storage.allow_cached_archive = true` 才允许归档。

**用量属性读不到 / `monitor_unavailable`**
个别账户配置或同步状态下，目标存储不暴露 `0x0E080014`。可先在 Outlook 中
确认账户同步正常、容量页面可见；工具会把原因写入日志并告警，不会把「读不到」
当成「容量很低」。

**PowerShell 报 STA / 版本错误**
引擎要求 Windows PowerShell 5.1（Desktop）且线程为 STA。编排层已自动加
`-STA`；若你手动运行，请使用 `powershell.exe -STA -File ...`，不要用 PowerShell 7
（Core）代替。

**中文乱码**
所有 `.ps1` 文件刻意保持纯 ASCII，以兼容 Windows PowerShell 5.1 对无 BOM
UTF-8 的解析。请**不要**往 `.ps1` 里加中文字符串；面向人的文案放在 Python
编排层（UTF-8）即可。

## 归档阶段

**`older_than_cutoff` 数量与预期不符**
探针只统计**收件箱顶层**的普通邮件（`Class = 43`），按 Exchange UTC 投递时间
判断。子文件夹、其他文件夹、恰好等于截止时间的邮件都不算。

**`archive_operation_failed` / `partial_archive = true`**
详情里有 `moved_count`、`eligible_count`、`pst_path`。已成功移动的邮件在 PST 里，
未移动的仍在收件箱。先看 `cause`（磁盘空间、PST 上限、连接切换、ACL 等），
处理后下一轮会重新评估；不会自动重复移动已归档的邮件。

**`usage_still_above_trigger_after_archive`**
缓存模式读数可能滞后于服务器。工具因此不谎报降线，而是额外告警。请对照服务器
容量页面确认。

**`overlapping_run_skipped`**
另一个实例持锁。请确认调度器没有并发触发；同一邮箱只在一台机器启用归档。

## 通知阶段

**完全没有收到消息**
先确认 `notify.enabled = true` 且 `provider` 不是 `none`。默认是关闭的，这是
刻意设计。再看本轮输出 JSON 的 `summary_notification.sent` 与 `detail`。

**webhook 返回成功却没收到**
各平台在 HTTP 200 里也可能带业务错误码（飞书 `code`、企微/钉钉 `errcode`）。
工具会解析并视为失败，详情写在 `detail`。检查机器人的关键词/白名单/签名设置。

**飞书 / 钉钉签名失败**
开启签名或加签时，`webhook_secret` 必须填对应密钥；未开启则留空。填错会报
业务错误码。

**`provider = command` 失败**
确认命令退出码为 0、可从 stdin 或 `MBM_MESSAGE` 读取正文、执行用户有权限。
失败会保留事件待下轮重试。

**通知失败会不会导致归档失败？**
不会。归档与通知解耦；通知失败只把本轮 `ok` 置为 `false` 并保留事件待重试，
已完成的归档不会被回滚或重复执行。

## 调度阶段

**定时任务运行了却什么都没发生**
多半是 Outlook 未在该会话打开，或调度器以服务/系统账户运行（无交互会话）。
勾选「仅当用户登录时运行」，并使用打开 Outlook 的那个 Windows 用户。

**任务卡住不结束**
给任务设置执行超时；工具自身也有 `probe_timeout_seconds` 与
`archive_timeout_seconds`。超时会产生告警而不是无限等待。
