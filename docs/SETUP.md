# 配置与部署详解

本文件面向人类使用者；AI 助手请优先阅读仓库根目录的 [AGENT.md](../AGENT.md)。

## 1. 准备

- Windows，已安装**经典版桌面 Outlook**（不是「新版 Outlook」/ 网页版）。
- 目标邮箱已添加到当前 Windows 用户的 Outlook 账户列表，且运行时 Outlook 打开。
- Windows PowerShell 5.1（Desktop，`powershell.exe -STA`）。
- Python 3.9+。默认路径**不需要安装任何第三方 Python 包**。

```text
cp config.example.json config.json
python run.py --self-test
python run.py --check-config
python run.py --dry-run
```

`config.json` 已被 `.gitignore` 忽略，切勿提交。

## 2. 配置项逐项说明

### mailbox

| 字段 | 说明 |
| --- | --- |
| `smtp_address` | 目标邮箱完整 SMTP 地址，必须与 Outlook 账户里的地址完全一致（区分大小写不敏感，但拼写要一致） |
| `archive_folder_name` | PST 内用于存放归档邮件的文件夹名；不存在时自动创建 |
| `pst_file_prefix` | PST 文件名前缀，只允许 `A-Z a-z 0-9 . _ -`；不含路径分隔符。生成的文件形如 `<prefix>-yyMMdd.pst`，同一天重复归档会追加到同一文件 |
| `timezone_id` | 用于 PST 按日期命名的 Windows 时区 ID，如 `China Standard Time`、`Taipei Standard Time`、`UTC` |

> 工具通过 Outlook 已登录会话读取邮箱，**不需要在配置里写密码或 OAuth 凭据**。

### capacity

| 字段 | 说明 |
| --- | --- |
| `quota_bytes` | 邮箱容量上限，单位**字节**。请从 Exchange / OWA / 管理员后台的容量页面读取。换算：`1 GB = 1024³ = 1073741824 字节`。示例：`5 GB` 对应 `5368709120`。此处的数值只是通用举例，与任何真实邮箱无关 |
| `trigger_ratio` | 触发比例，取值 `(0, 1]`。触发线 = `ceil(quota_bytes × trigger_ratio)` |
| `retention_hours` | 保留最近多少小时的邮件。归档只处理「投递时间早于 `现在 − retention_hours`」的收件箱顶层邮件 |
| `disk_reserve_bytes` | 移动前要求磁盘至少保留的余量（默认 128 MiB） |
| `pst_ceiling_bytes` | 单个 PST 接近该大小时暂停（默认 45 GiB） |

`quota_bytes` 必须填正数；填 0 或留占位符会被 `--check-config` 拒绝。

### storage

| 字段 | 说明 |
| --- | --- |
| `data_directory` | **绝对路径**，存放 PST、日志、事件、锁文件、已送达记录与临时运行时文件 |
| `allow_cached_archive` | 是否允许在 Outlook 完整缓存模式（`700`）下归档。联机模式（`800`）不受此项影响 |

目录要求：本地固定 NTFS/ReFS 卷、仅当前用户（及 SYSTEM / 管理员）可写、
不是联接 / 符号链接等重解析点。引擎启动时会检查，权限过宽或路径异常会直接
停止并给出原因，不会静默写入别处。

### runtime

| 字段 | 说明 |
| --- | --- |
| `powershell_path` | PowerShell 5.1 可执行文件。默认 `powershell.exe`（从 PATH 查找，找不到时回退到系统目录） |
| `launcher_command` | 可选的命令包装器，见第 5 节 |
| `probe_timeout_seconds` | 只读探针超时（默认 180 秒） |
| `archive_timeout_seconds` | 归档引擎超时（默认 900 秒） |

### notify

| 字段 | 说明 |
| --- | --- |
| `enabled` | **默认 `false`**：完全离线，不发任何网络请求 |
| `provider` | `none` / `stdout` / `webhook` / `feishu` / `wecom` / `dingtalk` / `command` |
| `on_every_run` | `true` 时每轮发一条结果回执（含：本次检测容量、总容量、是否达到清理标准、是否清理）；`false` 时只在归档完成或异常时发 |
| `webhook_url` | 机器人 / 服务的 incoming webhook 地址 |
| `webhook_secret` | 飞书 / 钉钉开启签名或加签时填写；不开启则留空 |
| `payload_template` | `provider=webhook` 时自定义报文 JSON 模板，用 `{text}` 作正文占位符 |
| `custom_command` | `provider=command` 时的命令（字符串或字符串数组） |
| `mention_all` | 企业微信 / 钉钉是否 @所有人 |
| `timeout_seconds` | 单次发送超时 |

`enabled=false` 时 `provider` 会被强制视为 `none`，任何代码路径都不会尝试联网。

## 3. 接入消息工具（可选）

### 不接（默认）

```json
"notify": { "enabled": false, "provider": "none" }
```

结果只写本地并输出到终端。工具功能不受影响。

### 飞书 / Lark 自定义机器人

1. 群设置 → 群机器人 → 添加「自定义机器人」。
2. 复制 Webhook 地址；若开启了「签名校验」，同时复制密钥。
3. 配置：

```json
"notify": {
  "enabled": true,
  "provider": "feishu",
  "on_every_run": true,
  "webhook_url": "https://open.feishu.cn/open-apis/bot/v2/hook/YOUR_TOKEN",
  "webhook_secret": "YOUR_SECRET_OR_EMPTY",
  "timeout_seconds": 15
}
```

开启密钥时，工具会按飞书规则用 `timestamp + \n + secret` 做 HMAC-SHA256 并
base64 生成 `sign` 一起发送。

### 企业微信群机器人

群聊 → 群机器人 → 添加，复制 Webhook 地址，`provider` 填 `wecom`。
`mention_all=true` 时会在报文里加入 `@all`。

### 钉钉自定义机器人

群设置 → 智能群助手 → 添加自定义机器人，复制 Webhook；若选择「加签」，复制
密钥填入 `webhook_secret`。工具会按钉钉规则在 URL 上追加 `timestamp` 与 `sign`。

### 通用 webhook / 自建服务

`provider=webhook` 默认发送 `{"text": "..."}`（兼容 Slack 及多数服务）。
若你的服务格式不同，用模板：

```json
"notify": {
  "enabled": true,
  "provider": "webhook",
  "webhook_url": "https://your-service.example/hook",
  "payload_template": "{\"content\": \"{text}\"}"
}
```

`{text}` 会被安全转义后替换，正文里的引号和换行不会破坏 JSON。

### 接入你已有的连接器 / CLI / 脚本

```json
"notify": {
  "enabled": true,
  "provider": "command",
  "custom_command": ["python", "my_notifier.py"],
  "on_every_run": true
}
```

约定：

- 消息正文通过 **stdin** 传入，同时也放在环境变量 `MBM_MESSAGE`；
- `MBM_IDEMPOTENCY_KEY` 为每次执行唯一的键，可用于去重；
- `MBM_PROVIDER` 为 provider 名称；
- 退出码 `0` 视为成功；非 0 记为失败，事件保留待下一轮重试。

`custom_command` 也可写成单条字符串，此时通过系统 shell 执行。这样任何平台的
命令行工具、内部机器人或自动化框架都能接入，**无需修改仓库源码**。

### 仅终端输出

`provider=stdout`：把消息打印到标准输出，不联网，方便管道或上层助手读取。

## 4. 结果回执

`on_every_run=true` 时，每轮结束都会发一条（无论是否达线）：

```text
【邮箱容量监测】本轮执行结果
本次检测容量：<已换算的 GB>
总容量：<已换算的 GB>
是否达到清理标准：是/否/未知（触发线 <已换算的 GB>，即总容量的 <比例>%）
是否清理：是/否/否（部分）
容量占比：<百分比>
时间(UTC)：<ISO8601>
```

容量按 `1024³` 换算，整数显示为整数 GB，与邮箱容量页面口径一致。若本轮没取到
读数（例如 Outlook 未打开），写「未知」而不是编造数值。归档完成与异常另发独立
通知，按 `event_id` 去重，失败保留重试。

## 5. 提升权限宿主的可选包装器

正常情况下，脚本以打开 Outlook 的同一个交互用户身份直接启动 PowerShell。
若你的自动化宿主以**管理员 / 提升令牌**运行，Outlook 的 COM 接口可能拒绝连接
（典型错误 `0x800401E3 MK_E_UNAVAILABLE`）。此时可以用任意降权或包装工具，
通过 `runtime.launcher_command` 传入，不需要改源码：

```json
"runtime": { "launcher_command": ["python", "deElevate.py"] }
```

两种写法：

- **纯前缀**：数组元素不含 `{command}`，PowerShell 命令会追加在其后。
- **占位符**：`{command}` 替换为完整 PowerShell 命令行，例如
  `["psexec", "-l", "{command}"]`；另外 `{stdout}`、`{stderr}` 会替换为本次运行的
  输出文件路径，`{cwd}` 替换为工作目录，便于包装器自行重定向输出，例如
  `["python", "deElevate.py", "--cwd", "{cwd}", "--stdout", "{stdout}", "--"]`。

若包装器在标准输出打印 `child_exit=N`，编排层会把 `N` 作为子进程真实退出码。
具体用哪种工具由你决定，本仓库不内置、也不依赖任何特定降权程序。

## 6. 定时执行

Windows 任务计划程序：

- 程序：`python`（或 Python 可执行文件完整路径）
- 参数：`run.py`
- 起始位置：仓库根目录
- 勾选「仅当用户登录时运行」（需要 Outlook 交互会话）
- 建议设置执行超时，并配置「已有实例运行时跳过新实例」

也可使用你所用助手 / 自动化平台的定时能力，或任意能定期执行命令的调度器，
调用的都是同一条 `python run.py`。**同一邮箱只在一台机器上启用归档**；引擎有
文件锁，重叠运行会被跳过并告警。

## 7. 首次上线检查清单

- [ ] `--self-test` 返回 `ok: true`
- [ ] `--check-config` 返回 `ok: true`
- [ ] `--dry-run` 的 `probe.ok` 为 `true`，连接模式为 `700` 或 `800`
- [ ] 与使用者确认 `probe.inbox.older_than_cutoff` 的封数符合预期
- [ ] 已明确告知「移动 = 从共享邮箱服务器移走」并取得确认
- [ ] 首次归档后打开 PST，核对邮件与附件可读、源邮件已离开收件箱
- [ ] `config.json`、PST、日志、`_runtime/` 均未被加入版本库
