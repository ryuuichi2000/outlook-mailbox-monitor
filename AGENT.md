# 给 AI 助手 / Agent 的接入引导

> 这份文件是写给**拿到本仓库后负责协助配置的 AI 助手**看的，不是给人看的部署手册。
> 人类使用者请直接阅读 [README.md](README.md) 与 [docs/SETUP.md](docs/SETUP.md)。
>
> 本文件不使用任何特定 AI 产品、自动化平台或厂商的专有术语。无论你运行在哪种
> 助手、IDE、命令行或自动化环境里，按下面的步骤引导使用者即可。

## 你的角色与铁律

使用者把本仓库交给你，希望你帮他跑通「邮箱容量监测 → 达线时归档到本地 PST →
可选地推送结果」。请遵守：

1. **先配置，后运行；先只读，后写入。** 在完成配置校验和只读预览之前，不要执行
   任何会移动邮件的命令。
2. **绝不自作主张填值。** 邮箱地址、配额、输出目录、通知渠道都属于使用者自己的
   环境，必须逐项向他确认。源码与示例配置里没有任何真实值，这是刻意设计的。
3. **通知是可选项，不是前置条件。** 使用者完全可以选择「不接任何消息工具」，
   此时工具照常工作，只写本地日志并在终端输出结果。
4. **不要修改 `src/` 里的引擎逻辑**来适配某个平台。所有差异都通过
   `config.json` 表达；若现有渠道不够用，用 `notify.provider = "command"` 接
   使用者自己的命令即可。
5. **提醒归档语义**：归档用的是「移动」而非「复制」。对共享 / 公用邮箱来说，
   邮件会从服务器邮箱移走、其他共用者不再看到。首次真实归档前必须让使用者
   明确知情并确认。

## 第 1 步：检查配置是否存在

- 查看仓库根目录有没有 `config.json`。
- 没有 → 从 `config.example.json` 复制一份，**不要**直接改示例文件：

  ```text
  cp config.example.json config.json
  ```

  （Windows 也可用 `copy config.example.json config.json`）

- 已存在 → 先读一遍，逐项和使用者核对，不要覆盖他的值。
- `config.json` 已被 `.gitignore` 忽略。提醒使用者：**不要把它、PST 文件、
  运行日志或任何凭据提交到版本库。**

## 第 2 步：逐项询问使用者（不要猜）

按顺序问，一次问完也可以，但每一项的答案都要来自使用者本人：

| 配置项 | 你要问的问题 | 备注 |
| --- | --- | --- |
| `mailbox.smtp_address` | 要监测哪个邮箱？请给出完整 SMTP 地址 | 必须与使用者的桌面版 Outlook 中配置的账户**完全一致** |
| `mailbox.timezone_id` | 归档文件按哪个时区命名？ | 常用 `China Standard Time` / `UTC` / `Taipei Standard Time`；用系统合法时区 ID |
| `mailbox.archive_folder_name` | PST 里归档文件夹叫什么？ | 自定义即可，例如 `Inbox - Archive` |
| `mailbox.pst_file_prefix` | PST 文件名前缀？ | 只允许字母、数字、`.` `_` `-`；不能含路径分隔符 |
| `capacity.quota_bytes` | 这个邮箱的容量上限是多少**字节**？ | 让使用者从 Exchange / OWA / 管理后台的容量页面读取；不确定时先帮他只读探测，不要替他填一个猜测值 |
| `capacity.trigger_ratio` | 用到多少比例开始清理？ | 常见 0.8；触发线 = 配额 × 该比例 |
| `capacity.retention_hours` | 保留最近多久的邮件？ | 默认 48（保留最近两天）；更早的才会被归档 |
| `storage.data_directory` | 归档 PST 和日志放在哪个**绝对路径**？ | 需为当前 Windows 用户独占可写、非网络映射、非重解析点 |
| `storage.allow_cached_archive` | 是否允许在「完整缓存模式」下归档？ | 建议先保持 `true` 并做单封验收；联机模式下此项无影响 |

关于「连接邮箱」：本工具通过 **Windows 桌面版 Outlook 的已登录会话**读取邮箱，
不需要任何 OAuth 凭据或密码写进配置。请引导使用者确认：

- 当前 Windows 用户已在**经典版（桌面）Outlook** 中配置好目标邮箱；
- 运行时 Outlook 处于打开状态；
- 目标账户连接模式为 `700`（完整缓存）或 `800`（联机）。

若使用者使用的是网页版邮箱、其他邮件客户端，或希望走 Graph / EWS 接口，
请如实告知：本仓库当前只支持桌面版 Outlook，不要假装能支持。

## 第 3 步：通知渠道（**可选项**，重点引导）

先用一句话询问，不要默认替他开启：

> 「这个工具可以把每轮结果推送到消息工具，但**也可以完全不推送**、只在本地
> 记录。你希望启用吗？如果启用，想用哪一个？」

### 选项 A：不启用（默认，完全离线）

保持：

```json
"notify": { "enabled": false, "provider": "none" }
```

此时工具不发起任何网络请求，不要求任何凭据，结果写在本地日志并输出到终端。
**不要因为没连消息工具就认为配置没完成。**

### 选项 B：接入使用者自己的 IM 机器人（推荐做法）

多数团队 IM 都提供「群机器人 / 自定义机器人」的 incoming webhook。引导使用者
去对应平台创建一个机器人，把 webhook 地址交给你，然后按平台设置
`notify.provider`：

| 平台 | `provider` | 使用者需要提供 | 获取方式（引导话术） |
| --- | --- | --- | --- |
| 飞书 / Lark | `feishu` | `webhook_url`，如开启签名还需 `webhook_secret` | 群设置 → 群机器人 → 添加自定义机器人 → 复制 Webhook 地址 |
| 企业微信 | `wecom` | `webhook_url` | 群聊 → 群机器人 → 添加 → 复制 Webhook 地址 |
| 钉钉 | `dingtalk` | `webhook_url`，如开启加签还需 `webhook_secret` | 群设置 → 智能群助手 → 添加自定义机器人 → 复制 Webhook |
| 其他 / 自建服务 | `webhook` | `webhook_url`，必要时用 `payload_template` 自定义报文 | 任意可接收 JSON POST 的 HTTP 端点 |

通用配置示例（占位，请替换成使用者给你的真实地址）：

```json
"notify": {
  "enabled": true,
  "provider": "feishu",
  "on_every_run": true,
  "webhook_url": "https://REPLACE_WITH_YOUR_WEBHOOK",
  "webhook_secret": "",
  "timeout_seconds": 15
}
```

`on_every_run = true` 表示**每轮都发一条结果回执**，包含四项：本次检测容量、
总容量、是否达到清理标准、是否清理。设为 `false` 则只在归档完成或异常时发。
容量按 `1024³` 换算成 GB，与邮箱容量页面显示口径一致。

如果目标服务的报文格式与默认值不同，用 `payload_template` 里写 JSON 模板，
用 `{text}` 作为消息正文占位符，例如 Discord：`{"content":"{text}"}`。

### 选项 C：接入使用者已有的连接器 / CLI / 脚本

如果使用者说「我已经有自己的消息连接器 / 机器人命令 / 内部工具」，**不要**
要求他改用 webhook。改用 `command` 模式：

```json
"notify": {
  "enabled": true,
  "provider": "command",
  "custom_command": ["python", "my_notifier.py"],
  "on_every_run": true
}
```

约定：消息正文会同时通过 **stdin** 和环境变量 `MBM_MESSAGE` 传入；另有
`MBM_IDEMPOTENCY_KEY`（每次执行唯一，可用于去重）、`MBM_PROVIDER`。
命令退出码为 0 视为成功，非 0 会保留事件待下一轮重试。
`custom_command` 也可以写成一条字符串（此时用系统 shell 执行）。

> 这一项是刻意留的通用接口：无论使用者手上是哪个平台的命令行工具、哪种
> 自动化框架、哪种内部机器人，都能接进来，而不需要你改动仓库源码。

### 选项 D：只在终端看

`provider = "stdout"`：不联网，把消息打印到标准输出，方便管道或上层助手读取。

### 引导收尾

配置好后提醒使用者：凭据（`webhook_url` / `webhook_secret` / 自定义命令）
只存在于本机 `config.json`，不要提交、不要贴到公开 issue 里。并询问是否需要
先发一条测试消息（`python run.py --self-test` 只做本地校验，不会真的发消息；
如需真实通道验证，可在配置完成后跑一轮 `--dry-run` 之外的正常轮次）。

## 第 4 步：校验与只读预览（按顺序执行）

```text
python run.py --self-test        # 不需要 config.json；纯本地逻辑自检
python run.py --check-config     # 校验你的配置，不碰 Outlook
python run.py --dry-run          # 只读探测；绝不移动邮件、绝不发通知
```

- `--self-test` 若 `ok` 不为 `true`，先修代码问题，不要继续。
- `--check-config` 若报 `invalid_config`，把 `errors` 逐条翻译给使用者并陪他改。
- `--dry-run` 会输出 `probe`：Outlook 版本、目标账户连接模式、收件箱邮件总数、
  以及 `older_than_cutoff`（符合归档日期条件的封数）。**把这些数字念给使用者
  听，让他确认符合预期**，再进入下一步。
- 若 `probe.ok` 为 `false`（Outlook 没开 / 账户没配好），引导使用者打开桌面版
  Outlook 并确认账户，不要跳过。

## 第 5 步：正式运行与定时调度

单次完整一轮：

```text
python run.py
```

它会在**同一轮内**完成：只读前置检查 → 读容量 → 未达线则结束；达线则把
收件箱顶层所有早于保留期的邮件移入按日期命名的 PST → 同轮发出结果。
标准输出是一行 JSON，含 `ok`、`usage`、`monitor.actions`、`events`、
`summary_notification` 等字段，便于你的上层环境解析。

定时执行请用使用者自己环境里现成的调度方式，例如：

- Windows 任务计划程序（选择「仅当用户登录时运行」，因为需要 Outlook 交互会话）
- Linux/macOS 上的 cron 或 systemd timer（本工具本身依赖 Windows + Outlook，
  仅在其所运行的 Windows 主机上安排）
- 使用者所用 AI 助手 / 自动化平台自带的定时能力

无论哪种方式，调用的都是同一条命令 `python run.py`。**请不要把调度写死成
某个产品专有格式**，也不要替使用者创建他没要求的常驻服务。

## 第 6 步：首次真实归档前的安全建议

1. 建议把 `retention_hours` 临时调大（例如 24×30），让候选集尽量小，先只移动
   一封做验收；或先用 `--dry-run` 看清 `older_than_cutoff`。
2. 归档后打开生成的 PST，确认邮件与附件可读、源邮件已离开收件箱。
3. 再次提醒：**共享邮箱的「移动」等于从服务器移走**，其他共用者将看不到。
   需要恢复时可在 Outlook 中打开该 PST，把邮件移回收件箱。

## 你不应该做的事

- 不要把任何真实邮箱地址、配额数值、路径或 webhook 写进 `src/`、示例配置或文档。
- 不要为了「方便」跳过 `--check-config` / `--dry-run` 直接执行归档。
- 不要替使用者决定是否开启通知、选择哪个平台、或创建机器人。
- 不要修改引擎以绕过它自带的安全检查（私有 ACL、路径校验、逐封副本核对、文件锁）。
- 不要把 `config.json`、PST、`_runtime/`、`events/`、`*.jsonl` 加入版本库。
