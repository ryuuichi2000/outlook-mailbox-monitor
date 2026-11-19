# Outlook 邮箱容量监测与本地归档工具

[![self-test](https://github.com/ryuuichi2000/outlook-mailbox-monitor/actions/workflows/self-test.yml/badge.svg)](https://github.com/ryuuichi2000/outlook-mailbox-monitor/actions/workflows/self-test.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

一个**配置驱动、可离线运行、通知可选**的小工具：定期读取桌面版 Outlook 中某个
邮箱的容量，达到设定阈值时，把收件箱顶层早于保留期的邮件**移动**到本地 PST
文件，并可选择把结果推送到你自己的消息工具。

原为某个具体邮箱定制，现已做成通用开源版本：

- **零硬编码**：源码、示例配置、文档中不含任何真实邮箱地址、配额、路径、
  人名、机器人 ID 或凭据；全部通过 `config.json` 提供。
- **通知可选**：默认完全离线，不发起任何网络请求。要通知时可选飞书 / 企业
  微信 / 钉钉 / 通用 webhook / 任意本地命令 / 仅终端输出。
- **不绑定平台**：没有依赖任何特定 AI 助手、自动化产品或调度器的专有格式，
  命令行、任务计划、cron 或你所用的助手定时能力都可以直接调用。
- **附带接入引导**：[AGENT.md](AGENT.md) 专门写给拿到本仓库后负责协助配置的
  AI 助手，引导它逐项询问使用者并帮忙接入使用者自己的连接器。

---

## ⚠️ 使用前必读

归档动作是 **移动（Move）**，不是复制。对**共享 / 公用邮箱**而言，邮件会从
服务器邮箱移走，其他共用者将不再看到该邮件。若这不符合你的预期，请不要启用
归档。需要恢复时，可在 Outlook 中打开本地 PST，把邮件移回收件箱。

首次使用请务必按顺序执行 `--self-test` → `--check-config` → `--dry-run`，并先做
单封验收。工具不会删除任何邮件：它只执行移动，且在移动前逐封校验，在移动后
逐封核对 PST 副本的父文件夹、附件数与投递时间（详见 [docs/SAFETY.md](docs/SAFETY.md)）。

---

## 运行环境

| 依赖 | 要求 |
| --- | --- |
| 操作系统 | Windows |
| 邮件客户端 | **经典版（桌面）Outlook**，且目标邮箱已配置在该 Windows 用户会话中、运行时保持打开 |
| Shell | Windows PowerShell 5.1（Desktop 版，STA） |
| Python | 3.9+，**默认路径无需任何第三方包** |
| 权限 | 运行用户需对输出目录有独占写权限 |

网页版邮箱、其他邮件客户端、Graph/EWS 接口当前**不支持**。

## 快速开始

```text
# 1. 生成配置
cp config.example.json config.json        # Windows: copy config.example.json config.json

# 2. 编辑 config.json，填入你自己的邮箱、配额、输出目录等
#    （逐项含义见 docs/SETUP.md，或让 AI 助手按 AGENT.md 引导你）

# 3. 校验
python run.py --self-test                  # 纯本地逻辑自检，不需要 config.json
python run.py --check-config               # 校验配置，不碰 Outlook
python run.py --dry-run                    # 只读探测，绝不移动邮件

# 4. 单次完整一轮
python run.py
```

标准输出是一行 JSON，例如：

```json
{"ok": true, "usage": {"used_bytes": 1234567890}, "monitor": {"actions": ["usage_sample", "below_trigger"]}, "notify_enabled": false, "summary": "本轮未达清理线，未执行归档。"}
```

## 配置一览

完整字段与含义见 [docs/SETUP.md](docs/SETUP.md)。最关键几项：

| 字段 | 含义 |
| --- | --- |
| `mailbox.smtp_address` | 目标邮箱完整 SMTP 地址 |
| `capacity.quota_bytes` | 邮箱容量上限（**字节**，从 Exchange/OWA 容量页面读取） |
| `capacity.trigger_ratio` | 触发清理的比例，触发线 = 配额 × 比例 |
| `capacity.retention_hours` | 保留最近多少小时；更早的邮件才会被归档 |
| `storage.data_directory` | PST、日志、事件的输出绝对路径 |
| `notify.enabled` | 是否启用通知，**默认 `false`** |
| `notify.provider` | `none` / `stdout` / `webhook` / `feishu` / `wecom` / `dingtalk` / `command` |
| `notify.on_every_run` | 是否每轮都发一条结果回执 |

> 换算口径：容量以 `1024³` 换算成 GB 显示，与邮箱容量页面一致（例如整数配额
> 显示为整数 GB）。触发线按配额 × 比例向上取整。**示例中的任何数字都只是举例，
> 与任何真实邮箱无关。**

## 通知：可选，不接也能用

不启用通知时，工具只写本地日志（`monitor-YYYY-MM.jsonl`、`rounds.jsonl`）
和结果事件（`events/*.json`），并在终端输出一行 JSON。启用方式：

```json
"notify": {
  "enabled": true,
  "provider": "feishu",
  "on_every_run": true,
  "webhook_url": "https://YOUR_WEBHOOK_URL",
  "webhook_secret": "",
  "timeout_seconds": 15
}
```

每轮回执包含四项：本次检测容量、总容量、是否达到清理标准、是否清理。
归档完成与异常会另外单独通知，按 `event_id` 去重，失败保留待下轮重试。

如果你已有自己的机器人 / 连接器 / CLI，用 `provider = "command"` 即可接入，
无需改源码：

```json
"notify": { "enabled": true, "provider": "command", "custom_command": ["python", "my_notifier.py"] }
```

消息正文通过 stdin 与环境变量 `MBM_MESSAGE` 传入，另有
`MBM_IDEMPOTENCY_KEY`、`MBM_PROVIDER`。详见 [docs/SETUP.md](docs/SETUP.md)。

## 定时执行

用你环境中现成的调度方式调用同一条命令即可：

- Windows 任务计划程序：程序 `python`，参数 `run.py`，起始目录设为仓库根目录，
  勾选「仅当用户登录时运行」（需要 Outlook 交互会话）。
- 你所用 AI 助手 / 自动化平台的定时能力。
- 其他能定期执行命令的调度器。

工具自带文件锁，同一邮箱请只在一台机器上启用归档。

## 目录结构

```text
├── run.py                    唯一入口
├── config.example.json       示例配置（全占位符，无真实值）
├── config.json               你的配置（.gitignore 忽略，不要提交）
├── AGENT.md                  给 AI 助手的接入引导
├── README.md                 本文件
├── src/
│   ├── config.py             配置加载 / 校验 / 容量换算
│   ├── ps_runner.py          通用 PowerShell 启动与 JSON-lines 解析
│   ├── notifier.py           可选通知层（纯标准库）
│   ├── monitor.py            同轮编排、每轮回执、防卡住兜底
│   ├── MailboxMonitor.ps1    归档引擎（参数化）
│   └── probe_outlook.ps1     只读前置探针（参数化）
├── docs/
│   ├── SETUP.md              配置与部署详解
│   ├── SAFETY.md             防误删 / 防卡住设计
│   ├── EVENT-CONTRACT.md     结果事件格式（供自定义消费者）
│   └── TROUBLESHOOTING.md    常见问题
└── tests/
    └── test_desensitize.py   脱敏与通用性自检
```

## 自检

```text
python run.py --self-test          # 逻辑自检：配置校验、通知报文、事件去重、回执四字段
python tests/test_desensitize.py   # 扫描源码，确认无真实值与专有依赖残留
```

两个自检都是纯本地、物理隔离，不连接 Outlook、不发送任何消息、不移动邮件。

## 隐私与脱敏声明

本仓库源码与文档**刻意不含**任何真实邮箱地址、公司域名、配额数值、用户路径、
人名、机器人 ID 或凭据。示例配置中的 `REPLACE_WITH_*`、`example.com` 都是占位符，
配置校验会拒绝未替换的占位符。运行产物（PST、日志、事件、`config.json`、
`_runtime/`）全部由 `.gitignore` 忽略，且只保存在你本机。

## 许可

MIT，见 [LICENSE](LICENSE)。使用前请自行确认符合你所在组织关于邮件归档、
数据留存与共享邮箱操作的制度；本工具不提供任何合规保证。
