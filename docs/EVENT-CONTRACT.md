# 结果事件契约（供自定义消费者）

引擎把**已完整写入**的结果事件放在 `<data_directory>/events/<event_id>.json`。
每个文件一条，UTF-8，文件名即 `event_id`。你可以不用内置通知层，改为让任何
脚本、助手或消息工具读取这个目录自行消费。

## 字段

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `schema_version` | int | 固定为 `1`。**消费者必须校验**，不认识的版本请跳过而不是猜测 |
| `event_id` | string | 唯一标识，格式 `^(archive-)?[a-f0-9]{32}$`；用于去重 |
| `event_type` | string | `archive_completed` 或 `attention_required` |
| `time_utc` | string | ISO 8601 UTC 时间 |
| `mailbox` | string | 目标邮箱 SMTP 地址（来自你的配置） |
| `quota_bytes` | int | 配置的容量上限（字节） |
| `trigger_bytes` | int | `ceil(quota_bytes × trigger_ratio)` |
| `details` | object | 事件专属字段，见下 |

普通用量采样、达线判定、批次进度**只写本地日志**（`monitor-YYYY-MM.jsonl`），
不产生事件文件。

## `event_type = archive_completed`

`details` 字段：

| 字段 | 说明 |
| --- | --- |
| `run_id` | 本轮归档 ID |
| `eligible_count` | 符合日期条件的候选邮件数 |
| `moved_count` | 实际移动并校验通过的封数 |
| `skipped_count` | 移动前复核不通过而跳过的封数 |
| `pst_path` | 本轮 PST 绝对路径 |
| `before_bytes` / `after_bytes` | 归档前后用量读数 |
| `cutoff_utc` | 本轮截止时间（ISO 8601 UTC） |
| `connection_mode` / `usage_source` | 连接模式与读数来源 |

示例（**数值仅为举例，与任何真实邮箱无关**）：

```json
{
  "schema_version": 1,
  "event_id": "0123456789abcdef0123456789abcdef",
  "event_type": "archive_completed",
  "time_utc": "2026-01-01T12:00:00+00:00",
  "mailbox": "user@example.org",
  "quota_bytes": 5368709120,
  "trigger_bytes": 4294967296,
  "details": {
    "run_id": "0123456789abcdef0123456789abcdef",
    "eligible_count": 12,
    "moved_count": 12,
    "skipped_count": 0,
    "pst_path": "C:\\Archive\\output\\mail-260101.pst",
    "before_bytes": 4400000000,
    "after_bytes": 3900000000,
    "cutoff_utc": "2025-12-30T12:00:00+00:00",
    "connection_mode": 700,
    "usage_source": "outlook_cached_or_offline"
  }
}
```

## `event_type = attention_required`

`details.reason` 取值：

| reason | 含义 |
| --- | --- |
| `monitor_unavailable` | 读数连续不可用 |
| `outlook_not_ready` | Outlook 未打开或目标账户不可用 |
| `archive_connection_not_eligible` | 连接模式不满足归档条件 |
| `old_mail_insufficient` | 已无可归档旧邮件，但用量仍超线 |
| `archive_operation_failed` | 归档失败或抛异常（`partial_archive=true` 表示部分归档） |
| `usage_still_above_trigger_after_archive` | 归档后缓存读数仍高于触发线 |
| `move_result_uncertain` | 跨存储移动结果不明 |
| `orchestration_blocked` | 编排层检测到某个节点受阻 |
| `orchestrator_exception` | 编排层未预期异常 |

其余字段按需出现：`cause`、`error_code`、`used_bytes`、`connection_mode`、
`moved_count`、`eligible_count`、`skipped_count`、`run_id`、`pst_path`、
`consecutive_checks`、`partial_archive`。

## 消费建议

1. 只处理 `schema_version == 1` 的文件；解析失败的记录留待人工查看，不要删。
2. 按 `event_id` 维护一份**已送达**记录；确认投递成功后再标记，失败保留重试。
3. 字节换算用 `1024³`，与邮箱容量页面口径一致。
4. 事件文件与 PST、日志一样属于运行产物，不要提交到版本库。

内置的 `src/notifier.py` 与 `src/monitor.py` 就是按这份契约实现的；如果你用
`notify.provider = "command"` 或完全关闭通知，可以自行实现等价消费者。
