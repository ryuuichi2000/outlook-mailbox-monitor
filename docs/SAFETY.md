# 安全设计：防误删、防错移、防卡住

本文件说明工具的安全边界与已实现的防护。它是通用说明，不含任何真实邮箱、
配额或部署环境信息。

## 最重要的语义：移动，不是复制

归档使用 Outlook 的 `MailItem.Move()`。对**共享 / 公用邮箱**，这会把邮件从
服务器邮箱移走，其他共用者不再看到。工具没有任何「删除」调用，但「移动」
本身就是不可逆的业务动作。请在使用前确认这符合你的组织制度。

恢复方式：在 Outlook 中打开生成的 PST，找到归档文件夹，把邮件移回收件箱。

## 一、不会误删 / 错移

源码中**不存在**任何显式删除 API（`.Delete()`、`.Remove()`、`Remove-Item`、
`Clear-Content`、`RemoveStore`、`.Clear()` 等）。唯一的写操作是 `Move()`，
且被多层校验包住：

### 移动前（源邮件）

1. 只处理 `Class = 43`（普通邮件），排除会议请求等其他类型。
2. 只扫描目标邮箱**默认收件箱顶层**，不递归子文件夹、不碰其他文件夹。
3. 按 Exchange 的 UTC 投递时间（`0x0E060040`）判断，必须**严格早于**
   本轮截止时间（现在 − `retention_hours`）。恰好等于截止时间的邮件保留。
4. 用 `Session.GetItemFromID()` 重新取回邮件，核对
   `Parent.EntryID` 与 `Parent.StoreID` 仍是目标收件箱与目标存储。
5. 校验 `Size > 0`；移动前检查磁盘余量并保留 `disk_reserve_bytes`。
6. 移动前再次确认目标账户连接状态仍允许归档。

> 任何一项不满足，该邮件被**跳过**并计数，不会抛错中断整轮。

### 移动后（PST 副本）

1. 用 `Move()` 返回的新 EntryID 从 PST 存储重新打开该邮件。
2. 核对副本的 `Parent.StoreID`、`Parent.EntryID` 指向目标 PST 文件夹。
3. 核对附件数量与移动前一致。
4. 核对 UTC 投递时间与移动前一致。
5. 全部通过才计入 `moved_count`；任一不符即抛出
   `PstDestinationVerificationFailed`，本轮记为部分归档并告警。

### 输出路径

- `data_directory` 必须是绝对路径，且不能是联接 / 符号链接等重解析点。
- 启动时检查目录 ACL：只允许当前用户、SYSTEM、本机管理员；权限过宽即停止
  （`UntrustedOutputAcl`）。
- PST 大小接近 `pst_ceiling_bytes` 时暂停。
- 事件文件名只接受 `^(archive-)?[a-f0-9]{32}$`，拒绝路径穿越。

## 二、不会因某个节点失败而卡住

工具采用**无状态单轮**设计：读用量 → 归档 → 结算 → 通知，全部在一次调用内
完成。它不读写 `state.json`，不产生跨轮的 `Pending` / `InFlight` 状态，也没有
「等下一轮再复核」的延迟，因此不会出现「完成通知拖到第二天」或「卡在待确认
状态需要人工介入」的情况。

| 风险节点 | 处理 |
| --- | --- |
| Outlook 未打开 / 账户未配置 / 连接模式异常 | 只读前置探针当轮判定，**立即告警并终止**，绝不尝试归档 |
| 探针超时 / 归档超时 | 分别由 `probe_timeout_seconds`、`archive_timeout_seconds` 限制；超时即告警 |
| 另一个实例正在运行 | 文件锁检测到重叠，输出 `overlapping_run_skipped` 并告警，不静默跳过 |
| 归档过程中异常 | 捕获后写 `archive_operation_failed` 事件；若已移动部分邮件，详情带 `moved_count`、`eligible_count`、`pst_path`、`partial_archive=true` |
| 达线却没有结算结果 | 编排层检测 `triggered_but_no_settlement` 并告警 |
| 缓存模式读数滞后 | 归档后用量仍高于触发线时，额外发 `usage_still_above_trigger_after_archive`，不谎报已降线 |
| 收件箱没有符合条件的旧邮件但用量仍超线 | 发 `old_mail_insufficient` 告警，不假装成功 |
| 通知发送失败 | 重试一次；仍失败则本轮 `ok=false`，事件保留待下轮按 `event_id` 重试 |
| 编排层未预期异常 | 兜底 `orchestrator_exception` 告警，并输出错误详情 |
| 日志半行 / 状态损坏 | 日志按行追加时自动补换行；单轮模式不依赖持久状态 |

**所有退出路径**都会输出一行 JSON；若启用了 `notify.on_every_run`，还会发一条
结果回执。不存在静默失败。

## 三、容量读数的诚实性

- 用量读取自目标存储的 MAPI 属性 `0x0E080014`，并记录 `connection_mode`、
  `connected`、`archive_allowed`、`usage_source`，便于判断读数来源。
- 完整缓存模式（`700`）的读数可能滞后于服务器；工具会标注来源，并在归档后
  仍超线时额外告警，而不是把「缓存读数下降」当成服务器已释放空间。
- 回执中的容量一律按 `1024³` 换算成 GB。取不到读数时写「未知」，不编造数字。

## 四、运行前的建议验证

1. `python run.py --self-test`：纯本地逻辑自检（配置校验、通知报文构造、
   事件去重、回执四字段、函数绑定）。
2. `python run.py --check-config`：确认配置无占位符残留。
3. `python tests/test_desensitize.py`：确认源码里没有真实值或专有依赖残留。
4. `python run.py --dry-run`：只读普查，确认 Outlook 就绪、账户匹配、候选封数
   符合预期。
5. 首次真实归档：把 `retention_hours` 临时调大以缩小候选集，只移动一封，
   打开 PST 核对邮件与附件，并确认源邮件已离开收件箱。

自检与只读预览都使用临时目录或纯本地计算，不连接真实邮箱、不发送任何消息、
不移动任何邮件。

## 五、仍然存在的边界（使用者需自行判断）

- 只处理**收件箱顶层**，不递归子文件夹。
- 依赖运行时 Outlook 已打开且处于交互会话；工具不会替你启动 Outlook。
- 只覆盖单台机器的并发；同一邮箱请勿在两台机器同时启用归档。
- 附件内容是否可读属于现场验收项；工具只核对附件**数量**与投递时间等元数据。
- PST 的备份、空间维护与长期留存由运行机器负责。
