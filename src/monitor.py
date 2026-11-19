"""One-round orchestrator for the Outlook mailbox capacity monitor.

Design goals (all provider-agnostic, no vendor-specific wording required):

  * One invocation does preflight -> read usage -> archive -> settle -> notify.
    Nothing is deferred to a later run.                                   (R1)
  * The retention cutoff is applied to ALL eligible messages; there is no
    per-run message cap in the daily path.                               (R2)
  * Outlook not open / account missing / bad connection mode -> immediate
    alert on this very round, and no archive is attempted.                (R4)
  * Every blocked node (timeout, lock overlap, partial archive, stalled
    settlement, delivery failure) produces an alert, never a silent exit.
  * Notifications are OPTIONAL. With notify.enabled=false the tool is fully
    offline: it only writes local logs/events and prints one JSON line.

A round receipt (optional, see notify.on_every_run) always carries four fields:
detected usage, total quota, whether the cleanup line was reached, whether
cleanup was performed. Sizes are shown in 1024^3 "GB" to match the Exchange/OWA
quota page. When a reading is unavailable it says "unknown" instead of
inventing a number.
"""

import json
import os
import shutil
import sys
import tempfile
import time
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config
import notifier
import ps_runner

PS1 = os.path.join(config.ROOT, 'src', 'MailboxMonitor.ps1')
PROBE_PS1 = os.path.join(config.ROOT, 'src', 'probe_outlook.ps1')

ALLOWED_MODES = (700, 800)  # 700 = full cached, 800 = online
NOTIFY_TYPES = ('archive_completed', 'attention_required')

REASON_CN = {
    'monitor_unavailable': 'Outlook 读数连续不可用（可能未打开或账户未配置）',
    'outlook_not_ready': 'Outlook 未打开或目标邮箱账户不可用',
    'archive_connection_not_eligible': 'Outlook 连接模式不满足归档条件',
    'old_mail_insufficient': '收件箱已无符合保留期的可归档邮件，但用量仍在触发线以上',
    'archive_operation_failed': '归档操作失败（可能为部分归档）',
    'usage_still_above_trigger_after_archive': '归档已完成，但缓存模式读数仍高于触发线',
    'move_result_uncertain': '跨存储移动结果不明，需人工核对',
    'orchestration_blocked': '本轮某个节点受阻，结果未确认',
    'orchestrator_exception': '编排层发生未预期异常',
}


def now_iso():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


def _paths(cfg):
    data = cfg['storage']['data_directory']
    return {
        'data': data,
        'events': config.events_dir(cfg),
        'delivered': os.path.join(data, 'delivered_events.json'),
        'run_log': os.path.join(data, 'rounds.jsonl'),
        'probe_json': os.path.join(config.runtime_dir(cfg), 'probe.json'),
    }


def append_run_log(cfg, record):
    try:
        path = _paths(cfg)['run_log']
        with open(path, 'a', encoding='utf-8') as f:
            f.write(json.dumps(record, ensure_ascii=False) + '\n')
    except OSError:
        pass


def _send(text, cfg, key, dry_run=False):
    """Send with one retry. A disabled provider is a successful no-op."""
    ok, detail = notifier.send(text, cfg, key, dry_run)
    if not ok and notifier.is_enabled(cfg) and not dry_run:
        ok, detail = notifier.send(text, cfg, key + '-retry', False)
    return bool(ok), detail


def build_round_summary(result, cfg):
    """The four mandatory fields, plus context. Never invents a capacity value."""
    usage = result.get('usage') or {}
    actions = (result.get('monitor') or {}).get('actions') or []
    details = result.get('archive_details')
    problems = result.get('problems') or []
    alerts = result.get('alerts') or []
    consume = result.get('events') or {}

    used = usage.get('used_bytes')
    quota = usage.get('quota_bytes') or cfg['capacity']['quota_bytes']
    trigger = usage.get('trigger_bytes') or config.trigger_bytes(cfg)
    at = usage.get('at_or_above_trigger')
    ratio_pct = '%g' % (float(cfg['capacity']['trigger_ratio']) * 100)

    if at is True:
        standard = '是（已达触发线 %s，即总容量的 %s%%）' % (config.gb_display(trigger), ratio_pct)
    elif at is False:
        standard = '否（触发线 %s，即总容量的 %s%%）' % (config.gb_display(trigger), ratio_pct)
    else:
        standard = '未知（本轮未取得容量读数）'

    if details is not None:
        moved = details.get('moved_count') or 0
        cleaned = ('是，已归档 %s 封到本机 PST' % moved) if moved > 0 \
            else '否（达线但收件箱无符合日期条件的旧邮件）'
    elif 'archive_operation_failed' in actions or 'move_result_uncertain' in actions:
        cleaned = '否/部分（归档未完成，详见告警）'
    elif 'below_trigger' in actions:
        cleaned = '否（未达清理标准，未移动任何邮件）'
    elif result.get('error') in ('outlook_not_ready', 'archive_connection_not_eligible'):
        cleaned = '否（Outlook 未就绪或连接模式不允许，未执行归档）'
    else:
        cleaned = '否（本轮未执行归档）'

    lines = [
        '【邮箱容量监测】本轮执行结果',
        '本次检测容量：%s' % (config.gb_display(used) if used is not None else '未知'),
        '总容量：%s' % config.gb_display(quota),
        '是否达到清理标准：%s' % standard,
        '是否清理：%s' % cleaned,
    ]
    if used is not None:
        try:
            lines.append('容量占比：%.2f%%' % (100.0 * float(used) / float(quota)))
        except (TypeError, ValueError, ZeroDivisionError):
            pass
    if details:
        if details.get('pst_path'):
            lines.append('归档位置：%s' % details['pst_path'])
        if details.get('before_bytes') is not None and details.get('after_bytes') is not None:
            lines.append('归档前后读数：%s → %s' % (
                config.gb_display(details['before_bytes']),
                config.gb_display(details['after_bytes'])))
    if problems:
        lines.append('异常节点：%s' % ', '.join(problems))
    if consume.get('failed'):
        lines.append('事件通知失败：%d 条（已保留待重试）' % consume['failed'])
    if alerts:
        lines.append('已发送告警：%d 条' % len(alerts))
    lines.append('时间(UTC)：%s' % result.get('time_utc', now_iso()))
    return '\n'.join(lines)


def build_text(event, cfg):
    et = event.get('event_type')
    d = event.get('details') or {}
    t = event.get('time_utc', '')
    if et == 'archive_completed':
        lines = ['【邮箱容量监测】归档完成（同轮结算）']
        if 'moved_count' in d:
            extra = '移动邮件：%s 封' % d['moved_count']
            if d.get('eligible_count') is not None:
                extra += '（符合日期条件 %s 封' % d['eligible_count']
                if d.get('skipped_count'):
                    extra += '，跳过 %s 封' % d['skipped_count']
                extra += '）'
            lines.append(extra)
        if d.get('before_bytes') is not None and d.get('after_bytes') is not None:
            lines.append('Outlook 读数：%s → %s' % (
                config.gb_display(d['before_bytes']), config.gb_display(d['after_bytes'])))
        elif 'used_bytes' in d:
            lines.append('Outlook 读数：%s' % config.gb_display(d['used_bytes']))
        if d.get('cutoff_utc'):
            lines.append('清理范围：%s 之前的收件箱顶层邮件' % d['cutoff_utc'])
        if d.get('pst_path'):
            lines.append('归档位置：%s' % d['pst_path'])
        lines.append('时间(UTC)：%s' % t)
    else:
        reason = d.get('reason', '(未注明)')
        lines = ['【邮箱容量监测】需人工关注',
                 '原因：%s' % REASON_CN.get(reason, reason)]
        if d.get('partial_archive'):
            lines.append('注意：本轮为部分归档，已成功移动的邮件保存在本机 PST，剩余邮件仍在收件箱。')
        for k in ('cause', 'error_code', 'consecutive_checks', 'used_bytes',
                  'connection_mode', 'moved_count', 'eligible_count',
                  'skipped_count', 'run_id', 'pst_path'):
            v = d.get(k)
            if v in (None, ''):
                continue
            lines.append('%s：%s' % (k, config.gb_display(v) if k == 'used_bytes' else v))
        lines.append('时间(UTC)：%s' % t)
    return '\n'.join(lines)


def notify_alert(reason, lines, key, result, cfg, dry_run=False):
    text = '\n'.join(['【邮箱容量监测】需人工关注',
                      '原因：%s' % REASON_CN.get(reason, reason)] +
                     list(lines) + ['时间(UTC)：%s' % now_iso()])
    ok, detail = _send(text, cfg, key, dry_run)
    result.setdefault('alerts', []).append(
        {'reason': reason, 'sent': ok, 'detail': detail})
    return ok


def load_delivered(path):
    if not os.path.exists(path):
        return set()
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return set(json.load(f).get('delivered', []))
    except (ValueError, OSError):
        return set()


def save_delivered(path, delivered):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump({'delivered': sorted(delivered), 'updated_utc': now_iso()},
                  f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def consume_events(cfg, dry_run=False, events_dir=None, delivered_path=None):
    """Read events, dedupe by event_id, notify only result-class events."""
    p = _paths(cfg)
    events_dir = events_dir or p['events']
    delivered_path = delivered_path or p['delivered']
    delivered = load_delivered(delivered_path)
    pending, skipped_schema = [], []
    if os.path.isdir(events_dir):
        for name in sorted(os.listdir(events_dir)):
            if not name.endswith('.json'):
                continue
            path = os.path.join(events_dir, name)
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    ev = json.load(f)
            except (ValueError, OSError) as e:
                skipped_schema.append('%s: %s' % (name, e))
                continue
            if ev.get('schema_version') != 1:
                skipped_schema.append('%s: schema_version != 1' % name)
                continue
            eid = ev.get('event_id') or os.path.splitext(name)[0]
            if eid in delivered:
                continue
            if ev.get('event_type') not in NOTIFY_TYPES:
                delivered.add(eid)
                continue
            pending.append((eid, ev))

    notified, failed = [], []
    for eid, ev in pending:
        ok, detail = _send(build_text(ev, cfg), cfg, eid, dry_run)
        if ok:
            delivered.add(eid)
            notified.append({'event_id': eid, 'event_type': ev.get('event_type'), 'result': detail})
        else:
            failed.append({'event_id': eid, 'event_type': ev.get('event_type'), 'error': detail})

    if pending or notified or failed:
        save_delivered(delivered_path, delivered)
    return {
        'events_dir': events_dir,
        'already_delivered': len(delivered) - len(notified),
        'to_notify': len(pending),
        'notified': len(notified),
        'failed': len(failed),
        'skipped_schema': skipped_schema,
        'notified_detail': notified,
        'failed_detail': failed,
    }


def finish(result, cfg, dry_run=False):
    """Every exit path ends here. Sends the optional round receipt."""
    if 'summary_notification' not in result:
        if cfg['notify'].get('on_every_run', True) and not dry_run:
            text = build_round_summary(result, cfg)
            key = 'summary-%s' % time.time_ns()
            ok, detail = _send(text, cfg, key, False)
            result['summary_notification'] = {
                'sent': bool(ok), 'detail': detail, 'text': text}
        else:
            result['summary_notification'] = {
                'sent': True,
                'detail': 'disabled_or_dry_run',
                'text': build_round_summary(result, cfg)}
    sn = result.get('summary_notification') or {}
    # A disabled / not-configured channel is NOT a failure.
    if not sn.get('sent') and notifier.is_enabled(cfg):
        result['ok'] = False
        result.setdefault('problems', []).append('summary_notification_failed')
    append_run_log(cfg, result)
    print(json.dumps(result, ensure_ascii=False))


def _engine_args(cfg):
    m, c, s = cfg['mailbox'], cfg['capacity'], cfg['storage']
    args = [
        '-Mode', 'Archive', '-DailyOnce',
        '-MailboxSmtp', str(m['smtp_address']),
        '-QuotaBytes', str(int(c['quota_bytes'])),
        '-TriggerRatio', str(float(c['trigger_ratio'])),
        '-DataDirectory', str(s['data_directory']),
        '-ArchiveFolderName', str(m['archive_folder_name']),
        '-PstFilePrefix', str(m['pst_file_prefix']),
        '-TimeZoneId', str(m['timezone_id']),
        '-RetentionHours', str(int(c['retention_hours'])),
        '-DiskReserveBytes', str(int(c['disk_reserve_bytes'])),
        '-PstCeilingBytes', str(int(c['pst_ceiling_bytes'])),
    ]
    if s.get('allow_cached_archive'):
        args.append('-AllowCachedArchive')
    return args


def _probe_args(cfg, out_file):
    return ['-MailboxSmtp', str(cfg['mailbox']['smtp_address']),
            '-RetentionHours', str(int(cfg['capacity']['retention_hours'])),
            '-OutFile', out_file]


def run(cfg, dry_run=False):
    p = _paths(cfg)
    r = cfg['runtime']
    day_key = time.strftime('%Y%m%d', time.gmtime())
    result = {'ok': True, 'version': '1.0', 'time_utc': now_iso(),
              'notify_enabled': notifier.is_enabled(cfg),
              'dry_run': bool(dry_run)}
    try:
        # ---- 1. READ-ONLY preflight: Outlook not open -> alert immediately ----
        probe = ps_runner.run_ps(
            PROBE_PS1, ps_args=_probe_args(cfg, p['probe_json']), cfg=cfg,
            tag='probe', timeout=int(r['probe_timeout_seconds']),
            result_file=p['probe_json'])
        pres = probe.get('result') or {}
        target = pres.get('target') or {}
        result['probe'] = {
            'ok_child': probe['ok_child'],
            'timed_out': probe['timed_out'],
            'elapsed': probe['elapsed'],
            'reason': pres.get('reason'),
            'outlook_version': pres.get('outlook_version'),
            'connection_mode': target.get('connection_mode'),
            'inbox': pres.get('inbox'),
        }
        if probe['timed_out'] or pres.get('ok') is not True:
            notify_alert('outlook_not_ready', [
                '探测结果：%s' % (pres.get('reason') or probe['launcher_stderr'] or 'no result'),
                '错误码：%s' % pres.get('error_code', ''),
                '本轮未执行任何归档动作。'
            ], 'outlook-not-ready-' + day_key, result, cfg, dry_run)
            result['ok'] = False
            result['error'] = 'outlook_not_ready'
            result['summary'] = 'Outlook 未就绪，已告警，本轮未执行归档。'
            finish(result, cfg, dry_run)
            return 1

        mode = target.get('connection_mode')
        if mode not in ALLOWED_MODES:
            notify_alert('archive_connection_not_eligible', [
                '连接模式：%s（允许 700 完整缓存 / 800 联机）' % mode,
                '本轮未执行任何归档动作。'
            ], 'connection-' + day_key, result, cfg, dry_run)
            result['ok'] = False
            result['error'] = 'archive_connection_not_eligible'
            result['summary'] = 'Outlook 连接模式不满足归档条件，已告警。'
            finish(result, cfg, dry_run)
            return 1

        if dry_run:
            result['summary'] = '只读预览完成，未执行归档、未发送通知。'
            result['preview_receipt'] = build_round_summary(result, cfg)
            result['summary_notification'] = {
                'sent': True, 'detail': 'dry_run',
                'text': result['preview_receipt']}
            append_run_log(cfg, result)
            print(json.dumps(result, ensure_ascii=False))
            return 0

        # ---- 2. Same-round stateless archive ----
        arch = ps_runner.run_ps(
            PS1, ps_args=_engine_args(cfg), cfg=cfg, tag='daily',
            timeout=int(r['archive_timeout_seconds']))
        actions = [x.get('action') for x in arch['records']
                   if isinstance(x, dict) and x.get('action')]
        usage = None
        for x in reversed(arch['records']):
            if isinstance(x, dict) and x.get('action') == 'usage_sample':
                usage = x.get('details')
                break
        result['monitor'] = {
            'ok_child': arch['ok_child'],
            'timed_out': arch['timed_out'],
            'child_exit': arch['child_exit'],
            'launcher_rc': arch['launcher_rc'],
            'elapsed': arch['elapsed'],
            'actions': actions,
            'stderr': arch['stderr_raw'][:400],
        }
        result['usage'] = usage
        result['archive_details'] = None
        for x in arch['records']:
            if isinstance(x, dict) and x.get('action') == 'archive_completed':
                result['archive_details'] = x.get('details')

        # ---- 3. Same-round event delivery ----
        consume = consume_events(cfg, dry_run=False)
        result['events'] = consume

        # ---- 4. Blocked-node detection: never exit silently on a stuck node ----
        problems = []
        if arch['timed_out']:
            problems.append('monitor_timeout_%ss' % r['archive_timeout_seconds'])
        if arch['child_exit'] not in (0, 2):
            problems.append('monitor_exit_%s' % arch['child_exit'])
        for a in ('overlapping_run_skipped', 'state_invalid_or_unreadable',
                  'move_result_uncertain', 'archive_operation_failed'):
            if a in actions:
                problems.append(a)
        if usage and usage.get('at_or_above_trigger') is True \
                and 'below_trigger' not in actions \
                and 'archive_completed' not in actions \
                and 'archive_operation_failed' not in actions:
            problems.append('triggered_but_no_settlement')
        result['problems'] = problems

        if problems and consume['notified'] == 0 and not result.get('alerts'):
            notify_alert('orchestration_blocked', [
                '异常节点：%s' % ', '.join(problems),
                '脚本输出：%s' % (arch['stdout_raw'][-300:] or '(empty)'),
                '本轮结果未确认，请查看输出目录下的 monitor-*.jsonl 日志。'
            ], 'blocked-' + day_key, result, cfg)

        result['ok'] = (not problems) and consume['failed'] == 0 and all(
            a['sent'] for a in result.get('alerts', []))

        if consume['notified']:
            kinds = sorted(set(x['event_type'] for x in consume['notified_detail']))
            result['summary'] = '本轮同轮结算完成，已通知 %d 条结果事件（%s）。' % (
                consume['notified'], ', '.join(kinds))
        elif consume['failed']:
            result['summary'] = '本轮有 %d 条事件通知失败，已保留待下轮重试。' % consume['failed']
        elif result.get('alerts'):
            result['summary'] = '本轮检测到异常，已发送 %d 条告警。' % len(result['alerts'])
        elif 'below_trigger' in actions:
            result['summary'] = '本轮未达清理线，未执行归档。'
        else:
            result['summary'] = '本轮完成，无结果事件、无告警。'

        finish(result, cfg)
        return 0 if result['ok'] else 1

    except Exception as e:
        try:
            notify_alert('orchestrator_exception', [
                '错误：%s: %s' % (type(e).__name__, e)
            ], 'exception-' + day_key, result, cfg, dry_run)
        except Exception:
            pass
        result['ok'] = False
        result['error'] = 'orchestrator_exception'
        result['problems'] = ['orchestrator_exception', '%s: %s' % (type(e).__name__, e)]
        result['summary'] = '编排层发生未预期异常，已尝试告警。'
        finish(result, cfg, dry_run)
        return 1


def self_test():
    """Physically isolated. No config.json, no Outlook, no network, no mail."""
    tmp = tempfile.mkdtemp(prefix='mailbox-monitor-selftest-')
    try:
        fake = {
            'mailbox': {'smtp_address': 'user@mail.invalid',
                        'archive_folder_name': 'Inbox - Test',
                        'pst_file_prefix': 'test', 'timezone_id': 'UTC'},
            'capacity': {'quota_bytes': 5 * config.GB, 'trigger_ratio': 0.8,
                         'retention_hours': 48,
                         'disk_reserve_bytes': 128 * 1024 * 1024,
                         'pst_ceiling_bytes': 45 * config.GB},
            'storage': {'data_directory': tmp, 'allow_cached_archive': True},
            'runtime': {'powershell_path': 'powershell.exe', 'launcher_command': [],
                        'probe_timeout_seconds': 180, 'archive_timeout_seconds': 900},
            'notify': {'enabled': True, 'provider': 'webhook',
                       'on_every_run': True,
                       'webhook_url': 'https://example.invalid/hook',
                       'webhook_secret': '', 'payload_template': '',
                       'custom_command': '', 'mention_all': False,
                       'timeout_seconds': 15},
        }

        # -- 1. config validation rejects placeholders / zero quota -------------
        bad = json.loads(json.dumps(fake))
        bad['capacity']['quota_bytes'] = 0
        bad['mailbox']['smtp_address'] = 'REPLACE_WITH_YOUR_MAILBOX@example.com'
        _, bad_errors = config.validate(bad)
        ok_config = (len(config.validate(fake)[1]) == 0 and len(bad_errors) >= 2)

        # -- 2. notification payloads (pure, no network) -----------------------
        payload_checks = {}
        wf = json.loads(json.dumps(fake)); wf['notify'].update(
            provider='wecom', webhook_url='https://example.invalid/wecom', mention_all=True)
        payload_checks['wecom_mention'] = (
            notifier.build_request('hello', wf)['body']['text']['mentioned_list'] == ['@all'])
        ff = json.loads(json.dumps(fake)); ff['notify'].update(
            provider='feishu', webhook_url='https://example.invalid/feishu',
            webhook_secret='secret-value')
        fbody = notifier.build_request('hello', ff)['body']
        payload_checks['feishu_signed'] = ('sign' in fbody and 'timestamp' in fbody
                                           and fbody['content']['text'] == 'hello')
        df = json.loads(json.dumps(fake)); df['notify'].update(
            provider='dingtalk', webhook_url='https://example.invalid/ding?token=x',
            webhook_secret='secret-value')
        durl = notifier.build_request('hello', df)['url']
        payload_checks['dingtalk_signed'] = ('timestamp=' in durl and 'sign=' in durl)
        gf = json.loads(json.dumps(fake)); gf['notify'].update(
            provider='webhook', payload_template='{"content":"{text}"}')
        payload_checks['generic_template'] = (
            notifier.build_request('a "quoted" line', gf)['body']['content'] == 'a "quoted" line')

        off = json.loads(json.dumps(fake)); off['notify']['enabled'] = False
        payload_checks['disabled_noop'] = (notifier.send('x', off) == (True, 'disabled'))
        payload_checks['dry_run_noop'] = (notifier.send('x', fake, dry_run=True) == (True, 'dry_run'))

        # -- 3. event consumption + dedupe (isolated temp dir) ------------------
        ev_dir = os.path.join(tmp, 'events')
        os.makedirs(ev_dir)
        state = os.path.join(tmp, 'delivered.json')
        samples = [
            {'schema_version': 1, 'event_id': 'a' * 32, 'event_type': 'archive_completed',
             'time_utc': '2026-01-01T00:00:00+00:00',
             'details': {'moved_count': 3, 'eligible_count': 3, 'skipped_count': 0,
                         'before_bytes': 4 * config.GB, 'after_bytes': 3 * config.GB,
                         'cutoff_utc': '2025-12-30T00:00:00+00:00',
                         'pst_path': os.path.join(tmp, 'test-260101.pst')}},
            {'schema_version': 1, 'event_id': 'b' * 32, 'event_type': 'attention_required',
             'time_utc': '2026-01-01T00:05:00+00:00',
             'details': {'reason': 'archive_operation_failed', 'partial_archive': True,
                         'moved_count': 1, 'cause': 'PstDiskSpaceLow'}},
            {'schema_version': 1, 'event_id': 'c' * 32, 'event_type': 'usage_sample',
             'time_utc': '2026-01-01T00:06:00+00:00', 'details': {'used_bytes': 1 * config.GB}},
        ]
        for s in samples:
            with open(os.path.join(ev_dir, s['event_id'] + '.json'), 'w', encoding='utf-8') as f:
                json.dump(s, f, ensure_ascii=False)
        r1 = consume_events(fake, dry_run=True, events_dir=ev_dir, delivered_path=state)
        r2 = consume_events(fake, dry_run=True, events_dir=ev_dir, delivered_path=state)

        # -- 4. mandatory round receipt with four fields ------------------------
        fake_round = {
            'time_utc': '2026-01-01T00:00:00+00:00',
            'usage': {'used_bytes': int(2.5 * config.GB),
                      'quota_bytes': 5 * config.GB,
                      'trigger_bytes': 4 * config.GB,
                      'at_or_above_trigger': False},
            'monitor': {'actions': ['usage_sample', 'below_trigger']},
            'events': {'failed': 0}, 'problems': [], 'alerts': [],
            'archive_details': None,
        }
        receipt = build_round_summary(fake_round, fake)
        static = {
            'functions_bound': all(callable(globals().get(n)) for n in (
                'build_text', 'notify_alert', 'consume_events', 'finish', 'run',
                'build_round_summary', '_engine_args', '_probe_args')),
            'ps_runner_bound': hasattr(ps_runner, 'run_ps'),
            'notifier_bound': hasattr(notifier, 'send') and hasattr(notifier, 'build_request'),
            'reasons_covered': all(k in REASON_CN for k in (
                'outlook_not_ready', 'archive_connection_not_eligible',
                'old_mail_insufficient', 'archive_operation_failed',
                'usage_still_above_trigger_after_archive',
                'orchestration_blocked', 'orchestrator_exception')),
            'text_fields': ('移动邮件：3 封' in build_text(samples[0], fake)
                            and '部分归档' in build_text(samples[1], fake)),
            'receipt_fields': all(k in receipt for k in (
                '本次检测容量', '总容量', '是否达到清理标准', '是否清理')),
            'receipt_values': ('2.50 GB' in receipt and '5 GB' in receipt
                               and '4 GB' in receipt),
            'receipt_no_archive': '是否清理：否' in receipt,
            'gb_whole': config.gb_display(5 * config.GB) == '5 GB',
            'unknown_capacity': '未知' in build_round_summary({'usage': {}}, fake),
            'engine_no_cap': '-DailyOnce' in _engine_args(fake),
            'engine_params': all(k in _engine_args(fake) for k in (
                '-MailboxSmtp', '-QuotaBytes', '-TriggerRatio', '-RetentionHours')),
        }

        passed = (ok_config and all(payload_checks.values()) and all(static.values())
                  and r1['to_notify'] == 2 and r1['notified'] == 2 and r1['failed'] == 0
                  and r2['to_notify'] == 0 and r2['notified'] == 0)
        result = {
            'ok': passed, 'self_test': True,
            'note': 'physical isolation; no config.json, no Outlook, no network, no real mail',
            'config_validation': {'valid_cfg_ok': len(config.validate(fake)[1]) == 0,
                                  'bad_cfg_errors': len(bad_errors)},
            'payload_checks': payload_checks,
            'round1': {k: r1[k] for k in ('to_notify', 'notified', 'failed')},
            'round2': {k: r2[k] for k in ('to_notify', 'notified', 'failed')},
            'static_assertions': static,
            'sample_receipt': receipt,
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if passed else 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
