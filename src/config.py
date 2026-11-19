"""Configuration loading and validation.

Every environment-specific value lives in ``config.json`` (copied from
``config.example.json``). No mailbox address, quota, filesystem path, credential
or vendor-specific identifier is hard-coded in the source code.

This module is plain Python 3 with no third-party dependency.
"""

import copy
import json
import math
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXAMPLE_PATH = os.path.join(ROOT, 'config.example.json')
CONFIG_PATH = os.path.join(ROOT, 'config.json')

GB = 1024 ** 3

PLACEHOLDER_RE = re.compile(
    r'REPLACE_WITH|YOUR_|<[^>]*>', re.IGNORECASE)
SAFE_PREFIX_RE = re.compile(r'^[A-Za-z0-9._-]+$')
EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')

PROVIDERS = ('none', 'stdout', 'webhook', 'feishu', 'wecom', 'dingtalk', 'command')
WEBHOOK_PROVIDERS = ('webhook', 'feishu', 'wecom', 'dingtalk')

# Neutral defaults. A missing config.json or an unedited placeholder is
# rejected by validate() instead of silently running against a wrong mailbox.
DEFAULTS = {
    'mailbox': {
        'smtp_address': '',
        'archive_folder_name': '',
        'pst_file_prefix': '',
        'timezone_id': 'UTC',
    },
    'capacity': {
        'quota_bytes': 0,
        'trigger_ratio': 0.8,
        'retention_hours': 48,
        'disk_reserve_bytes': 128 * 1024 * 1024,
        'pst_ceiling_bytes': 45 * 1024 * 1024 * 1024,
    },
    'storage': {
        'data_directory': '',
        'allow_cached_archive': True,
    },
    'runtime': {
        'powershell_path': 'powershell.exe',
        'launcher_command': [],
        'probe_timeout_seconds': 180,
        'archive_timeout_seconds': 900,
    },
    'notify': {
        'enabled': False,
        'provider': 'none',
        'on_every_run': True,
        'webhook_url': '',
        'webhook_secret': '',
        'payload_template': '',
        'custom_command': '',
        'mention_all': False,
        'timeout_seconds': 15,
    },
}


def _merge(base, override):
    out = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if key.startswith('_'):
            continue
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = _merge(out[key], value)
        else:
            out[key] = value
    return out


def load(path=None):
    """Load and normalise the config. Raises FileNotFoundError / ValueError."""
    path = path or CONFIG_PATH
    if not os.path.isfile(path):
        raise FileNotFoundError(
            'Configuration file not found: %s\n'
            'Copy config.example.json to config.json and fill in your own values.' % path)
    with open(path, 'r', encoding='utf-8-sig') as f:
        raw = json.load(f)
    if not isinstance(raw, dict):
        raise ValueError('The config root must be a JSON object.')
    cfg = _merge(DEFAULTS, raw)
    # Notifications are optional. When disabled, force the no-op provider so
    # no code path can accidentally try to reach a network endpoint.
    if not cfg['notify']['enabled']:
        cfg['notify']['provider'] = 'none'
    return cfg


def validate(cfg):
    """Return (ok, errors). Never raises for user-facing mistakes."""
    errors = []
    m = cfg.get('mailbox', {})
    c = cfg.get('capacity', {})
    s = cfg.get('storage', {})
    r = cfg.get('runtime', {})
    n = cfg.get('notify', {})

    smtp = str(m.get('smtp_address') or '')
    if not smtp:
        errors.append('mailbox.smtp_address is empty')
    elif (PLACEHOLDER_RE.search(smtp) or 'example.com' in smtp.lower()
          or 'example.org' in smtp.lower() or not EMAIL_RE.match(smtp)):
        errors.append('mailbox.smtp_address must be your real SMTP address (placeholder not allowed)')

    for key in ('archive_folder_name', 'pst_file_prefix', 'timezone_id'):
        v = str(m.get(key) or '')
        if not v:
            errors.append('mailbox.%s is empty' % key)
        elif PLACEHOLDER_RE.search(v):
            errors.append('mailbox.%s still contains a placeholder' % key)

    prefix = str(m.get('pst_file_prefix') or '')
    if prefix and not SAFE_PREFIX_RE.match(prefix):
        errors.append('mailbox.pst_file_prefix may only contain A-Z a-z 0-9 . _ - (no path separators)')

    try:
        quota = int(c.get('quota_bytes'))
    except (TypeError, ValueError):
        quota = 0
    if quota <= 0:
        errors.append('capacity.quota_bytes must be a positive integer (bytes); '
                      'read the quota from your Exchange/OWA mailbox settings')

    try:
        ratio = float(c.get('trigger_ratio'))
    except (TypeError, ValueError):
        ratio = -1
    if not (0 < ratio <= 1):
        errors.append('capacity.trigger_ratio must be in (0, 1]')

    try:
        hours = int(c.get('retention_hours'))
    except (TypeError, ValueError):
        hours = 0
    if hours <= 0:
        errors.append('capacity.retention_hours must be a positive integer')

    data_dir = str(s.get('data_directory') or '')
    if not data_dir:
        errors.append('storage.data_directory is empty')
    elif PLACEHOLDER_RE.search(data_dir):
        errors.append('storage.data_directory still contains a placeholder')
    elif not os.path.isabs(data_dir):
        errors.append('storage.data_directory must be an absolute path (e.g. C:\\MailboxArchive\\output)')

    if not str(r.get('powershell_path') or ''):
        errors.append('runtime.powershell_path is empty')
    launcher = r.get('launcher_command') or []
    if launcher and not isinstance(launcher, (list, str)):
        errors.append('runtime.launcher_command must be a JSON array of strings or a string')

    provider = str(n.get('provider') or 'none')
    if provider not in PROVIDERS:
        errors.append('notify.provider must be one of: %s' % ', '.join(PROVIDERS))
    if n.get('enabled'):
        if provider in WEBHOOK_PROVIDERS:
            url = str(n.get('webhook_url') or '')
            if not url.startswith(('http://', 'https://')):
                errors.append('notify.webhook_url must be an http(s) URL when a webhook provider is enabled')
            elif PLACEHOLDER_RE.search(url):
                errors.append('notify.webhook_url still contains a placeholder')
        elif provider == 'command':
            cmd = n.get('custom_command') or ''
            if not cmd:
                errors.append('notify.custom_command is required when provider=command')
            elif PLACEHOLDER_RE.search(str(cmd)):
                errors.append('notify.custom_command still contains a placeholder')

    return (len(errors) == 0), errors


def trigger_bytes(cfg):
    quota = int(cfg['capacity']['quota_bytes'])
    ratio = float(cfg['capacity']['trigger_ratio'])
    return int(math.ceil(quota * ratio))


def gb_display(value):
    """Human-readable size using 1024^3, matching the Exchange/OWA quota page.

    Whole numbers print as '3 GB' instead of '3.00 GB'.
    Returns 'unknown' when the value is not available (never invents a number).
    """
    try:
        v = float(value) / GB
    except (TypeError, ValueError):
        return 'unknown'
    if abs(v - round(v)) < 0.005:
        return '%.0f GB' % v
    return '%.2f GB' % v


def runtime_dir(cfg):
    """Directory for transient launcher output files (safe to delete)."""
    return os.path.join(cfg['storage']['data_directory'], '_runtime')


def events_dir(cfg):
    return os.path.join(cfg['storage']['data_directory'], 'events')
