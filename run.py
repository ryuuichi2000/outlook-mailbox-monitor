#!/usr/bin/env python3
"""Single entry point for the Outlook mailbox capacity monitor.

    python run.py --check-config   validate config.json, touch nothing
    python run.py --self-test      isolated logic checks, no config needed
    python run.py --dry-run        READ-ONLY Outlook probe, no archiving
    python run.py                  one full monitoring + archiving round

Requires: Windows, classic Outlook (desktop), Windows PowerShell 5.1, Python 3.9+.
No third-party Python packages are needed for the default path.
"""

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src'))
import config
import monitor


def _force_utf8_stdio():
    """Make stdout/stderr emit UTF-8 regardless of the console code page.

    Windows consoles default to a legacy code page (cp1252, cp936, ...), while
    every message this tool prints is UTF-8 JSON that may contain Chinese text.
    Without this the process dies with UnicodeEncodeError on such machines --
    e.g. a CI runner or a machine with a non-Chinese locale. errors='replace'
    keeps a diagnostic message readable instead of raising.
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except (AttributeError, ValueError):
            # Python < 3.7, or a stream that cannot be reconfigured (pipe wrappers).
            pass


def main():
    _force_utf8_stdio()
    ap = argparse.ArgumentParser(
        description='Outlook mailbox capacity monitor / local PST archiver')
    ap.add_argument('--config', default=None,
                    help='path to config.json (default: next to this script)')
    ap.add_argument('--check-config', action='store_true',
                    help='validate the configuration and exit; does not touch Outlook')
    ap.add_argument('--self-test', action='store_true',
                    help='run isolated self checks; no config.json or Outlook required')
    ap.add_argument('--dry-run', action='store_true',
                    help='read-only Outlook probe; never moves mail or sends messages')
    args = ap.parse_args()

    if args.self_test:
        return monitor.self_test()

    if args.config and not os.path.isfile(args.config):
        print(json.dumps({'ok': False, 'error': 'config_not_found',
                          'message': 'Config file not found: %s' % args.config},
                         ensure_ascii=False))
        return 2
    try:
        cfg = config.load(args.config)
    except (FileNotFoundError, ValueError) as e:
        print(json.dumps({'ok': False, 'error': 'config', 'message': str(e)},
                         ensure_ascii=False))
        return 2

    ok, errors = config.validate(cfg)
    if not ok:
        print(json.dumps({'ok': False, 'error': 'invalid_config', 'errors': errors,
                          'hint': 'Copy config.example.json to config.json and fill in '
                                  'your own values; see docs/SETUP.md and AGENT.md.'},
                         ensure_ascii=False))
        return 2

    if args.check_config:
        print(json.dumps({
            'ok': True, 'check': 'config',
            'config_path': args.config or config.CONFIG_PATH,
            'quota_bytes': cfg['capacity']['quota_bytes'],
            'trigger_bytes': config.trigger_bytes(cfg),
            'quota_display': config.gb_display(cfg['capacity']['quota_bytes']),
            'trigger_display': config.gb_display(config.trigger_bytes(cfg)),
            'retention_hours': cfg['capacity']['retention_hours'],
            'notify_enabled': bool(cfg['notify']['enabled']),
            'notify_provider': cfg['notify']['provider'],
            'data_directory': cfg['storage']['data_directory'],
        }, ensure_ascii=False, indent=2))
        return 0

    return monitor.run(cfg, dry_run=args.dry_run)


if __name__ == '__main__':
    sys.exit(main())
