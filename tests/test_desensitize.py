#!/usr/bin/env python3
"""Read-only desensitization / portability check for this repository.

Run from anywhere:  python tests/test_desensitize.py

Verifies that no real mailbox, quota, user path, person, robot id, credential,
vendor-specific tool name or product-specific wording leaked into the shipped
sources and docs. It also checks that every .ps1 file is pure ASCII (Windows
PowerShell 5.1 mis-decodes BOM-less UTF-8) and that the example config really
contains only placeholders.

The tests/ directory (and .git/, caches) is excluded from the scan so this
file's own banned-token list cannot match itself.
"""

import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

EXCLUDE_DIRS = {
    '.git', 'tests', '__pycache__', '_runtime', 'events',
    '.venv', 'venv', '.idea', '.vscode', '.pytest_cache',
}
SCAN_EXT = ('.py', '.ps1', '.md', '.json', '.txt', '.yml', '.yaml',
            '.toml', '.cfg', '.ini', '.gitignore', '.sh', '.psd1', '.psm1')

# Built from fragments so this file never contains a complete real identifier.
BANNED = [
    '17173' + '-jkzx',
    'cyou' + '-inc.com',
    'ou_2c11' + 'a630dafa64e0ab9c4a1fb544359d',
    'cli_' + 'aaeade0039385bb3',
    'bai' + 'yurou',
    'd:' + '\\mailbox output',
    'd:' + '\\mailbox-monitor',
    'medium' + '_launch.py',
    'lark' + '-cli',
    'work' + 'buddy',
    'present' + '_files',
    'automation' + '_update',
    'high ' + 'il',
    '2,576,' + '980,378',
    '2.40 ' + 'gb',
    '3221225' + '472',
    '2576980' + '378',
    '1.51 ' + 'gb',
    '1.48 ' + 'gb',
    '李' + '龙一',
]
BANNED = [b.lower() for b in BANNED]

USER_PATH_RE = re.compile(r'c:\\users\\[A-Za-z0-9._-]+', re.IGNORECASE)


def iter_files():
    for base, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
        for name in files:
            if name.endswith(SCAN_EXT) or name == '.gitignore' or name == 'LICENSE':
                yield os.path.join(base, name)


def scan():
    findings = []
    scanned = 0
    for path in iter_files():
        rel = os.path.relpath(path, ROOT)
        try:
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                text = f.read()
        except OSError:
            continue
        scanned += 1
        low = text.lower()
        for token in BANNED:
            if token in low:
                for lineno, line in enumerate(text.splitlines(), 1):
                    if token in line.lower():
                        findings.append({'file': rel, 'line': lineno,
                                        'token': token,
                                        'excerpt': line.strip()[:160]})
        for m in USER_PATH_RE.finditer(text):
            line = text[:m.start()].count('\n') + 1
            findings.append({'file': rel, 'line': line,
                             'token': 'hard-coded user profile path',
                             'excerpt': m.group(0)})
    return scanned, findings


def ps1_ascii():
    bad = []
    src = os.path.join(ROOT, 'src')
    for name in sorted(os.listdir(src)):
        if not name.endswith('.ps1'):
            continue
        with open(os.path.join(src, name), 'rb') as f:
            data = f.read()
        non_ascii = sum(1 for b in data if b > 127)
        bad.append({'file': 'src/' + name, 'non_ascii_bytes': non_ascii})
    return bad


def example_config():
    path = os.path.join(ROOT, 'config.example.json')
    with open(path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    return {
        'quota_bytes_is_zero': cfg['capacity']['quota_bytes'] == 0,
        'notify_disabled': cfg['notify']['enabled'] is False,
        'notify_provider_none': cfg['notify']['provider'] == 'none',
        'smtp_is_placeholder': 'REPLACE_WITH' in cfg['mailbox']['smtp_address'],
        'data_dir_is_placeholder': 'REPLACE_WITH' in cfg['storage']['data_directory'],
        'webhook_url_empty': cfg['notify']['webhook_url'] == '',
    }


def main():
    scanned, findings = scan()
    ps1 = ps1_ascii()
    ex = example_config()
    checks = {
        'no_banned_tokens': len(findings) == 0,
        'ps1_pure_ascii': all(x['non_ascii_bytes'] == 0 for x in ps1),
        'example_config_placeholders_only': all(ex.values()),
        'scanned_enough_files': scanned >= 10,
    }
    ok = all(checks.values())
    result = {
        'ok': ok,
        'check': 'desensitization_and_portability',
        'files_scanned': scanned,
        'checks': checks,
        'findings': findings[:40],
        'ps1_ascii': ps1,
        'example_config': ex,
        'note': 'read-only; excludes tests/ and caches; no Outlook, no network',
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
