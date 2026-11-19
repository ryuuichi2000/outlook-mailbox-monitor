#!/usr/bin/env python3
"""Parse-check every .ps1 with the Windows PowerShell 5.1 parser.

Pure syntax parsing: the scripts are NOT executed, Outlook is never touched,
no file is written. Run: python tests/check_ps_syntax.py
"""

import json
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGETS = ('src/MailboxMonitor.ps1', 'src/probe_outlook.ps1')

PS_CODE = r'''
$ErrorActionPreference = 'Stop'
$targets = @('src/MailboxMonitor.ps1', 'src/probe_outlook.ps1')
$errs = @()
foreach ($p in $targets) {
    $full = Join-Path $PWD.Path $p
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $full, [ref]$tokens, [ref]$errors)
    foreach ($e in $errors) {
        $errs += ("{0}:{1}: {2}" -f $p, $e.Extent.StartLineNumber, $e.Message)
    }
}
if ($errs.Count -eq 0) { '[]' } else { $errs | ConvertTo-Json -Compress }
'''


def powershell():
    found = shutil.which('powershell.exe')
    if found:
        return found
    win = os.environ.get('WINDIR') or os.environ.get('SystemRoot') or os.path.join('C:' + os.sep, 'Windows')
    return os.path.join(win, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')


def main():
    try:
        proc = subprocess.run(
            [powershell(), '-NoLogo', '-NoProfile', '-NonInteractive',
             '-Command', PS_CODE],
            cwd=ROOT, capture_output=True, text=True,
            encoding='utf-8', errors='replace', timeout=120)
    except Exception as e:
        print(json.dumps({'ok': False, 'error': 'cannot_run_powershell',
                          'detail': '%s: %s' % (type(e).__name__, e)},
                         ensure_ascii=False, indent=2))
        return 1
    out = (proc.stdout or '').strip()
    try:
        errs = json.loads(out) if out else []
    except ValueError:
        errs = ['unparsable output: %s %s' % (out[:300], (proc.stderr or '')[:300])]
    if isinstance(errs, str):
        errs = [errs]
    ok = proc.returncode == 0 and len(errs) == 0
    print(json.dumps({
        'ok': ok, 'check': 'powershell_syntax_parse_only',
        'targets': list(TARGETS),
        'returncode': proc.returncode,
        'errors': errs[:30],
        'stderr': (proc.stderr or '').strip()[:400],
        'note': 'parser only; no execution, no Outlook, no file writes',
    }, ensure_ascii=False, indent=2))
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
