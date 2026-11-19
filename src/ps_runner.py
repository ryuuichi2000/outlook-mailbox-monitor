"""Run a Windows PowerShell 5.1 script and decode its JSON-lines output.

Portable by design: no user names, no installation paths, no vendor-specific
launcher. In the normal case the script is started directly as the same
interactive Windows user that has classic Outlook open.

If your automation host runs with an elevated token and Outlook COM cannot be
reached (MK_E_UNAVAILABLE), ``runtime.launcher_command`` lets you prepend ANY
de-elevation / wrapper tool of your choice. Examples in docs/TROUBLESHOOTING.md.
A launcher may either be a plain command prefix, or contain the token
``{command}`` which is replaced with the full PowerShell command line. If the
launcher prints ``child_exit=N`` on stdout, that value is used as the real exit
code of the child process.
"""

import json
import os
import shutil
import subprocess
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def decode_file(path):
    if not os.path.exists(path):
        return ''
    with open(path, 'rb') as f:
        raw = f.read()
    for enc in ('utf-8-sig', 'utf-8', 'utf-16-le', 'utf-16', 'gbk', 'cp936', 'latin1'):
        try:
            text = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        if '\x00' not in text[:200]:
            return text
    return raw.decode('latin1')


def _powershell(cfg):
    configured = str((cfg.get('runtime') or {}).get('powershell_path') or 'powershell.exe')
    if os.path.isfile(configured) or os.path.sep in configured or '/' in configured:
        return configured
    found = shutil.which(configured)
    if found:
        return found
    # Standard Windows fallback, built from environment variables (no hard-coded
    # user profile or product directory).
    win = os.environ.get('WINDIR') or os.environ.get('SystemRoot') or os.path.join('C:' + os.sep, 'Windows')
    return os.path.join(win, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')


def _build_cmd(script, ps_args, cfg, out=None, err=None, cwd=ROOT):
    launcher = (cfg.get('runtime') or {}).get('launcher_command') or []
    if isinstance(launcher, str):
        launcher = [launcher]
    ps_cmd = [_powershell(cfg), '-NoLogo', '-NoProfile', '-NonInteractive',
              '-STA', '-ExecutionPolicy', 'Bypass', '-File', script]
    ps_cmd += [str(a) for a in ps_args]
    if not launcher:
        return ps_cmd
    quoted = subprocess.list2cmdline(ps_cmd)
    # A wrapper may need the full command line, the output file paths or the
    # working directory. All four tokens are optional.
    rendered, used = [], False
    for part in launcher:
        part = str(part)
        if '{command}' in part:
            used = True
        part = (part.replace('{command}', quoted)
                    .replace('{stdout}', str(out or ''))
                    .replace('{stderr}', str(err or ''))
                    .replace('{cwd}', str(cwd)))
        rendered.append(part)
    if not used:
        rendered += ps_cmd
    return rendered


def run_ps(script, ps_args=None, cfg=None, tag='ps', timeout=600, result_file=None):
    """Return a dict describing the run. Never raises on child failure."""
    cfg = cfg or {}
    ps_args = [str(a) for a in (ps_args or [])]
    data_dir = (cfg.get('storage') or {}).get('data_directory') or ROOT
    out_dir = os.path.join(data_dir, '_runtime')
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError:
        out_dir = tempfile.gettempdir()
    out = os.path.join(out_dir, tag + '.out')
    err = out + '.err'

    for path in (out, err, result_file):
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                pass

    cmd = _build_cmd(script, ps_args, cfg, out, err, ROOT)
    t0 = time.time()
    try:
        proc = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True,
                              encoding='utf-8', errors='replace', timeout=timeout)
        launcher_rc = proc.returncode
        launcher_stdout = proc.stdout or ''
        launcher_stderr = proc.stderr or ''
        timed_out = False
    except subprocess.TimeoutExpired:
        launcher_rc, launcher_stdout, launcher_stderr = -1, '', 'launcher_timeout'
        timed_out = True
    except Exception as exc:
        launcher_rc, launcher_stdout, launcher_stderr = -1, '', '%s: %s' % (type(exc).__name__, exc)
        timed_out = True

    parsed_child = None
    for line in launcher_stdout.splitlines():
        if line.strip().startswith('child_exit='):
            try:
                parsed_child = int(line.strip().split('=', 1)[1])
            except ValueError:
                pass
    child_exit = parsed_child if parsed_child is not None else launcher_rc

    raw = decode_file(out)
    records, unparsed = [], []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            obj = json.loads(line)
            if isinstance(obj, dict):
                records.append(obj)
            else:
                unparsed.append(line)
        except ValueError:
            unparsed.append(line)

    result = None
    if result_file and os.path.exists(result_file):
        try:
            result = json.loads(decode_file(result_file))
        except ValueError:
            result = None

    return {
        'ok_child': (not timed_out and launcher_rc == 0 and child_exit == 0),
        'timed_out': timed_out,
        'launcher_rc': launcher_rc,
        'child_exit': child_exit,
        'launcher_stdout': launcher_stdout.strip()[:800],
        'launcher_stderr': launcher_stderr.strip()[:800],
        'stdout_raw': raw[:4000],
        'stderr_raw': decode_file(err)[:2000],
        'records': records,
        'unparsed': unparsed,
        'result': result,
        'elapsed': round(time.time() - t0, 1),
    }
