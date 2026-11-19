"""Optional outbound notifications.

Notifications are OFF by default (``notify.enabled = false``). The monitor
works fully offline: it writes local logs/events and prints a one-line JSON
result. Turning notifications on is a configuration choice, never a runtime
requirement, and no provider is tied to a specific AI product.

Supported providers (``notify.provider``):

  none      - no network call at all (default)
  stdout    - print the message to the console, useful for piping/agents
  webhook   - generic HTTP POST, payload configurable via payload_template
  feishu    - Feishu / Lark custom bot (incoming webhook, optional signed mode)
  wecom     - WeCom group robot (incoming webhook)
  dingtalk  - DingTalk custom robot (incoming webhook, optional signed mode)
  command   - call ANY local CLI / script / connector you already have

For ``command``, the message text is passed BOTH on stdin and through the
environment variables MBM_MESSAGE / MBM_IDEMPOTENCY_KEY / MBM_PROVIDER, so any
CLI, agent tool or custom connector can consume it. A non-zero exit code is
treated as a delivery failure (the event stays pending for the next round).

Only the Python standard library is used; no SDK and no vendor CLI required.
"""

import base64
import hashlib
import hmac
import json
import os
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def is_enabled(cfg):
    n = cfg.get('notify') or {}
    return bool(n.get('enabled')) and (n.get('provider') or 'none') != 'none'


def _feishu_sign(secret):
    ts = str(int(time.time()))
    string_to_sign = ts + '\n' + secret
    digest = hmac.new(string_to_sign.encode('utf-8'), b'', hashlib.sha256).digest()
    return ts, base64.b64encode(digest).decode('ascii')


def _dingtalk_sign(secret, url):
    ts = str(round(time.time() * 1000))
    digest = hmac.new(secret.encode('utf-8'),
                      (ts + '\n' + secret).encode('utf-8'),
                      hashlib.sha256).digest()
    sign = urllib.parse.quote_plus(base64.b64encode(digest).decode('ascii'))
    sep = '&' if '?' in url else '?'
    return '%s%stimestamp=%s&sign=%s' % (url, sep, ts, sign)


def build_request(text, cfg):
    """Build (url, body, headers) for a webhook provider. Pure; no network."""
    n = cfg.get('notify') or {}
    provider = n.get('provider') or 'webhook'
    url = str(n.get('webhook_url') or '')
    secret = str(n.get('webhook_secret') or '')
    mention = bool(n.get('mention_all'))
    headers = {'Content-Type': 'application/json; charset=utf-8'}

    if provider == 'feishu':
        body = {'msg_type': 'text', 'content': {'text': text}}
        if secret:
            ts, sign = _feishu_sign(secret)
            body['timestamp'] = ts
            body['sign'] = sign
    elif provider == 'wecom':
        body = {'msgtype': 'text', 'text': {'content': text}}
        if mention:
            body['text']['mentioned_list'] = ['@all']
    elif provider == 'dingtalk':
        body = {'msgtype': 'text', 'text': {'content': text}}
        if mention:
            body['at'] = {'isAtAll': True}
        if secret:
            url = _dingtalk_sign(secret, url)
    else:  # generic webhook
        template = str(n.get('payload_template') or '')
        if template:
            # Escape the text once so it is safe to inject into a JSON string.
            escaped = json.dumps(text, ensure_ascii=False)[1:-1]
            body = json.loads(template.replace('{text}', escaped))
        else:
            # Slack-compatible default. Override with payload_template for
            # Discord ({"content":"{text}"}), Teams, or your own service.
            body = {'text': text}

    return {'url': url, 'body': body, 'headers': headers, 'provider': provider}


def _as_int(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _post(req, timeout):
    data = json.dumps(req['body'], ensure_ascii=False).encode('utf-8')
    request = urllib.request.Request(
        req['url'], data=data, headers=req['headers'], method='POST')
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            status = response.status
            raw = response.read().decode('utf-8', 'replace')[:800]
    except urllib.error.HTTPError as e:
        body = ''
        try:
            body = e.read().decode('utf-8', 'replace')[:300]
        except Exception:
            pass
        return False, 'http_%d %s' % (e.code, body)
    except Exception as e:
        return False, '%s: %s' % (type(e).__name__, e)

    if not (200 <= status < 300):
        return False, 'http_%d %s' % (status, raw[:300])

    # IM platforms report application-level failures inside a 200 response.
    try:
        obj = json.loads(raw)
    except ValueError:
        obj = {}
    if isinstance(obj, dict):
        for key in ('errcode', 'code', 'StatusCode'):
            if key in obj and _as_int(obj[key]) not in (0, None):
                return False, '%s=%s %s' % (key, obj[key], raw[:200])
        detail = obj.get('message_id') or obj.get('msgid') or obj.get('data', {}).get('message_id')
        return True, str(detail or ('http_%d' % status))
    return True, 'http_%d' % status


def _command(text, cfg, key):
    n = cfg['notify']
    command = n.get('custom_command')
    env = dict(os.environ)
    env['MBM_MESSAGE'] = text
    env['MBM_IDEMPOTENCY_KEY'] = str(key)
    env['MBM_PROVIDER'] = str(n.get('provider') or '')
    env['MBM_TEXT_IS_UTF8'] = '1'
    timeout = int(n.get('timeout_seconds') or 15)
    kwargs = dict(input=text, text=True, capture_output=True,
                  encoding='utf-8', errors='replace', timeout=timeout,
                  env=env, cwd=ROOT)
    if isinstance(command, str):
        proc = subprocess.run(command, shell=True, **kwargs)
    else:
        proc = subprocess.run([str(x) for x in command], shell=False, **kwargs)
    if proc.returncode == 0:
        return True, (proc.stdout or '').strip()[:300] or 'command_ok'
    return False, 'exit_%d %s' % (proc.returncode, (proc.stderr or '').strip()[:300])


def send(text, cfg, key='', dry_run=False):
    """Deliver one message. Returns (ok, detail).

    A disabled provider is a successful no-op, so the orchestrator never marks
    a round as failed just because the user chose not to connect any IM tool.
    """
    n = cfg.get('notify') or {}
    provider = n.get('provider') or 'none'
    if not n.get('enabled') or provider == 'none':
        return True, 'disabled'
    if provider == 'stdout':
        print('[notify]\n' + text)
        return True, 'stdout'
    if dry_run:
        return True, 'dry_run'
    try:
        if provider == 'command':
            return _command(text, cfg, key)
        return _post(build_request(text, cfg), int(n.get('timeout_seconds') or 15))
    except Exception as e:
        return False, '%s: %s' % (type(e).__name__, e)
