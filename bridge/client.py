"""Local CLI for the assistant and PowerShell; reads the token locally."""
import argparse
import json
from pathlib import Path
import sys
import time
import uuid
from urllib.request import Request, urlopen
from urllib.error import HTTPError

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')


def call(action, args=None, polls=60):
    state = json.loads((Path(__file__).parent / 'runtime/session.json').read_text('utf-8'))
    request_id = str(uuid.uuid4())
    payload = {'action': action, 'args': args or {}, 'request_id': request_id}
    for _ in range(polls):
        req = Request(state['url'] + '/api/call', json.dumps(payload).encode('utf-8'),
                      {'Content-Type': 'application/json', 'X-Local-Token': state['token']})
        try:
            with urlopen(req, timeout=20) as response:
                result = json.load(response)
                if response.status != 202:
                    return result
        except HTTPError as exc:
            raise RuntimeError(json.load(exc).get('error', str(exc))) from None
        payload = {'action': 'poll', 'request_id': request_id}
        time.sleep(0.5)
    raise RuntimeError('Outlook is still busy. Request code: ' + request_id + '. Do not save the draft again before you have checked Outlook.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Local mail API: status, folders, messages, read, draft, sync ...')
    parser.add_argument('action')
    parser.add_argument('--json', default='{}', help='Command arguments as JSON')
    parser.add_argument('--file', help='Read arguments from a local JSON file')
    options = parser.parse_args()
    try:
        args = json.loads(Path(options.file).read_text('utf-8-sig') if options.file else options.json)
        print(json.dumps(call(options.action, args), ensure_ascii=False, indent=2))
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)
