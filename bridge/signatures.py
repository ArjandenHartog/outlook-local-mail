"""Mailbox-scoped Outlook HTML signatures, including embedded logo/photo images."""
import base64
import hashlib
import html
import json
import os
from pathlib import Path
import re
import tempfile

from inspect_signatures import read_file

ROOT = Path(__file__).resolve().parent
SIGNATURES = Path(os.environ.get('APPDATA', '')) / 'Microsoft/Signatures'
SETTINGS = ROOT / 'signature-settings.json'
CID_PROPERTY = 'http://schemas.microsoft.com/mapi/proptag/0x3712001F'
HIDDEN_PROPERTY = 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'
MIME_PROPERTY = 'http://schemas.microsoft.com/mapi/proptag/0x370E001F'


def catalog(email):
    suffix = '(' + email.lower() + ')'
    defaults = json.loads(SETTINGS.read_text('utf-8')).get(email.lower(), {}) if SETTINGS.exists() else {}
    choices = []
    for path in sorted(SIGNATURES.glob('*.htm')):
        if not path.stem.lower().endswith(suffix):
            continue
        text_file = path.with_suffix('.txt')
        text = read_file(text_file).strip() if text_file.exists() else ''
        match = re.search(r'(?i)(met\s+vriendelijke\s+groet|kind\s+regards|best\s+regards|sincerely)', text)
        template = bool(match and len(text[:match.start()].strip()) > 200)
        choices.append({'name': path.stem, 'template': template,
                        'default_new': path.stem == defaults.get('new'),
                        'default_reply': path.stem == defaults.get('reply')})
    return {'email': email, 'default_new': defaults.get('new'), 'default_reply': defaults.get('reply'), 'signatures': choices}


def select(email, reply=False, name=None):
    options = catalog(email)
    chosen = options['default_reply' if reply else 'default_new'] if name is None else name
    if chosen in ('', 'none'):
        return None
    if chosen is None:
        return None
    found = next((s for s in options['signatures'] if s['name'] == chosen), None)
    if found is None:
        raise ValueError('The signature does not belong to this mailbox or no longer exists.')
    if found['template']:
        raise ValueError('This is a complete mail template. Use a plain signature and write the message text separately.')
    return SIGNATURES / (chosen + '.htm')


def fragment(path):
    source = read_file(path)
    body = re.search(r'(?is)<body\b[^>]*>(.*?)</body\s*>', source)
    content = body.group(1) if body else source
    # Correct mismatched mailto links in the connector's copy, retaining the
    # original signature file. Some signatures display one address but link another.
    replacements = {}
    def link(match):
        start, address, rest, label = match.groups()
        visible = html.unescape(re.sub(r'<[^>]+>', '', label)).strip()
        if re.fullmatch(r'[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}', visible):
            replacements[address] = visible
            return start + 'mailto:' + visible + rest + label + '</a>'
        return match.group(0)
    content = re.sub(r'(?is)(<a\b[^>]*?href=["\x27])mailto:([^"\x27]+)(["\x27][^>]*>)(.*?)</a>', link, content)
    # Word reconstitutes hyperlinks from conditional HYPERLINK field comments.
    for old, new in replacements.items():
        content = re.sub(r'(?i)mailto:'+re.escape(old),lambda _: 'mailto:'+new,content)
    return content


def attach_images(draft, content):
    """Convert inline data images to actual Outlook CID attachments."""
    attached = {}

    def replace(match):
        mime, payload = match.groups()
        raw = base64.b64decode(re.sub(r'\s+', '', html.unescape(payload)), validate=True)
        digest = hashlib.sha256(raw).hexdigest()[:24]
        if mime == 'image/svg+xml':
            prepared = ROOT / 'runtime/signature-images' / (digest + '.png')
            if not prepared.exists():
                raise ValueError('The SVG logo must be prepared once as a PNG.')
            raw, mime = prepared.read_bytes(), 'image/png'
        cid = 'outlook-local-signature-' + digest
        if cid not in attached:
            extension = {'image/png': '.png', 'image/jpeg': '.jpg', 'image/gif': '.gif', 'image/svg+xml': '.svg'}.get(mime)
            if extension is None:
                raise ValueError('This image format is not supported in signatures.')
            # Place short-lived files inside the user-restricted runtime directory.
            private_runtime = ROOT / 'runtime'
            with tempfile.TemporaryDirectory(prefix='signature-', dir=private_runtime) as directory:
                path = Path(directory) / ('signature' + extension)
                path.write_bytes(raw)
                attachment = draft.Attachments.Add(str(path), 1)
                attachment.PropertyAccessor.SetProperty(CID_PROPERTY, cid)
                attachment.PropertyAccessor.SetProperty(HIDDEN_PROPERTY, True)
                attachment.PropertyAccessor.SetProperty(MIME_PROPERTY, mime)
            attached[cid] = mime
        return 'cid:' + cid

    converted = re.sub(r'data:(image/[\w.+-]+);base64,([^"\x27>]+)', replace, content)
    return converted, len(attached)


def compose(draft, body, signature_path, reply=False):
    content, image_count = attach_images(draft, fragment(signature_path))
    # data-outlook-local-body lets update-draft find the authored text later.
    message = '<div data-outlook-local-body="true" style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' +html.escape(str(body)).replace('\r\n', '\n').replace('\r', '\n').replace('\n', '<br>') + '</div><br>'
    authored = message + '<div data-outlook-local-signature="true">' + content + '</div>'
    if reply:
        original = str(draft.HTMLBody or '')
        body_start = re.search(r'(?is)<body\b[^>]*>', original)
        if body_start:
            draft.HTMLBody = original[:body_start.end()] + authored + '<br>' + original[body_start.end():]
        else:
            draft.HTMLBody = '<html><body>' + authored + '<br>' + original + '</body></html>'
    else:
        draft.BodyFormat = 2
        draft.HTMLBody = '<html><body>' + authored + '</body></html>'
    return image_count
