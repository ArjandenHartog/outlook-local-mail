"""Explicit local files and private browser uploads. No mail is sent here."""
import mimetypes
import re
import time
import uuid
from pathlib import Path

UPLOADS = Path(__file__).resolve().parent / 'runtime/uploads'
MAX_BYTES = 20 * 1024 * 1024
MAX_FILES = 20
TTL_SECONDS = 24 * 60 * 60


def filename(value):
    name = str(value)
    if (not name or len(name) > 180 or name in ('.', '..') or name.endswith((' ', '.'))
            or any(ord(c) < 32 or c in '/\\:<>"|?*' for c in name)
            or name.split('.')[0].upper() in {'CON','PRN','AUX','NUL',*[f'COM{i}' for i in range(1,10)],*[f'LPT{i}' for i in range(1,10)]}):
        raise ValueError('Invalid file name for an attachment.')
    return name


def upload(stream, size, name):
    name = filename(name)
    if not 0 <= size <= MAX_BYTES:
        raise ValueError('An attachment may be at most 20 MB.')
    UPLOADS.mkdir(parents=True, exist_ok=True)
    # Only remove expired files from our own UUID directories; never recurse.
    for directory in UPLOADS.iterdir():
        if not re.fullmatch(r'[0-9a-f]{32}', directory.name) or directory.is_symlink() or not directory.is_dir():
            continue
        if time.time() - directory.stat().st_mtime > TTL_SECONDS:
            for old in directory.iterdir():
                if old.is_file() and not old.is_symlink(): old.unlink()
            try: directory.rmdir()
            except OSError: pass
    data = stream.read(size)
    if len(data) != size:
        raise ValueError('The upload was not received completely.')
    identity = uuid.uuid4().hex
    directory = UPLOADS / identity
    directory.mkdir()
    (directory / name).write_bytes(data)
    return {'upload_id': identity, 'name': name, 'size': size}


def prepare(values=None):
    if values is None: return []
    if not isinstance(values, list) or len(values) > MAX_FILES:
        raise ValueError('Choose at most 20 attachments.')
    result = []; total = 0
    for value in values:
        if isinstance(value, str):
            path = Path(value)
            if not path.is_absolute(): raise ValueError('An attachment requires a full local file path.')
        elif isinstance(value, dict) and set(value) == {'upload_id'}:
            identity = value['upload_id']
            if not isinstance(identity, str) or not re.fullmatch(r'[0-9a-f]{32}', identity):
                raise ValueError('Invalid upload ID.')
            directory = UPLOADS / identity
            if directory.is_symlink() or not directory.is_dir() or time.time()-directory.stat().st_mtime > TTL_SECONDS:
                raise ValueError('This attachment has expired. Choose the file again.')
            files = list(directory.iterdir())
            if len(files) != 1 or files[0].is_symlink(): raise ValueError('Invalid local upload.')
            path = files[0]
        else: raise ValueError('Use a local file path or upload ID for an attachment.')
        if not path.is_file(): raise ValueError('Attachment not found: ' + path.name)
        name = filename(path.name)
        if path.stat().st_size > MAX_BYTES-total: raise ValueError('Attachments may total at most 20 MB.')
        with path.open('rb') as stream: data = stream.read(MAX_BYTES-total+1)
        total += len(data)
        if total > MAX_BYTES: raise ValueError('Attachments may total at most 20 MB.')
        mime = mimetypes.guess_type(name)[0] or 'application/octet-stream'
        result.append({'path': str(path.resolve()), 'name': name, 'size': len(data), 'data': data, 'mime': mime})
    return result


def receipt(files):
    return [{'name': file['name'], 'size': file['size']} for file in files]
