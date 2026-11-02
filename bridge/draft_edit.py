"""Shared validation for editing or deleting an already saved draft. Never sends mail."""
import html
import re

FIELDS = ('subject', 'body', 'to', 'cc', 'bcc', 'attachments_add', 'attachments_remove')
LIMITS = {'subject': 998, 'body': 500000, 'to': 10000, 'cc': 10000, 'bcc': 10000}
BODY_MARKER = 'data-outlook-local-body'


def parse_update(spec):
    """Return only the fields to change. Absent key = unchanged; empty string = clear."""
    if not isinstance(spec, dict):
        raise ValueError('Give the fields to change as a JSON object.')
    unknown = sorted(set(spec) - set(FIELDS))
    if unknown:
        raise ValueError('Unknown or unchangeable fields: ' + ', '.join(map(str, unknown))
                         + '. Changeable: ' + ', '.join(FIELDS)
                         + '. Signature, sender, reply link and quoted text stay unchanged.')
    result = {}
    for key, value in spec.items():
        if key in LIMITS:
            if not isinstance(value, str):
                raise ValueError(f'{key} must be text (an empty string clears the field; omitting it leaves it unchanged).')
            if len(value) > LIMITS[key]:
                raise ValueError('The draft is too large.')
            if key != 'body' and ('\r' in value or '\n' in value):
                raise ValueError('Invalid line break in ' + key + '.')
        elif key == 'attachments_add':
            if not isinstance(value, list):
                raise ValueError('attachments_add must be a list of file paths.')
        else:
            if (not isinstance(value, list) or len(value) > 20
                    or any(not isinstance(v, str) or not v.strip() for v in value)
                    or len({v.lower() for v in value}) != len(value)):
                raise ValueError('attachments_remove must be a list of unique file names.')
        result[key] = value
    if not any(value not in ([],) for value in result.values()):
        raise ValueError('There is nothing to change.')
    return result


def plain_to_html(text):
    return html.escape(str(text)).replace('\r\n', '\n').replace('\r', '\n').replace('\n', '<br>')


def check_removals(names, present):
    """present: list of (name, ...) visible attachments. Every name must match exactly one."""
    lowered = [str(p[0]).lower() for p in present]
    for name in names:
        hits = lowered.count(name.lower())
        if hits == 0:
            raise ValueError('Attachment not found in the draft: ' + name)
        if hits > 1:
            raise ValueError('Multiple attachments are named ' + name + '; removal is ambiguous.')


def check_totals(existing_sizes, added_sizes):
    import mail_attachments as files
    if len(existing_sizes) + len(added_sizes) > files.MAX_FILES:
        raise ValueError('A draft may have at most 20 attachments.')
    if sum(existing_sizes) + sum(added_sizes) > files.MAX_BYTES:
        raise ValueError('Attachments may total at most 20 MB.')
