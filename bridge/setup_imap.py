"""One-time local provisioning of a direct IMAP/SMTP mailbox.

The password is read without echo, verified against the server (no mail is sent) and stored
encrypted with the Windows user DPAPI in bridge/runtime/. Only non-secret settings go into
bridge/imap-accounts.json.

Example:
  python setup_imap.py --email me@example.com --host mail.example.com --alias alias@example.com \
      --display-name "Jane Doe"
"""
import argparse
import getpass
import imaplib
import json
import re
import smtplib
import ssl
from pathlib import Path
import win32crypt
from supervisor import private_runtime

ROOT = Path(__file__).resolve().parent


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--email', required=True, help='Primary mailbox address (also the default login)')
    parser.add_argument('--username', help='Login name if it differs from the address')
    parser.add_argument('--host', required=True, help='IMAP host')
    parser.add_argument('--smtp-host', help='SMTP host if it differs from the IMAP host')
    parser.add_argument('--imap-port', type=int, default=993)
    parser.add_argument('--smtp-port', type=int, default=465)
    parser.add_argument('--alias', action='append', default=[], help='Allowed extra From address (repeatable)')
    parser.add_argument('--display-name', default='')
    parser.add_argument('--signature-file', help='HTML signature file name in bridge/ or bridge/runtime/ (use {{email}} as placeholder)')
    parser.add_argument('--signature-name', default='Default')
    args = parser.parse_args()
    email = args.email.lower()
    username = args.username or args.email
    slug = re.sub(r'[^a-z0-9]+', '-', email)
    private_runtime()
    password = getpass.getpass('IMAP password (not displayed): ')
    context = ssl.create_default_context()
    with imaplib.IMAP4_SSL(args.host, args.imap_port, ssl_context=context, timeout=20) as conn:
        conn.login(username, password)
        _, rows = conn.list()
        print(json.dumps({'imap_authenticated': True, 'folders': len(rows)}))
    with smtplib.SMTP_SSL(args.smtp_host or args.host, args.smtp_port, context=context, timeout=20) as conn:
        conn.login(username, password)
        print(json.dumps({'smtp_authenticated': True, 'mail_sent': False}))
    secret = ROOT / 'runtime' / f'imap-{slug}.dpapi'
    secret.write_bytes(win32crypt.CryptProtectData(password.encode('utf-8'), 'Local IMAP ' + email, None, None, None, 1))
    del password
    config = {'email': email, 'username': username, 'host': args.host, 'imap_port': args.imap_port,
              'smtp_port': args.smtp_port, 'aliases': [a.lower() for a in args.alias],
              'credential_file': secret.name, 'display_name': args.display_name}
    if args.smtp_host:
        config['smtp_host'] = args.smtp_host
    if args.signature_file:
        config.update(signature_file=args.signature_file, signature_name=args.signature_name)
    path = ROOT / 'imap-accounts.json'
    accounts = json.loads(path.read_text('utf-8')) if path.exists() else []
    accounts = [a for a in accounts if a['email'].lower() != email] + [config]
    path.write_text(json.dumps(accounts, indent=2), encoding='utf-8')
    print(json.dumps({'configured': email, 'credential_storage': 'Windows user DPAPI', 'aliases': config['aliases']}))


if __name__ == '__main__':
    main()
