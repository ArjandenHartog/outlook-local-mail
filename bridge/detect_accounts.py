"""Write bridge/accounts.json from the accounts that are signed in to Outlook Classic (read-only).

Usage: python detect_accounts.py [--force]
Nothing is changed in Outlook. Existing accounts.json is kept unless --force is given.
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TARGET = ROOT / 'accounts.json'


def main():
    if TARGET.exists() and '--force' not in sys.argv:
        print(json.dumps({'kept': str(TARGET.name), 'note': 'accounts.json already exists; use --force to overwrite.'}))
        return 0
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.GetActiveObject('Outlook.Application')
        except Exception:
            app = win32com.client.Dispatch('Outlook.Application')  # starts Outlook Classic if it is installed
        accounts = []
        for account in app.Session.Accounts:
            address = str(account.SmtpAddress or '').strip().lower()
            if address and account.DeliveryStore is not None and all(a['email'] != address for a in accounts):
                accounts.append({'email': address, 'type': 'Outlook', 'description': address})
    except Exception as exc:
        print(json.dumps({'error': 'Could not read the Outlook profile: ' + str(exc)[:200],
                          'hint': 'Open Outlook Classic, sign in, then run this again. Or copy config/accounts.example.json to bridge/accounts.json by hand.'}))
        return 1
    finally:
        pythoncom.CoUninitialize()
    TARGET.write_text(json.dumps(accounts, indent=2), encoding='utf-8')
    print(json.dumps({'written': TARGET.name, 'accounts': [a['email'] for a in accounts]}))
    return 0


if __name__ == '__main__':
    sys.exit(main())
