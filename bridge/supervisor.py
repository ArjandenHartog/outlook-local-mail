"""Console-free, idempotent watchdog for the per-user Outlook bridge."""
import json
import subprocess
import sys
import time
from pathlib import Path
from urllib.error import URLError
from urllib.request import urlopen
import win32api
import win32event
import win32security
import ntsecuritycon

ROOT = Path(__file__).resolve().parent
RUNTIME = ROOT / 'runtime'
# Run background processes with the same environment (and no console window).
PYTHONW = Path(sys.executable).with_name('pythonw.exe')

def healthy():
    try:
        with urlopen('http://127.0.0.1:38473/health', timeout=2) as response:
            state = json.load(response)
        if state.get('service') != 'outlook-local-manager':
            raise RuntimeError('Port 38473 belongs to another service')
        return True
    except (URLError, ConnectionError, OSError):
        return False

def private_runtime():
    RUNTIME.mkdir(exist_ok=True)
    token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32security.TOKEN_QUERY)
    user = win32security.GetTokenInformation(token, win32security.TokenUser)[0]
    system = win32security.ConvertStringSidToSid('S-1-5-18')
    acl = win32security.ACL()
    for sid in (user, system):
        acl.AddAccessAllowedAceEx(win32security.ACL_REVISION,
            win32security.OBJECT_INHERIT_ACE | win32security.CONTAINER_INHERIT_ACE,
            ntsecuritycon.FILE_ALL_ACCESS, sid)
    win32security.SetNamedSecurityInfo(str(RUNTIME), win32security.SE_FILE_OBJECT,
        win32security.DACL_SECURITY_INFORMATION | win32security.PROTECTED_DACL_SECURITY_INFORMATION,
        None, None, acl, None)

def schedule_index(force=False):
    import mail_index
    import datetime as dt
    state=mail_index.status()
    if not force and state['available'] and not state['stale']:
        since_start=(dt.datetime.now(dt.timezone.utc)-dt.datetime.fromisoformat(state['started_at'])).total_seconds()
        if since_start<240:return
    # The indexer has its own named mutex. It reads metadata without taking focus.
    subprocess.Popen([str(PYTHONW),str(ROOT/'index_mail.py')],
        cwd=ROOT,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL,close_fds=True,creationflags=subprocess.CREATE_NO_WINDOW)
    return {'requested':True,'note':'Search data is being updated silently in the background. The current index stays available.'}

def main():
    mutex = win32event.CreateMutex(None, False, 'Local\\OutlookLocalMailStartup')
    acquired = win32event.WaitForSingleObject(mutex, 25000)
    if acquired not in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED):
        raise RuntimeError('Another startup is still running')
    try:
        if healthy():
            schedule_index()
            return
        private_runtime()
        with (RUNTIME / 'service-output.log').open('ab') as output, (RUNTIME / 'service-error.log').open('ab') as error:
            subprocess.Popen([str(PYTHONW), str(ROOT / 'server.py')],
                cwd=ROOT, stdout=output, stderr=error, stdin=subprocess.DEVNULL,
                close_fds=True, creationflags=subprocess.CREATE_NO_WINDOW)
        for _ in range(40):
            time.sleep(.25)
            if healthy():
                schedule_index()
                return
        raise RuntimeError('Local service did not start')
    finally:
        win32event.ReleaseMutex(mutex)
        win32api.CloseHandle(mutex)

if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        # No session token or message data is written by the watchdog.
        if RUNTIME.exists():
            with (RUNTIME / 'supervisor-error.log').open('a', encoding='utf-8') as log:
                log.write(f'{type(exc).__name__}: {exc}\n')
        raise
