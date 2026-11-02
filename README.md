# outlook-local-mail

A local mail bridge and agent skill for **Windows**: let an AI agent (Claude Code, Codex, ...)
read, search, file, draft, edit drafts and — only when you explicitly say so — send mail from
your own mailboxes, without an OAuth app or any cloud service.

- **Outlook Classic** mailboxes through the already signed-in COM profile (any number of accounts).
- **Direct IMAP/SMTP** mailboxes (TLS, normal certificate checks), password stored with Windows DPAPI.
- A loopback-only service (`127.0.0.1:38473`, per-run token) serialises Outlook COM calls.
- Fast local metadata index (subject/sender/folder) plus full-text search.
- Drafts first: `draft`, `update-draft`, `delete-draft`; `send` only for an existing draft you point at.
- Delivery verification (`verify-sent`): Sent copy, bounces/NDRs, receipts, own-mailbox arrival.
- Send journal that prevents double submission, also after a restart.

> This is a personal-automation tool, not a hardened multi-user service. It trusts the local
> Windows user. Read the safety notes below before pointing an agent at your mail.

## Requirements

- Windows 10/11, Python 3.11+ (tested on 3.12)
- Outlook Classic (the "classic" desktop app, not "new Outlook") for the Outlook accounts
- `pip install -r bridge/requirements.txt` (only `pywin32`)

[![tests](https://github.com/ArjandenHartog/outlook-local-mail/actions/workflows/test.yml/badge.svg)](https://github.com/ArjandenHartog/outlook-local-mail/actions/workflows/test.yml)

## Quick install

```powershell
git clone https://github.com/ArjandenHartog/outlook-local-mail
cd outlook-local-mail
.\install.ps1 -ScheduleTask -LinkSkill both
```

`install.ps1` creates `.venv`, installs the dependencies, **detects the accounts signed in to Outlook Classic**
(read-only, writes `bridge\accounts.json`), runs the tests, and prints `status`. Optional switches:
`-ScheduleTask` registers a logon task that keeps the service running, `-LinkSkill claude|codex|both`
links `skill\` into `%USERPROFILE%\.claude\skills` / `.codex\skills`, `-SkipTests`, `-SkipAccounts`.

## Updating

```powershell
.\update.ps1            # latest release tag (tests, then restarts the service)
.\update.ps1 -Ref main  # or follow main
```

## Manual setup / extra mailboxes

1. **Outlook accounts** — normally done by `install.ps1`. By hand: copy
   `config/accounts.example.json` to `bridge/accounts.json` and list the addresses signed in to Outlook Classic.
2. **Direct IMAP (optional)** — store the password (not echoed, verified, nothing is sent):

   ```powershell
   cd bridge
   ..\.venv\Scripts\python.exe setup_imap.py --email you@example.com --host imap.example.com `
       --smtp-host smtp.example.com --alias alias@example.com --display-name "Your Name"
   ```

   This writes `bridge/imap-accounts.json` (no secrets) and an encrypted file in `bridge/runtime/`.
   See `config/imap-accounts.example.json` for all options, including an HTML signature
   (`signature_file`, `{{email}}` placeholder; see `config/signature.example.html`).
3. **Outlook signatures (optional)** — signatures are read from `%APPDATA%\Microsoft\Signatures`
   (file names must end with `(address)`). Copy `config/signature-settings.example.json` to
   `bridge/signature-settings.json` to choose defaults per mailbox.
4. Try it:

   ```powershell
   .\.venv\Scripts\python.exe skill\scripts\mail.py status
   ```

   The first call starts the local service in the background (`pythonw`, no window).
   To start it at logon, create a Task Scheduler task that runs `pythonw.exe bridge\supervisor.py`
   (it is idempotent and restarts the service when it is down).

## Use it as an agent skill

Copy or link `skill/` to your agent's skills folder as `outlook-local-mail`
(for Claude Code: `%USERPROFILE%\.claude\skills\outlook-local-mail`). `skill/SKILL.md` tells the
agent how to use the CLI; `skill/references/commands.md` lists every command. If you keep the skill
outside this repository, set `OUTLOOK_LOCAL_BRIDGE` to the `bridge` directory.

## Editing drafts

```powershell
.\.venv\Scripts\python.exe skill\scripts\mail.py update-draft --account you@example.com --id <draft-id> --file changes.json
```

`changes.json` contains only what changes (`subject`, `body`, `to`, `cc`, `bcc`, `attachments_add`,
`attachments_remove`). Signature, quoted thread, reply headers and sender are kept. Outlook edits the
item in place; IMAP appends the new version, verifies it and removes the old UID (a new ID is returned).
Only unsent drafts in the mailbox's Drafts folder are accepted. Nothing is ever sent by these commands.

## Monitoring (read-only)

`skill/scripts/watch.py` is meant for a scheduled agent run or a manual check; it never sends, edits or deletes:

```powershell
.\.venv\Scripts\python.exe skill\scripts\watch.py drafts            # duplicate drafts + drafts older than 7 days
.\.venv\Scripts\python.exe skill\scripts\watch.py bounces --days 3  # new bounces on recently sent mail
```

- `drafts` groups drafts per mailbox by subject (ignoring `RE:`/`FW:`) and compares recipients and text;
  only near-identical ones appear in `suggested_delete`. You decide, then run `delete-draft --expect-subject`.
  Drafts you want left alone go in `bridge/runtime/watch-config.json` (`ignore_drafts`: account + subject regex).
- `bounces` runs `verify-sent` on recent Sent items (max 40 per run, new ones first, state in
  `bridge/runtime/watch-state.json`) and reports each bounce once. An NDR that cannot be tied to one message
  (a mail merge with identical subjects) is reported separately as `ambiguous_ndrs`.
- The background service also builds the folder-filing analysis for `suggest-folder` by itself on first start
  and refreshes it weekly (`bridge/runtime/filing-analysis.json`, local only).

## Releases

Pushing a tag such as `v0.2.0` runs the tests and publishes a GitHub release with a zip
(`.github/workflows/release.yml`). CI (`test.yml`) runs the tests on every push on Windows with Python 3.11 and 3.12.

## Safety notes

- The service listens on loopback only and requires a random per-run token (`bridge/runtime/session.json`,
  readable by your Windows user only). Any program running as you can read that token — treat it as
  full access to your mail.
- Nothing is sent unless you (or an agent acting on your explicit instruction) call `send` on a draft.
  Tests never send; the Outlook `Send` method raises in the test doubles.
- `trash` / `delete-draft` move to the mailbox's trash; there is no permanent-delete command.
- Message bodies are never stored in the index, logs or the send journal (subject, recipients and
  attachment names/sizes are, so delivery can be verified).
- Agents can be tricked by text inside mail. The skill tells them to treat message content as data;
  keep an eye on what you let an agent send.

## Layout

```text
bridge/   local service, Outlook COM + IMAP adapters, index, tests (test_*.py)
skill/    SKILL.md, references/commands.md, scripts/mail.py (the CLI)
config/   example configuration files
```

## Tests

```powershell
cd bridge
..\.venv\Scripts\python.exe -m unittest discover -p "test_*.py"
```

The tests use fakes for Outlook COM, IMAP and SMTP; no mailbox is touched.

## Limitations

- Windows + Outlook Classic only for Outlook accounts; IMAP accounts need only Windows (DPAPI).
- Outlook's local cache is the source for Outlook search and the index, so very old mail may be missing
  unless Outlook keeps it offline.
- `update-draft` can replace the body only for drafts whose own text the connector wrote
  (it marks it); otherwise it refuses and changes nothing.
- Folder-filing suggestions for Outlook use `bridge/runtime/filing-analysis.json`, built locally by the
  background service (or `python bridge/analyze_filing.py`); IMAP suggestions use the index.

## License

MIT
