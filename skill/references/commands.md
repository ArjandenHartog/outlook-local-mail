# Command reference

Call: `python skill/scripts/mail.py <action> ...` (use the environment with pywin32 installed).
The helper starts the local service when needed and prints JSON.

| Action | Arguments | Result |
| --- | --- | --- |
| `status` | none | Every configured account, connection state, offline flag, errors |
| `overview` | `--limit 1..200` (default 15) | Recent Inbox items per connected mailbox plus missing accounts |
| `folders` | `--account`, optional `--refresh` | Folder ID, full path, depth, counts; Inbox/Drafts/Sent IDs (cached 30 s) |
| `signatures` | `--account` | Signature names, template flag, defaults for new mail and replies |
| `messages` | `--account`, optional `--folder`, `--query`, `--offset`, `--limit` | Message metadata, `total`, `next_offset`, `read_at` |
| `read` | `--account`, `--id` | Plain text, recipients (`to`, `cc`, `bcc`), attachment names; no external HTML |
| `quick-search` | optional `--query`, `--account`, `--sender`, `--folder`, `--after`, `--before`, `--unread`, `--include-deleted`, `--offset`, `--limit` | Fast metadata search (subject, sender, folder path) |
| `search` | `--query`, optional `--account`, `--include-deleted`, `--offset`, `--limit`, `--cursor` | Full-text search of physical mail folders |
| `index-status` / `refresh-index` | none | Index age and progress / request a silent rebuild |
| `suggest-folder` | `--account`, `--id` | Filing suggestions from earlier archive examples; no mutation |
| `move` | `--account`, `--id`, `--folder`, optional `--source-folder` | Archive into an existing folder of the same mailbox; read state kept; returns `new_id` |
| `trash` | `--account`, `--id`, `--source-folder <inbox-id>`, `--reason` | Inbox message to the mailbox's own trash; `permanent_delete=false` |
| `draft` | `--file <json>` | Saved draft: ID, account, folder, `sent=false` |
| `update-draft` | `--account`, `--id`, `--file <json with changes only>` | Changed draft: `id`, `previous_id`, `changed`, `attachments`, `sent=false` |
| `delete-draft` | `--account`, `--id`, `--reason`, optional `--expect-subject` | Draft moved to trash; `permanent_delete=false` |
| `send` | `--account`, `--id`, optional `--verify-wait 0..900` | `submitted=true` = accepted by Outlook/SMTP, not delivered |
| `verify-sent` | `--account` plus `--id`, `--message-id` or `--subject`; optional `--expect-to`, `--expect-attachments`, `--deep`, `--wait`, `--interval` | Sent copy, bounces, receipts, own-mailbox arrival, `delivery_verified` |
| `sync` | none | Requests Outlook send/receive groups (`completed=false`) |

Searching: `quick-search` words are literal, all must match, word prefixes count; dates are UTC
(`--after` inclusive, `--before` exclusive); `--folder` needs `--account`. `search` returns
`complete`, `errors`, `missing`, `truncated`; the CLI resumes bounded scans automatically.

## New draft

```json
{
  "email": "you@example.com",
  "subject": "Subject",
  "body": "The draft text.",
  "to": "recipient@example.org",
  "cc": "",
  "attachments": ["C:\\Users\\you\\Documents\\report.pdf"],
  "request_delivery_receipt": false,
  "request_read_receipt": false
}
```

Reply: `{"email": "...", "reply_id": "<id of the original message>", "body": "Text above the quote."}`.
Leave `subject`, `to`, `cc` out unless a change is requested. `signature` is optional (a name from
`signatures`, or `none`). IMAP mailboxes also accept `from_address` for a configured alias.
A read-receipt request is visible to the recipient.

## Editing a draft (`update-draft`) and removing one (`delete-draft`)

```text
update-draft --account you@example.com --id <draft-id> --file C:\path\changes.json
delete-draft --account you@example.com --id <draft-id> --reason "duplicate draft" --expect-subject "Exact subject"
```

```json
{
  "subject": "New subject",
  "body": "New own text (replaces only the text above signature/quote).",
  "to": "a@example.org, b@example.org",
  "cc": "",
  "bcc": "hidden@example.org",
  "attachments_remove": ["old.pdf"],
  "attachments_add": ["C:\\Users\\you\\Documents\\new.pdf"]
}
```

| Rule | Behaviour |
| --- | --- |
| Omitted field | Unchanged (signature, quoted thread, In-Reply-To/References, sender) |
| Empty string | Clears `subject`, `to`, `cc` or `bcc` |
| `attachments_remove` | By file name (case-insensitive); unknown or ambiguous = error; inline/signature images cannot be removed; removed before adding |
| `attachments_add` | Same rules as `draft` (20 files, 20 MiB including existing ones; duplicate name = error) |
| `body` | Only connector-written text; otherwise error with no change |
| Unknown fields | `reply_id`, `signature`, `email`, `from_address`, ... are refused |
| Validation | Everything is checked before the draft is touched |

Result fields: `id`, `previous_id`, `id_changed`, `email`, `subject`, `folder`, `sent=false`, `changed`,
`sender_email`, `attachments` (Outlook reports the stored size, which can differ from the file size).
IMAP only: `uid_changed=true`, `uidvalidity_changed`, `old_removed`, optional `note`, and
`already_updated=true` when the exact same request is repeated (no second copy). If `old_removed=false`
or the outcome is uncertain, check Drafts for duplicates and never repeat blindly.

## Delivery status after sending

```text
send --account you@example.com --id <draft-id> --verify-wait 120
verify-sent --account you@example.com --id <same-draft-id> --wait 300
verify-sent --account you@example.com --message-id "<...>"
verify-sent --account you@example.com --subject "Subject" --expect-to hr@example.org --expect-attachments 2
```

| Field | Meaning |
| --- | --- |
| `state` | `delivered`, `partially_verified`, `bounced`, `sent_no_bounce_seen`, `in_outbox`, `not_found` |
| `submitted`, `in_sent_items`, `outbox_empty` | Journal/Sent Items/Outbox evidence |
| `checks` | Subject/recipients/attachments compared with what was intended |
| `bounce_detected`, `bounces`, `failed_recipients` | NDR that references the Message-ID (or same original subject within 2 min) |
| `possible_bounces` | NDR with only the same subject: not proof, review manually |
| `delivery_receipt`, `read_receipt` | Requested and received confirmations |
| `recipient_mailboxes` | Recipients that are own mailboxes: message with the same Message-ID found |
| `delivery_verified`, `explanation` | `true` only when every recipient is proven and nothing bounced |

"No bounce seen" is never proof of delivery; a remote server may bounce later.
