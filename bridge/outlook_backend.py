"""Outlook OOM adapter. All methods run on the same STA worker thread."""
import datetime as dt
import json
import hashlib
import os
import re
import subprocess
import time
from pathlib import Path
import signatures
import mail_attachments
import draft_edit
from workflow import WorkflowMixin
from delivery import OutlookDeliveryMixin


def now():
    return dt.datetime.now(dt.timezone.utc).isoformat()


def timestamp(value):
    try:
        return value.astimezone(dt.timezone.utc).isoformat()
    except (ValueError, AttributeError):
        return str(value)


def received_timestamp(item):
    # OOM ReceivedTime is local wall time, but pywin32 can label that DATE as
    # GMT. Read the underlying UTC delivery timestamp instead.
    try:
        return timestamp(item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x0E060040'))
    except Exception:
        try:
            return timestamp(item.PropertyAccessor.LocalTimeToUTC(item.ReceivedTime))
        except Exception:
            return timestamp(item.ReceivedTime)


class MailError(Exception):
    pass


class OutlookBackend(WorkflowMixin, OutlookDeliveryMixin):
    def __init__(self, expected_path, client=None, journal_path=None):
        self.expected = json.loads(Path(expected_path).read_text(encoding='utf-8-sig'))
        accounts_root=Path(expected_path).resolve().parent
        runtime=accounts_root/'runtime'
        self.index_invalidation_path=runtime/'index-invalidated.json'
        self.client = client
        self.shared = {}
        self.last_sync_request = None
        self.sync_progress = {}
        self.sync_events = []
        self.last_classic_launch = None
        self.journal_path = Path(journal_path) if journal_path else Path(__file__).parent / 'runtime/send-journal.json'

    def session(self):
        if self.client is None:
            import win32com.client
            try:
                self.client = win32com.client.GetActiveObject('Outlook.Application')
            except Exception:
                # An unattended -Embedding startup can hang before publishing its
                # COM object. Use the ordinary installed client, without focus,
                # then attach through the ROT with a bounded readiness wait.
                if self.last_classic_launch is None or time.monotonic() - self.last_classic_launch >= 120:
                    outlook = Path(os.environ.get('ProgramFiles', r'C:\Program Files')) / 'Microsoft Office/Root/Office16/OUTLOOK.EXE'
                    if not outlook.is_file():
                        raise MailError('Outlook Classic was not found on this computer.')
                    startup = subprocess.STARTUPINFO()
                    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startup.wShowWindow = 0  # SW_HIDE; attached Explorer is minimized below.
                    subprocess.Popen([str(outlook), '/recycle'], startupinfo=startup)
                    self.last_classic_launch = time.monotonic()
                for _ in range(60):
                    try:
                        self.client = win32com.client.GetActiveObject('Outlook.Application')
                        for explorer in self.client.Explorers:
                            explorer.WindowState = 1  # olMinimized
                        break
                    except Exception:
                        import pythoncom
                        pythoncom.PumpWaitingMessages()
                        time.sleep(0.5)
                if self.client is None:
                    raise MailError('Outlook Classic is not available yet. Check your Outlook profile and sign-in; the connector will try again later.') from None
        try:
            session = self.client.Session
            _ = session.Accounts.Count
            return session
        except Exception:
            self.client = None
            self.shared.clear()
            self.folder_cache = {}
            self.searches = {}
            raise MailError('The connection to Outlook Classic was closed. Open Outlook Classic and refresh.') from None

    def bindings(self, session):
        result = {}
        for account in session.Accounts:
            email = str(account.SmtpAddress or '').lower()
            if email and account.DeliveryStore is not None:
                result[email] = (account.DeliveryStore, account, None)
        result.update({k: v for k, v in self.shared.items() if k not in result})
        return result

    def mailbox(self, email):
        email = str(email).lower()
        if email not in {a['email'].lower() for a in self.expected}:
            raise MailError('This mailbox is not in the list of allowed accounts.')
        session = self.session()
        binding = self.bindings(session).get(email)
        if not binding:
            raise MailError('This account is not yet available in Outlook Classic.')
        return session, binding

    def default_folder(self, binding, number):
        store, account, recipient = binding
        if recipient is not None:
            return self.session().GetSharedDefaultFolder(recipient, number)
        return store.GetDefaultFolder(number)

    def status(self):
        result = {'checked_at': now(), 'source': 'Outlook Classic', 'connected': 0,
                  'expected': len(self.expected), 'offline': None,
                  'last_sync_request': self.last_sync_request,
                  'sync_verified': False, 'sync_progress': self.sync_progress,
                  'accounts': [], 'stores': [], 'error': None}
        bindings = {}
        try:
            session = self.session()
            result['offline'] = bool(session.Offline)
            bindings = self.bindings(session)
            for store in session.Stores:
                result['stores'].append({'name': str(store.DisplayName), 'id': str(store.StoreID)})
        except Exception as exc:
            result['error'] = str(exc)
        for expected in self.expected:
            row = {'email': expected['email'], 'connected': False, 'inbox_count': None,
                   'latest_received': None, 'mode': None, 'error': None}
            binding = bindings.get(expected['email'].lower())
            if binding:
                try:
                    inbox = self.default_folder(binding, 6)
                    row.update(connected=True, inbox_count=inbox.Items.Count,
                               mode='gedeeld' if binding[2] else 'account')
                    items = inbox.Items
                    items.Sort('[ReceivedTime]', True)
                    if items.Count:
                        try:
                            row['latest_received'] = received_timestamp(items.Item(1))
                        except Exception:
                            pass
                    result['connected'] += 1
                except Exception as exc:
                    row['error'] = str(exc)
            result['accounts'].append(row)
        return result

    def folders(self, email, refresh=False):
        _, binding = self.mailbox(email)
        key=(str(email).lower(),str(binding[0].StoreID))
        cache=getattr(self,'folder_cache',{})
        if not refresh and key in cache and time.monotonic()-cache[key][0]<30:
            return dict(cache[key][1],cached=True)
        rows = []
        errors = []
        limited = False
        root = binding[0].GetRootFolder() if binding[2] is None else self.default_folder(binding, 6)

        def walk(folder, depth):
            nonlocal limited
            if depth > 20 or len(rows) >= 1000:
                limited = True
                return
            try:
                rows.append({'id': str(folder.EntryID), 'name': str(folder.Name),
                             'path': str(folder.FolderPath), 'depth': depth,
                             'count': folder.Items.Count, 'unread': folder.UnReadItemCount})
                for child in folder.Folders:
                    walk(child, depth + 1)
            except Exception as exc:
                errors.append({'folder':str(getattr(folder,'FolderPath','unknown')),'error':str(exc)})
                return
        walk(root, 0)
        result = {'folders': rows, 'inbox_id': str(self.default_folder(binding, 6).EntryID),
                  'drafts_id': str(self.default_folder(binding, 16).EntryID)}
        try:
            result['sent_id'] = str(self.default_folder(binding, 5).EntryID)
        except Exception:
            result['sent_id'] = None
        result.update(read_at=now(),cached=False,cache_seconds=30,errors=errors,complete=not errors and not limited,truncated=limited)
        if not hasattr(self,'folder_cache'):self.folder_cache={}
        self.folder_cache[key]=(time.monotonic(),result)
        return result

    def signatures(self, email):
        self.mailbox(email)
        return signatures.catalog(email)

    def folder(self, session, binding, folder_id):
        if not folder_id:
            return self.default_folder(binding, 6)
        folder = session.GetFolderFromID(folder_id, binding[0].StoreID)
        if str(folder.StoreID) != str(binding[0].StoreID):
            raise MailError('The folder does not belong to the chosen mailbox.')
        return folder

    def summary(self, item):
        sender=str(item.SenderEmailAddress or '')
        if '@' not in sender:
            try:
                smtp=item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x5D01001F')
                if isinstance(smtp,str) and '@' in smtp:sender=smtp
            except Exception:pass
        return {'id': str(item.EntryID), 'subject': str(item.Subject or ''),
                'sender': str(item.SenderName or ''), 'sender_email':sender,
                'received': received_timestamp(item), 'unread': bool(item.UnRead)}

    def messages(self, email, folder_id=None, offset=0, limit=50, query=''):
        session, binding = self.mailbox(email)
        offset, limit = int(offset), int(limit)
        if offset < 0 or not 1 <= limit <= 200:
            raise MailError('Invalid page size.')
        query = str(query).strip()
        if len(query) > 300:
            raise MailError('Search text is too long.')
        selected_folder = self.folder(session, binding, folder_id)
        # Read one metadata table instead of opening each of 50 COM MailItems.
        # Keep the Items path for providers which do not expose these columns.
        try:
            restriction = '"http://schemas.microsoft.com/mapi/proptag/0x001A001F" like \'IPM.Note%\''
            if query:
                literal = query.replace("'", "''")
                restriction += ' AND (' + ' OR '.join(f'"{field}" like \'%{literal}%\'' for field in
                    ('urn:schemas:httpmail:subject', 'urn:schemas:httpmail:textdescription')) + ')'
            table = selected_folder.GetTable('@SQL=' + restriction)
            table.Columns.RemoveAll()
            for column in ('EntryID','Subject','SenderName','SenderEmailAddress',
                    'http://schemas.microsoft.com/mapi/proptag/0x5D01001F',
                    'http://schemas.microsoft.com/mapi/proptag/0x0E060040','UnRead'):
                table.Columns.Add(column)
            table.Sort('[ReceivedTime]', True)
            total = table.GetRowCount()
            remaining = min(offset, total)
            while remaining:
                skipped = table.GetArray(min(remaining, 500))
                if not skipped: raise ValueError('Outlook tabelpagina ontbreekt.')
                remaining -= len(skipped)
            rows = []
            for identifier, subject, name, raw, smtp, received, unread in table.GetArray(min(limit, max(0,total-offset))) if offset < total else []:
                sender = smtp if isinstance(smtp,str) and '@' in smtp else raw
                rows.append({'id':str(identifier),'subject':str(subject or ''),'sender':str(name or ''),
                    'sender_email':str(sender or ''),'received':timestamp(received),'unread':bool(unread)})
            end = min(offset+len(rows), total)
            return {'items':rows,'total':total,'offset':offset,
                    'next_offset':end if end<total else None,'read_at':now(),'source':'Outlook metadata table'}
        except Exception:
            pass
        items = selected_folder.Items
        if query:
            # DASL literals must escape apostrophes. Search subject/body in this folder.
            literal = query.replace("'", "''")
            filters = ['"urn:schemas:httpmail:subject"', '"urn:schemas:httpmail:textdescription"']
            # The Windows content index can report enabled while newly synced
            # subfolders are not indexed. DASL LIKE searches the actual store.
            items = items.Restrict('@SQL=' + ' OR '.join(f"{f} like '%{literal}%'" for f in filters))
        items.Sort('[ReceivedTime]', True)
        rows = []
        # Offset indexes refer to Outlook's collection, including non-mail items.
        end = min(offset + limit, items.Count)
        for index in range(offset + 1, end + 1):
            item = items.Item(index)
            if item.Class == 43:
                rows.append(self.summary(item))
        return {'items': rows, 'total': items.Count, 'offset': offset,
                'next_offset': end if end < items.Count else None, 'read_at': now()}

    def message(self, email, item_id):
        session, binding = self.mailbox(email)
        item = session.GetItemFromID(item_id, binding[0].StoreID)
        if item.Class != 43 or str(item.Parent.StoreID) != str(binding[0].StoreID):
            raise MailError('This message does not belong to the chosen mailbox.')
        return item, binding

    def read(self, email, item_id):
        item, binding = self.message(email, item_id)
        result = self.summary(item)
        # Plain text only: no tracking pixels or active HTML in the local dashboard.
        result.update(body=str(item.Body or ''), to=str(item.To or ''), cc=str(item.CC or ''), bcc=str(getattr(item, 'BCC', '') or ''),
                      email=str(email),folder_id=str(item.Parent.EntryID),folder_path=str(item.Parent.FolderPath),
                      attachments=[{'name': str(a.FileName), 'size': a.Size} for a in item.Attachments],
                      read_at=now(), sent=bool(item.Sent),
                      is_draft=str(item.Parent.EntryID)==str(self.default_folder(binding, 16).EntryID))
        return result

    def draft(self, email, subject='', body='', to='', cc='', reply_id=None, signature=None, attachments=None,
              request_delivery_receipt=False, request_read_receipt=False):
        if not isinstance(request_delivery_receipt, bool) or not isinstance(request_read_receipt, bool):
            raise MailError('request_delivery_receipt and request_read_receipt must be true or false.')
        files = mail_attachments.prepare(attachments)
        _, binding = self.mailbox(email)
        drafts = self.default_folder(binding, 16)
        if binding[1] is None:
            raise MailError('Drafts for shared mailboxes require your own sender account. Add that account first.')
        if any(len(str(v)) > n for v, n in [(subject, 998), (body, 500000), (to, 10000), (cc, 10000)]):
            raise MailError('The draft is too large.')
        try:
            signature_path = signatures.select(email, reply=bool(reply_id), name=signature)
        except ValueError as exc:
            raise MailError(str(exc)) from None
        signature_images = 0
        if reply_id:
            original, _ = self.message(email, reply_id)
            draft = original.Reply()
            self.set_sender(draft, binding[1])
            draft.Save()
            # Preserve native conversation metadata and quoted original.
            if signature_path:
                signature_images = signatures.compose(draft, body, signature_path, reply=True)
            else:
                draft.Body = str(body) + '\r\n\r\n' + str(draft.Body or '')
            if subject:
                draft.Subject = str(subject)
            if to:
                draft.To = str(to)
            if cc:
                draft.CC = str(cc)
            self.set_sender(draft, binding[1])
            draft.Save()
            if str(draft.Parent.EntryID) != str(drafts.EntryID):
                draft = draft.Move(drafts)
        else:
            draft = drafts.Items.Add('IPM.Note')
            draft.Subject = str(subject)
            draft.To, draft.CC = str(to), str(cc)
            # Bind the account before HTML initializes Outlook's From field.
            self.set_sender(draft, binding[1])
            draft.Save()
            draft.Body = str(body)
            if signature_path:
                signature_images = signatures.compose(draft, body, signature_path)
        # Commit HTML/attachments before binding the account. Outlook's first
        # HTML save can initialize From and otherwise clear an account set earlier.
        for file in files:
            draft.Attachments.Add(file['path'], 1)  # olByValue: embed a copy.
        if request_delivery_receipt:
            draft.OriginatorDeliveryReportRequested = True
        if request_read_receipt:
            draft.ReadReceiptRequested = True
        draft.Save()
        if hasattr(draft, '_oleobj_'):
            identifier = str(draft.EntryID)
            draft = None  # Release the cached HTML compose object before reopening.
            draft = self.session().GetItemFromID(identifier, binding[0].StoreID)
        self.set_draft_sender(draft, binding[1])
        draft.Save()
        if hasattr(draft, '_oleobj_'):
            identifier = str(draft.EntryID)
            draft = None
            draft = self.session().GetItemFromID(identifier, binding[0].StoreID)
        sender_email = self.sender_email(draft)
        if (bool(draft.Sent) or str(draft.Parent.StoreID) != str(binding[0].StoreID)
                or sender_email.lower() != str(binding[1].SmtpAddress).lower()):
            raise MailError('The draft could not be confirmed in the chosen mailbox: '
                            + str(draft.Parent.FolderPath) + '; afzender ' + sender_email + '.')
        self.invalidate_index('draft')
        return {'id': str(draft.EntryID), 'email': email, 'subject': str(draft.Subject),
                'saved_at': now(), 'sent': False, 'folder': str(draft.Parent.FolderPath),
                'signature': signature_path.stem if signature_path else None, 'signature_images': signature_images,
                'sender_email': sender_email, 'attachments': mail_attachments.receipt(files),
                'delivery_receipt_requested': bool(getattr(draft, 'OriginatorDeliveryReportRequested', False)),
                'read_receipt_requested': bool(getattr(draft, 'ReadReceiptRequested', False))}

    def _editable_draft(self, email, item_id):
        """Open an unsent draft in this mailbox's Drafts folder; refuse anything else."""
        item, binding = self.message(email, item_id)
        drafts = self.default_folder(binding, 16)
        if bool(item.Sent) or str(item.Parent.EntryID) != str(drafts.EntryID):
            raise MailError('Only an unsent draft in the Drafts folder of the chosen mailbox can be changed or deleted.')
        if binding[1] is None:
            raise MailError('Drafts for shared mailboxes require your own sender account.')
        key = hashlib.sha256((str(email).lower() + '\0' + str(item_id)).encode()).hexdigest()
        if self.journal_path.exists() and key in json.loads(self.journal_path.read_text('utf-8')):
            raise MailError('This draft was already submitted for sending; it will no longer be changed or deleted.')
        return item, binding

    @staticmethod
    def _visible_attachments(item):
        """(name, 1-based index, size) of ordinary attachments; hidden signature images are excluded."""
        rows = []
        for index in range(1, item.Attachments.Count + 1):
            attachment = item.Attachments.Item(index)
            # A reopened draft no longer reports the hidden flag for signature images; they keep a
            # content-id and the MHTML-reference attach flag. Treat any such part as inline, not removable.
            inline = False
            for prop, test in ((signatures.HIDDEN_PROPERTY, bool), (signatures.CID_PROPERTY, lambda v: bool(str(v or '').strip())),
                               ('http://schemas.microsoft.com/mapi/proptag/0x37140003', lambda v: bool(int(v or 0) & 4))):
                try:
                    inline = inline or test(attachment.PropertyAccessor.GetProperty(prop))
                except Exception:
                    pass
            if not inline:
                rows.append((str(attachment.FileName), index, int(attachment.Size)))
        return rows

    @staticmethod
    def _replace_body(item, body):
        """New body for the authored text only; signature and quoted thread stay as they are."""
        try:
            source = str(item.HTMLBody or '')
        except Exception:
            source = ''
        found = re.search(r'(?is)(<div\b[^>]*\b' + draft_edit.BODY_MARKER + r'=["\']?true["\']?[^>]*>)(.*?)(</div>)', source)
        if found:
            if '<div' in found.group(2).lower():
                raise MailError("This draft's own text contains formatting that cannot be replaced safely. Edit the draft in Outlook.")
            return 'html', source[:found.start(2)] + draft_edit.plain_to_html(body) + source[found.end(2):]
        try:
            replied = bool(str(item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1042001F') or ''))
        except Exception:
            replied = None
        if getattr(item, 'BodyFormat', None) == 1 and replied is False:
            return 'plain', str(body)
        raise MailError("This draft's own text cannot be delimited (no recognizable connector structure, "
                        'or it is a reply/formatted draft). Nothing was changed; edit the text in Outlook.')

    def update_draft(self, email, item_id, **fields):
        spec = draft_edit.parse_update(fields)
        files = mail_attachments.prepare(spec.get('attachments_add'))
        item, binding = self._editable_draft(email, item_id)
        before_sender = self.sender_email(item)
        visible = self._visible_attachments(item)
        removals = spec.get('attachments_remove', [])
        draft_edit.check_removals(removals, visible)
        gone = {name.lower() for name in removals}
        kept = [row for row in visible if row[0].lower() not in gone]
        kept_names = {row[0].lower() for row in kept}
        for file in files:
            if file['name'].lower() in kept_names or sum(f['name'].lower() == file['name'].lower() for f in files) > 1:
                raise MailError('This draft already has an attachment named ' + file['name'] + '.')
        draft_edit.check_totals([row[2] for row in kept], [file['size'] for file in files])
        replacement = self._replace_body(item, spec['body']) if 'body' in spec else None
        # Everything is validated; mutate in one pass and save once.
        for _, index, _ in sorted((row for row in visible if row[0].lower() in gone), key=lambda r: r[1], reverse=True):
            item.Attachments.Remove(index)
        if 'subject' in spec:
            item.Subject = spec['subject']
        for key, attribute in (('to', 'To'), ('cc', 'CC'), ('bcc', 'BCC')):
            if key in spec:
                setattr(item, attribute, spec[key])
        if replacement:
            if replacement[0] == 'html':
                item.HTMLBody = replacement[1]
            else:
                item.Body = replacement[1]
        for file in files:
            item.Attachments.Add(file['path'], 1)  # olByValue: embed a copy.
        item.Save()
        store_id = binding[0].StoreID
        if hasattr(item, '_oleobj_'):
            identifier = str(item.EntryID)
            item = None  # Release the cached compose object before reopening.
            item = self.session().GetItemFromID(identifier, store_id)
        sender_email = self.sender_email(item)
        if before_sender and sender_email.lower() != before_sender.lower() and before_sender.lower() == str(binding[1].SmtpAddress).lower():
            # An HTML save can clear the persisted From; restore the original identity.
            self.set_draft_sender(item, binding[1])
            item.Save()
            if hasattr(item, '_oleobj_'):
                identifier = str(item.EntryID)
                item = None
                item = self.session().GetItemFromID(identifier, store_id)
            sender_email = self.sender_email(item)
        if (bool(item.Sent) or str(item.Parent.StoreID) != str(store_id)
                or str(item.Parent.EntryID) != str(self.default_folder(binding, 16).EntryID)
                or (before_sender and sender_email.lower() != before_sender.lower())):
            raise MailError('The changed draft could not be confirmed in Drafts (sender ' + sender_email
                            + '). Check Drafts; do not retry automatically.')
        self.invalidate_index('draft')
        new_id = str(item.EntryID)
        return {'id': new_id, 'previous_id': str(item_id), 'id_changed': new_id != str(item_id), 'email': email,
                'subject': str(item.Subject), 'saved_at': now(), 'sent': False, 'folder': str(item.Parent.FolderPath),
                'changed': sorted(spec), 'sender_email': sender_email,
                'attachments': [{'name': row[0], 'size': row[2]} for row in self._visible_attachments(item)]}

    def delete_draft(self, email, item_id, reason, expect_subject=None):
        """Move an unsent draft to Deleted Items (reversible). Never touches sent or received mail."""
        if not str(reason or '').strip():
            raise MailError('Give the reason for deleting the draft.')
        item, _ = self._editable_draft(email, item_id)
        subject = str(item.Subject or '')
        if expect_subject is not None and subject != str(expect_subject):
            raise MailError('The draft subject does not match --expect-subject. Nothing was deleted.')
        item.Delete()  # Outlook moves the item to Deleted Items; no permanent deletion.
        self.invalidate_index('draft')
        return {'deleted': True, 'permanent_delete': False, 'email': email, 'id': str(item_id), 'subject': subject,
                'moved_to': 'Deleted Items', 'reason': str(reason).strip(), 'deleted_at': now()}

    def invalidate_index(self,reason):
        import mail_index
        try:mail_index.invalidate(reason,self.index_invalidation_path)
        except OSError:pass

    @staticmethod
    def sender_email(draft):
        account = draft.SendUsingAccount
        if account is not None:
            return str(account.SmtpAddress)
        if hasattr(draft, '_oleobj_'):
            try:
                sender = draft.Sender
                if sender is not None:
                    if sender.Type == 'EX':
                        address = str(sender.GetExchangeUser().PrimarySmtpAddress or '')
                    else:
                        address = str(sender.Address or '')
                    if address:
                        return address
            except Exception:
                pass
        # OOM may return None for reopened drafts. Read the persisted account
        # name rather than incorrectly falling back to the profile default.
        if hasattr(draft, 'PropertyAccessor'):
            try:
                address = str(draft.PropertyAccessor.GetProperty(
                    'http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/8580001F'))
                if address:
                    return address
            except Exception:
                pass
        # Outlook omits the explicit stamp for its primary account. Accept
        # that implicit identity only when BOTH the first account and default
        # store match this draft's own store, never for another mailbox.
        try:
            session = draft.Session
            first = session.Accounts.Item(1)
            if (str(draft.Parent.StoreID) == str(session.DefaultStore.StoreID)
                    == str(first.DeliveryStore.StoreID)):
                return str(first.SmtpAddress)
        except Exception:
            pass
        return ''

    @staticmethod
    def set_draft_sender(draft, account):
        # Persist the visible From identity in a saved draft. The transport
        # account is applied separately immediately before actual Send; setting
        # Sender after that would clear Outlook's in-memory transport choice.
        OutlookBackend.set_sender(draft, account)
        if hasattr(draft, '_oleobj_'):
            import pythoncom
            draft._oleobj_.Invoke(draft._oleobj_.GetIDsOfNames('Sender'), 0,
                                 pythoncom.DISPATCH_PROPERTYPUTREF, 0,
                                 account.CurrentUser.AddressEntry)

    @staticmethod
    def set_sender(draft, account):
        if hasattr(draft, '_oleobj_'):
            import pythoncom
            # Outlook expects an interface reference (PROPERTYPUTREF). The
            # dynamic Python assignment silently leaves SendUsingAccount=None.
            property_id = draft._oleobj_.GetIDsOfNames('SendUsingAccount')
            draft._oleobj_.Invoke(property_id, 0, pythoncom.DISPATCH_PROPERTYPUTREF, 0, account)
            # Read back the native selection before Save. Outlook otherwise
            # may leave the account change pending while committing new HTML.
            selected = draft.SendUsingAccount
            if selected is None or str(selected.SmtpAddress).lower() != str(account.SmtpAddress).lower():
                raise MailError('Outlook did not accept the chosen sender account.')
            # Transport callers must not set Sender after this: that clears
            # the native account. Draft From and send transport are separate.
        else:
            draft.SendUsingAccount = account

    def sync(self):
        session = self.session()
        names = []
        self.sync_events = []
        self.sync_progress = {}
        for sync in session.SyncObjects:
            name = str(sync.Name)
            progress = {'state':'requested','events_available':False}
            self.sync_progress[name] = progress
            try:
                import win32com.client
                def handler_class(state):
                    class Events:
                        def OnSyncStart(self):
                            state.update(state='running', started_at=now())
                        def OnSyncEnd(self):
                            if state['state'] != 'error':
                                state.update(state='completed', completed_at=now())
                        def OnError(self, Code, Description):
                            state.update(state='error', error_code=Code)
                    return Events
                self.sync_events.append(win32com.client.WithEvents(sync, handler_class(progress)))
                progress['events_available'] = True
            except Exception:
                pass
            sync.Start()
            names.append(name)
        self.last_sync_request = now()
        return {'requested_at': self.last_sync_request, 'groups': names,
                'completed': False, 'note': 'Synchronization was requested. Check in Outlook that all folders are up to date.'}

    def send(self, email, item_id):
        # Exact draft IDs are required. No background rule sends mail on its own.
        key = hashlib.sha256((str(email).lower()+'\0'+str(item_id)).encode()).hexdigest()
        journal = json.loads(self.journal_path.read_text('utf-8')) if self.journal_path.exists() else {}
        previous = journal.get(key)
        if previous:
            if previous['state'] == 'submitted':
                return previous
            raise MailError('This send command was started before. Check the Outbox and Sent Items before trying again; the outcome is uncertain.')
        item, binding = self.message(email, item_id)
        drafts = self.default_folder(binding, 16)
        if bool(item.Sent) or str(item.Parent.EntryID) != str(drafts.EntryID):
            raise MailError('Only an unsent draft from the chosen mailbox can be sent.')
        if item.Recipients.Count == 0 or not item.Recipients.ResolveAll():
            raise MailError('The draft has no recipients, or a recipient cannot be resolved.')
        self.set_sender(item, binding[1])
        if item.SendUsingAccount is None or str(item.SendUsingAccount.SmtpAddress).lower() != str(binding[1].SmtpAddress).lower():
            raise MailError('The sender account could not be set. Nothing was sent.')
        item.Save()
        # Intended subject/recipients/attachments and the search key, so
        # verify_sent can find this exact Sent copy later (no body stored).
        receipt = {'email':email,'state':'submitting','requested_at':now(),'submitted':False,'delivery_verified':False,
                   'send_ref':key,'fingerprint':self.fingerprint(item)}

        def persist():
            self.journal_path.parent.mkdir(exist_ok=True)
            temp = self.journal_path.with_suffix('.tmp')
            temp.write_text(json.dumps(journal,ensure_ascii=False),encoding='utf-8')
            os.replace(temp,self.journal_path)
        journal[key] = receipt
        persist()
        try:
            # Reapply the in-memory transport account immediately before Send.
            self.set_sender(item, binding[1])
            item.Send()
        except Exception:
            receipt['state'] = 'uncertain'
            persist()
            raise MailError('Outlook could not confirm the send outcome. Check the Outbox and Sent Items; this command is not repeated automatically.') from None
        receipt.update(state='submitted',submitted=True,submitted_at=now(),
                       note='Submitted to Outlook for sending. Delivery to the recipient is not yet confirmed; check with verify-sent (same --account and --id).')
        persist()
        self.invalidate_index('send')
        return receipt
