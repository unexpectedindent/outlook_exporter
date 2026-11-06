"""Shared COM/SQLite helpers. Execute on Windows with classic Outlook installed."""
from __future__ import annotations
import hashlib
import json
import logging
import re
import sqlite3
from datetime import datetime
from pathlib import Path

LOG = logging.getLogger('outlook_exporter')
ROOT = Path(__file__).resolve().parent
DATA = ROOT / 'data'
DB_FILE = DATA / 'outlook.sqlite3'
ATTACH_DIR = DATA / 'attachments'

# Outlook enums: OlDefaultFolders, OlObjectClass
DEFAULT_FOLDERS = {'deleted': 3, 'outbox': 4, 'sent': 5, 'inbox': 6,
                   'calendar': 9, 'drafts': 16, 'junk': 23, 'sync_issues': 20}
OL_MAIL = 43
OL_APPOINTMENT = 26
OL_MEETING_REQUEST = 53
PR_INTERNET_MESSAGE_ID = 'http://schemas.microsoft.com/mapi/proptag/0x1035001F'
PR_IN_REPLY_TO_ID = 'http://schemas.microsoft.com/mapi/proptag/0x1042001F'
PR_SEARCH_KEY = 'http://schemas.microsoft.com/mapi/proptag/0x300B0102'
PR_SMTP_ADDRESS = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001F'


def load_config(path: str | Path) -> dict:
    with open(path, encoding='utf-8') as f:
        cfg = json.load(f)
    if not cfg.get('primary_account_smtp'):
        raise ValueError('Set primary_account_smtp in config.json')
    return cfg


def safe(obj, attr, default=None):
    try:
        value = getattr(obj, attr)
        return default if value is None else value
    except Exception:
        return default


def mapi(obj, prop):
    try:
        return obj.PropertyAccessor.GetProperty(prop)
    except Exception:
        return None


def dt(value):
    if value is None:
        return None
    try:
        return value.isoformat(sep=' ')
    except (AttributeError, ValueError, TypeError):
        return str(value)


def uid(store_id: str, entry_id: str) -> str:
    return hashlib.sha256((store_id + '\0' + entry_id).encode('utf-8')).hexdigest()[:24]


def clean_filename(name: str, limit=140) -> str:
    name = (name or 'unnamed').replace('\\', '/').split('/')[-1]
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip(' .')
    if not name:
        name = 'unnamed'
    return name[:limit]


def connect_db() -> sqlite3.Connection:
    DATA.mkdir(exist_ok=True, parents=True)
    ATTACH_DIR.mkdir(exist_ok=True, parents=True)
    conn = sqlite3.connect(DB_FILE)
    conn.execute('PRAGMA foreign_keys = ON')
    conn.execute('PRAGMA journal_mode = WAL')
    conn.executescript('''
    CREATE TABLE IF NOT EXISTS messages (
      uid TEXT PRIMARY KEY, store_id TEXT NOT NULL, entry_id TEXT NOT NULL,
      store_name TEXT, folder_path TEXT, internet_message_id TEXT, in_reply_to_id TEXT,
      conversation_id TEXT, conversation_topic TEXT, conversation_index TEXT,
      subject TEXT, sender_email TEXT, sender_name TEXT,
      sent_at TEXT, received_at TEXT, modified_at TEXT,
      is_reply INTEGER, is_forward INTEGER,
      body_raw TEXT, body_new TEXT, body_edited TEXT, body_effective TEXT,
      has_attachments INTEGER, attachment_count INTEGER, exported_at TEXT NOT NULL,
      UNIQUE (store_id, entry_id)
    );
    CREATE TABLE IF NOT EXISTS recipients (
      owner_kind TEXT NOT NULL, owner_uid TEXT NOT NULL, role TEXT NOT NULL,
      ordinal INTEGER NOT NULL, email TEXT, display_name TEXT,
      PRIMARY KEY (owner_kind, owner_uid, role, ordinal)
    );
    CREATE TABLE IF NOT EXISTS meetings (
      uid TEXT PRIMARY KEY, store_id TEXT NOT NULL, entry_id TEXT NOT NULL,
      store_name TEXT, folder_path TEXT, global_appointment_id TEXT,
      conversation_id TEXT, subject TEXT, organizer_name TEXT, organizer_email TEXT,
      start_at TEXT, end_at TEXT, created_at TEXT, modified_at TEXT,
      is_recurring INTEGER, recurrence_state INTEGER, meeting_status INTEGER,
      body_raw TEXT, has_attachments INTEGER, attachment_count INTEGER,
      exported_at TEXT NOT NULL, UNIQUE (store_id, entry_id)
    );
    CREATE TABLE IF NOT EXISTS attachments (
      owner_kind TEXT NOT NULL, owner_uid TEXT NOT NULL, ordinal INTEGER NOT NULL,
      original_name TEXT, relative_path TEXT, size_bytes INTEGER,
      content_id TEXT, is_inline INTEGER, export_error TEXT,
      PRIMARY KEY (owner_kind, owner_uid, ordinal)
    );
    CREATE TABLE IF NOT EXISTS export_errors (
      id INTEGER PRIMARY KEY AUTOINCREMENT, stage TEXT, store_name TEXT,
      folder_path TEXT, entry_id TEXT, error TEXT, logged_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_messages_conv ON messages(conversation_id, sent_at);
    CREATE INDEX IF NOT EXISTS idx_messages_mid ON messages(internet_message_id);
    CREATE INDEX IF NOT EXISTS idx_meetings_start ON meetings(start_at);
    ''')
    return conn


def error(conn, stage, store='', folder='', entry='', exc=None):
    LOG.exception('Error %s store=%s folder=%s entry=%s: %s', stage, store, folder, entry, exc)
    conn.execute('INSERT INTO export_errors(stage, store_name, folder_path, entry_id, error, logged_at) VALUES (?,?,?,?,?,?)',
                 (stage, store, folder, entry, str(exc), datetime.now().isoformat()))


def connect_outlook():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError('Install pywin32 and execute using Windows Python') from exc
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch('Outlook.Application')
    return app.GetNamespace('MAPI')


def list_stores(ns):
    for index in range(1, ns.Stores.Count + 1):
        store = ns.Stores.Item(index)
        yield store


def choose_stores(ns, config):
    """No implicit selection of other accounts or archives."""
    target = config['primary_account_smtp'].casefold().strip()
    account_store_id = None
    for i in range(1, ns.Accounts.Count + 1):
        acc = ns.Accounts.Item(i)
        if (safe(acc, 'SmtpAddress', '') or '').casefold() == target:
            delivery = safe(acc, 'DeliveryStore')
            if delivery:
                account_store_id = safe(delivery, 'StoreID')
            break
    if not account_store_id:
        raise RuntimeError(f'Account {target!r} or its DeliveryStore not found. Run list_sources.py')

    extra_names = {x.casefold() for x in config.get('additional_archive_store_names', [])}
    extra_paths = {str(Path(x)).casefold() for x in config.get('additional_archive_store_paths', [])}
    stores = []
    for store in list_stores(ns):
        store_id = safe(store, 'StoreID', '')
        name = safe(store, 'DisplayName', '')
        filepath = safe(store, 'FilePath', '') or ''
        primary = store_id == account_store_id
        extra = name.casefold() in extra_names or filepath.casefold() in extra_paths
        if (primary and config.get('include_primary_store', True)) or extra:
            stores.append(store)
    if not stores:
        raise RuntimeError('No stores selected; check config.json')
    return stores


def excluded_folder_ids(store, cfg):
    mapping = cfg.get('exclude_default_folders', [])
    ids = set()
    for name in mapping:
        folder_type = DEFAULT_FOLDERS.get(name.casefold())
        if folder_type is None:
            raise ValueError(f'Unknown default folder: {name}')
        try:
            folder = store.GetDefaultFolder(folder_type)
            if folder is not None:
                ids.add(folder.EntryID)
        except Exception:
            pass
    return ids


def iter_folders(store, cfg):
    blocked_ids = excluded_folder_ids(store, cfg)
    blocked_names = {x.casefold() for x in cfg.get('exclude_folder_names', [])}
    blocked_paths = [x.casefold() for x in cfg.get('exclude_folder_path_contains', [])]
    root = store.GetRootFolder()
    stack = [root]
    while stack:
        folder = stack.pop()
        name = safe(folder, 'Name', '') or ''
        path = safe(folder, 'FolderPath', '') or ''
        if (safe(folder, 'EntryID') in blocked_ids or name.casefold() in blocked_names
            or any(x in path.casefold() for x in blocked_paths)):
            LOG.info('SKIP folder %s', path)
            continue
        yield folder
        try:
            children = folder.Folders
            for i in range(children.Count, 0, -1):
                stack.append(children.Item(i))
        except Exception as exc:
            LOG.warning('Cannot enumerate children of %s: %s', path, exc)


def smtp_from_address_entry(entry):
    if entry is None:
        return None
    typ = (safe(entry, 'Type', '') or '').upper()
    if typ == 'EX':
        user = safe(entry, 'GetExchangeUser')
        if callable(user):
            try:
                ex = user()
                email = safe(ex, 'PrimarySmtpAddress')
                if email:
                    return email
            except Exception:
                pass
        dist = safe(entry, 'GetExchangeDistributionList')
        if callable(dist):
            try:
                ex = dist()
                email = safe(ex, 'PrimarySmtpAddress')
                if email:
                    return email
            except Exception:
                pass
    return mapi(entry, PR_SMTP_ADDRESS) or safe(entry, 'Address')


def sender_info(item):
    name = safe(item, 'SenderName') or safe(item, 'Organizer')
    email = safe(item, 'SenderEmailAddress')
    entry = safe(item, 'Sender')
    if entry is not None:
        email = smtp_from_address_entry(entry) or email
    if email and str(email).upper().startswith('/O='):
        email = None
    return email, name


def recipient_rows(item, owner_kind, owner_uid, meeting=False):
    """Mail type 1=To 2=Cc 3=Bcc; Meeting 1=required 2=optional 3=resource."""
    rows = []
    recipients = safe(item, 'Recipients')
    if recipients is None:
        return rows
    try:
        count = recipients.Count
    except Exception:
        return rows
    for i in range(1, count + 1):
        try:
            r = recipients.Item(i)
            kind = int(safe(r, 'Type', 0) or 0)
            roles = {1:'required', 2:'optional', 3:'resource'} if meeting else {1:'to', 2:'cc', 3:'bcc'}
            email = smtp_from_address_entry(safe(r, 'AddressEntry')) or safe(r, 'Address')
            if email and str(email).upper().startswith('/O='):
                email = None
            rows.append((owner_kind, owner_uid, roles.get(kind, 'other'), i, email, safe(r, 'Name')))
        except Exception as exc:
            LOG.warning('Recipient #%d of %s: %s', i, owner_uid, exc)
    return rows


def save_recipients(conn, rows, owner_kind, owner_uid):
    conn.execute('DELETE FROM recipients WHERE owner_kind=? AND owner_uid=?', (owner_kind, owner_uid))
    conn.executemany('INSERT INTO recipients(owner_kind,owner_uid,role,ordinal,email,display_name) VALUES (?,?,?,?,?,?)', rows)


def save_attachments(conn, item, kind, item_uid, cfg):
    conn.execute('DELETE FROM attachments WHERE owner_kind=? AND owner_uid=?', (kind, item_uid))
    atts = safe(item, 'Attachments')
    if atts is None:
        return 0
    try:
        count = atts.Count
    except Exception:
        return 0
    for i in range(1, count + 1):
        path, problem, original, sz, cid, inline = None, None, None, None, None, False
        try:
            a = atts.Item(i)
            original = safe(a, 'FileName') or f'attachment_{i}'
            sz = safe(a, 'Size')
            cid = mapi(a, 'http://schemas.microsoft.com/mapi/proptag/0x3712001F')
            inline = bool(cid)
            if cfg.get('export_attachments', True) and (cfg.get('save_inline_attachments', True) or not inline):
                dest_name = f'{item_uid}__{i:03d}__{clean_filename(original)}'
                rel = Path(kind) / dest_name
                dest = ATTACH_DIR / rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                a.SaveAsFile(str(dest))
                path = str(rel).replace('\\', '/')
        except Exception as exc:
            problem = str(exc)
            LOG.warning('Attachment %s/%s: %s', item_uid, i, exc)
        conn.execute('''INSERT INTO attachments(owner_kind,owner_uid,ordinal,original_name,relative_path,
                      size_bytes,content_id,is_inline,export_error) VALUES(?,?,?,?,?,?,?,?,?)''',
                     (kind, item_uid, i, original, path, sz, cid, int(inline), problem))
    return count


def now():
    return datetime.now().isoformat(timespec='seconds')
