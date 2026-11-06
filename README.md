# Outlook → SQLite (client-side, read-only)

## Requirements

- **Windows + classic Outlook desktop** (not the *new Outlook*). Outlook must already have the Exchange account and accessible archives configured; user has legal/organizational permission to export corporate mail.
- Python 3.10+ **on that Windows machine**, with architecture compatible with installed Outlook if COM issues arise.
- Only third-party dependency: `pywin32`. No Graph, EWS, Exchange Admin, administrator rights or server credentials used directly.
- Disk space for attachment copies and full-text database; export can be slow, and Outlook can block automation through corporate policy.

## Setup / first export

```powershell
cd C:\path\to\outlook_exporter
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python list_sources.py
Copy-Item config.example.json config.json
notepad config.json
python run_all.py --config config.json
```

In `config.json`, fill `primary_account_smtp` with the actual SMTP returned by the discovery script. The primary account **must match**; there is no silent fallback. Archives mapped as **additional Outlook stores** must be explicitly listed by exact `DisplayName` or `FilePath`; an archive that isn't mounted/visible to this Outlook profile cannot be exported. Other accounts/stores remain untouched. If several stores have the same display name, prefer explicit file paths (for PST files); no name-based assumption about owner is made. Be careful when choosing archives from other accounts.

## Commands

- `python list_sources.py` — list accounts and available stores; does not export messages.
- `python export_mail.py --config config.json` — email, nested folders and attachments.
- `python export_calendar.py --config config.json` — calendar appointments/meeting objects, including non-default calendar folders and archives.
- `python thread_text.py --config config.json` — derived text extraction/diff for all emails.
- `python run_all.py --config config.json` — all three steps.
- `python -m unittest discover -s tests -v` — text unit tests (works without Outlook).

## Data

`data/outlook.sqlite3` with tables `messages`, `meetings`, `recipients`, `attachments`, `export_errors`.

- `uid` is a deterministic 24-character SHA-256 prefix of **StoreID + EntryID**; `entry_id` remains stored exactly. Message IDs may change if mail is moved to a different store; InternetMessageID (when available) can help recognize duplicates across stores.
- `conversation_id`: Outlook ConversationID. Some historical mails may lack it; they are intentionally **not grouped solely by subject**.
- `recipients`: `mail` roles `to`, `cc`, `bcc`; `meeting` roles `required`, `optional`, `resource`. Exchange SMTP lookup is best-effort; some unresolved distribution-list or legacy EX recipients may have missing SMTP.
- `body_raw` is verbatim Outlook plain-text Body (not HTML); `body_new` strips detected quoted chain; `body_edited` holds newly added or modified lines *inside the quote*, when confidence is high; `body_effective` combines the two. **Not a guaranteed semantic diff**: Outlook quote formatting and top-posted/inline replies are not always machine-distinguishable. Review diff results and raw mail for important decisions. No original body is overwritten.
- File names: `data/attachments/mail/<UID>__NNN__original.ext` and `data/attachments/meeting/...`. The DB stores the relative paths. Embedded files/images may appear as attachments. A failed save is recorded in `attachments.export_error`.
- `is_reply` / `is_forward`: subject-prefix heuristics, not a MAPI-authenticated action flag. Forwarded mail may have a new conversation and must not automatically be treated as a prior-message quote.
- Calendar exports **recurrence series and stored exceptions**, not materialized future instances; repeating series are **not expanded** (to avoid infinite series). Calendar `start_at`/`end_at` reflect Outlook local date/time. Meetings/appointments are separate from incoming invitation messages (MeetingItem objects), which this MVP does not export.
- `max_items_per_folder: 0` = all items, for a trial use e.g. `20`; reset to `0` for full export. Existing records with the same `LastModificationTime` are skipped; changes to filtering or attachment export settings require deleting/rebuilding the database to refresh old records.
- Excluded folder subtrees are skipped based on default folder IDs, names and path substring patterns. Custom named folders matching an exclusion name are also excluded, so review config.

## SQLite queries

```sql
SELECT uid, conversation_id, subject, sender_email, sent_at, body_effective
FROM messages ORDER BY sent_at;

SELECT m.subject, m.start_at, r.role, r.email, r.display_name
FROM meetings m JOIN recipients r
ON r.owner_kind='meeting' AND r.owner_uid=m.uid
ORDER BY m.start_at;

SELECT a.owner_uid, a.original_name, a.relative_path
FROM attachments a WHERE a.export_error IS NULL;

SELECT * FROM export_errors ORDER BY logged_at DESC;
```

## Limitations and operating safely

1. Works only with content accessible in **classic Outlook** (cached `.ost`, open `.pst`, online accessible folders). If Outlook cache includes only recent mail, older server data might require adjusting sync/caching settings or opening a supported archive; this script cannot bypass Exchange/client restrictions.
2. Avoid running simultaneously with mass mail moves, and keep Outlook open. Outlook COM object model can be blocked by IT policies. On very large mailboxes, export can take a long time.
3. This tool is **read-only toward Outlook**; it writes only into its own SQLite DB and file directory. It never sends mail, calls Exchange administration APIs or changes Outlook items.
4. `EntryID` isn't globally immutable; cross-store moves can create duplicates. Real de-duplication requires a subsequent policy based on InternetMessageID + date/hash, not solely UID.
5. Plain text does not preserve HTML formatting and inline graphics layout; attachments are saved separately. Embedded `.msg` files are copied, not recursively parsed.
6. Message body quote handling is heuristic; signatures or quoted content may be detected incorrectly. Full text is retained for recovery/audit.
7. Database and attachment copies are sensitive corporate data: store on an approved encrypted corporate drive; respect IT policy and personal-data protections. Do **not** upload the export to personal cloud services.
