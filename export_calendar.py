"""Read-only export of actual calendar AppointmentItem entries, not invitation mail."""
import argparse
import logging
from common import (OL_APPOINTMENT, choose_stores, connect_db, connect_outlook,
    dt, error, iter_folders, load_config, now, recipient_rows, safe,
    save_attachments, save_recipients, uid)

LOG = logging.getLogger('outlook_exporter')

def run(cfg):
    ns, db = connect_outlook(), connect_db()
    completed = 0
    for store in choose_stores(ns, cfg):
        name, store_id = store.DisplayName, store.StoreID
        for folder in iter_folders(store, cfg):
            path = safe(folder, 'FolderPath', '')
            try:
                # Do not use IncludeRecurrences=True: that creates virtual instances
                # and may be unbounded for no-end recurring series.
                items, count = folder.Items, folder.Items.Count
            except Exception as exc:
                error(db, 'calendar_folder', name, path, exc=exc)
                db.commit()
                continue
            limit = cfg.get('max_items_per_folder', 0) or count
            for index in range(1, min(count, limit) + 1):
                entry_id = ''
                try:
                    item = items.Item(index)
                    if safe(item, 'Class') != OL_APPOINTMENT:
                        continue
                    entry_id = safe(item, 'EntryID')
                    if not entry_id:
                        continue
                    item_uid = uid(store_id, entry_id)
                    modified = dt(safe(item, 'LastModificationTime'))
                    existing = db.execute('SELECT modified_at FROM meetings WHERE uid=?', (item_uid,)).fetchone()
                    if existing and existing[0] == modified:
                        continue
                    count_attachments = save_attachments(db,item,'meeting',item_uid,cfg)
                    rows = recipient_rows(item,'meeting',item_uid,meeting=True)
                    save_recipients(db,rows,'meeting',item_uid)
                    organizer = safe(item,'Organizer')
                    organizer_email = None
                    # Organizer is only guaranteed to be a display name, not SMTP.
                    # If organizer is among attendees, resolve by matching the name.
                    for row in rows:
                        if (row[5] or '').casefold() == (organizer or '').casefold():
                            organizer_email = row[4]
                            break
                    db.execute('''INSERT INTO meetings(uid,store_id,entry_id,store_name,folder_path,
                        global_appointment_id,conversation_id,subject,organizer_name,organizer_email,
                        start_at,end_at,created_at,modified_at,is_recurring,recurrence_state,meeting_status,
                        body_raw,has_attachments,attachment_count,exported_at)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        ON CONFLICT(uid) DO UPDATE SET
                        folder_path=excluded.folder_path,global_appointment_id=excluded.global_appointment_id,
                        conversation_id=excluded.conversation_id,subject=excluded.subject,
                        organizer_name=excluded.organizer_name,organizer_email=excluded.organizer_email,
                        start_at=excluded.start_at,end_at=excluded.end_at,created_at=excluded.created_at,
                        modified_at=excluded.modified_at,is_recurring=excluded.is_recurring,
                        recurrence_state=excluded.recurrence_state,meeting_status=excluded.meeting_status,
                        body_raw=excluded.body_raw,has_attachments=excluded.has_attachments,
                        attachment_count=excluded.attachment_count,exported_at=excluded.exported_at''',
                        (item_uid,store_id,entry_id,name,path,safe(item,'GlobalAppointmentID'),
                         safe(item,'ConversationID'),safe(item,'Subject'),organizer,organizer_email,
                         dt(safe(item,'Start')),dt(safe(item,'End')),
                         dt(safe(item,'CreationTime')),modified,int(bool(safe(item,'IsRecurring',False))),
                         safe(item,'RecurrenceState'),safe(item,'MeetingStatus'),safe(item,'Body') or '',
                         int(count_attachments>0),count_attachments,now()))
                    completed += 1
                    if completed % cfg.get('commit_every', 100) == 0:
                        db.commit()
                        LOG.info('Exported %d calendar items',completed)
                except Exception as exc:
                    error(db,'calendar_item',name,path,entry_id,exc)
            db.commit()
    LOG.info('DONE: %d appointments/meetings updated', completed)
    db.close()

if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('--config', default='config.json')
    args = p.parse_args()
    logging.basicConfig(level=logging.INFO,format='%(asctime)s %(levelname)s %(message)s')
    run(load_config(args.config))
