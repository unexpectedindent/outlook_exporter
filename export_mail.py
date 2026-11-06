"""Read-only export of MailItem objects to SQLite and the attachment directory."""
import argparse
import logging
from common import (OL_MAIL, OL_MEETING_REQUEST, choose_stores, connect_db,
    connect_outlook, dt, error, iter_folders, load_config, mapi, now,
    PR_INTERNET_MESSAGE_ID, PR_IN_REPLY_TO_ID, safe, save_attachments,
    recipient_rows, save_recipients, sender_info, uid)

LOG = logging.getLogger('outlook_exporter')

def run(cfg):
    ns = connect_outlook()
    db = connect_db()
    completed = 0
    for store in choose_stores(ns, cfg):
        name, store_id = store.DisplayName, store.StoreID
        LOG.info('STORE %s', name)
        for folder in iter_folders(store, cfg):
            path = safe(folder, 'FolderPath', '')
            LOG.info('FOLDER %s', path)
            try:
                items = folder.Items
                count = items.Count
            except Exception as exc:
                error(db, 'folder_items', name, path, exc=exc)
                db.commit()
                continue
            limit = cfg.get('max_items_per_folder', 0) or count
            for index in range(1, min(count, limit) + 1):
                entry_id = ''
                try:
                    item = items.Item(index)
                    objclass = safe(item, 'Class')
                    if objclass != OL_MAIL:
                        if objclass == OL_MEETING_REQUEST and cfg.get('include_meeting_requests_in_mail'):
                            LOG.info('Meeting request skipped as MailItem: %s', safe(item, 'Subject'))
                        continue
                    entry_id = safe(item, 'EntryID')
                    if not entry_id:
                        continue
                    item_uid = uid(store_id, entry_id)
                    existing = db.execute('SELECT modified_at FROM messages WHERE uid=?',(item_uid,)).fetchone()
                    modified = dt(safe(item, 'LastModificationTime'))
                    if existing and existing[0] == modified:
                        continue
                    mail_sender, display_sender = sender_info(item)
                    count_attachments = save_attachments(db, item, 'mail', item_uid, cfg)
                    topic = safe(item, 'ConversationTopic')
                    subject = safe(item, 'Subject') or ''
                    import re
                    is_reply = bool(re.match(r'^\s*(RE|AW|SV|ОТВ)\s*:', subject, re.I))
                    is_forward = bool(re.match(r'^\s*(FW|FWD|WG|ПЕРЕСЛ)\s*:', subject, re.I))
                    ci = safe(item, 'ConversationIndex')
                    if isinstance(ci, (bytes, bytearray)):
                        ci = ci.hex()
                    body = safe(item, 'Body') or ''
                    db.execute('''INSERT INTO messages(uid,store_id,entry_id,store_name,folder_path,
                       internet_message_id,in_reply_to_id,conversation_id,conversation_topic,conversation_index,
                       subject,sender_email,sender_name,sent_at,received_at,modified_at,is_reply,is_forward,
                       body_raw,body_new,body_edited,body_effective,has_attachments,attachment_count,exported_at)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                       ON CONFLICT(uid) DO UPDATE SET
                       folder_path=excluded.folder_path, internet_message_id=excluded.internet_message_id,
                       in_reply_to_id=excluded.in_reply_to_id, conversation_id=excluded.conversation_id,
                       conversation_topic=excluded.conversation_topic, conversation_index=excluded.conversation_index,
                       subject=excluded.subject,sender_email=excluded.sender_email,sender_name=excluded.sender_name,
                       sent_at=excluded.sent_at,received_at=excluded.received_at,modified_at=excluded.modified_at,
                       is_reply=excluded.is_reply,is_forward=excluded.is_forward,body_raw=excluded.body_raw,
                       body_new=NULL,body_edited=NULL,body_effective=NULL,
                       has_attachments=excluded.has_attachments,attachment_count=excluded.attachment_count,
                       exported_at=excluded.exported_at''',
                       (item_uid,store_id,entry_id,name,path,mapi(item,PR_INTERNET_MESSAGE_ID),
                        mapi(item,PR_IN_REPLY_TO_ID),safe(item,'ConversationID'),topic,ci,
                        subject,mail_sender,display_sender,dt(safe(item,'SentOn')),
                        dt(safe(item,'ReceivedTime')),modified,int(is_reply),int(is_forward),
                        body,None,None,None,int(count_attachments>0),count_attachments,now()))
                    rows = recipient_rows(item, 'mail', item_uid)
                    save_recipients(db, rows, 'mail', item_uid)
                    completed += 1
                    if completed % cfg.get('commit_every', 100) == 0:
                        db.commit()
                        LOG.info('Exported %d mail items', completed)
                except Exception as exc:
                    error(db, 'mail_item', name, path, entry_id, exc)
            db.commit()
    LOG.info('DONE: %d mail records updated', completed)
    db.close()

if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('--config', default='config.json')
    args = p.parse_args()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    run(load_config(args.config))
