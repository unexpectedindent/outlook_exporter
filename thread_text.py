"""Post-process email threads without mutating Outlook or discarding originals.

Conservative heuristic: identify Outlook-style nested header blocks and quoted
material, then compare quoted body with a prior mail in the same conversation.
Inline edits are surfaced separately with context; raw body is never destroyed.
"""
from __future__ import annotations
import argparse
import difflib
import logging
import re
from collections import defaultdict
from common import connect_db, load_config

LOG = logging.getLogger('outlook_exporter')
HEADER = re.compile(r'^\s*(?:From|От|Von|De)\s*:\s*\S+', re.I)
META = re.compile(r'^\s*(?:Sent|Отправлено|Date|Дата|To|Кому|Cc|Копия|Subject|Тема|From|От|Від|Von)\s*:', re.I)
ON_WROTE = re.compile(r'^\s*(?:On\s+.+\s+wrote:|.+\s+писал\s*\(?[а-я]*\)?:|.+\s+написал\s*\(?[а-я]*\)?:)\s*$', re.I)
QUOTED = re.compile(r'^\s*>\s?')
SEPARATOR = re.compile(r'^\s*(?:-{4,}\s*(?:Original Message|Исходное сообщение)\s*-*|_{10,}|-{12,})\s*$', re.I)


def normalize(line):
    """Ignore quote characters, whitespace and transport metadata when comparing."""
    line = QUOTED.sub('', line)
    return re.sub(r'\s+', ' ', line).strip().casefold()


def drop_signatures(lines, config):
    begins = [s.casefold() for s in config.get('ignore_signature_from', [])]
    for i, line in enumerate(lines):
        if normalize(line) in begins:
            return lines[:i]
    return lines


def strip_meta(lines):
    return [line for line in lines if not META.match(line)]


def split_thread(text):
    """Return (new leading content, first quoted message body, detected boundary).

    Header blocks need at least 2 recognized metadata header lines close together,
    including From and either Sent/To/Subject. This reduces false cuts in prose.
    """
    lines = (text or '').splitlines()
    boundary = None
    quote_start = None
    for i, line in enumerate(lines):
        if ON_WROTE.match(line):
            boundary, quote_start = i, i+1
            break
        if SEPARATOR.match(line):
            boundary, quote_start = i, i+1
            break
        if HEADER.match(line):
            candidates = lines[i:i+7]
            nmeta = sum(bool(META.match(x)) for x in candidates)
            if nmeta >= 3:
                boundary = i
                # Consume contiguous transport header block (+ empty lines)
                j = i
                while j < len(lines) and (META.match(lines[j]) or not lines[j].strip()):
                    j += 1
                quote_start = j
                break
        if QUOTED.match(line) and sum(bool(QUOTED.match(x)) for x in lines[i:i+5]) >= 3:
            boundary, quote_start = i, i
            break
    if boundary is None:
        return text or '', '', False
    leading = '\n'.join(lines[:boundary]).strip()
    quoted = lines[quote_start:]
    # Stop at a second nested metadata block if present.
    for j in range(1, len(quoted)):
        if HEADER.match(quoted[j]) and sum(bool(META.match(x)) for x in quoted[j:j+7]) >= 3:
            quoted = quoted[:j]
            break
    return leading, '\n'.join(quoted).strip(), True


def compare_quote(current_quote, older_raw, diff_cfg):
    """Return added/changed quoted lines, only if alignment is convincing.

    This is deliberately conservative: reformatting and inline signature changes
    cannot be reliably separated from semantic edits in every mail client.
    """
    if not current_quote or not older_raw:
        return ''
    _, old_lead, _ = None, None, None
    # Previous message's own top content is most often quoted by Outlook.
    old_lead, _, _ = split_thread(older_raw)
    old_lines = drop_signatures(strip_meta(old_lead.splitlines()), diff_cfg)
    current_lines = drop_signatures(strip_meta(current_quote.splitlines()), diff_cfg)
    old_clean = [normalize(x) for x in old_lines if normalize(x)]
    cur_clean = [normalize(x) for x in current_lines if normalize(x)]
    if min(len(cur_clean), len(old_clean)) < diff_cfg.get('min_quoted_lines', 3):
        return ''
    # Use a local window to prevent comparing against an entirely different chain.
    # Ratios are sensitive to long edits. We intentionally favor false negatives.
    matcher = difflib.SequenceMatcher(None, old_clean, cur_clean, autojunk=False)
    if matcher.ratio() < diff_cfg.get('similarity_threshold', 0.72):
        return ''
    edits = []
    for operation, _a, _b, start, end in matcher.get_opcodes():
        if operation in ('insert', 'replace'):
            edits.extend(cur_clean[start:end])
    return '\n'.join(edits).strip()


def process(config):
    db = connect_db()
    rows = db.execute('''SELECT uid, conversation_id, conversation_topic,
                                internet_message_id, in_reply_to_id, sent_at, received_at,
                                body_raw, subject, conversation_index
                         FROM messages''').fetchall()
    grouped = defaultdict(list)
    for r in rows:
        # Never group different independent conversations solely by subject.
        conv = r[1] or f'no-conversation:{r[0]}'
        grouped[conv].append(r)
    for group in grouped.values():
        # SentOn order; ConversationIndex is auxiliary and can vary by store.
        group.sort(key=lambda r: (r[5] or r[6] or '', r[9] or '', r[0]))
        by_mid = {r[3]: r for r in group if r[3]}
        for i, row in enumerate(group):
            uid, _cid, _topic, _mid, reply_to, _sent, _recv, body, _subject, _idx = row
            leading, quoted, detected = split_thread(body)
            older = by_mid.get(reply_to)
            if older is None and i:
                older = group[i-1]
            edited = compare_quote(quoted, older[7], config['diff']) if detected and older else ''
            parts = [x for x in (leading.strip(), edited.strip()) if x]
            effective = '\n\n[ИЗМЕНЕНИЯ В ЦИТИРУЕМОМ ТЕКСТЕ]\n' .join(parts)
            db.execute('UPDATE messages SET body_new=?, body_edited=?, body_effective=? WHERE uid=?',
                       (leading, edited, effective, uid))
    db.commit()
    LOG.info('Updated text extraction for %d messages in %d conversations',len(rows),len(grouped))
    db.close()

if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--config', default='config.json')
    args=parser.parse_args()
    logging.basicConfig(level=logging.INFO)
    process(load_config(args.config))
