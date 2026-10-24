"""Dump EVERY raw message from the group, unfiltered, to JSON + readable TXT."""
import json
import sys
import placement_monitor as pm

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

rows = []
name_cache = {}

with pm.make_client() as client:
    client.start(phone=pm.PHONE)
    entity = pm.resolve_group(client)
    for m in client.iter_messages(entity, reverse=True):   # oldest -> newest
        sender = ""
        if m.sender_id is not None:
            if m.sender_id not in name_cache:
                try:
                    s = m.get_sender()
                    name_cache[m.sender_id] = (
                        getattr(s, "title", None)
                        or " ".join(filter(None, [getattr(s, "first_name", ""),
                                                   getattr(s, "last_name", "")])).strip()
                        or f"id{m.sender_id}")
                except Exception:
                    name_cache[m.sender_id] = f"id{m.sender_id}"
            sender = name_cache[m.sender_id]
        rows.append({
            "id": m.id,
            "date": m.date.astimezone().strftime("%Y-%m-%d %H:%M:%S"),
            "sender": sender,
            "has_media": m.media is not None,
            "is_offer": pm.is_offer(m.message),
            "text": m.message or "",
        })

with open("raw_messages.json", "w", encoding="utf-8") as f:
    json.dump(rows, f, ensure_ascii=False, indent=2)

with open("raw_messages.txt", "w", encoding="utf-8") as f:
    f.write(f"RAW TELEGRAM DUMP  |  {len(rows)} messages  |  group {pm.GROUP}\n")
    f.write("=" * 80 + "\n\n")
    for r in rows:
        tag = "OFFER" if r["is_offer"] else "     "
        media = " [media]" if r["has_media"] else ""
        f.write(f"[{tag}] #{r['id']}  {r['date']}  — {r['sender']}{media}\n")
        f.write((r["text"] or "(no text)") + "\n")
        f.write("-" * 80 + "\n")

print(f"Wrote {len(rows)} messages to raw_messages.json and raw_messages.txt")
print(f"Offers: {sum(r['is_offer'] for r in rows)}  |  "
      f"With media: {sum(r['has_media'] for r in rows)}")
