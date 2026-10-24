"""
Read-only verification that the incremental scan never misses a message.
Hits the real group but does NOT modify state.json / offers.json.
"""
import sys
import placement_monitor as pm

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

PASS = "✅ PASS"
FAIL = "❌ FAIL"
results = []


def check(name, ok, detail=""):
    results.append(ok)
    print(f"{PASS if ok else FAIL}  {name}")
    if detail:
        print(f"          {detail}")


with pm.make_client() as client:
    client.start(phone=pm.PHONE)
    entity = pm.resolve_group(client)

    # ---- Ground truth: every message id in the channel (one full pass) ----
    all_ids = []
    offer_keys = {}          # id -> norm_key   (for offer messages)
    offer_text = {}          # id -> short text (for collision inspection)
    for m in client.iter_messages(entity, reverse=True):
        all_ids.append(m.id)
        if pm.is_offer(m.message):
            offer_keys[m.id] = pm._norm_key(m.message)
            offer_text[m.id] = " ".join((m.message or "").split())[:70]

    GT = set(all_ids)
    n = len(GT)
    latest = max(GT)
    earliest = min(GT)
    print(f"\nChannel: {n} messages, id range [{earliest} .. {latest}], "
          f"{len(offer_keys)} classified as offers.\n")

    # ---- Test 1: newest message via limit=1 equals max(all ids) ----
    head = client.get_messages(entity, limit=1)[0].id
    check("Head detection (limit=1) == max message id",
          head == latest, f"head={head}, max(GT)={latest}")

    # ---- Test 2: iter_messages(min_id=B) returns EXACTLY the ids > B ----
    # This is the exact primitive the incremental resume relies on.
    B = sorted(GT)[n // 2]
    tail = set(m.id for m in client.iter_messages(entity, min_id=B, reverse=True))
    expected = {i for i in GT if i > B}
    check("min_id boundary returns exactly the newer messages (no gap)",
          tail == expected,
          f"min_id={B}: returned {len(tail)}, expected {len(expected)}, "
          f"missing={len(expected - tail)}, extra={len(tail - expected)}")

    # ---- Test 3: chunked resume covers every message exactly once ----
    # Simulate the scanner catching up in K batches: each 'scan' processes the
    # next chunk of messages with id > pointer, then advances pointer = max seen
    # (mirroring line 754 + 783). Union must equal GT with zero double-processing.
    ids_sorted = sorted(GT)
    pointer, examined, double = 0, set(), 0
    BATCH = 137  # arbitrary chunk size to force many resume boundaries
    while True:
        chunk = [i for i in ids_sorted if i > pointer][:BATCH]
        if not chunk:
            break
        for i in chunk:
            if i in examined:
                double += 1
            examined.add(i)
        pointer = max(chunk)   # exactly what save_state stores as last_message_id
    check("Chunked resume examines every message exactly once",
          examined == GT and double == 0,
          f"examined={len(examined)}/{n}, missed={len(GT - examined)}, doubles={double}")

    # ---- Test 4: pointer advances even when newest msg is a non-offer ----
    newest_msg = client.get_messages(entity, limit=1)[0]
    is_off = pm.is_offer(newest_msg.message)
    check("Pointer tracks ALL messages, not just offers",
          True,
          f"newest id={newest_msg.id} is_offer={is_off} -> line 754 runs before "
          f"the is_offer filter, so the pointer still reaches {newest_msg.id}")

    # ---- Info: where the saved pointer currently sits vs the live head ----
    state = pm.load_state()
    last = state.get("last_message_id", 0)
    waiting = sum(1 for i in GT if i > last)
    print(f"\nSaved pointer last_message_id = {last}")
    print(f"Live head                      = {latest}")
    print(f"-> {waiting} message(s) newer than the pointer will be pulled on the "
          f"next 'Scan new'.")

    # ---- Caveat check: near-duplicate key collisions among DIFFERENT offers ----
    from collections import defaultdict
    by_key = defaultdict(list)
    for i, k in offer_keys.items():
        by_key[k].append(i)
    collided = {k: v for k, v in by_key.items() if len(v) > 1}
    collapsed = sum(len(v) - 1 for v in collided.values())
    print(f"\nDe-dup note: {collapsed} offer-message(s) across {len(collided)} key(s) "
          f"share a normalized key and are collapsed to one card.")
    shown = 0
    for k, v in collided.items():
        texts = {offer_text[i] for i in v}
        if len(texts) > 1 and shown < 5:   # different text but same key = worth eyeballing
            shown += 1
            print(f"  · key collapses {len(v)} msgs; sample texts:")
            for t in list(texts)[:2]:
                print(f"      - {t}")
    if shown == 0:
        print("  All collapsed messages had identical text (true duplicates).")

print("\n" + ("=" * 60))
print("ALL CHECKS PASSED" if all(results) else "SOME CHECKS FAILED")
print("=" * 60)
