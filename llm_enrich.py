"""
Local LLM enrichment via Ollama — 100% on your machine, no API, no tokens.

For each offer message it asks a local model to pull out clean structured fields
(company, role, eligibility, pay, skills, location, deadline, min CGPA, a one-line
summary). Results are cached by message hash so each unique post is only ever
processed once. If Ollama isn't running, everything silently falls back to the
regex extractor — the tool still works.
"""
import hashlib
import json
import os
import re
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
OLLAMA_URL = os.getenv("OLLAMA_URL", "http://localhost:11434").rstrip("/")
MODEL = os.getenv("OLLAMA_MODEL", "llama3.2")
CACHE_FILE = os.path.join(HERE, "llm_cache.json")

PROMPT = """You extract structured data from a message posted in a college \
placement/jobs Telegram group. Return ONLY valid JSON (no markdown) with EXACTLY \
these keys:
- company: string or null (the hiring company/organization)
- role: string or null (job/intern title)
- type: one of "Job", "Internship", "Hackathon", "Other"
- eligibility: string or null (batch years, branches, backlog/CGPA rules, in short)
- min_cgpa: number or null (minimum CGPA required, e.g. 6.5)
- pay: string or null (CTC / stipend / prize exactly as written, e.g. "12 LPA")
- skills: array of strings (tech/skills required; [] if none stated)
- location: string or null (city/cities or "Remote")
- mode: "Remote", "Hybrid", "On-site", or null
- deadline: string "YYYY-MM-DD" or null (the LAST DATE TO APPLY only, not event dates)
- summary: one concise sentence (max 22 words) describing the opportunity

Rules: Use ONLY facts present in the message. Never invent. Missing field -> null \
(or [] for skills). Output JSON only.

MESSAGE:
\"\"\"
{msg}
\"\"\"
"""


def available():
    try:
        with urllib.request.urlopen(OLLAMA_URL + "/api/tags", timeout=3) as r:
            return r.status == 200
    except Exception:
        return False


def model_present():
    try:
        with urllib.request.urlopen(OLLAMA_URL + "/api/tags", timeout=3) as r:
            tags = json.loads(r.read())
        names = [m.get("name", "") for m in tags.get("models", [])]
        return any(n == MODEL or n.startswith(MODEL + ":") for n in names)
    except Exception:
        return False


def _load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)


def _hash(text):
    return hashlib.sha1((MODEL + "|" + (text or "")).encode("utf-8")).hexdigest()


def _call(text):
    payload = {
        "model": MODEL,
        "prompt": PROMPT.replace("{msg}", (text or "")[:4000]),
        "stream": False,
        "format": "json",
        "options": {"temperature": 0, "num_ctx": 4096},
    }
    req = urllib.request.Request(
        OLLAMA_URL + "/api/generate",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=180) as r:
        resp = json.loads(r.read())
    raw = resp.get("response", "").strip()
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r"\{.*\}", raw, re.S)
        return json.loads(m.group(0)) if m else None


_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def _apply(offer, data):
    """Overlay LLM fields onto an offer, but ONLY when they're grounded in the
    source text (small models hallucinate). Regex stays authoritative for any
    value it already found; the LLM fills gaps + owns the one-line summary."""
    if not isinstance(data, dict):
        return
    import placement_monitor as pm

    text = offer.get("Message") or ""
    low = text.lower()
    ntext = _norm(text)

    # Company: keep if it's a known brand or literally appears in the text.
    comp = (data.get("company") or "").strip()
    if comp and (not offer.get("Company") or pm.is_junk_company(offer.get("Company", ""))):
        if pm.canon_company(comp) or _norm(comp) in ntext:
            offer["Company"] = pm.canon_company(comp) or comp[:60]

    # Role: at least one 4+ letter word of it must appear in the post.
    role = (data.get("role") or "").strip()
    if role and not offer.get("Role / Position"):
        if any(w in low for w in re.findall(r"[a-z]{4,}", role.lower())):
            offer["Role / Position"] = role[:60]

    # Pay: needs a pay unit in the text AND the leading digits grounded.
    pay = (data.get("pay") or "").strip()
    if pay and not offer.get("CTC / Stipend"):
        digits = re.sub(r"[^0-9]", "", pay)
        has_unit = re.search(r"lpa|lakh|stipend|/month|per month|ctc|package|salary|₹|\$|\bk\b", low)
        if has_unit and (not digits or digits[:2] in ntext):
            offer["CTC / Stipend"] = pay[:60]

    # Eligibility: must share a concrete token (year/branch/backlog/cgpa) with text.
    elig = (data.get("eligibility") or "").strip()
    if elig and not offer.get("Eligibility / Batch"):
        toks = re.findall(r"20\d\d|cse|ise|ece|eee|mech|civil|b\.?e|b\.?tech|"
                          r"m\.?tech|mca|mba|backlog|cgpa|branch", elig.lower())
        if any(_norm(t) in ntext for t in toks):
            offer["Eligibility / Batch"] = elig[:140]

    # Location: the first named place must appear in the text.
    loc = (data.get("location") or "").strip()
    if loc and not offer.get("Location"):
        first = re.split(r"[,/]", loc)[0].strip()
        if _norm(first) in ntext:
            offer["Location"] = loc[:50]

    # Mode: only if a work-mode word is actually in the text.
    mode = (data.get("mode") or "").strip()
    if mode and not offer.get("mode"):
        if re.search(r"remote|hybrid|on[\s-]?site|onsite|wfh|work from home|in[\s-]?office", low):
            offer["mode"] = mode

    # Skills: keep ONLY real technical skills from the known vocabulary, grounded
    # in the text. This drops soft-skill sentences ("strong analytical skills").
    base = pm.extract_skills(text)                      # clean regex tech skills
    merged = list(base)
    if isinstance(data.get("skills"), list):
        for s in data["skills"]:
            canon = pm._SKILL_CANON.get(str(s).strip().lower())
            if canon and canon not in merged and _norm(canon) in ntext:
                merged.append(canon)
    offer["skills"] = merged[:12]

    # Deadline: only if the post actually contains a date and that year.
    dl = data.get("deadline")
    if (isinstance(dl, str) and _DATE_RE.match(dl) and not offer.get("deadline")
            and pm.DATE_TOKEN.search(text) and dl[:4] in text):
        offer["deadline"] = dl

    # Min CGPA: only if 'cgpa'/'gpa' is mentioned.
    cg = data.get("min_cgpa")
    if (offer.get("cgpa") is None and isinstance(cg, (int, float))
            and 4 <= cg <= 10 and re.search(r"cgpa|gpa", low)):
        offer["cgpa"] = float(cg)

    if data.get("summary"):
        offer["summary_llm"] = str(data["summary"]).strip()[:160]


def enrich_offers(offers, log=None):
    """Enrich a list of offer dicts in place. Returns (enriched, new_calls)."""
    log = log or (lambda *a: None)
    if not available():
        return offers, 0
    cache = _load_cache()
    calls = 0
    for i, o in enumerate(offers):
        text = o.get("Message") or ""
        h = _hash(text)
        if h in cache:
            data = cache[h]
        else:
            try:
                data = _call(text)
            except Exception as e:  # noqa
                data = None
                log(f"  llm error on msg {o.get('id')}: {e}")
            cache[h] = data
            calls += 1
            if calls % 10 == 0:
                _save_cache(cache)
                log(f"  enriched {calls} new messages…")
        _apply(o, data)
    _save_cache(cache)
    return offers, calls
