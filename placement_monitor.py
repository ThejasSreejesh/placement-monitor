"""
Placement Offer Monitor for Telegram
====================================

Logs into your Telegram account (via Telethon / MTProto so it can read the
group's *entire* history), scans a group for placement / job offers, extracts
the useful fields, and writes a clean, organized Excel workbook.

Usage
-----
  python placement_monitor.py --list          # list your groups (find the right one)
  python placement_monitor.py --once          # scan once, update the Excel file
  python placement_monitor.py --once --full   # rescan the ENTIRE history from scratch
  python placement_monitor.py --schedule      # keep running, scan every N minutes

First run will ask for the code Telegram sends you (and 2FA password if set).
That login is cached in a .session file so later runs are non-interactive.
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, timezone

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

from dotenv import load_dotenv
from dateutil import parser as dateparser
from telethon.sync import TelegramClient
from telethon.tl.types import Channel, Chat

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
HERE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(HERE, ".env"))

API_ID = os.getenv("API_ID", "").strip()
API_HASH = os.getenv("API_HASH", "").strip()
PHONE = os.getenv("PHONE", "").strip()
GROUP = os.getenv("GROUP", "").strip()
INTERVAL = int(os.getenv("SCAN_INTERVAL_MINUTES", "60"))
OUTPUT_XLSX = os.path.join(HERE, os.getenv("OUTPUT_XLSX", "placement_offers.xlsx"))

SESSION = os.path.join(HERE, "placement_monitor.session")
STATE_FILE = os.path.join(HERE, "state.json")

COLUMNS = [
    "Date", "Company", "Role / Position", "CTC / Stipend", "Location",
    "Eligibility / Batch", "Apply Link", "Posted By", "Message",
]

# --------------------------------------------------------------------------- #
# Offer detection + field extraction (heuristics)
# --------------------------------------------------------------------------- #
# Strong signals that a message is an actual hiring / opportunity post.
STRONG_OFFER = re.compile(
    r"\b(campus hiring|off[\s-]?campus|on[\s-]?campus|placement drive|"
    r"recruitment drive|we are hiring|is hiring|are hiring|hiring for|"
    r"job opening|openings?|walk[\s-]?in|apply (?:now|here|link|before|by|through)|"
    r"internship (?:opportunit|program|drive)|job opportunit|"
    r"hackathon|referral|full[\s-]?time (?:role|opportunit|position))\b",
    re.IGNORECASE,
)

# Messages that merely reference placements but are NOT offers themselves.
NEGATIVE_OFFER = re.compile(
    r"\b(reminder|registration list|list of (?:selected|registered)|"
    r"selected students|shortlist(?:ed)?|results? (?:are|announced|out)|"
    r"nominated|pre[\s-]?placement talk (?:is )?(?:starting|scheduled|today)|"
    r"attend the (?:session|talk)|venue|hall ticket|deadline (?:extended|today)|"
    r"kind attn|kind attention)\b",
    re.IGNORECASE,
)

URL_RE = re.compile(r"https?://[^\s)>\]]+", re.IGNORECASE)
EMOJI_RE = re.compile("[\U0001F000-\U0001FAFF\U00002600-\U000027BF\U0001F1E6-\U0001F1FF←-⇿⬀-⯿]")

# CTC like "12 LPA", "8-12 LPA", "₹8,00,000", "CTC 22 LPA", "40k / month stipend"
CTC_RE = re.compile(
    r"(₹\s?[\d,]+(?:\.\d+)?(?:\s*(?:lpa|lakhs?|/month|per month))?|"
    r"\$\s?[\d,]+(?:\.\d+)?|"
    r"\d+(?:\.\d+)?\s*(?:-\s*\d+(?:\.\d+)?\s*)?(?:lpa|lakhs?|cr)"
    r"(?:\s*(?:per\s*(?:annum|year)|p\.?a\.?))?|"
    r"\d{2,3}\s?k\s*(?:/\s*month|per month|stipend))",
    re.IGNORECASE,
)

# Common role titles found in free-form posts.
ROLE_RE = re.compile(
    r"\b(software (?:development )?engineer(?:ing)?(?:[\s-]?[I1-3]+)?|sde[\s-]?[1-3]?|"
    r"software developer|full[\s-]?stack developer|backend developer|frontend developer|"
    r"data (?:analyst|scientist|engineer)|business analyst|system(?:s)? engineer|"
    r"ml engineer|devops engineer|qa engineer|test engineer|"
    r"research intern(?:ship)?|summer intern(?:ship)?|graduate (?:engineer )?trainee|"
    r"associate (?:software )?engineer|technical (?:analyst|consultant)|"
    r"analyst|developer|consultant|trainee|intern(?:ship)?)\b",
    re.IGNORECASE,
)

# Company name patterns for label-less posts.
COMPANY_PATTERNS = [
    re.compile(r"campus hiring[:\s]+([A-Za-z][A-Za-z0-9&.\-' ]{1,40})", re.I),
    re.compile(r"hiring (?:for )?(?:at )?([A-Z][A-Za-z0-9&.\-' ]{1,40})\b", re.I),
    re.compile(r"([A-Z][A-Za-z0-9&.\-' ]{1,40}?) is (?:hiring|hosting|conducting|offering|inviting)", re.I),
    re.compile(r"greetings from (?:team )?([A-Z][A-Za-z0-9&.\-' ]{1,40})", re.I),
    re.compile(r"from team ([A-Z][A-Za-z0-9&.\-' ]{1,40})", re.I),
    re.compile(r"opportunit(?:y|ies) (?:at|with) ([A-Z][A-Za-z0-9&.\-' ]{1,40})", re.I),
    re.compile(r"([A-Z][A-Za-z0-9.&' ]{1,40}?)\s*[-–—]\s*campus drive", re.I),
    re.compile(r"([A-Z][A-Za-z0-9.&' ]{1,40}?)\s+(?:campus drive|off[- ]campus|recruitment drive)", re.I),
    re.compile(r"career (?:with|at) ([A-Z][A-Za-z0-9.&' ]{1,40})", re.I),
    re.compile(r"welcome to ([A-Z][A-Za-z0-9.&' ]{1,40})", re.I),
    re.compile(r"([A-Z][A-Za-z0-9.&' ]{1,40}?)\s+hiring\b", re.I),
    re.compile(r"([A-Z][A-Za-z0-9.&' ]{1,40}?)\s+internship\s+(?:opportunit|program|recruitment|drive|\+)", re.I),
    re.compile(r"subject\s*[:\-]\s*([A-Z][A-Za-z0-9.&' ]{1,40}?)\s+(?:internship|webinar|hiring|off[\s-]?campus|opportunit|recruitment|drive)", re.I),
    re.compile(r"opportunity (?:to join|at) ([A-Z][A-Za-z0-9.&' ]{1,40})", re.I),
]

# Hosts where the domain is an aggregator/ATS, not the employer.
AGG_HOSTS = {
    "unstop.com", "techgig.com", "forms.gle", "docs.google.com", "drive.google.com",
    "google.com", "bit.ly", "tinyurl.com", "linkedin.com", "lnkd.in", "t.co",
    "eej.at", "notion.so", "typeform.com", "airtable.com", "tally.so", "gle",
    "instagram.com", "facebook.com", "youtu.be", "youtube.com", "wa.me",
    "chat.whatsapp.com", "whatsapp.com", "t.me", "telegram.me", "naukri.com",
    "indeed.com", "monster.com", "hackerearth.com", "dare2compete.com",
    "cutshort.io", "hirist.com", "apna.co", "internshala.com", "zoho.com",
}
_URL_SUFFIX = {"com", "in", "co", "org", "net", "io", "ai", "me", "us", "uk",
               "info", "biz", "gov", "edu", "ac", "app", "dev", "tech"}


def company_from_link(url):
    """Best-effort employer name from an apply URL's domain/path."""
    m = re.match(r"https?://([^/]+)(/[^?#]*)?", url or "", re.I)
    if not m:
        return ""
    host = m.group(1).lower().split(":")[0]
    if host.startswith("www."):
        host = host[4:]
    path = m.group(2) or ""
    # Workday: <company>.wdN.myworkdayjobs.com
    if host.endswith("myworkdayjobs.com"):
        return canon_company(host.split(".")[0]) or host.split(".")[0].title()
    # Company in the path for these ATSes.
    if any(h in host for h in ("greenhouse.io", "lever.co", "smartrecruiters.com")):
        seg = [s for s in path.split("/") if s]
        return (canon_company(seg[0]) or seg[0].title()) if seg else ""
    if host in AGG_HOSTS:
        return ""
    labels = [l for l in host.split(".") if l]
    while len(labels) > 1 and labels[-1] in _URL_SUFFIX:
        labels.pop()
    if not labels:
        return ""
    main = labels[-1]
    main = re.sub(r"(group|careers|jobs|india|global|inc|corp|technolog\w*)$", "", main)
    if len(main) < 2:
        main = labels[-1]
    return canon_company(main) or main[:1].upper() + main[1:]

# Prefixes to strip off the front of a candidate line.
LINE_PREFIX_RE = re.compile(
    r"^\s*(?:subject\s*[:\-]\s*|re\s*[:\-]\s*|fwd\s*[:\-]\s*|\d+[.)]\s*|[-*•]\s*)+",
    re.IGNORECASE,
)


# A curated list of companies that recruit in Indian campus groups. If any
# appears in the message, it becomes the company name (big accuracy win).
KNOWN_COMPANIES = [
    "BNY Mellon", "BNY", "Infosys", "TCS", "Wipro", "Accenture", "Cognizant",
    "Capgemini", "Deloitte", "Amazon", "Microsoft", "Google", "Oracle", "SAP Labs",
    "SAP", "IBM", "Bosch", "Siemens", "Cisco", "Adobe", "Intel", "Qualcomm",
    "Nvidia", "NVIDIA", "PayPal", "Walmart", "Goldman Sachs", "JPMorgan",
    "JP Morgan", "Morgan Stanley", "HSBC", "Standard Chartered", "ZS Associates",
    "ZS", "Mercedes-Benz", "Mercedes", "Continental", "Harman", "Philips",
    "Samsung", "Nokia", "Ericsson", "Juniper", "VMware", "Dell", "ServiceNow",
    "Salesforce", "Uber", "Swiggy", "Zomato", "Flipkart", "Myntra", "Paytm",
    "PhonePe", "Razorpay", "Freshworks", "Zoho", "LTIMindtree", "Mindtree",
    "Tech Mahindra", "Mphasis", "Hexaware", "Persistent", "Nagarro",
    "Thoughtworks", "Publicis Sapient", "EPAM", "GlobalLogic", "HCLTech", "HCL",
    "Genpact", "Fidelity", "Optum", "UnitedHealth", "Micron", "Texas Instruments",
    "Analog Devices", "NXP", "MediaTek", "AMD", "Synopsys", "Cadence",
    "Wells Fargo", "Citi", "Barclays", "Societe Generale", "Nomura", "Sabre",
    "Epsilon", "Tredence", "Fractal", "Sigmoid", "LatentView", "Netradyne",
    "KLA", "Applied Materials", "Honeywell", "Schneider", "Groww", "Zerodha",
    "CRED", "Meesho", "ShareChat", "Ola", "Juspay", "Sprinklr", "Postman",
    "Chargebee", "BrowserStack", "Netskope", "Rubrik", "Nutanix", "Palo Alto",
    "Arista", "Airbus", "Boeing", "Caterpillar", "Cummins", "Emertxe",
    "TenZorX", "Varroc", "USAII", "Cargill", "Verzeo", "Zscaler", "Atlassian",
    "SanDisk", "Western Digital", "Seagate", "Broadcom", "Marvell", "Infineon",
    "Apple", "Sony", "Keyence", "eInfochips", "Odoo", "FOSSEE", "Baker Hughes",
    "Dr. Reddy's", "Dr Reddys", "Span.io", "Arrow", "Cummins", "Tiger Analytics",
    "Sprinklr", "Zeta", "Hitachi", "Toshiba", "Denso", "GE", "L&T", "Larsen",
    "Stripe", "Snowflake", "Rolls-Royce", "Rolls Royce", "Merck", "CRMIT",
    "SBI", "AXA", "Databricks", "MongoDB", "Confluent", "HashiCorp", "Twilio",
    "Cloudflare", "Datadog", "GitHub", "GitLab", "Atlassian", "Notion", "Figma",
    "Siemens Healthineers", "Philips", "Medtronic", "Shell", "ExxonMobil",
    "Schlumberger", "SLB", "John Deere", "Deere", "Mastercard", "Visa",
    "American Express", "Amex", "BlackRock", "Deutsche Bank", "UBS", "Credit Suisse",
]
COMPANY_LOOKUP = sorted(KNOWN_COMPANIES, key=len, reverse=True)
KNOWN_RE = re.compile(
    r"(?<![A-Za-z])(" + "|".join(re.escape(c) for c in COMPANY_LOOKUP) + r")(?![A-Za-z])",
    re.IGNORECASE,
)
_CANON = {c.lower(): c for c in KNOWN_COMPANIES}

# Skills / tech stack vocabulary.
SKILLS = [
    "Python", "Java", "C++", "C#", "JavaScript", "TypeScript", "Golang",
    "Rust", "Kotlin", "Swift", "PHP", "Ruby", "Scala",
    "React", "Angular", "Vue", "Next.js", "Node.js", "Express", "Django",
    "Flask", "FastAPI", "Spring Boot", "Spring", ".NET", "Rails",
    "SQL", "MySQL", "PostgreSQL", "MongoDB", "Redis", "Cassandra",
    "AWS", "Azure", "GCP", "Docker", "Kubernetes", "Terraform", "Jenkins",
    "Git", "REST", "GraphQL", "Microservices", "Kafka", "Spark", "Hadoop",
    "HTML", "CSS", "Tailwind", "DSA", "Data Structures", "Algorithms",
    "Machine Learning", "Deep Learning", "NLP", "Computer Vision",
    "TensorFlow", "PyTorch", "Keras", "Pandas", "NumPy", "Scikit-learn",
    "OpenCV", "LLM", "Generative AI", "Tableau", "Power BI", "Excel",
    "Linux", "Selenium", "Android", "iOS", "Flutter",
    # CS core + hardware/EC skills common in campus posts
    "DBMS", "Operating Systems", "Computer Networks", "System Design", "OOP",
    "OOPS", "REST API", "REST APIs", "CI/CD", "Networking", "Cybersecurity",
    "Embedded C", "Embedded Systems", "Verilog", "VHDL", "VLSI", "RTL",
    "MATLAB", "Firebase", "Unity", "Bootstrap", "jQuery", "Sass", "Figma",
    "Power Automate", "Data Analysis", "Data Structures & Algorithms",
    "Microcontrollers", "IoT", "Arduino", "Raspberry Pi", "PLC", "AutoCAD",
    "SolidWorks", "CATIA", "Ansys",
]
SKILL_RE = re.compile(
    r"(?<![A-Za-z0-9+#.])(" +
    "|".join(re.escape(s) for s in sorted(SKILLS, key=len, reverse=True)) +
    r")(?![A-Za-z0-9])",
    re.IGNORECASE,
)
_SKILL_CANON = {s.lower(): s for s in SKILLS}

AIML_RE = re.compile(
    r"\b(a\.?i\.?|artificial intelligence|machine learning|\bml\b|deep learning|"
    r"data scien|data analyst|data analytics|\bnlp\b|natural language|"
    r"computer vision|tensorflow|pytorch|keras|opencv|\bllm\b|generative ai|"
    r"gen ai|neural network|ai/ml|ai engineer|ml engineer)\b",
    re.IGNORECASE,
)
BACKEND_RE = re.compile(
    r"\b(back[\s-]?end|server[\s-]?side|\bapi\b|micro[\s-]?service|node\.?js|"
    r"django|flask|fastapi|spring|\bsql\b|database|\bsde\b|java developer|"
    r"golang|rest api|software engineer|software developer|backend developer|"
    r"\.net|full[\s-]?stack)\b",
    re.IGNORECASE,
)


def canon_company(text):
    """Return the earliest-appearing known company in the text, or ''."""
    m = KNOWN_RE.search(text or "")
    if m:
        return _CANON.get(m.group(1).lower(), m.group(1))
    return ""


def extract_skills(text):
    seen, out = set(), []
    for m in SKILL_RE.finditer(text or ""):
        canon = _SKILL_CANON.get(m.group(1).lower(), m.group(1))
        if canon.lower() not in seen:
            seen.add(canon.lower())
            out.append(canon)
    return out[:12]


def classify(text):
    cats = []
    if AIML_RE.search(text or ""):
        cats.append("AIML")
    if BACKEND_RE.search(text or ""):
        cats.append("BACKEND")
    return cats


INTERN_RE = re.compile(
    r"\b(intern(?:ship)?s?|summer intern|winter intern|trainee|stipend|"
    r"industrial training|apprentice)\b", re.IGNORECASE)
JOB_RE = re.compile(
    r"\b(full[\s-]?time|\bfte\b|permanent role|permanent position|"
    r"\blpa\b|per annum|\bp\.?a\.?\b|\bctc\b|lakhs? per|annual (?:ctc|package)|"
    r"software engineer|associate engineer|developer|analyst)\b", re.IGNORECASE)


def detect_types(text):
    """Return a list among ['JOB','INTERN']. Defaults to JOB when ambiguous."""
    t = text or ""
    intern = bool(INTERN_RE.search(t))
    job = bool(JOB_RE.search(t))
    types = []
    if job:
        types.append("JOB")
    if intern:
        types.append("INTERN")
    if not types:              # campus drives are usually full-time roles
        types = ["JOB"]
    return types


JUNK_COMPANY_RE = re.compile(
    r"^(subject|dear|greetings|invitation|win\b|unlock|your|our|opportunit|list|"
    r"we\b|hi\b|hello|note|update|important|attention|reminder|remainder|final|"
    r"registration|selected|congrat|the\b|this\b|all\b|as\b|kindly|please|link|apply|"
    r"eligible|interested|last date|deadline|requirement|skills?\b|associate|"
    r"sponsored|kickstart|career|role|position|about|team\b|here\b|good\b|"
    r"placement|internship|job\b|hiring|recruit|drive\b|exciting|great\b)",
    re.IGNORECASE,
)


ROLE_TAIL_RE = re.compile(
    r"\b(interns?|developers?|engineers?|trainees?|analysts?|associates?|"
    r"executives?|managers?|programmers?)$", re.IGNORECASE)
JUNK_EXTRA_RE = re.compile(
    r"^(one\b|packages?\b|cab\b|rtl\b|up\b|full stack|front[\s-]?end|"
    r"back[\s-]?end|not\b|no\b|yes\b|available|provided)", re.IGNORECASE)


def is_junk_company(name):
    if not name or len(name) < 3:
        return True
    name = name.strip()
    if canon_company(name):          # a real, known brand — never junk
        return False
    if ":" in name:
        return True
    if JUNK_COMPANY_RE.match(name) or JUNK_EXTRA_RE.match(name):
        return True
    if ROLE_TAIL_RE.search(name):    # "…Developer", "…Interns" = a role, not a company
        return True
    if "student" in name.lower() or "batch" in name.lower():
        return True
    if not re.search(r"[A-Za-z]", name):
        return True
    return False


def _field(text, labels):
    """Grab 'Label: value' style fields (value is rest of that line)."""
    for label in labels:
        m = re.search(
            rf"(?im)^\s*{label}\s*[:\-–]\s*(.+)$", text
        )
        if m:
            return m.group(1).strip(" *_`")
    return ""


def _clean(s):
    s = EMOJI_RE.sub("", s)
    s = re.sub(r"\s+", " ", s).strip(" *_`#>-•:|\t")
    return s


def _norm_company(s):
    """Drop a leading preposition/article ('For Blinkit,' -> 'Blinkit') and
    trailing punctuation so we don't invent bogus companies."""
    s = re.sub(r"^(for|at|with|from|to|by|the|a|an)\s+", "", s or "", flags=re.I)
    return s.strip(" ,.:;-–—")


def _guess_company(text):
    # 1) Labeled field wins.
    v = _field(text, [r"company", r"company name", r"organi[sz]ation", r"firm"])
    if v and not is_junk_company(_clean(v)):
        known = canon_company(v)
        return known or _clean(v)[:60]
    # 2) Known company anywhere in the text.
    known = canon_company(text)
    if known:
        return known
    for pat in COMPANY_PATTERNS:
        m = pat.search(text)
        if m:
            cand = _clean(m.group(1))
            # Trim trailing filler words.
            cand = re.split(r"\b(is|are|for|to|and|the|a national|hackathon)\b",
                            cand, 1, flags=re.I)[0].strip()
            cand = _norm_company(cand)
            if 2 <= len(cand) <= 60 and not is_junk_company(cand):
                return canon_company(cand) or cand
    # Fallback: scan the first few lines for a clean, non-junk name.
    for line in text.splitlines()[:6]:
        s = _norm_company(_clean(LINE_PREFIX_RE.sub("", line)))
        if not s or URL_RE.match(s):
            continue
        known = canon_company(s)
        if known:
            return known
        # A short Title-Case-ish phrase that isn't junk.
        words = s.split()
        if 1 <= len(words) <= 6 and s[0].isupper() and not is_junk_company(s):
            return s[:60]
    return ""


def _guess_role(text):
    v = _field(text, [r"role", r"position", r"profile", r"designation", r"job title", r"title"])
    if v:
        return _clean(v)[:50]
    m = ROLE_RE.search(text)
    return m.group(1).strip().title() if m else ""


CITIES = [
    "Bengaluru", "Bangalore", "Hyderabad", "Chennai", "Pune", "Mumbai",
    "Delhi", "New Delhi", "Noida", "Gurgaon", "Gurugram", "Kolkata",
    "Ahmedabad", "Coimbatore", "Kochi", "Trivandrum", "Thiruvananthapuram",
    "Mysuru", "Mysore", "Chandigarh", "Jaipur", "Indore", "Nagpur",
    "Visakhapatnam", "Vizag", "Bhubaneswar", "Mangaluru", "Mangalore",
    "Pan India", "Remote", "Bhopal", "Vadodara", "Lucknow", "Kanpur",
]
CITY_RE = re.compile(r"\b(" + "|".join(sorted(CITIES, key=len, reverse=True)) + r")\b", re.I)
_CITY_CANON = {c.lower(): c for c in CITIES}

MODE_RE = re.compile(
    r"\b(work from home|wfh|remote|hybrid|on[\s-]?site|in[\s-]?office|onsite)\b", re.I)

# What kind of opportunity is this (used as a headline when there's no role).
KIND_PATTERNS = [
    (re.compile(r"\bhackathon\b", re.I), "Hackathon"),
    (re.compile(r"\bwebinar\b", re.I), "Webinar"),
    (re.compile(r"\bscholarship\b", re.I), "Scholarship"),
    (re.compile(r"\bfellowship\b", re.I), "Fellowship"),
    (re.compile(r"\b(pre[\s-]?placement talk|ppt)\b", re.I), "Pre-Placement Talk"),
    (re.compile(r"\breferral\b", re.I), "Referral"),
    (re.compile(r"\b(contest|challenge|competition)\b", re.I), "Contest"),
    (re.compile(r"\b(recruitment drive|placement drive|campus drive|drive)\b", re.I), "Recruitment Drive"),
    (re.compile(r"\bintern(ship)?\b", re.I), "Internship"),
    (re.compile(r"\b(hiring|opening|opportunit|recruit)\b", re.I), "Job Opening"),
]


def _guess_location(text):
    v = _field(text, [r"location", r"job location", r"place", r"work location", r"base location"])
    if v:
        return _clean(v)[:40]
    found = []
    for m in CITY_RE.finditer(text or ""):
        c = _CITY_CANON.get(m.group(1).lower(), m.group(1))
        if c not in found:
            found.append(c)
    return ", ".join(found[:2])


def guess_mode(text):
    m = MODE_RE.search(text or "")
    if not m:
        return ""
    v = m.group(1).lower()
    if v in ("wfh", "work from home", "remote"):
        return "Remote"
    if v == "hybrid":
        return "Hybrid"
    return "On-site"


def guess_kind(text):
    for pat, label in KIND_PATTERNS:
        if pat.search(text or ""):
            return label
    return ""


_GREETING_RE = re.compile(
    r"^(dear|greetings|hi|hello|hey|namaste|note|important|subject|re|fwd|"
    r"we are|we're|we would|team|good (?:morning|afternoon|evening))\b", re.I)


def smart_summary(role, kind, ctc, location, mode, type_label):
    """Compose a clean one-liner from structured facts (no AI needed).
    e.g. 'Software Engineer Intern · 6 LPA · Bengaluru · Remote'."""
    headline = role or kind or (type_label if type_label else "Opportunity")
    parts = [headline]
    if ctc:
        parts.append(ctc)
    if location:
        parts.append(location)
    if mode and (not location or mode.lower() not in location.lower()):
        parts.append(mode)
    out = " · ".join(p for p in parts if p)
    return out[:120]


BATCH_RE = re.compile(r"\b(20\d\d)(?:\s*[/&,and ]+\s*(20\d\d))*\s*batch\b", re.I)
BRANCH_RE = re.compile(
    r"\b(B\.?E|B\.?Tech|M\.?Tech|MCA|MBA|MSc|BSc|"
    r"CSE|ISE|ECE|EEE|AIML|CSD|IT|Mechanical|Civil|Information Science|"
    r"Computer Science)\b", re.I)


def _guess_eligibility(text):
    v = _field(text, [
        r"eligibility", r"eligible(?:\s*branches)?", r"qualification",
        r"who can apply", r"passout", r"pass\s*out", r"eligible batch",
    ])
    if v:
        return _clean(v)[:120]
    # Build from detected batch years + branches.
    parts = []
    mb = BATCH_RE.search(text)
    if mb:
        parts.append(_clean(mb.group(0)))
    branches = []
    for m in BRANCH_RE.finditer(text):
        b = m.group(1).upper().replace(".", "")
        if b not in branches:
            branches.append(b)
    if branches:
        parts.append("/".join(branches[:6]))
    return ", ".join(parts)


CGPA_RE = re.compile(
    r"(\d{1,2}(?:\.\d{1,2})?)\s*\+?\s*(?:cgpa|c\.g\.p\.a|gpa)\b"
    r"|(?:cgpa|c\.g\.p\.a|gpa)\s*(?:of|:|-|–|=|>=|>|above|min(?:imum)?|and above|"
    r"or above|greater than|atleast|at least)?\s*(\d{1,2}(?:\.\d{1,2})?)",
    re.IGNORECASE,
)


def extract_cgpa(text):
    """Minimum CGPA a candidate needs, as a float, or None if not stated.
    Restricted to a plausible cutoff band (4.0–9.5) so scale mentions like
    'out of 10' or percentages don't get misread as a cutoff."""
    if not text:
        return None
    cands = []
    for m in CGPA_RE.finditer(text):
        num = m.group(1) or m.group(2)
        try:
            v = float(num)
        except (TypeError, ValueError):
            continue
        if 4.0 <= v <= 9.5:
            cands.append(v)
    return min(cands) if cands else None


def _guess_ctc(text):
    v = _field(text, [r"ctc", r"package", r"salary", r"stipend", r"compensation"])
    if v:
        return v
    m = CTC_RE.search(text)
    return m.group(1).strip() if m else ""


def is_offer(text):
    if not text or len(text.strip()) < 25:
        return False
    strong = bool(STRONG_OFFER.search(text))
    negative = bool(NEGATIVE_OFFER.search(text))
    has_money = bool(CTC_RE.search(text))
    has_link = bool(URL_RE.search(text))
    if not strong:
        return False
    # A reminder/result post only counts if it still carries an apply link + pay.
    if negative and not (has_money and has_link):
        return False
    return True


# --- Application deadline ("last date to apply") ---------------------------- #
DEADLINE_CUE = re.compile(
    r"(?:last date(?:\s*(?:to|for)\s*(?:apply|applying|registration|register))?|"
    r"last day(?:\s*to\s*(?:apply|register))?|"
    r"apply(?:ing)?\s*(?:by|before|on or before|latest by|no later than|last date)|"
    r"register(?:ing|ration)?\s*(?:by|before|on or before|latest by|no later than)|"
    r"on or before|latest by|no later than|"
    r"deadline|due date|"
    r"registration\s*(?:closes?|ends?|deadline|last date)|"
    r"application(?:s)?\s*(?:closing date|closes?|close by|deadline|end(?:s|ing)?)|"
    r"closing date|closes?\s*on|closes?\s*by|window closes|"
    r"submit(?:\s*by|\s*before))\s*[:\-–,]?\s*",
    re.IGNORECASE,
)
_MONTHS = (r"jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|"
           r"january|february|march|april|june|july|august|september|"
           r"october|november|december")
DATE_TOKEN = re.compile(
    r"(\d{1,2}[/\-.]\d{1,2}(?:[/\-.]\d{2,4})?|"                       # 05/09/2026
    r"\d{1,2}(?:st|nd|rd|th)?\s+(?:" + _MONTHS + r")[a-z]*"           # 5th September
    r"(?:,?\s+\d{2,4})?|"
    r"(?:" + _MONTHS + r")[a-z]*\s+\d{1,2}(?:st|nd|rd|th)?"           # September 5
    r"(?:,?\s+\d{2,4})?)",
    re.IGNORECASE,
)


_TIME_TOKEN = re.compile(r"^\d{1,2}[.:]\d{2}$")  # e.g. 9.00 / 09:30 -> a time, not a date


def extract_deadline(text):
    """Return the 'last date to apply' as YYYY-MM-DD, or '' if none.
    Scans a window after each cue, tries every date token, and keeps the
    latest plausible one (times like '9.00 AM' are skipped)."""
    if not text:
        return ""
    yr = datetime.now().year
    best = ""
    for cue in DEADLINE_CUE.finditer(text):
        window = text[cue.end():cue.end() + 42]
        for dm in DATE_TOKEN.finditer(window):
            tok = dm.group(1)
            if _TIME_TOKEN.match(tok):            # skip "9.00" / "10:30" times
                continue
            try:
                dt = dateparser.parse(tok, dayfirst=True, fuzzy=True,
                                      default=datetime(yr, 1, 1))
            except (ValueError, OverflowError):
                continue
            # Plausibility: a real apply-by date sits within ~1yr back to 2yr ahead.
            if not (yr - 1 <= dt.year <= yr + 2):
                continue
            iso = dt.strftime("%Y-%m-%d")
            if iso > best:
                best = iso
    return best


def _guess_requirements(text, skills):
    v = _field(text, [
        r"requirements?", r"required skills?", r"skills? required",
        r"must[\s-]?have", r"skills?", r"tech stack", r"prerequisites?",
    ])
    if v:
        return _clean(v)[:160]
    return ", ".join(skills)


def extract_offer(msg, sender_name):
    text = msg.message or ""
    links = URL_RE.findall(text)
    skills = extract_skills(text)
    company = _guess_company(text)
    if not company and links:                 # last resort: the apply-link domain
        company = company_from_link(links[0])
    return {
        "id": msg.id,
        "date_iso": msg.date.astimezone().strftime("%Y-%m-%d"),
        "Date": msg.date.astimezone().strftime("%Y-%m-%d %H:%M"),
        "Company": company,
        "Role / Position": _guess_role(text),
        "CTC / Stipend": _guess_ctc(text),
        "Location": _guess_location(text),
        "mode": guess_mode(text),
        "kind": guess_kind(text),
        "Eligibility / Batch": _guess_eligibility(text),
        "Requirements / Skills": _guess_requirements(text, skills),
        "deadline": extract_deadline(text),
        "cgpa": extract_cgpa(text),
        "skills": skills,
        "categories": classify(text),
        "types": detect_types(text),
        "Apply Link": links[0] if links else "",
        "Posted By": sender_name,
        "Message": re.sub(r"\s+", " ", text).strip()[:1600],
    }


# --------------------------------------------------------------------------- #
# State (so scheduled runs are incremental, not re-scanning everything)
# --------------------------------------------------------------------------- #
def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_message_id": 0, "seen_ids": []}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


OFFERS_FILE = os.path.join(HERE, "offers.json")


def load_offers():
    if os.path.exists(OFFERS_FILE):
        with open(OFFERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_offers(offers):
    with open(OFFERS_FILE, "w", encoding="utf-8") as f:
        json.dump(offers, f, ensure_ascii=False, indent=2)


def _norm_key(text):
    """Collapse near-duplicate posts (reminders etc.) to one key."""
    t = EMOJI_RE.sub("", text or "").lower()
    t = re.sub(r"\b(final|urgent|reminder|remainder|high priority|today|"
               r"\d{1,2}[:.]\d{2}\s*(?:am|pm)?)\b", "", t)
    t = re.sub(r"[^a-z0-9]", "", t)
    return t[:120]


def _longest(values):
    vals = [v for v in values if v]
    return max(vals, key=len) if vals else ""


def grouped_offers(offers=None):
    """Group raw per-message offers by company, aggregating every date, the
    union of skills, categories, and keeping one clean record per company."""
    offers = load_offers() if offers is None else offers
    today_iso = datetime.now().strftime("%Y-%m-%d")
    groups = {}
    for o in offers:
        company = (o.get("Company") or "").strip()
        key = re.sub(r"[^a-z0-9]", "", company.lower()) or f"_{o['id']}"
        g = groups.setdefault(key, {
            "Company": company, "posts": [], "skills": [], "categories": set(),
            "dates": set(), "links": [], "deadlines": [], "types": set(),
        })
        if o.get("deadline"):
            g["deadlines"].append(o["deadline"])
        g["types"].update(o.get("types", []))
        # Prefer a known/cleaner company label for the group.
        if canon_company(company) and not canon_company(g["Company"]):
            g["Company"] = canon_company(company)
        g["posts"].append(o)
        for s in o.get("skills", []):
            if s not in g["skills"]:
                g["skills"].append(s)
        g["categories"].update(o.get("categories", []))
        if o.get("date_iso"):
            g["dates"].add(o["date_iso"])
        if o.get("Apply Link"):
            g["links"].append(o["Apply Link"])

    out = []
    for gid, g in groups.items():
        posts = g["posts"]
        company = g["Company"]
        link = g["links"][0] if g["links"] else ""
        skills = g["skills"][:12]
        requirements = ", ".join(skills) or _longest(
            [p.get("Requirements / Skills", "") for p in posts])
        eligibility = _longest([p.get("Eligibility / Batch", "") for p in posts])
        role = _longest([p.get("Role / Position", "") for p in posts])
        ctc = next((p.get("CTC / Stipend") for p in posts if p.get("CTC / Stipend")), "")
        cgpas = [p.get("cgpa") for p in posts if p.get("cgpa") is not None]
        cgpa = min(cgpas) if cgpas else None
        message = _longest([p.get("Message", "") for p in posts])
        location = next((p.get("Location") for p in posts if p.get("Location")), "")
        mode = next((p.get("mode") for p in posts if p.get("mode")), "")
        # Derive kind from the representative (shown) message so the tag matches it.
        kind = guess_kind(message)
        dates = sorted(g["dates"], reverse=True)
        cats = sorted(g["categories"])
        has_link = bool(link)
        # Drop noise: junk company with nothing useful attached.
        if is_junk_company(company) and not has_link and not skills:
            continue
        # Deadline = latest "last date to apply" found; status vs today.
        deadline = max(g["deadlines"]) if g["deadlines"] else ""
        days_left = None
        if deadline:
            status = "OPEN" if deadline >= today_iso else "CLOSED"
            try:
                days_left = (datetime.strptime(deadline, "%Y-%m-%d").date()
                             - datetime.now().date()).days
            except ValueError:
                days_left = None
        else:
            status = "UNKNOWN"
        rank = {"OPEN": 2, "UNKNOWN": 1, "CLOSED": 0}[status]
        types = sorted(g["types"]) or ["JOB"]
        type_label = ("Job + Intern" if len(types) > 1
                      else ("Internship" if types[0] == "INTERN" else "Job"))
        llm_sum = _longest([p.get("summary_llm", "") for p in posts])
        summary = llm_sum or smart_summary(role, kind, ctc, location, mode, type_label)
        out.append({
            "oid": gid,
            "last_id": max((p.get("id", 0) for p in posts), default=0),
            "types": types,
            "type_label": type_label,
            "Company": company or "(unknown)",
            "Role / Position": role,
            "summary": summary,
            "Location": location,
            "mode": mode,
            "kind": kind,
            "Eligibility / Batch": eligibility,
            "Requirements / Skills": requirements,
            "CTC / Stipend": ctc,
            "Apply Link": link,
            "skills": skills,
            "categories": cats,
            "primary": "AIML" if "AIML" in cats else ("BACKEND" if "BACKEND" in cats else "General"),
            "deadline": deadline,
            "status": status,
            "days_left": days_left,
            "cgpa": cgpa,
            "dates": dates,
            "last_date": dates[0] if dates else "",
            "times_posted": len(g["dates"]),
            "has_link": has_link,
            "_rank": rank,
            "Message": message,
        })
    # Sort: OPEN first, then still-useful (with link), then most recent.
    out.sort(key=lambda x: (x["_rank"], x["has_link"], x["last_date"]), reverse=True)
    return out


# --------------------------------------------------------------------------- #
# Telegram
# --------------------------------------------------------------------------- #
def make_client():
    if not API_ID or not API_HASH:
        sys.exit("ERROR: API_ID / API_HASH missing. Copy .env.example to .env and fill it in.")
    return TelegramClient(SESSION, int(API_ID), API_HASH)


def resolve_group(client):
    """Match GROUP against username, id, or title among your dialogs."""
    target = GROUP
    # Numeric id?
    try:
        if target.lstrip("-").isdigit():
            return client.get_entity(int(target))
    except Exception:
        pass
    try:
        return client.get_entity(target)  # handles @username
    except Exception:
        pass
    # Fall back to title match across joined dialogs.
    for dialog in client.iter_dialogs():
        if dialog.name and dialog.name.strip().lower() == target.strip().lower():
            return dialog.entity
    sys.exit(f"ERROR: could not find group '{GROUP}'. Run --list to see options.")


def list_groups(client):
    print("\nYour groups / channels:\n" + "-" * 60)
    for dialog in client.iter_dialogs():
        ent = dialog.entity
        if isinstance(ent, (Channel, Chat)):
            uname = getattr(ent, "username", None)
            handle = f"@{uname}" if uname else "(private)"
            print(f"{dialog.name!r:45} id={ent.id:<15} {handle}")
    print("-" * 60)
    print("Put the @username, the id, or the exact title into GROUP in your .env\n")


def fetch_offers(client, full=False):
    entity = resolve_group(client)
    state = load_state()
    seen = set(state.get("seen_ids", []))
    stored = [] if full else load_offers()
    seen_keys = set(state.get("seen_keys", [])) if not full else set()
    if full:
        seen = set()

    new_offers, max_id = [], (0 if full else state.get("last_message_id", 0))
    min_id = 0 if full else state.get("last_message_id", 0)
    name_cache = {}

    for msg in client.iter_messages(entity, min_id=min_id, reverse=True):
        max_id = max(max_id, msg.id)
        if msg.id in seen or not is_offer(msg.message):
            continue
        key = _norm_key(msg.message)
        if key in seen_keys:
            seen.add(msg.id)
            continue
        seen_keys.add(key)
        # Resolve sender name (cached).
        sender = "—"
        if msg.sender_id is not None:
            if msg.sender_id not in name_cache:
                try:
                    s = msg.get_sender()
                    name_cache[msg.sender_id] = (
                        getattr(s, "title", None)
                        or " ".join(filter(None, [getattr(s, "first_name", ""),
                                                   getattr(s, "last_name", "")])).strip()
                        or f"id{msg.sender_id}"
                    )
                except Exception:
                    name_cache[msg.sender_id] = f"id{msg.sender_id}"
            sender = name_cache[msg.sender_id]
        new_offers.append(extract_offer(msg, sender))
        seen.add(msg.id)

    all_offers = stored + new_offers
    save_offers(all_offers)
    save_state({
        "last_message_id": max_id,
        "seen_ids": sorted(seen),
        "seen_keys": sorted(seen_keys),
    })
    return new_offers


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
FONT = "Calibri"
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10.5)
TITLE_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=15)
SUB_FONT = Font(name=FONT, color="D6E2F2", size=9.5)
BODY_FONT = Font(name=FONT, size=10.5, color="1B2430")
HAIR = Side(style="thin", color="E4E9F0")
BORDER = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)
BAND_FILL = PatternFill("solid", fgColor="F5F8FC")
TITLE_FILL = PatternFill("solid", fgColor="1F3A5F")
CAT_STYLE = {
    "AIML":    ("5B2FB0", "EFE7FB"),
    "BACKEND": ("0A7C4A", "E1F4EB"),
    "General": ("55607A", "EEF1F6"),
}
STATUS_STYLE = {
    "OPEN":    ("0A7C4A", "E1F4EB"),
    "CLOSED":  ("B42318", "FBE7E5"),
    "UNKNOWN": ("55607A", "EEF1F6"),
}

# (Header, offer-key, width, wrap?)
XL_COLUMNS = [
    ("#", "_num", 5, False),
    ("Company", "Company", 24, True),
    ("Category", "primary", 11, False),
    ("Type", "type_label", 12, False),
    ("Role / Position", "Role / Position", 20, True),
    ("Eligibility", "Eligibility / Batch", 22, True),
    ("Min CGPA", "_cgpa_str", 9, False),
    ("Requirements / Skills", "Requirements / Skills", 34, True),
    ("CTC / Stipend", "CTC / Stipend", 14, True),
    ("Last Date to Apply", "_deadline_str", 16, False),
    ("Status", "status", 11, False),
    ("Apply", "Apply Link", 14, False),
    ("Seen", "times_posted", 7, False),
]


def build_excel(path=None, category="ALL"):
    """Rebuild a polished workbook from grouped offers (open first, newest first)."""
    path = path or OUTPUT_XLSX
    rows = grouped_offers()
    if category and category != "ALL":
        rows = [r for r in rows if category in r["categories"]]
    for i, r in enumerate(rows, start=1):
        r["_num"] = i
        r["_deadline_str"] = r["deadline"] or "—"
        r["_cgpa_str"] = f"{r['cgpa']:.1f}" if r.get("cgpa") is not None else "—"

    ncol = len(XL_COLUMNS)
    last_col = get_column_letter(ncol)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Placement Offers"
    ws.sheet_view.showGridLines = False

    # --- Title band (rows 1-2) ---
    ws.merge_cells(f"A1:{last_col}1")
    ws.merge_cells(f"A2:{last_col}2")
    t = ws.cell(row=1, column=1,
                value=f"Placement Offers  ·  {category} view")
    t.font = TITLE_FONT
    t.fill = TITLE_FILL
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    opened = sum(1 for r in rows if r["status"] == "OPEN")
    sub = ws.cell(row=2, column=1, value=(
        f"Generated {datetime.now():%d %b %Y, %H:%M}   |   {len(rows)} offers   |   "
        f"{opened} open now   |   Source: Telegram group {GROUP}"))
    sub.font = SUB_FONT
    sub.fill = TITLE_FILL
    sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 16

    # --- Header row (row 3) ---
    hrow = 3
    for c, (title, _k, _w, _wrap) in enumerate(XL_COLUMNS, start=1):
        cell = ws.cell(row=hrow, column=c, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[hrow].height = 22
    ws.freeze_panes = f"A{hrow + 1}"

    # --- Data rows ---
    for i, offer in enumerate(rows):
        row = hrow + 1 + i
        band = BAND_FILL if i % 2 else None
        for c, (title, key, _w, wrap) in enumerate(XL_COLUMNS, start=1):
            val = offer.get(key, "")
            cell = ws.cell(row=row, column=c)
            cell.font = BODY_FONT
            cell.border = BORDER
            center = key in ("_num", "primary", "status", "_deadline_str",
                             "times_posted", "Apply Link")
            cell.alignment = Alignment(
                horizontal="center" if center else "left",
                vertical="center", wrap_text=wrap, indent=0 if center else 1)
            if band:
                cell.fill = band

            if key == "Apply Link":
                if val:
                    cell.value = "Apply ↗"
                    cell.hyperlink = val
                    cell.font = Font(name=FONT, size=10.5, color="1155CC", underline="single")
                else:
                    cell.value = "—"
            elif key == "primary":
                cell.value = val
                fg, bg = CAT_STYLE.get(val, CAT_STYLE["General"])
                cell.font = Font(name=FONT, bold=True, size=9.5, color=fg)
                cell.fill = PatternFill("solid", fgColor=bg)
            elif key == "status":
                cell.value = val.title()
                fg, bg = STATUS_STYLE.get(val, STATUS_STYLE["UNKNOWN"])
                cell.font = Font(name=FONT, bold=True, size=9.5, color=fg)
                cell.fill = PatternFill("solid", fgColor=bg)
            else:
                cell.value = val
        # Row height scales a little with skills/eligibility length.
        longest = max(len(str(offer.get("Requirements / Skills", ""))),
                      len(str(offer.get("Eligibility / Batch", ""))))
        ws.row_dimensions[row].height = 30 if longest < 34 else 44

    for c, (_t, _k, w, _wr) in enumerate(XL_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.auto_filter.ref = f"A{hrow}:{last_col}{hrow + len(rows)}"
    wb.save(path)
    return len(rows)


# Backwards-compat alias used by run_once/scheduler.
def write_excel(_new_offers):
    return build_excel()


# --------------------------------------------------------------------------- #
# Run modes
# --------------------------------------------------------------------------- #
def run_once(full=False):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with make_client() as client:
            client.start(phone=PHONE)
            offers = fetch_offers(client, full=full)
        # Local LLM enrichment via Ollama (silently skipped if not running).
        enriched = 0
        try:
            import llm_enrich
            if llm_enrich.available():
                all_offers = load_offers()
                _, enriched = llm_enrich.enrich_offers(all_offers, log=print)
                save_offers(all_offers)
        except Exception as e:  # noqa
            print(f"  llm enrichment skipped: {e}")
        n = build_excel()
        tail = f"; {enriched} llm-enriched" if enriched else ""
        line = f"[{stamp}] OK  {len(offers)} new offer msg(s){tail}; {n} total offers -> {OUTPUT_XLSX}"
    except Exception as e:  # noqa
        n = 0
        line = f"[{stamp}] ERROR  {e}"
    print(line)
    try:
        with open(os.path.join(HERE, "scan.log"), "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    return n


def run_scheduler():
    import schedule
    import time
    print(f"Scheduler started: scanning every {INTERVAL} min. Ctrl+C to stop.")
    run_once()  # immediate first scan
    schedule.every(INTERVAL).minutes.do(run_once)
    while True:
        schedule.run_pending()
        time.sleep(15)


def main():
    ap = argparse.ArgumentParser(description="Telegram placement-offer monitor")
    ap.add_argument("--list", action="store_true", help="list your groups and exit")
    ap.add_argument("--once", action="store_true", help="scan once and update Excel")
    ap.add_argument("--full", action="store_true", help="rescan entire history")
    ap.add_argument("--schedule", action="store_true", help="run continuously on a timer")
    args = ap.parse_args()

    if args.list:
        with make_client() as client:
            client.start(phone=PHONE)
            list_groups(client)
    elif args.schedule:
        run_scheduler()
    elif args.once or args.full:
        run_once(full=args.full)
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
