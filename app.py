"""
Financial Statements Extractor — Streamlit app.

Upload annual report PDFs; get the balance sheet, profit and loss and cash flow statement,
year by year, as an Excel workbook. Every figure is copied from the PDF's text layer, and a
statement is only shown or exported after it passes verification against the report.

Run locally:   streamlit run app.py
"""
from __future__ import annotations

import hashlib
import html
import io
import math
import os
import re
import statistics
from collections import Counter
from datetime import datetime

import pymupdf

# =====================================================================================
#  1. Patterns
# =====================================================================================
MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}
MON_NAMES = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

TYPES = {
    "BS": {"name": "Balance Sheet", "title": re.compile(r"balance\s*sheet"),
           "body": re.compile(r"total\s+assets|equity\s+and\s+liabilities|capital\s+and\s+liabilities|liabilities\s+and\s+equity")},
    "PL": {"name": "Profit and Loss", "title": re.compile(r"statement\s+of\s+profit|profit\s*(and|&)\s*loss|income\s+statement|statement\s+of\s+income"),
           "body": re.compile(r"(revenue\s+from\s+operations|total\s+income|interest\s+earned|revenue)[\s\S]*(profit|loss)")},
    "CF": {"name": "Cash Flow", "title": re.compile(r"cash\s*flows?\s*statement|statement\s+of\s+cash\s*flows?"),
           "body": re.compile(r"operating\s+activities|investing\s+activities|financing\s+activities")},
}
EXCLUDE_TOP = re.compile(r"notes?\s+(to|forming)|accounting\s+polic|independent\s+auditor|auditor'?s'?\s+report|annexure|board'?s'?\s+report|"
                         r"directors'?\s+report|management\s+discussion|ten[- ]year|financial\s+highlights|key\s+ratios|statement\s+of\s+changes\s+in\s+equity")
END_MARK = re.compile(r"accompanying\s+notes|as\s+per\s+our\s+report|for\s+and\s+on\s+behalf|notes\s+referred\s+to\s+above|see\s+accompanying|"
                      r"summary\s+of\s+(significant|material)\s+accounting|in\s+terms\s+of\s+our\s+report|the\s+notes\s+are\s+an\s+integral|"
                      r"forms?\s+an?\s+integral\s+part|schedules\s+referred\s+to", re.I)
NUM_RE = re.compile(r"^\(?[-–−]?(?:₹|rs\.?|`)?\s?\d{1,3}(?:,\d{2,3})*(?:\.\d+)?\)?$|^\(?[-–−]?\d+(?:\.\d+)?\)?$", re.I)
DASH_RE = re.compile(r"^[-–—−]+$|^nil$", re.I)
NOTE_RE = re.compile(r"^\d{1,2}(?:\.\d{1,2})?(?:\s?\(?[a-z]{1,3}\)?)?$|^\d{1,2}[a-z]$", re.I)
NO_CONVERT = re.compile(r"earnings?\s+per|per\s+(equity\s+)?share|\beps\b|^basic\b|^diluted\b|basic\s*(and|&)\s*diluted|face\s+value|"
                        r"number\s+of\s+(equity\s+)?shares|no\.\s*of\s+shares|weighted\s+average\s+(number|no)", re.I)
MONTH_RE = re.compile(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b", re.I)
UNIT_FACTORS = {"rupees": 1, "thousand": 1e3, "lakh": 1e5, "million": 1e6, "crore": 1e7, "billion": 1e9}
UNIT_LABEL = {"rupees": "₹", "thousand": "₹ thousand", "lakh": "₹ lakh", "million": "₹ million", "crore": "₹ crore", "billion": "₹ billion"}
PAT_RE = re.compile(r"^(net )?(profit|loss)( loss)?( for the (year|period)| after tax(es|ation)?)( from continuing operations)?$|^profit after tax")
AFTER_PAT = re.compile(r"other comprehensive|total comprehensive|earnings? per|attributable to|^owners|non controlling|appropriation")
NET_CHANGE_RE = re.compile(r"^net (increase|decrease|change|movement|inflow|outflow).*\bcash\b")


def is_num(s: str) -> bool:
    return bool(NUM_RE.match(s) or DASH_RE.match(s))


def is_amount(s: str) -> bool:
    return is_num(s) and bool(re.search(r"\d[,.]\d", s) or DASH_RE.match(s))


def parse_num(raw: str):
    s = raw.strip()
    if DASH_RE.match(s):
        return 0.0
    neg = bool(re.match(r"^\(.*\)$", s) or re.match(r"^[-–−]", s))
    t = re.sub(r"[()₹`,\s]|rs\.?", "", s, flags=re.I)
    t = re.sub(r"^[-–−]", "", t)
    try:
        n = float(t)
    except ValueError:
        return None
    return -n if neg else n


def decimals_of(raw: str) -> int:
    m = re.search(r"\.(\d+)", raw)
    return len(m.group(1)) if m else 0


def median(a):
    if not a:
        return 0
    b = sorted(a)
    return b[len(b) // 2]


NOTE_TAIL = re.compile(r"\s+(?:note\s*)?\d{1,2}(?:\.\d{1,2})?\s?(?:\([a-z]{1,3}\)|[a-z])?(?:\s*(?:,|&|and)\s*\d{1,2}(?:\.\d{1,2})?\s?(?:\([a-z]{1,3}\))?)*$", re.I)


# "increase / (decrease)" and "(decrease) / increase" name the same line: each pair maps to its first word
ALT_PAIRS = [("generated from", "used in"), ("proceeds from", "repayment of"), ("increase", "decrease"), ("profit", "loss"),
             ("gain", "loss"), ("inflow", "outflow"), ("generated", "used"), ("proceeds", "repayment"), ("surplus", "deficit"),
             ("income", "expenses"), ("income", "expense"), ("credit", "charge"), ("receipts", "payments")]
_ALT_RX = []
for _a, _b in ALT_PAIRS:
    _pa, _pb = _a.replace(" ", r"\s+"), _b.replace(" ", r"\s+")
    for _x, _y in ((_pa, _pb), (_pb, _pa)):
        _ALT_RX.append((re.compile(r"\(?\b" + _x + r"\b\)?\s*/\s*\(?\b" + _y + r"\b\)?"), _a))


def norm_label(s: str) -> str:
    """Canonical form of a line name: no '(a)'/'(i)'/'A.' prefixes, no note numbers, no punctuation,
    and 'X / (Y)' alternatives read the same whichever way round the report prints them."""
    s = s.lower().strip()
    for rx, canon in _ALT_RX:
        s = rx.sub(canon, s)
    s = re.sub(r"^net\s+(increase|decrease|change|movement)\b", "net change", s)
    s = re.sub(r"^[-–•*]\s*", "", s)
    for _ in range(2):
        s = re.sub(r"^\(?(?:[ivx]{1,4}|[a-z]|\d{1,2})[).]\s+", "", s)
    s = NOTE_TAIL.sub("", s)
    s = re.sub(r"/\s*\(?(loss|profit|decrease|increase|used in|generated from)\)?", "", s)
    s = re.sub(r"\((loss|profit)\)", "", s)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


# =====================================================================================
#  2. Reading pages into positioned words
# =====================================================================================
def read_pages(data: bytes, progress=None) -> dict:
    """Returns {'pages': [...], 'pdf_pages': n}. Each page: words with x0, x1, y (from bottom), h."""
    doc = pymupdf.open(stream=data, filetype="pdf")
    pages = []
    for pno in range(doc.page_count):
        page = doc[pno]
        W, H = page.rect.width, page.rect.height
        raw = sorted(page.get_text("words"), key=lambda w: (w[5], w[6], w[7]))
        words, cur, key = [], None, None
        for x0, y0, x1, y1, s, b, ln, _ in raw:
            s = s.replace(" ", " ").strip()
            if not s:
                continue
            h = max(y1 - y0, 1.0)
            w = {"s": s, "x0": x0, "x1": x1, "y": H - y1, "h": h}
            if cur is not None and key == (b, ln):
                gap = x0 - cur["x1"]
                glued = cur["s"] + s
                if gap < 2.5 and NUM_RE.match(re.sub(r"\s+", "", glued)):          # "(" "1,234.56" ")" -> one number
                    cur["s"] = re.sub(r"\s+", "", glued); cur["x1"] = x1; continue
                if gap < h * 0.55 and not is_amount(cur["s"]) and not is_amount(s):  # words of one phrase
                    cur["s"] += " " + s; cur["x1"] = x1; continue
            if cur is not None:
                words.append(cur)
            cur, key = w, (b, ln)
        if cur is not None:
            words.append(cur)
        pages.extend(split_spread({"pdfPage": pno + 1, "half": None, "xOff": 0.0, "W": W, "H": H, "fullW": W, "words": words}))
        if progress:
            progress(pno + 1, doc.page_count)
    for i, pg in enumerate(pages):
        pg["n"] = i + 1
        pg["text"] = " ".join(w["s"] for w in pg["words"]).lower()
        pg["folio"] = folio_of(pg)
    n = doc.page_count
    doc.close()
    return {"pages": pages, "pdf_pages": n}


def split_spread(pg):
    """A landscape PDF page holding two printed pages is split at the empty gutter."""
    W, H, ws = pg["W"], pg["H"], pg["words"]
    if W < H * 1.15 or len(ws) < 6:
        return [pg]
    best_x, best_cov = None, 10 ** 9
    x = W * 0.40
    while x <= W * 0.60:
        cov = sum(1 for w in ws if w["x0"] < x < w["x1"] and H * 0.06 < w["y"] < H * 0.94)
        if cov < best_cov or (cov == best_cov and abs(x - W / 2) < abs(best_x - W / 2)):
            best_cov, best_x = cov, x
        x += 1
    if best_cov > 0:
        return [pg]
    L = [w for w in ws if (w["x0"] + w["x1"]) / 2 < best_x]
    R = [w for w in ws if (w["x0"] + w["x1"]) / 2 >= best_x]
    if len(L) < 3 or len(R) < 3:
        return [pg]
    return [
        {**pg, "half": "left", "xOff": 0.0, "W": best_x, "words": L},
        {**pg, "half": "right", "xOff": best_x, "W": W - best_x,
         "words": [{**w, "x0": w["x0"] - best_x, "x1": w["x1"] - best_x} for w in R]},
    ]


def page_name(pg) -> str:
    if not pg:
        return ""
    return f"PDF page {pg['pdfPage']}" + (f" ({pg['half']} half)" if pg["half"] else "")


def folio_of(pg):
    c = [w for w in pg["words"] if (w["y"] < pg["H"] * 0.07 or w["y"] > pg["H"] * 0.95) and re.fullmatch(r"\d{1,3}", w["s"])]
    if not c:
        return None
    return int(sorted(c, key=lambda w: w["y"])[0]["s"])


def to_rows(page):
    ws = sorted(page["words"], key=lambda w: (-w["y"], w["x0"]))
    rows = []
    tol = max(2.0, median([w["h"] for w in ws]) * 0.45)
    for w in ws:
        if rows and abs(rows[-1]["y"] - w["y"]) <= tol:
            rows[-1]["words"].append(w)
        else:
            rows.append({"y": w["y"], "words": [w]})
    for r in rows:
        r["words"].sort(key=lambda w: w["x0"])
        r["text"] = " ".join(w["s"] for w in r["words"])
    return rows


def top_text(page, frac=0.3) -> str:
    ws = [w for w in page["words"] if w["y"] > page["H"] * (1 - frac)]
    ws.sort(key=lambda w: (-w["y"], w["x0"]))
    return " ".join(w["s"] for w in ws).lower()


# =====================================================================================
#  3. Contents page and finding the statement pages
# =====================================================================================
def is_section(t: str) -> bool:
    return bool(re.search(r"(standalone|consolidated)", t) and re.search(r"(financial\s+statements?|financials|accounts)", t)
                and not re.search(r"balance|profit|cash|notes?\b|auditor", t))


def read_contents(pages):
    out = {"entries": [], "found": False, "sections": [], "targets": {}}
    toc_pages = [pg for pg in pages[:16] if re.search(r"\b(contents|index|inside this report|what'?s inside)\b", pg["text"])]
    num_seg = re.compile(r"^\d{1,3}(\s*[-–]\s*\d{1,3})?$")
    for pg in toc_pages:
        for r in to_rows(pg):
            segs = []
            for w in r["words"]:
                if segs and w["x0"] - segs[-1]["x1"] < 22:
                    segs[-1]["s"] += " " + w["s"]; segs[-1]["x1"] = w["x1"]
                else:
                    segs.append({"s": w["s"], "x0": w["x0"], "x1": w["x1"]})
            i = 0
            while i < len(segs):
                t = re.sub(r"[.·…_]{2,}", " ", segs[i]["s"]).strip()
                label = num = None
                m1 = re.match(r"^(.*[a-z].*?)[\s.·…_-]+(\d{1,3})(?:\s*[-–]\s*\d{1,3})?$", t, re.I)
                m2 = re.match(r"^(\d{1,3})(?:\s*[-–]\s*\d{1,3})?\s+(.*[a-z].*)$", t, re.I)
                if m1:
                    label, num = m1.group(1), int(m1.group(2))
                elif m2:
                    num, label = int(m2.group(1)), m2.group(2)
                elif num_seg.match(t) and i + 1 < len(segs) and not num_seg.match(segs[i + 1]["s"].strip()) and re.search(r"[a-z]", segs[i + 1]["s"], re.I):
                    num = int(re.match(r"\d+", t).group()); i += 1; label = segs[i]["s"]
                elif re.search(r"[a-z]", t, re.I) and i + 1 < len(segs) and num_seg.match(segs[i + 1]["s"].strip()):
                    label = t; i += 1; num = int(re.match(r"\d+", segs[i]["s"].strip()).group())
                if label and num is not None:
                    out["entries"].append({"label": label.strip(), "printed": num, "low": label.lower()})
                i += 1
    if not out["entries"]:
        return out
    out["entries"].sort(key=lambda e: (e["printed"], 0 if is_section(e["low"]) else 1))
    basis = None
    for e in out["entries"]:
        low = e["low"]
        if is_section(low):
            basis = "Consolidated" if "consolidated" in low else "Standalone"
            out["sections"].append({"basis": basis, "printed": e["printed"]})
            continue
        if re.search(r"notes?\b", low):
            continue
        typ = ("BS" if re.search(r"balance\s*sheet", low) else
               "PL" if re.search(r"profit\s*(and|&)\s*loss|income\s+statement|statement\s+of\s+income", low) else
               "CF" if re.search(r"cash\s*flows?", low) else None)
        if not typ:
            continue
        b = "Consolidated" if "consolidated" in low else "Standalone" if "standalone" in low else basis
        k = f"{typ}|{b or 'Standalone'}"
        out["targets"].setdefault(k, e["printed"])
    out["found"] = bool(out["targets"] or out["sections"])
    return out


def printed_to_index(pages, printed):
    direct = [p for p in pages if p["folio"] == printed]
    votes = Counter(p["n"] - p["folio"] for p in pages if p["folio"] is not None)
    off = None
    if votes:
        o, c = votes.most_common(1)[0]
        off = o if c >= 3 else None
    if len(direct) == 1:
        return direct[0]["n"]
    if len(direct) > 1 and off is not None:
        return sorted(direct, key=lambda p: abs(p["n"] - (printed + off)))[0]["n"]
    if direct:
        return direct[0]["n"]
    return printed + off if off is not None else None


def classify(pages, toc):
    cands, ranges = [], []
    if toc and toc["sections"]:
        secs = [{"basis": s["basis"], "start": printed_to_index(pages, s["printed"])} for s in toc["sections"]]
        secs = sorted([s for s in secs if s["start"]], key=lambda s: s["start"])
        for i, s in enumerate(secs):
            ranges.append({"basis": s["basis"], "start": s["start"], "end": secs[i + 1]["start"] - 1 if i + 1 < len(secs) else len(pages)})

    def basis_from_toc(n):
        for r in ranges:
            if r["start"] <= n <= r["end"]:
                return r["basis"]
        return None

    for pg in pages:
        top = top_text(pg)
        if EXCLUDE_TOP.search(top):
            continue
        nums = sum(1 for w in pg["words"] if re.search(r"\d,\d|\d\.\d", w["s"]) and is_num(w["s"]))
        if nums < 12:
            continue
        for typ, d in TYPES.items():
            if not d["title"].search(top) or not d["body"].search(pg["text"]):
                continue
            if typ == "BS" and TYPES["PL"]["title"].search(top) and not re.search(r"balance\s*sheet", top[:120]):
                continue
            score = nums + 50
            if re.search(r"as\s+at|year\s+ended|period\s+ended", top):
                score += 15
            if re.search(r"particulars|\bnote", pg["text"]):
                score += 10
            basis = ("Consolidated" if "consolidated" in top else "Standalone" if "standalone" in top else (basis_from_toc(pg["n"]) or "Standalone"))
            cands.append({"type": typ, "basis": basis, "page": pg["n"], "score": score})
    best = {}
    for c in cands:
        k = f"{c['type']}|{c['basis']}"
        if k not in best or c["score"] > best[k]["score"]:
            best[k] = c
    if toc:
        for k, printed in toc["targets"].items():
            idx = printed_to_index(pages, printed)
            if not idx:
                continue
            typ, basis = k.split("|")
            near = [c for c in cands if c["type"] == typ and idx - 1 <= c["page"] <= idx + 2]
            near.sort(key=lambda c: (abs(c["page"] - idx), -c["score"]))
            if near:
                best[k] = {**near[0], "basis": basis, "via": "contents"}
    return best


# =====================================================================================
#  4. Rebuilding a statement table
# =====================================================================================
_M = r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)"
DATE_PATTERNS = [
    (re.compile(r"(\d{1,2})(?:st|nd|rd|th)?[\s\-./,]*" + _M + r"[a-z]*[\s\-./,]*(\d{4})", re.I), lambda m: (int(m[3]), MONTHS[m[2].lower()], int(m[1]))),
    (re.compile(_M + r"[a-z]*[\s.]*(\d{1,2})(?:st|nd|rd|th)?,?\s*(\d{4})", re.I), lambda m: (int(m[3]), MONTHS[m[1].lower()], int(m[2]))),
    (re.compile(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b"), lambda m: (int(m[3]), int(m[2]), int(m[1]))),
    (re.compile(r"\b" + _M + r"[a-z]*\s*[-'’]\s*(\d{2})\b", re.I), lambda m: (2000 + int(m[2]), MONTHS[m[1].lower()], 0)),
    (re.compile(r"\b" + _M + r"[a-z]*[\s,]+(\d{4})\b", re.I), lambda m: (int(m[2]), MONTHS[m[1].lower()], 0)),
    (re.compile(r"\b(?:fy\s*)?(20\d{2})\s*[-–/]\s*(\d{2,4})\b", re.I), lambda m: (2000 + int(m[2]) if len(m[2]) == 2 else int(m[2]), 3, 31)),
    (re.compile(r"\bfy\s*'?(\d{2,4})\b", re.I), lambda m: (2000 + int(m[1]) if len(m[1]) == 2 else int(m[1]), 3, 31)),
    (re.compile(r"\b(20\d{2})\b"), lambda m: (int(m[1]), 3, 31)),
]


def parse_date(text, patterns=DATE_PATTERNS):
    for rx, f in patterns:
        m = rx.search(text)
        if m:
            y, mo, d = f(m)
            if 1990 < y < 2100 and mo and 1 <= mo <= 12:
                return {"y": y, "m": mo, "d": d}
    return None


def is_header_row(t):
    return bool(re.search(r"particulars|^\s*note|as\s+at|year\s+ended|period\s+ended|for\s+the\s+year", t, re.I)
                or (re.search(_M + r"[a-z]*", t, re.I) and re.search(r"\b(19|20)\d{2}\b", t) and len(t.split(" ")) < 14))


def infer_years(cols, title_text, fy_hint):
    """Columns headed 'Current Year / Previous Year' get their year from a dated sibling, the title, or the report year."""
    anchor, ai = None, 0
    di = next((i for i, p in enumerate(cols) if p["date"]), -1)
    if di >= 0:
        anchor, ai = cols[di]["date"], di
    else:
        anchor = parse_date(title_text, DATE_PATTERNS[:6])
        if not anchor and fy_hint:
            anchor = {"y": fy_hint, "m": 3, "d": 31}
    if not anchor:
        return
    for i, p in enumerate(cols):
        if p["date"]:
            continue
        p["date"] = {"y": anchor["y"] - (i - ai), "m": anchor["m"], "d": anchor["d"]}
        p["key"] = f"{p['date']['y']}-{p['date']['m']}"
        p["inferred"] = True


def detect_report_year(pages, file_name):
    def from_range(t):
        m = re.search(r"(?:^|[^\d])((?:20)?\d{2})\s*[-–_ ]\s*((?:20)?\d{2})(?!\d)", t)
        if m:
            a = 2000 + int(m[1]) if len(m[1]) == 2 else int(m[1])
            b = 2000 + int(m[2]) if len(m[2]) == 2 else int(m[2])
            if b == a + 1:
                return b
        return None
    if file_name:
        n = re.sub(r"\.pdf$", "", file_name, flags=re.I)
        r = from_range(n)
        if r:
            return r
        fy = re.search(r"fy\s*[-_ ]?\s*'?(\d{2}|20\d{2})(?!\d)", n, re.I)
        if fy:
            return 2000 + int(fy[1]) if len(fy[1]) == 2 else int(fy[1])
    counts = Counter()
    for pg in pages[:8]:
        for m in re.finditer(r"(20\d{2})\s*[-–]\s*(\d{2})(?!\d)", pg["text"]):
            a, b = int(m[1]), 2000 + int(m[2])
            if b == a + 1:
                counts[b] += 1
    if counts:
        return sorted(counts.items(), key=lambda kv: (-kv[1], -kv[0]))[0][0]
    if file_name:
        y = re.search(r"(?:^|[^\d])(20\d{2})(?!\d)", file_name)
        if y:
            return int(y[1])
    return None


def detect_unit(text):
    t = text.lower()
    U = r"(crores?|cr\b|lakhs?|lacs?|millions?|mn\b|billions?|thousands?|'000)"
    pats = [r"(?:₹|rs\.?|inr|rupees|`|■)\s*(?:in\s+)?" + U,
            r"(?:amounts?|figures?)\s+(?:are\s+)?in\s+(?:\S{1,4}\s*)?" + U,
            r"\(\s*(?:\S{1,4}\s+)?in\s+" + U,
            r"\bin\s+(?:\S{1,4}\s+)?" + U + r"\b",
            U + r"\s+of\s+(?:₹|rupees|inr)"]
    for p in pats:
        m = re.search(p, t)
        if m:
            u = m.group(1)
            if u.startswith("cr"):
                return "crore"
            if u.startswith("la"):
                return "lakh"
            if u.startswith("m"):
                return "million"
            if u.startswith("b"):
                return "billion"
            if u.startswith("th") or "000" in u:
                return "thousand"
    return None


def is_total(label):
    l = re.sub(r"^([ivx]+|[a-z]|\d+)\.\s*|^\([a-z]+\)\s*", "", label, flags=re.I)
    return bool(re.match(r"^(total|net\s+(cash|increase|decrease|profit|loss)|profit|loss|cash\s+and\s+cash\s+equivalents\s+at\s+the\s+end|grand\s+total|sub[- ]?total)", l, re.I))


def parse_statement(pages_data, typ, fy_hint):
    out = {"rows": [], "periods": [], "warnings": [], "decimals": 0, "unit": None}
    unit, ended, col_model = None, False, None
    for pg in pages_data:
        if ended:
            break
        unit = unit or detect_unit(top_text(pg, 0.4))
        body = [r for r in to_rows(pg) if 32 < r["y"] < pg["H"] - 40]
        title_re = TYPES[typ]["title"]
        ttl = next((i for i, r in enumerate(body) if title_re.search(r["text"].lower())), -1)
        hdr_end = -1
        for i in range(max(ttl, 0), min(len(body), 18)):
            amounts = sum(1 for w in body[i]["words"] if re.search(r"\d[,.]\d", w["s"]) and is_num(w["s"]))
            if amounts >= 1 and not MONTH_RE.search(body[i]["text"]):
                break
            if is_header_row(body[i]["text"]):
                hdr_end = i
        title_idx = next((i for i, r in enumerate(body) if title_re.search(r["text"].lower()) or re.search(r"continued|contd", r["text"], re.I)), -1)
        start = max(hdr_end, title_idx) + 1
        end = len(body)
        for i in range(start, len(body)):
            if END_MARK.search(body[i]["text"]):
                end, ended = i, True
                break
        data_rows, hdr_rows = body[start:end], body[title_idx + 1:start]

        # numeric columns from right edges
        nw = sorted([w for r in data_rows for w in r["words"] if (is_num(w["s"]) or NOTE_RE.match(w["s"])) and w["x1"] > pg["W"] * 0.3], key=lambda w: w["x1"])
        clusters = []
        for w in nw:
            if clusters and w["x1"] - clusters[-1]["last"] <= 12:
                clusters[-1]["ws"].append(w); clusters[-1]["last"] = w["x1"]
            else:
                clusters.append({"ws": [w], "last": w["x1"]})
        cols = []
        for c in clusters:
            if len(c["ws"]) < max(2, len(data_rows) * 0.08):
                continue
            ws = c["ws"]
            note = (sum(1 for w in ws if NOTE_RE.match(w["s"]) and "," not in w["s"]) / len(ws) >= 0.8
                    and not any(re.search(r",\d{3}", w["s"]) or re.search(r"\.\d{2}$", w["s"]) for w in ws))
            cols.append({"right": median([w["x1"] for w in ws]), "left": min(w["x0"] for w in ws), "note": note})
        note_col = None
        first_val = next((i for i, c in enumerate(cols) if not c["note"]), -1)
        if first_val > 0:
            note_col = cols[first_val - 1]
        val_cols = [c for c in cols if not c["note"] and (not note_col or c["right"] > note_col["right"])]
        if not val_cols:
            if not col_model:
                out["warnings"].append(f"{page_name(pg)}: no numeric columns found")
            continue

        periods = []
        for i, c in enumerate(val_cols):
            parts = []
            for r in hdr_rows:
                if detect_unit(r["text"]) and not re.search(r"\b(19|20)\d{2}\b", re.sub(r"\(.*?\)", "", r["text"])):
                    continue
                best, bo = None, -26
                for w in r["words"]:
                    ov = min(w["x1"], c["right"] + 4) - max(w["x0"], c["left"] - 4)
                    if ov > bo:
                        bo, best = ov, w
                if best and bo > -25 and best["x1"] > pg["W"] * 0.3:
                    parts.append(best["s"])
            txt = " ".join(parts).strip()
            periods.append({"text": txt, "date": parse_date(txt)})
        pkeys = [f"{p['date']['y']}-{p['date']['m']}" if p["date"] else f"col{i}" for i, p in enumerate(periods)]
        if any(p["date"] for p in periods) and any(not p["date"] for p in periods):
            fixed = []
            for i, k in enumerate(pkeys):
                if periods[i]["date"]:
                    fixed.append(k); continue
                nk = next((pkeys[j] for j in range(i + 1, len(periods)) if periods[j]["date"]), None)
                nk = nk or next((pkeys[j] for j in range(i - 1, -1, -1) if periods[j]["date"]), None)
                fixed.append(nk or k)
            pkeys = fixed
        if col_model is None:
            uniq = list(dict.fromkeys(pkeys))
            col_model = []
            for k in uniq:
                p = periods[pkeys.index(k)]
                col_model.append({"key": k, "colKey": k, "date": p["date"], "header": p["text"], "inferred": False})
            if any(not p["date"] for p in col_model):
                infer_years(col_model, " ".join(r["text"] for r in body[:start]) + " " + top_text(pg, 0.4), fy_hint)
            for i, p in enumerate(col_model):
                p["label"] = f"{MON_NAMES[p['date']['m']]}-{str(p['date']['y'])[2:]}" if p["date"] else f"Column {i + 1}"
            out["periods"] = col_model
        col_to_period = []
        for i, c in enumerate(val_cols):
            k = pkeys[i]
            pi = next((j for j, p in enumerate(col_model) if p["key"] == k or p["colKey"] == k), -1)
            col_to_period.append(pi if pi >= 0 else min(i, len(col_model) - 1))
        label_limit = (note_col["left"] if note_col else val_cols[0]["left"]) - 4

        lines = []
        for r in data_rows:
            lab, vals, note = [], [None] * len(col_model), None
            for w in r["words"]:
                if (is_num(w["s"]) or NOTE_RE.match(w["s"])) and w["x1"] > label_limit:
                    if note_col and abs(w["x1"] - note_col["right"]) < 14 and NOTE_RE.match(w["s"]):
                        note = f"{note}, {w['s']}" if note else w["s"]
                        continue
                    best, bd = -1, 1e9
                    for i, c in enumerate(val_cols):
                        d = abs(w["x1"] - c["right"])
                        if d < bd:
                            bd, best = d, i
                    if best >= 0 and bd < 40:
                        vals[col_to_period[best]] = {"raw": w["s"], "v": parse_num(w["s"])}
                        out["decimals"] = max(out["decimals"], decimals_of(w["s"]))
                        continue
                if w["x0"] < label_limit or not is_num(w["s"]):
                    lab.append(w)
            # a note reference printed just after the line name ("... equipment   4(a)") is a note, not part of the name
            if len(lab) >= 2 and NOTE_RE.match(lab[-1]["s"]) and lab[-1]["x0"] - lab[-2]["x1"] > 12:
                tail = lab.pop()["s"]
                note = f"{tail}, {note}" if note else tail
            label = re.sub(r"\s+", " ", " ".join(w["s"] for w in lab)).strip()
            m_note = re.search(r"\s+(\d{1,2}(?:\.\d{1,2})?\s?\([a-z]{1,3}\))$", label, re.I)
            if m_note:
                label = label[:m_note.start()].strip()
                note = note or m_note.group(1)
            if label or any(vals) or note:
                lines.append({"label": label, "note": note, "vals": vals, "hasVals": any(vals)})

        merged = []
        for i, L in enumerate(lines):
            prev = merged[-1] if merged else None
            starts_lower = bool(re.match(r"^[a-z]", L["label"]) or re.match(r"^(and|or|of|in|to|from|on|for)\b", L["label"], re.I))
            if prev and not prev["hasVals"] and L["hasVals"] and (L["label"] == "" or starts_lower):
                prev["label"] = (prev["label"] + " " + L["label"]).strip()
                prev["vals"], prev["hasVals"], prev["note"] = L["vals"], True, prev["note"] or L["note"]
                continue
            if prev and prev["hasVals"] and not L["hasVals"] and starts_lower and L["label"]:
                prev["label"] += " " + L["label"]
                continue
            merged.append(dict(L))
        for r in merged:
            r["kind"] = ("total" if is_total(r["label"]) else "item") if r["hasVals"] else "heading"
        out["rows"].extend(r for r in merged if r["label"] or r["hasVals"])
    out["unit"] = unit
    section, seen = "", Counter()
    for r in out["rows"]:
        if r["kind"] == "heading":
            section = norm_label(r["label"])
        base = (section + "|" + norm_label(r["label"]))[:160]
        seen[base] += 1
        r["key"] = base + (f"#{seen[base]}" if seen[base] > 1 else "")
        r["noConvert"] = bool(NO_CONVERT.search(r["label"]) or (re.search(r"earnings?\s+per", section) and r["kind"] != "heading"))
    if not out["rows"]:
        out["warnings"].append("No rows extracted")
    return out


# =====================================================================================
#  5. Arithmetic ties
# =====================================================================================
def find_row(st, rx):
    return next((r for r in st["rows"] if r["hasVals"] and re.search(rx, norm_label(r["label"]))), None)


def run_checks(st, typ):
    res = []
    tol = (0.06 if st["decimals"] else 1.5) * st.get("tol_scale", 1)

    def v(r, i):
        return r["vals"][i]["v"] if r and r["vals"][i] else None

    def cmp(name, a, b, hard=True):
        for i, p in enumerate(st["periods"]):
            try:
                x, y = a(i), b(i)
            except TypeError:
                continue
            if x is None or y is None:
                continue
            diff = round(x - y, 4)
            res.append({"name": name, "period": p["label"], "expected": y, "got": x, "diff": diff, "ok": abs(diff) <= tol, "hard": hard})

    if typ == "BS":
        ta = find_row(st, r"^total assets$")
        tel = find_row(st, r"^total (equity and liabilities|liabilities and equity|capital and liabilities|equity liabilities)")
        if ta and tel:
            cmp("Total assets = Total equity and liabilities", lambda i: v(ta, i), lambda i: v(tel, i))
        else:
            tots = [r for r in st["rows"] if re.fullmatch(r"total", norm_label(r["label"]))]
            if len(tots) >= 2:
                cmp("Total (liabilities side) = Total (assets side)", lambda i: v(tots[0], i), lambda i: v(tots[1], i))
    if typ == "PL":
        rev, oi, ti = find_row(st, r"^revenue from operations$"), find_row(st, r"^other income$"), find_row(st, r"^total (income|revenue)$")
        if rev and oi and ti:
            cmp("Revenue + Other income = Total income", lambda i: v(rev, i) + v(oi, i), lambda i: v(ti, i))
        te, pb = find_row(st, r"^total expenses$"), find_row(st, r"^profit before (exceptional|tax|share)")
        if ti and te and pb:
            cmp("Total income − Total expenses = Profit before tax line", lambda i: v(ti, i) - v(te, i), lambda i: v(pb, i))
        tots = [r for r in st["rows"] if r["hasVals"] and re.fullmatch(r"total", norm_label(r["label"]))]
        np_ = find_row(st, r"^net profit")
        if not ti and len(tots) == 2 and np_:
            cmp("Total income − Total expenditure = Net profit", lambda i: v(tots[0], i) - v(tots[1], i), lambda i: v(np_, i))
    if typ == "CF":
        a, b, c = find_row(st, r"net cash.*operating"), find_row(st, r"net cash.*investing"), find_row(st, r"net cash.*financing")
        n = find_row(st, r"^net (increase|decrease|change).*cash")
        if a and b and c and n:
            cmp("Operating + Investing + Financing = Net change in cash", lambda i: v(a, i) + v(b, i) + v(c, i), lambda i: v(n, i))
        o = find_row(st, r"cash and cash equivalents at the (beginning|start|commencement)")
        cl = find_row(st, r"cash and cash equivalents at the (end|close)")
        if n and o and cl:
            cmp("Opening cash + Net change = Closing cash", lambda i: v(o, i) + v(n, i), lambda i: v(cl, i), hard=False)
    return res


# =====================================================================================
#  6. Mandatory verification — nothing is shown or exported unless all checks pass
# =====================================================================================
FOREIGN = {
    "BS": re.compile(r"^(revenue from operations|cost of materials consumed|profit before tax|employee benefits? expenses?|total expenses|net cash (flow )?(generated|used|from))"),
    "PL": re.compile(r"^(property plant and equipment|total assets|total equity and liabilities|trade payables|capital work in progress|equity share capital|"
                     r"other equity|total equity|total current liabilities|total non current assets|net cash (flow )?(generated|used|from))"),
    "CF": re.compile(r"^(total assets|total equity and liabilities|equity share capital|other equity|total equity|total income|total expenses)$"),
}
MUST = {
    "BS": [(r"^total assets$|^total$|^total equity and liabilities$|^total capital and liabilities$", "a total assets line")],
    "PL": [(r"revenue|income|interest earned", "a revenue or income line"), (PAT_RE.pattern, "a profit after tax line")],
    "CF": [(r"operating activities", "operating activities"), (r"investing activities|financing activities", "investing or financing activities"),
           (NET_CHANGE_RE.pattern, "a net change in cash line")],
}


def verify(st, pages):
    res = []
    first = pages[st["pages"][0] - 1] if st["pages"] and st["pages"][0] - 1 < len(pages) else None
    top = top_text(first) if first else ""
    own = TYPES[st["type"]]["title"].search(top)
    others = [d["name"] for t, d in TYPES.items() if t != st["type"] and d["title"].search(top)]
    if not first:
        detail = "page missing"
    elif not own:
        detail = f"{page_name(first)} isn't headed \"{st['name']}\""
    elif others:
        detail = f"{page_name(first)} also carries the {' and '.join(others)} heading"
    else:
        detail = f"{page_name(first)} is headed {st['name']}"
    res.append({"name": "Statement heading on the page", "ok": bool(first and own and not others), "detail": detail})

    labels = [norm_label(r["label"]) for r in st["rows"]]
    foreign = [r for r in st["rows"] if FOREIGN[st["type"]].search(norm_label(r["label"]))]
    missing = [msg for rx, msg in MUST[st["type"]] if not any(re.search(rx, l) for l in labels)]
    if foreign:
        detail = "found " + ", ".join(f"\"{r['label']}\"" for r in foreign[:2]) + ", which belong to another statement"
    elif missing:
        detail = "no " + " or ".join(missing)
    elif len(st["rows"]) < 5:
        detail = "too few lines read"
    else:
        detail = f"{len(st['rows'])} lines read"
    res.append({"name": "Line items belong to this statement", "ok": not foreign and not missing and len(st["rows"]) >= 5, "detail": detail})

    on_page = {w["s"] for n in st["pages"] if n - 1 < len(pages) for w in pages[n - 1]["words"]}
    total = traced = 0
    for r in st["rows"]:
        for c in r["vals"]:
            if c:
                total += 1
                traced += c["raw"] in on_page
    res.append({"name": "Every number traced to the page", "ok": total > 0 and traced == total, "detail": f"{traced} of {total} figures found on {st['where']}"})

    keys = [p["key"] for p in st["periods"]]
    dated = bool(st["periods"]) and all(p["date"] for p in st["periods"]) and len(set(keys)) == len(keys)
    res.append({"name": "Year columns identified", "ok": dated,
                "detail": ", ".join(p["label"] + (" (from title)" if p.get("inferred") else "") for p in st["periods"]) if dated else "could not tell which year each column is"})

    hard = [c for c in st["checks"] if c["hard"]]
    bad = [c for c in hard if not c["ok"]]
    if not hard:
        detail = "no totals found to tie, so the figures could not be proven"
    elif bad:
        detail = "; ".join(f"{c['name']} is off by {abs(c['diff']):g} for {c['period']}" for c in bad)
    else:
        detail = f"{len(hard)} ties checked"
    res.append({"name": "Totals tie", "ok": bool(hard) and not bad, "detail": detail})
    st["verification"] = res
    st["verified"] = all(r["ok"] for r in res)
    return st


# =====================================================================================
#  7. Pipeline
# =====================================================================================
def build_statement(pages, typ, basis, page_nums, fy_hint):
    use = [pages[n - 1] for n in page_nums if 0 < n <= len(pages)]
    st = parse_statement(use, typ, fy_hint)
    if len(page_nums) == 1:
        n = page_nums[0]
        for _ in range(2):
            last = pages[n - 1]
            tail = " ".join(r["text"] for r in to_rows(last)[-14:])
            if n >= len(pages) or END_MARK.search(tail):
                break
            nx = pages[n]
            top = top_text(nx)
            other = any(t != typ and d["title"].search(top) for t, d in TYPES.items())
            if EXCLUDE_TOP.search(top) or other:
                break
            if sum(1 for w in nx["words"] if is_num(w["s"]) and re.search(r"\d,\d|\d\.\d", w["s"])) < 6:
                break
            n += 1
            use.append(nx)
        if len(use) > 1:
            st = parse_statement(use, typ, fy_hint)
    st.update(type=typ, basis=basis, name=TYPES[typ]["name"], pages=[p["n"] for p in use],
              pdfPages=list(dict.fromkeys(p["pdfPage"] for p in use)), where=", ".join(page_name(p) for p in use))
    st["checks"] = run_checks(st, typ)  # ties are checked on the full statement as printed
    if typ == "PL":
        trim_after_profit(st)
    if typ == "CF":
        trim_after_cash_change(st)
    return st


def trim_after_profit(st):
    """A P&L ends at profit after tax: other comprehensive income, EPS and anything after are dropped."""
    labels = [norm_label(r["label"]) for r in st["rows"]]
    cut = next((i for i, l in enumerate(labels) if AFTER_PAT.search(l)), len(labels))
    pats = [i for i, l in enumerate(labels[:cut]) if PAT_RE.search(l) and st["rows"][i]["hasVals"]]
    if not pats:
        pats = [i for i, l in enumerate(labels) if PAT_RE.search(l) and st["rows"][i]["hasVals"]][:1]
    if pats:
        st["rows"] = st["rows"][:pats[-1] + 1]
        st["rows"][-1]["kind"] = "total"
        st["endsAtProfit"] = True
    return st


def trim_after_cash_change(st):
    """A cash flow statement ends at the net change in cash: opening and closing balances and anything after are dropped."""
    for i, r in enumerate(st["rows"]):
        if r["hasVals"] and NET_CHANGE_RE.search(norm_label(r["label"])):
            st["rows"] = st["rows"][:i + 1]
            st["rows"][-1]["kind"] = "total"
            st["endsAtNetChange"] = True
            break
    return st


def build_from_pdf_pages(rep, typ, basis, pdf_nums):
    """The person types PDF page numbers; on a spread, take the half headed with this statement."""
    pages, idx = rep["pages"], []
    for p in pdf_nums:
        halves = [pg for pg in pages if pg["pdfPage"] == p]
        titled = [pg for pg in halves if TYPES[typ]["title"].search(top_text(pg))]
        idx.extend(pg["n"] for pg in (titled or halves))
    st = build_statement(pages, typ, basis, idx or [1], rep["fy_hint"])
    st["foundVia"] = "pages you entered"
    sib = next((s for s in rep["statements"] if s["basis"] == basis and s.get("unit")), None)
    if not st["unit"] and sib:
        st["unit"], st["unitInferred"] = sib["unit"], True
    return verify(st, pages)


def analyze(data: bytes, file_name: str, progress=None) -> dict:
    rep = read_pages(data, progress)
    pages = rep["pages"]
    rep["fy_hint"] = detect_report_year(pages, file_name)
    rep["toc"] = read_contents(pages)
    best = classify(pages, rep["toc"])
    rep["statements"] = []
    for basis in ("Standalone", "Consolidated"):
        for typ in ("BS", "PL", "CF"):
            b = best.get(f"{typ}|{basis}")
            if not b:
                continue
            st = build_statement(pages, typ, basis, [b["page"]], rep["fy_hint"])
            st["foundVia"] = "contents page" if b.get("via") == "contents" else "page scan"
            rep["statements"].append(st)
    for basis in ("Standalone", "Consolidated"):
        sts = [s for s in rep["statements"] if s["basis"] == basis]
        u = next((s["unit"] for s in sts if s["unit"]), None)
        for s in sts:
            if not s["unit"] and u:
                s["unit"], s["unitInferred"] = u, True
    for st in rep["statements"]:
        verify(st, pages)
    return rep


# =====================================================================================
#  8. Units, multi-year merge and the Excel workbook
# =====================================================================================
def unit_of(st):
    return st.get("unitOverride") or st.get("unit")


def view(st, out_unit):
    src = unit_of(st)
    can = out_unit != "reported" and src
    rows = []
    for r in st["rows"]:
        vals = []
        for c in r["vals"]:
            if c is None:
                vals.append(None)
            elif can and not r["noConvert"]:
                vals.append({"raw": c["raw"], "v": round(c["v"] * UNIT_FACTORS[src] / UNIT_FACTORS[out_unit], 2)})
            else:
                vals.append({"raw": c["raw"], "v": c["v"]})
        rows.append({**r, "vals": vals})
    target = out_unit if can else src
    note = ""
    if can and src != out_unit:
        note = f"converted from {UNIT_LABEL[src]}"
    if out_unit != "reported" and not src:
        note = "unit not found in the report, so figures are shown as reported"
    return {"periods": st["periods"], "rows": rows, "unit": target, "unitText": UNIT_LABEL[target] if target else "Unit not stated in report",
            "decimals": 2 if can else st["decimals"], "note": note}


def date_key(p):
    return p["date"]["y"] * 100 + p["date"]["m"] if p and p.get("date") else 0


def merge_years(items, out_unit):
    """items: [(report_name, statement)]. One row per line item, years oldest → latest.
    Newer reports win where years overlap; the older figure is kept as a 'restated' note."""
    views = sorted([{"file": f, "v": view(st, out_unit)} for f, st in items],
                   key=lambda x: -date_key(x["v"]["periods"][0] if x["v"]["periods"] else None))
    for x in views:
        x["keys"] = [p["key"] if p["date"] else f"{x['file']}::{p['key']}" for p in x["v"]["periods"]]
    periods, seen = [], set()
    for x in views:
        for p, k in zip(x["v"]["periods"], x["keys"]):
            if k not in seen:
                seen.add(k)
                periods.append({**p, "key": k, "header": "" if p.get("inferred") else p.get("header", "")})
    periods.sort(key=date_key)  # oldest on the left, latest on the right

    rows, idx, by_label = [], {}, {}
    for x in views:
        last_pos = -1
        for r in x["v"]["rows"]:
            nl = norm_label(r["label"])
            m = idx.get(r["key"])
            if m is not None and x["file"] in m["files"]:
                m = None
            if m is None:  # same line under a slightly different heading or name in another year
                cands = [c for c in by_label.get(nl, []) if x["file"] not in c["files"] and c["kind"] == r["kind"]]
                if cands:
                    after = [c for c in cands if rows.index(c) > last_pos]
                    m = (after or cands)[0]
            if m is None:
                m = {"label": r["label"], "note": r["note"], "kind": r["kind"], "key": r["key"], "noConvert": r["noConvert"],
                     "vals": {}, "restated": {}, "files": set()}
                rows.insert(last_pos + 1, m)
                idx[r["key"]] = m
                by_label.setdefault(nl, []).append(m)
            m["files"].add(x["file"])
            last_pos = rows.index(m)
            for c, k in zip(r["vals"], x["keys"]):
                if not c:
                    continue
                if k not in m["vals"]:
                    m["vals"][k] = {"v": c["v"], "file": x["file"]}
                elif abs(m["vals"][k]["v"] - c["v"]) > 0.011:
                    m["restated"][k] = {"v": c["v"], "file": x["file"]}

    # the same line split across two rows (one per set of years) is joined into one
    i = 0
    while i < len(rows):
        a = rows[i]
        for j in range(i + 1, min(i + 4, len(rows))):
            b = rows[j]
            same_name = norm_label(b["label"]) == norm_label(a["label"])
            if b["kind"] == a["kind"] and ((same_name and not (set(a["vals"]) & set(b["vals"]))) or
                                          (split_pair(a, b) and name_similarity(a["label"], b["label"]) >= 0.6)):
                a["vals"].update(b["vals"]); a["restated"].update(b["restated"]); a["files"] |= b["files"]
                if len(b["label"]) < len(a["label"]) and a["kind"] != "heading":
                    a["label"] = b["label"]
                del rows[j]
                break
        else:
            i += 1
    units = list(dict.fromkeys(x["v"]["unit"] for x in views))
    # what each report says, line by line, for the final trace check
    sources = {}
    for x in views:
        src = sources.setdefault(x["file"], {})
        for r in x["v"]["rows"]:
            for c, k in zip(r["vals"], x["keys"]):
                if c:
                    src.setdefault((norm_label(r["label"]), k), []).append(c["v"])
    # printed reports round each line; the allowed rounding gap grows when figures are converted to a larger number
    scale = 1.0
    for f, st in items:
        src = unit_of(st)
        factor = UNIT_FACTORS[src] / UNIT_FACTORS[out_unit] if (out_unit != "reported" and src) else 1.0
        scale = max(scale, factor * 10 ** (2 - st["decimals"]))
    m = {"periods": periods, "rows": rows, "units": units, "files": [x["file"] for x in views],
         "decimals": max(x["v"]["decimals"] for x in views), "sources": sources, "tol_scale": scale}
    return m


def split_pair(a, b):
    """Two rows that look like one line under two names: they agree on every year they share (and it isn't all zero),
    and each has years the other lacks."""
    shared = set(a["vals"]) & set(b["vals"])
    if not shared or not (set(a["vals"]) - set(b["vals"])) or not (set(b["vals"]) - set(a["vals"])):
        return False
    if any(abs(a["vals"][k]["v"] - b["vals"][k]["v"]) > 0.011 for k in shared):
        return False
    return any(abs(a["vals"][k]["v"]) > 1e-9 for k in shared)


def name_similarity(x, y):
    a, b = set(norm_label(x).split()), set(norm_label(y).split())
    return len(a & b) / max(1, len(a | b))


def cell(m, row, i):
    """Value shown for a line and year: the reported figure, or 0 where the report has none."""
    if row["kind"] == "heading":
        return None
    c = row["vals"].get(m["periods"][i]["key"])
    return c["v"] if c else 0.0


KEY_LINES = {"BS": (r"^total assets$|^total$", "total assets"), "PL": (PAT_RE.pattern, "profit after tax"),
             "CF": (r"net cash.*operating", "net cash from operating activities")}


def integrity(m, typ):
    """Final check before anything is published: no repeated lines, every year complete, totals tie in the combined table."""
    issues = []
    heads, seen = ["", ""], {}
    for r in m["rows"]:
        if r["kind"] == "heading":
            heads = [heads[-1], norm_label(r["label"])]
            continue
        section = heads[-1]
        k = (heads[0], heads[1], norm_label(r["label"]))
        if k in seen:
            other = seen[k]
            same_vals = set(other["vals"]) == set(r["vals"]) and all(abs(other["vals"][p]["v"] - r["vals"][p]["v"]) < 0.011 for p in r["vals"])
            if same_vals or not (set(other["vals"]) & set(r["vals"])):
                issues.append(f"\"{r['label']}\" appears twice under \"{section or 'the top'}\"")
        else:
            seen[k] = r
    # every figure must match the same line, for the same year, in the report it was taken from
    for r in m["rows"]:
        for p in m["periods"]:
            c = r["vals"].get(p["key"])
            if not c:
                continue
            said = m.get("sources", {}).get(c["file"], {}).get((norm_label(r["label"]), p["key"]))
            if said is None or not any(abs(v - c["v"]) < 0.011 for v in said):
                issues.append(f"{p['label']} figure for \"{r['label']}\" ({c['v']:g}) doesn't match that line in {c['file']}")
    body = [r for r in m["rows"] if r["kind"] != "heading"]
    for i, a in enumerate(body):
        for b in body[i + 1:i + 4]:
            if split_pair(a, b):
                issues.append(f"\"{a['label']}\" and \"{b['label']}\" look like one line split in two (same figure for the year they share)")
    rx, name = KEY_LINES[typ]
    key_rows = [r for r in m["rows"] if r["kind"] != "heading" and re.search(rx, norm_label(r["label"]))]
    for p in m["periods"]:
        if not any(p["key"] in r["vals"] for r in key_rows):
            issues.append(f"{p['label']} has no {name} figure, so its values can't be placed with confidence")
    # every "Total …" line must still equal the lines above it in each year; a value in the wrong year breaks this
    prev_total = -1
    for t, r in enumerate(m["rows"]):
        if r["kind"] != "total":
            continue
        block = [x for x in m["rows"][prev_total + 1:t] if x["kind"] == "item" and not x["noConvert"]]
        prev_total = t
        if len(block) < 2 or not norm_label(r["label"]).startswith("total"):
            continue
        tol = (0.011 * len(block) + 0.02) * m.get("tol_scale", 1)
        held, broke = [], []
        for i, p in enumerate(m["periods"]):
            if p["key"] not in r["vals"]:
                continue
            diff = sum(cell(m, x, i) for x in block) - r["vals"][p["key"]]["v"]
            (held if abs(diff) <= tol else broke).append((p["label"], diff))
        if held and broke:
            issues.append(f"\"{r['label']}\" no longer equals the lines above it for {', '.join(b[0] for b in broke)}, so a value may be in the wrong year or line")
    pseudo = {"rows": [{"label": r["label"], "hasVals": r["kind"] != "heading",
                        "vals": [r["vals"].get(p["key"]) for p in m["periods"]]} for r in m["rows"]],
              "periods": m["periods"], "decimals": m["decimals"], "tol_scale": m.get("tol_scale", 1)}
    for c in run_checks(pseudo, typ):
        if c["hard"] and not c["ok"]:
            issues.append(f"after combining the reports, {c['name']} is off by {abs(c['diff']):g} for {c['period']}")
    m["issues"] = list(dict.fromkeys(issues))
    m["ok"] = not m["issues"]
    return m


def combine(items, out_unit, typ):
    return integrity(merge_years(items, out_unit), typ)


SHEET_NAMES = {"BS": "Balance Sheet", "PL": "Income Statement", "CF": "Cash Flow Statement"}


def build_workbook(reports, out_unit, basis) -> bytes:
    """Four tabs: Balance Sheet, Income Statement, Cash Flow Statement, Checks.
    Each statement tab holds every year from every report, oldest to latest. A statement is written only if it
    passed verification and the final check; the workbook is then read back and compared figure by figure."""
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)
    bold = Font(bold=True)
    red, green = Font(bold=True, color="FFD70015"), Font(bold=True, color="FF1A7F37")
    top_line = Border(top=Side(style="thin", color="FF1D1D1F"))
    head_fill = PatternFill("solid", fgColor="FFF5F5F7")
    out_label = "as reported" if out_unit == "reported" else UNIT_LABEL[out_unit]

    def fmt(d):
        z = "0." + "0" * d if d > 0 else "0"
        return f"#,##{z};(#,##{z});{z}"

    final, written = {}, {}
    for typ in ("BS", "PL", "CF"):
        ws = wb.create_sheet(SHEET_NAMES[typ])
        items = [(f, s) for f, rep in reports for s in rep["statements"] if s["basis"] == basis and s["type"] == typ and s["verified"]]
        ws["A1"] = f"{basis} {SHEET_NAMES[typ]}"
        ws["A1"].font = Font(bold=True, size=13)
        ws.column_dimensions["A"].width = 64
        if not items:
            ws["A2"] = f"Not included: no {basis.lower()} {SHEET_NAMES[typ].lower()} passed verification. See the Checks tab."
            ws["A2"].font = red
            final[typ] = None
            continue
        m = combine(items, out_unit, typ)
        final[typ] = m
        if not m["ok"]:
            ws["A2"] = "Not included: the final check failed. See the Checks tab."
            ws["A2"].font = red
            continue
        has_note = len(items) == 1
        unit = (UNIT_LABEL[m["units"][0]] if m["units"][0] else "as reported in each report") if len(m["units"]) == 1 else "mixed, choose ₹ crore or ₹ million to align them"
        ws["A2"] = f"Units: {unit} · from {', '.join(m['files'])} · 0 means the report shows no figure for that year"
        ws["A2"].font = Font(italic=True, color="FF6E6E73")
        first = 2 + (1 if has_note else 0)
        hdr = ["Particulars"] + (["Note"] if has_note else []) + [p["label"] for p in m["periods"]]
        for j, h in enumerate(hdr, start=1):
            c = ws.cell(row=4, column=j, value=h)
            c.font, c.fill = bold, head_fill
            c.alignment = Alignment(horizontal="left" if j == 1 else "right")
        r = 5
        for row in m["rows"]:
            ws.cell(row=r, column=1, value=row["label"])
            if has_note and row.get("note"):
                ws.cell(row=r, column=2, value=row["note"]).alignment = Alignment(horizontal="center")
            is_bold = row["kind"] != "item"
            ws.cell(row=r, column=1).font = bold if is_bold else Font()
            if not is_bold:
                ws.cell(row=r, column=1).alignment = Alignment(indent=1)
            if row["kind"] != "heading":
                d = 2 if row["noConvert"] else m["decimals"]
                for i, p in enumerate(m["periods"]):
                    c = ws.cell(row=r, column=first + i, value=cell(m, row, i))
                    c.number_format = fmt(d)
                    if is_bold:
                        c.font = bold
                    if row["kind"] == "total":
                        c.border = top_line
                    rs = row["restated"].get(p["key"])
                    if rs:
                        c.comment = Comment(f"Shown as {rs['v']} in {rs['file']}. The newer report's figure is used here.", "Extractor")
            r += 1
        if has_note:
            ws.column_dimensions["B"].width = 8
        for j in range(first, len(hdr) + 1):
            ws.column_dimensions[get_column_letter(j)].width = 14
        ws.freeze_panes = ws.cell(row=5, column=first)
        written[typ] = (ws, m, first)

    # ---- read the sheets back: every line, in order, with every figure in its year ----
    readback = {}
    for typ, (ws, m, first) in written.items():
        problems, figures = [], 0
        hdr = [ws.cell(row=4, column=first + i).value for i in range(len(m["periods"]))]
        if hdr != [p["label"] for p in m["periods"]]:
            problems.append("year columns are not in the expected order")
        labels = []
        for k, row in enumerate(m["rows"]):
            rr = 5 + k
            if ws.cell(row=rr, column=1).value != row["label"]:
                problems.append(f"row {rr} should be \"{row['label']}\"")
                continue
            labels.append(row["label"])
            if row["kind"] == "heading":
                continue
            for i in range(len(m["periods"])):
                got, want = ws.cell(row=rr, column=first + i).value, cell(m, row, i)
                figures += 1
                if got is None or abs(got - want) > 1e-9:
                    problems.append(f"{row['label']} {m['periods'][i]['label']}: sheet has {got}, expected {want}")
        if ws.cell(row=5 + len(m["rows"]), column=1).value is not None:
            problems.append("extra rows found below the statement")
        readback[typ] = (figures, problems)

    # ---- Checks tab ----
    ck = wb.create_sheet("Checks")
    ck["A1"] = "Checks"
    ck["A1"].font = Font(bold=True, size=13)
    ck["A2"] = f"{basis} statements · output units: {out_label} · per-share figures are never converted · generated {datetime.now():%d %b %Y %H:%M}"
    ck["A2"].font = Font(italic=True, color="FF6E6E73")
    rr = 4

    def heading(text, cols):
        nonlocal rr
        rr += 1
        ck.cell(row=rr, column=1, value=text).font = Font(bold=True, size=12)
        rr += 1
        for j, h in enumerate(cols, start=1):
            c = ck.cell(row=rr, column=j, value=h)
            c.font, c.fill = bold, head_fill
        rr += 1

    def line(values, result_col=None, ok=True):
        nonlocal rr
        for j, v in enumerate(values, start=1):
            ck.cell(row=rr, column=j, value=v)
        if result_col:
            ck.cell(row=rr, column=result_col).font = green if ok else red
        rr += 1

    heading("1. Each tab", ["Tab", "Result", "Detail"])
    for typ in ("BS", "PL", "CF"):
        m = final[typ]
        if m is None:
            line([SHEET_NAMES[typ], "Not included", "no statement passed verification"], 2, False)
        elif not m["ok"]:
            line([SHEET_NAMES[typ], "Not included", "; ".join(m["issues"])], 2, False)
        else:
            figures, problems = readback[typ]
            ok = not problems
            detail = (f"{len(m['periods'])} years ({m['periods'][0]['label']} to {m['periods'][-1]['label']}), {sum(1 for x in m['rows'] if x['kind'] != 'heading')} lines; "
                      "no repeated lines; every figure traced to its report; every year complete; totals and subtotals tie; "
                      f"all {figures} figures read back from this file and matched") if ok else "; ".join(problems[:5])
            line([SHEET_NAMES[typ], "Pass" if ok else "FAIL", detail], 2, ok)

    heading("2. Where each statement came from", ["Report", "Statement", "Taken from", "Unit in report", "Years", "Verified", "Detail"])
    for fname, rep in reports:
        sts = [s for s in rep["statements"] if s["basis"] == basis]
        if not sts:
            line([fname, "—", "", "", "", "No", f"no {basis.lower()} statements found in this report"], 6, False)
        for st in sts:
            u = unit_of(st)
            detail = "; ".join(x["detail"] for x in st["verification"] if (x["ok"] if st["verified"] else not x["ok"]))
            line([fname, SHEET_NAMES[st["type"]], f"{st['where']} (via the {st.get('foundVia', 'page scan')})",
                  (UNIT_LABEL[u] + (" (set by you)" if st.get("unitOverride") else "")) if u else "not found",
                  ", ".join(p["label"] for p in sorted(st["periods"], key=date_key)), "Yes" if st["verified"] else "No",
                  detail if st["verified"] else "left out: " + detail], 6, st["verified"])

    heading("3. Totals in each report, as printed", ["Report", "Statement", "Check", "Year", "Expected", "Extracted", "Difference", "Result"])
    for fname, rep in reports:
        for st in rep["statements"]:
            if st["basis"] != basis:
                continue
            for c in sorted(st["checks"], key=lambda c: (c["name"], c["period"])):
                res = "Pass" if c["ok"] else ("FAIL" if c["hard"] else "Review")
                line([fname, SHEET_NAMES[st["type"]], c["name"], c["period"], c["expected"], c["got"], c["diff"], res], 8, c["ok"])
    for j, w in enumerate([24, 20, 52, 16, 16, 12, 70, 10], start=1):
        ck.column_dimensions[get_column_letter(j)].width = w
    for row in ck.iter_rows(min_row=4):
        for c in row:
            if c.column in (5, 6, 7) and isinstance(c.value, float):
                c.number_format = "#,##0.00;(#,##0.00);0.00"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def render_page_png(data: bytes, pg, width_px=900) -> bytes:
    """Image of the printed page a statement came from (the right half only, on a spread)."""
    doc = pymupdf.open(stream=data, filetype="pdf")
    page = doc[pg["pdfPage"] - 1]
    r = page.rect
    clip = pymupdf.Rect(r.x0 + pg["xOff"], r.y0, r.x0 + pg["xOff"] + pg["W"], r.y1)
    zoom = width_px / pg["W"]
    png = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip).tobytes("png")
    doc.close()
    return png


# =====================================================================================
#  9. Streamlit interface
# =====================================================================================
CSS = """
<style>
#MainMenu, footer, header[data-testid="stHeader"] {visibility: hidden; height: 0;}
.block-container {max-width: 1120px; padding-top: 3rem;}
html, body, [class*="css"] {font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Helvetica Neue", Inter, "Segoe UI", sans-serif;}
.fse-hero {text-align: center; margin: 1.5rem 0 1.25rem;}
.fse-hero h1 {font: 600 clamp(34px, 5.4vw, 52px)/1.06 -apple-system, BlinkMacSystemFont, "SF Pro Display", "Helvetica Neue", Inter, sans-serif;
  letter-spacing: -.025em; margin: 0; padding: 0;}
.fse-hero p {color: #6e6e73; font-size: 19px; margin: .8rem auto 0; max-width: 34em;}
.fse-files {font-size: 14px; color: #6e6e73; display: flex; flex-wrap: wrap; gap: 6px 20px; margin: .25rem 0 1rem;}
.fse-files b {font-weight: 500; color: inherit; filter: contrast(2);}
.fse-status {font-size: 14px; color: #6e6e73; margin: .25rem 0 1rem;}
.fse-status.bad {color: #d70015;}
.fse-card {border-radius: 18px; padding: 8px 22px 14px; background: rgba(127,127,127,.06); overflow-x: auto; margin-bottom: 1.25rem;}
.fse-table {border-collapse: collapse; width: 100%; font-size: 14px; font-variant-numeric: tabular-nums;}
.fse-table th, .fse-table td {padding: 9px 12px; text-align: right; white-space: nowrap; border: 0;}
.fse-table th {font-weight: 500; color: #8e8e93; font-size: 13px; border-bottom: 1px solid rgba(127,127,127,.25);}
.fse-table td:first-child, .fse-table th:first-child {text-align: left; white-space: normal; min-width: 240px; padding-left: 0;}
.fse-table tbody tr + tr td {border-top: 1px solid rgba(127,127,127,.12);}
.fse-table tr.h td {font-weight: 600; padding-top: 18px;}
.fse-table tr.t td {font-weight: 600; border-top: 1px solid rgba(127,127,127,.35);}
.fse-table tr.i td:first-child {padding-left: 14px;}
.fse-table td.y {color: #8e8e93;}
.fse-ok {color: #1a7f37; font-weight: 500;} .fse-no {color: #d70015; font-weight: 500;}
.fse-src {font-size: 14px;}
.fse-why {color: #8e8e93; font-size: 13px; margin: 2px 0 6px;}
div[data-testid="stFileUploader"] {max-width: 560px; margin: 0 auto;}
.stDownloadButton button, div[data-testid="stFileUploader"] button {border-radius: 980px;}
.stDownloadButton button[kind="primary"], button[kind="primary"] {background: #0071e3; border-color: #0071e3; color: #fff;}
[data-testid="stBaseButton-segmented_controlActive"] {background: #0071e3 !important; border-color: #0071e3 !important; color: #fff !important;}
[data-testid="stBaseButton-segmented_controlActive"] p {color: #fff !important;}
[data-testid="stBaseButton-segmented_control"] {border-color: rgba(127,127,127,.3) !important;}
</style>
"""

CODE_VERSION = hashlib.sha1(open(__file__, "rb").read()).hexdigest()[:12]  # changes with every update of this file
TYPE_TABS = {"BS": "Balance Sheet", "PL": "Income Statement", "CF": "Cash Flow Statement"}
UNIT_TABS = {"crore": "₹ Crore", "million": "₹ Million", "reported": "As reported"}


def fmt_num(v, d, unit):
    if v is None:
        return ""
    if abs(v) < 1e-9:
        v = 0.0
    a = abs(v)
    if unit in ("million", "billion", "thousand"):
        s = f"{a:,.{d}f}"
    else:  # Indian grouping: 12,34,567.89
        ip, _, dp = f"{a:.{d}f}".partition(".")
        if len(ip) > 3:
            head, tail = ip[:-3], ip[-3:]
            parts = []
            while len(head) > 2:
                parts.insert(0, head[-2:]); head = head[:-2]
            if head:
                parts.insert(0, head)
            ip = ",".join(parts) + "," + tail
        s = ip + ("." + dp if d else "")
    return f"({s})" if v < 0 else s


def table_html(m, unit, unit_text):
    e = html.escape
    np_ = len(m["periods"])
    h = [f'<table class="fse-table"><thead><tr><th>{e(unit_text)}</th>' + "".join(f"<th>{e(p['label'])}</th>" for p in m["periods"]) + "</tr></thead><tbody>"]
    for r in m["rows"]:
        if r["kind"] == "heading":
            h.append(f'<tr class="h"><td colspan="{np_ + 1}">{e(r["label"])}</td></tr>')
            continue
        d = 2 if r["noConvert"] else m["decimals"]
        h.append(f'<tr class="{"t" if r["kind"] == "total" else "i"}"><td>{e(r["label"])}</td>'
                 + "".join(f"<td>{fmt_num(cell(m, r, i), d, unit)}</td>" for i in range(np_)) + "</tr>")
    h.append("</tbody></table>")
    return "".join(h)


def keep_choice(key):
    """Clicking the option that is already selected would clear it; keep it selected instead."""
    import streamlit as st
    ss = st.session_state
    if ss.get(key) is None:
        ss[key] = ss.get(f"_{key}_last")
    else:
        ss[f"_{key}_last"] = ss[key]


def main():
    import streamlit as st

    st.set_page_config(page_title="Financial Statements Extractor", page_icon="📊", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown('<div class="fse-hero"><h1>Financial Statements Extractor</h1>'
                '<p>Balance sheet, profit and loss, and cash flow from annual reports, year by year, in Excel.</p></div>', unsafe_allow_html=True)

    files = st.file_uploader("Upload annual reports", type=["pdf"], accept_multiple_files=True, label_visibility="collapsed",
                             help="PDF files. Add several years of one company at once.")
    if not files:
        st.markdown('<p style="text-align:center;color:#8e8e93;font-size:13px">PDF files · add several years at once</p>', unsafe_allow_html=True)
        return

    @st.cache_data(show_spinner=False, max_entries=24)
    def cached_analyze(data: bytes, name: str, code_version: str):
        return analyze(data, name)

    ss = st.session_state
    ss.setdefault("fixes", {})
    ss.setdefault("units", {})

    reports, raw = [], {}
    bar = st.progress(0.0, text="Reading reports…")
    for i, f in enumerate(files):
        data = f.getvalue()
        key = hashlib.sha1(data).hexdigest()[:12]
        name = re.sub(r"\.pdf$", "", f.name, flags=re.I)
        bar.progress(i / len(files), text=f"Reading {name}…")
        try:
            rep = cached_analyze(data, f.name, CODE_VERSION)
        except Exception as exc:  # unreadable or encrypted PDF
            st.error(f"{name}: could not read this PDF ({exc}).")
            continue
        rep["key"], rep["name"] = key, name
        raw[key] = data
        for (k, typ, basis), nums in ss.fixes.items():  # pages the person pointed us to
            if k == key:
                rep["statements"] = [s for s in rep["statements"] if not (s["type"] == typ and s["basis"] == basis)]
                rep["statements"].append(build_from_pdf_pages(rep, typ, basis, nums))
        for (k, basis), u in ss.units.items():
            if k == key:
                for s in rep["statements"]:
                    if s["basis"] == basis:
                        s["unitOverride"] = u
        reports.append(rep)
    bar.empty()
    if not reports:
        return

    def latest(rep):
        ds = [date_key(s["periods"][0]) for s in rep["statements"] if s["periods"] and s["periods"][0].get("date")]
        return max(ds) if ds else (rep.get("fy_hint") or 0) * 100

    reports.sort(key=latest, reverse=True)

    has = lambda b: any(s["basis"] == b and s["verified"] for r in reports for s in r["statements"])
    default_basis = "Consolidated" if has("Consolidated") or not has("Standalone") else "Standalone"

    c1, c2, c3 = st.columns([1.1, 1.3, 1], vertical_alignment="center")
    with c1:
        basis = st.segmented_control("Statements", ["Consolidated", "Standalone"], default=default_basis, key="basis", label_visibility="collapsed",
                                     on_change=keep_choice, args=("basis",)) or default_basis
    with c2:
        out_unit = st.segmented_control("Units", list(UNIT_TABS), default="crore", format_func=UNIT_TABS.get, key="unit", label_visibility="collapsed",
                                        on_change=keep_choice, args=("unit",)) or "crore"

    verified_n = sum(1 for r in reports for s in r["statements"] if s["basis"] == basis and s["verified"])
    left_n = sum(1 for r in reports for s in r["statements"] if s["basis"] == basis and not s["verified"])
    with c3:
        if verified_n:
            xlsx = build_workbook([(r["name"], r) for r in reports], out_unit, basis)
            fname = f"Financial_Statements_{basis}_{'as_reported' if out_unit == 'reported' else 'INR_' + out_unit}.xlsx"
            st.download_button("Download Excel", xlsx, file_name=fname, type="primary", use_container_width=True,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.button("Download Excel", disabled=True, use_container_width=True, type="primary")

    files_line = []
    for r in reports:
        s0 = next((s for s in r["statements"] if s["basis"] == basis and s["verified"] and s["periods"]), None)
        files_line.append(f"<span><b>{html.escape(r['name'])}</b> {html.escape(', '.join(p['label'] for p in sorted(s0['periods'], key=date_key))) if s0 else ''}</span>")
    st.markdown(f'<div class="fse-files">{"".join(files_line)}</div>', unsafe_allow_html=True)
    if left_n:
        st.markdown(f'<p class="fse-status bad">{left_n} statement{"s" if left_n > 1 else ""} didn\'t pass verification and '
                    f'{"are" if left_n > 1 else "is"} left out. See Sources below.</p>', unsafe_allow_html=True)
    elif verified_n:
        st.markdown('<p class="fse-status">Every statement was checked against the report.</p>', unsafe_allow_html=True)

    typ = st.segmented_control("Statement", list(TYPE_TABS), default="BS", format_func=TYPE_TABS.get, key="type", label_visibility="collapsed",
                               on_change=keep_choice, args=("type",)) or "BS"

    items = [(r["name"], s) for r in reports for s in r["statements"] if s["basis"] == basis and s["type"] == typ and s["verified"]]
    if not items:
        st.markdown(f'<div class="fse-card"><p style="text-align:center;color:#8e8e93;padding:28px 0">No verified {basis.lower()} {TYPE_TABS[typ].lower()} yet.</p></div>',
                    unsafe_allow_html=True)
    else:
        m = combine(items, out_unit, typ)
        if not m["ok"]:
            st.markdown('<div class="fse-card"><p class="fse-no" style="padding-top:14px">Not published: the combined table failed the final check.</p>'
                        + "".join(f'<div class="fse-why">{html.escape(i)}.</div>' for i in m["issues"]) + "</div>", unsafe_allow_html=True)
        else:
            unit = m["units"][0] if len(m["units"]) == 1 else None
            ut = (UNIT_LABEL[unit] if unit else "as printed") if len(m["units"]) == 1 else "mixed units — choose ₹ Crore or ₹ Million"
            if len(items) == 1:
                vv = view(items[0][1], out_unit)
                ut = vv["unitText"] + (" (" + vv["note"] + ")" if vv["note"] else "")
            st.markdown(f'<div class="fse-card">{table_html(m, unit, ut)}</div>', unsafe_allow_html=True)

    # ---- sources: where each statement came from, and a way to fix it ----
    st.markdown("##### Sources")
    for r in reports:
        s = next((x for x in r["statements"] if x["basis"] == basis and x["type"] == typ), None)
        with st.container(border=True):
            a, b = st.columns([1, 3], vertical_alignment="center")
            a.markdown(f"**{html.escape(r['name'])}**")
            if not s:
                b.markdown(f'<span class="fse-no">No {basis.lower()} {TYPE_TABS[typ].lower()} found</span>', unsafe_allow_html=True)
            elif s["verified"]:
                b.markdown(f'<span class="fse-src"><span class="fse-ok">Verified</span> &nbsp; {html.escape(s["where"])}</span>', unsafe_allow_html=True)
            else:
                why = ". ".join(x["detail"] for x in s["verification"] if not x["ok"])
                b.markdown(f'<span class="fse-src"><span class="fse-no">Not verified, left out</span> &nbsp; {html.escape(s["where"])}</span>'
                           f'<div class="fse-why">{html.escape(why)}.</div>', unsafe_allow_html=True)
            if s and s["verified"] and out_unit != "reported" and not unit_of(s):
                u = st.selectbox("Unit not stated in this report. Choose it:", ["", "crore", "lakh", "million", "thousand", "rupees"],
                                 format_func=lambda x: UNIT_LABEL.get(x, "Choose"), key=f"u_{r['key']}_{basis}")
                if u:
                    ss.units[(r["key"], basis)] = u
                    st.rerun()
            if not s or not s["verified"]:
                f1, f2 = st.columns([3, 1], vertical_alignment="bottom")
                pg_txt = f1.text_input("Right PDF page", placeholder="e.g. 152", key=f"p_{r['key']}_{basis}_{typ}")
                if f2.button("Use", key=f"b_{r['key']}_{basis}_{typ}", use_container_width=True):
                    nums = [int(x) for x in re.findall(r"\d+", pg_txt or "") if 0 < int(x) <= r["pdf_pages"]]
                    if nums:
                        ss.fixes[(r["key"], typ, basis)] = nums
                        st.rerun()
                    else:
                        st.warning(f"That report has {r['pdf_pages']} PDF pages.")
            if s and st.toggle("View page", key=f"v_{r['key']}_{basis}_{typ}"):
                st.image(render_page_png(raw[r["key"]], r["pages"][s["pages"][0] - 1]), width=640)


if os.environ.get("FSE_NO_UI") != "1":
    main()
