"""Shared core: load workbook stints, load engine equity Reg+Growth universe, match names."""
import difflib
import json
import re

import openpyxl

SRC = "equity_fund_manager_history.xlsx"
ENGINE = "scratch/full_master_engine_data.json"

# AMC wording differences between the ACE MF short names (workbook) and AMFI (engine).
# Applied as a leading-prefix rewrite only, so it cannot fire mid-name.
AMC_ALIASES = [
    ("aditya birla sl", "aditya birla sun life"),
    ("icici pru", "icici prudential"),
    ("canara rob", "canara robeco"),
    ("lic mf", "lic"),
    ("invesco india", "invesco"),
    ("parag parikh", "ppfas"),
    ("woc", "whiteoak capital"),
    ("trustmf", "trust"),
    ("franklin india", "franklin"),
    ("templeton india", "franklin"),
    ("jioblackrock", "jio blackrock"),
    ("boi", "bank of india"),
]

STOP = {
    "fund", "funds", "scheme", "schemes", "plan", "plans", "option", "options",
    "growth", "regular", "reg", "direct", "the", "of", "and", "a", "an", "in",
    "mutual", "open", "ended", "payout", "reinvestment", "income", "distribution",
    "cum", "capital", "withdrawal", "idcw", "cumulative",
}

COMPOUNDS = [
    ("largecap", "large cap"), ("midcap", "mid cap"), ("smallcap", "small cap"),
    ("flexicap", "flexi cap"), ("multicap", "multi cap"),
    ("largemidcap", "large mid cap"), ("largemid", "large mid"),
    ("taxsaver", "tax saver"), ("intl", "international"),
    ("infra", "infrastructure"), ("mfg", "manufacturing"),
    ("opp", "opportunities"), ("opportunity", "opportunities"),
    ("serv", "services"), ("fin", "financial"), ("res", "resources"),
    ("teck", "technology"), ("eco", "economic"),
]

# Tokens that define the fund's mandate / SEBI sub-category. If the two names
# disagree on any of these, they are DIFFERENT funds however similar the rest is.
# This is what stops "Large Cap" matching "Large & Mid Cap".
DISCRIMINATIVE = {
    "large", "mid", "small", "cap", "flexi", "multi", "focused", "value", "contra",
    "elss", "tax", "saver", "yield", "dividend", "index", "etf", "quant", "quantamental",
    "esg", "momentum", "quality", "alpha", "volatility", "equal", "weight", "low",
    "sensex", "nifty", "bse", "msci", "nasdaq", "sectoral", "thematic",
    "banking", "financial", "services", "pharma", "pharmaceutical", "healthcare",
    "technology", "digital", "internet", "infrastructure", "manufacturing",
    "consumption", "energy", "power", "transportation", "logistics", "defence",
    "housing", "psu", "mnc", "business", "cycle", "ethical", "resources", "natural",
    "international", "asian", "japan", "china", "us", "global", "emerging",
    "innovation", "innovative", "special", "situations", "opportunities", "rural",
    "commodities", "metal", "auto", "automotive", "realty", "media", "fmcg", "it",
    "bluechip", "vision", "equity", "hybrid", "arbitrage", "balanced", "retirement",
    "children", "gold", "silver", "long", "short", "bharat", "india", "comma",
    "exports", "pioneer", "active", "passive", "factor", "bfsi", "teck", "ex",
}

# Renames / wording the matcher cannot infer safely. workbook name -> engine name.
MANUAL_MAP = {
    "Baroda BNP Paribas ESG Best-in-class Strategy Fund-Reg(G)":
        "Baroda BNP Paribas Best-in-Class Strategy Fund - Regular - Growth Option",
}


def strip_parens(s: str) -> str:
    """Drop parentheticals that carry rename history, e.g. '(erstwhile Bluechip Fund)'."""
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r"\((?:[^()]*\b(?:erstwhile|formerly|previously|known as)\b[^()]*)\)", " ", s, flags=re.I)
    return s


def strip_plan_option(name: str) -> str:
    s = " " + name.strip() + " "
    s = strip_parens(s)
    s = re.sub(r"-\s*(reg|regular|direct|dir)\s*\([^)]*\)", " ", s, flags=re.I)
    s = re.sub(r"\((g|growth|idcw[^)]*|d|bonus)\)", " ", s, flags=re.I)
    s = re.sub(r"-\s*(reg|regular|direct|dir)\b", " ", s, flags=re.I)
    s = re.sub(r"\((regular|direct)\)", " ", s, flags=re.I)
    s = re.sub(r"\b(regular|direct)\s+(plan|option)\b", " ", s, flags=re.I)
    s = re.sub(r"\bgrowth\s+(plan|option)\b", " ", s, flags=re.I)
    s = re.sub(r"\b(plan|option)\b", " ", s, flags=re.I)
    # any remaining parenthetical is a qualifier, not identity: "(30 stocks)", "(60:40)"
    s = re.sub(r"\([^()]*\)", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def canon(name: str, is_engine: bool = False) -> frozenset:
    s = strip_plan_option(name).lower()
    s = s.replace("&", " and ").replace("–", " ").replace("—", " ")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    if is_engine:
        s = re.sub(r"\bmutual fund\b", " ", s)
    for a, b in AMC_ALIASES:
        if s.startswith(a + " ") or s == a:
            s = b + s[len(a):]
            break
    for a, b in COMPOUNDS:
        s = re.sub(rf"\b{a}\b", b, s)
    return frozenset(t for t in s.split() if t not in STOP and len(t) > 1)


def load_stints():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["Sheet1"]
    out = []
    for idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), 1):
        scheme, incep, fm, frm, to = (list(row) + [None] * 5)[:5]
        out.append({"row": idx, "scheme": scheme, "incep": incep, "fm": fm, "from": frm, "to": to})
    return out


def is_data_row(r) -> bool:
    s = r["scheme"]
    if not isinstance(s, str) or s.strip().startswith("Scheme Name:"):
        return False
    if r["fm"] is None or r["from"] is None or r["to"] is None:
        return False
    # the repeated column-header band ('Scheme Name' | 'Inception Date' | ...)
    if s.strip() == "Scheme Name" and r["fm"] == "Fund Manager":
        return False
    # real stints carry real dates; the header band carries strings
    if not hasattr(r["from"], "year") or not hasattr(r["to"], "year"):
        return False
    return len(s) <= 200            # excludes the Accord Fintech disclaimer paragraph


def load_engine():
    with open(ENGINE, encoding="utf-8") as fh:
        d = json.load(fh)
    eq_subs = d["hierarchy"]["Open Ended"]["Equity"]
    funds = {}
    for sub in eq_subs:
        for f in d["open_ended"].get(sub, []):
            if f["plan"] == "Regular" and f["option"] == "Growth" and f["code"] not in funds:
                funds[f["code"]] = dict(f, sub=sub)
    return d, funds


def load_engine_all():
    """Every engine entry, for explaining why a scheme fell out of scope."""
    with open(ENGINE, encoding="utf-8") as fh:
        d = json.load(fh)
    out = []
    for grp in ("open_ended", "close_ended", "interval"):
        for sub, lst in d[grp].items():
            for f in lst:
                out.append(dict(f, grp=grp, sub=sub))
    return out


def build_index(funds):
    exact, cands = {}, []
    for code, f in funds.items():
        c = canon(f["name"], is_engine=True)
        exact.setdefault("".join(sorted(c)), []).append(code)
        cands.append((c, code))
    return exact, cands


def _prefer_regular(codes, funds):
    """Among share-class duplicates prefer the one actually named 'Regular Plan'."""
    for c in codes:
        if re.search(r"regular\s*plan", funds[c]["name"], re.I):
            return c
    return sorted(codes)[0]


def match_scheme(scheme_name, funds, exact, cands):
    """-> (code, score, method). code is None when no safe match exists."""
    if scheme_name in MANUAL_MAP:
        want = MANUAL_MAP[scheme_name]
        for code, f in funds.items():
            if f["name"] == want:
                return code, 1.0, "manual"

    c = canon(scheme_name)
    if not c:
        return None, 0.0, "no-match"

    key = "".join(sorted(c))
    if key in exact:
        codes = exact[key]
        return (_prefer_regular(codes, funds), 1.0,
                "exact" if len(codes) == 1 else f"exact-shareclass({len(codes)})")

    best, best_s = None, 0.0
    for cc, code in cands:
        inter = len(c & cc)
        if not inter:
            continue
        j = inter / len(c | cc)
        if j > best_s:
            best_s, best = j, code

    if best is None or best_s < 0.60:
        return None, round(best_s, 3), "no-match"

    # mandate guard: any disagreement on a mandate-defining token means different funds
    cb = canon(funds[best]["name"], is_engine=True)
    blocking = (c ^ cb) & DISCRIMINATIVE
    if blocking:
        return None, round(best_s, 3), "rejected:" + ",".join(sorted(blocking))

    sm = difflib.SequenceMatcher(None, " ".join(sorted(c)), " ".join(sorted(cb))).ratio()
    if best_s >= 0.75 or sm >= 0.88:
        return best, round(max(best_s, sm), 3), "fuzzy"
    return None, round(best_s, 3), "no-match"


# ===========================================================================
# Shared tenure + active-manager helpers
#
# Used by build_fund_manager_tenure_report.py (the enriched workbook) and by
# build_complete_dashboard_with_recency.py (the dashboard's fund-detail card),
# so both read tenure the same way.
# ===========================================================================

# Passive vehicles: by name, and by how the NAV master classifies them.
PASSIVE_NAME = re.compile(r"\bETF\b|\bIndex Fund\b|\bBeES\b|\bIndex Plan\b|"
                          r"\bFund of Fund\b|\bFoF\b", re.I)
PASSIVE_SUBS = {
    "Index Funds", "Indexs", "Equity ETF", "Other ETF", "Other ETFs", "Debt ETF",
    "Gold ETF", "Silver ETF", "Hybrid ETF", "ETFs investing overseas",
    "FoF Domestic", "FoF Overseas", "Fund of Funds Scheme (Domestic)",
    "Fund ofs Scheme (Domestic)", "Fund of Funds investing overseas",
}


# Share-class wording that survives canon() and would otherwise stop a fund's
# IDCW / Bonus rows collapsing onto its Growth row, e.g.
# "... - IDCW - Payout & Re-investment of Income Distribution ...".
SHARECLASS_NOISE = re.compile(
    r"\b(annual|quarterly|monthly|fortnightly|weekly|daily|half\s*-?\s*yearly|yearly|"
    r"re\s*-?\s*investment|reinvest(?:ment)?|payout|bonus|cumulative|standard|"
    r"institutional|retail)\b", re.I)


def ckey(name: str, is_engine: bool = False) -> str:
    """Canonical key shared by every share class of the same fund.

    canon() already strips plan and option wording; this additionally drops
    share-class noise, so the Regular-Growth, Direct-Growth, IDCW and Bonus rows
    of one fund all collapse to a single key. Applied to both sides of a lookup,
    so it stays symmetric. canon() itself is left alone, because match_scheme()
    is tuned against it.
    """
    return "".join(sorted(canon(SHARECLASS_NOISE.sub(" ", name), is_engine)))


def tenure_months(frm, to) -> int:
    m = (to.year - frm.year) * 12 + (to.month - frm.month)
    if to.day < frm.day:
        m -= 1
    return max(m, 0)


def tenure_text(months: int) -> str:
    y, m = divmod(months, 12)
    return f"{y} {'year' if y == 1 else 'years'} and {m} {'month' if m == 1 else 'months'}"


def build_full_index():
    """Canonical key -> every engine entry under it (all groups, plans, options)."""
    idx = {}
    for f in load_engine_all():
        idx.setdefault(ckey(f["name"], True), []).append(f)
    return idx


def is_passive(scheme: str, full_index) -> bool:
    if PASSIVE_NAME.search(scheme):
        return True
    subs = {h["sub"] for h in full_index.get(ckey(scheme), [])}
    return bool(subs) and subs <= PASSIVE_SUBS


def active_manager_map():
    """-> ({canonical key: [(manager, tenure text, months), ...]}, as-of date).

    Only managers whose stint is still open at the source file's as-of date, and
    only for actively managed schemes - index funds, ETFs and FoFs are left out.
    Each fund's list is ordered longest-serving first.
    """
    data = [r for r in load_stints() if is_data_row(r)]
    if not data:
        return {}, None
    asof = max(r["to"] for r in data)
    full_index = build_full_index()

    best = {}                      # key -> {manager: months}
    for r in data:
        if r["to"] != asof:        # stint already closed -> not an active manager
            continue
        scheme = r["scheme"].strip()
        if is_passive(scheme, full_index):
            continue
        name = str(r["fm"]).strip()
        if not name or name.lower() == "none":
            continue
        k = ckey(scheme)
        months = tenure_months(r["from"], r["to"])
        # a manager listed twice for one fund keeps the longer spell
        cur = best.setdefault(k, {})
        if months > cur.get(name, -1):
            cur[name] = months

    out = {}
    for k, mgrs in best.items():
        ordered = sorted(mgrs.items(), key=lambda kv: (-kv[1], kv[0]))
        out[k] = [(n, tenure_text(m), m) for n, m in ordered]
    return out, asof
