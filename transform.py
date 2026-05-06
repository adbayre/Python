"""
Initial Screen Transformer (label-based, surgical XML edit)
-----------------------------------------------------------
Reads a raw deal screen workbook, finds each field by LABEL (not by cell
position), applies value transformations, and writes the result into a
copy of the target template by editing the underlying XML directly.

Why direct XML?  openpyxl's load-and-save round-trip rewrites embedded
images, macros, calc chains, and printer settings, which causes Excel to
flag the result as corrupt.  This script preserves the template byte-for-
byte and only changes the cell values it needs to change.

Resilient to:
  - rows inserted, deleted, or reordered in the raw file
  - extra/missing blank rows
  - columns shifted left/right
  - which sheet the data sits on (auto-detected by label content)
  - sheet renamed in the raw file

Usage:
    python transform.py <raw_file> [<template_file>] [<output_file>]

Defaults:
    template_file -> input/initial-screen-template.xlsx
    output_file   -> output/<raw_basename> - Formatted.xlsx
"""

import sys
import shutil
import re
import zipfile
import os
from datetime import datetime, date
from xml.sax.saxutils import escape as xml_escape
import openpyxl


# ============================================================
#  TRANSFORMATIONS
# ============================================================

def fmt_amount(v, symbol="$"):
    if v in (None, ""): return ""
    try:
        n = float(v)
        return f"{symbol}{int(n) if n == int(n) else f'{n:g}'}m"
    except (TypeError, ValueError):
        return str(v)


def bucket(value, threshold, low_label, high_label, scale_pct=False):
    if value in (None, ""): return ""
    try:
        v = float(value)
        if scale_pct and v > 1: v = v / 100
        return low_label if v < threshold else high_label
    except (TypeError, ValueError):
        return str(value)


PLACEHOLDER_WORDS = {"name", "name(s)", "names", "date", "country", "value", "y/n", "tbd", "n/a"}

def to_initials(name):
    if not name: return ""
    s = str(name).strip()
    if s.lower() in PLACEHOLDER_WORDS: return ""
    # Take the first letter of every whitespace-separated word, uppercased.
    # Particles like "von" are kept (Urs von Bueren -> UVB).
    return "".join(w[0].upper() for w in s.split() if w and w[0].isalpha())


def parse_date(v):
    """Parse to a datetime, then format as Swiss DD.MM.YYYY string."""
    d = None
    if isinstance(v, datetime):
        d = v
    elif isinstance(v, date):
        d = datetime(v.year, v.month, v.day)
    elif v:
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d.%m.%Y", "%d-%m-%Y"):
            try:
                d = datetime.strptime(str(v).strip(), fmt); break
            except ValueError:
                continue
    if d is None: return "" if v in (None, "") else str(v)
    return d.strftime("%d.%m.%Y")


def yn(v):
    if v in (None, ""): return ""
    s = str(v).strip().upper()
    if s in ("Y", "YES", "TRUE"): return "Y"
    if s in ("N", "NO", "FALSE"): return "N"
    return s


def passthrough(v):
    return "" if v is None else v


# ============================================================
#  FIELD MAP
#  Each entry: target_cell, must (all keywords appear in label),
#              must_not (keywords that disqualify the row),
#              transform fn
# ============================================================

FIELDS = [
    ("gp_name",               "B4",  ["gp name"],                          [],                                        passthrough),
    ("gp_approved",           "C4",  ["gp name"],                          [],                                        yn),
    ("sponsor",                "C5", ["sponsor"],                          ["co-invest"],                              passthrough),
    # Available amount: passthrough here; currency symbol applied in post-processing
    ("available_amount",      "C8",  ["available", "amount"],              [],                                        passthrough),
    ("first_lien",            "C10", ["first lien"],                       [],                                        yn),
    ("ltv",                   "C11", ["ltv"],                              [],                                        lambda v: bucket(v, 0.55, "< 55%", "≥ 55%", scale_pct=True)),
    ("leverage_tranche",      "C12", ["leverage", "tranche"],              ["total"],                                  lambda v: bucket(v, 5.5, "< 5.5x", "> 5.5x")),
    ("prior_debt",            "C13", ["prior", "ranking", "debt"],         [],                                        lambda v: bucket(v, 1.5, "< 1.5x", "> 1.5x")),
    ("ebitda_adj",            "C14", ["ebitda", "adjustment"],             [],                                        lambda v: "N/A" if v in (None, "") else bucket(v, 0.30, "< 30%", "> 30%", scale_pct=True)),
    ("cyclical",              "C15", ["cyclical"],                         [],                                        yn),
    ("esg",                   "C16", ["esg"],                              ["alcohol", "forestry", "palm"],            yn),
    ("pe_co_invest",          "C17", ["pe", "co-invest"],                  [],                                        yn),
    # 3-year historical EBITDA negative — value derived from the financial table on
    # the One-Page Review Memo sheet via historical_ebitda_status() in post-processing.
    ("maint_covenants",       "C19", ["maintenance", "covenant"],          [],                                        yn),
    ("gross_yield",           "C20", ["gross", "asset", "yield"],          [],                                        lambda v: bucket(v, 800, "< 800bps", "> 800bps")),
    ("initial_screen_date",   "C23", ["initial screen"],                   ["analyst", "ic", "decision"],              parse_date),
    ("analyst",               "C24", ["analyst"],                          [],                                        to_initials),
    ("lender_of_record",      "C27", ["lender", "of", "record"],           [],                                        passthrough),
    # Participation sub-fields (column C) — extracted as actuals; will be overridden
    # to "N/A" downstream when the deal is Lender of Record (the common case).
    ("commingled_fund",       "C29", ["(i)", "commingled"],                [],                                        passthrough),
    ("participation_allowed", "C30", ["(ii)", "participation"],            [],                                        passthrough),
    ("max_fund_leverage",     "C31", ["(iii)", "max", "fund"],             [],                                        passthrough),
    ("legal_signoff",         "C32", ["(iv)", "legal"],                    [],                                        passthrough),
    ("ssg_participation",     "C33", ["(v)", "ssg"],                       [],                                        passthrough),
    ("fund_in_investment",    "C34", ["(vi)", "fund", "investment"],       ["known"],                                  passthrough),
    ("borrower_jurisdiction", "C37", ["borrower", "jurisdiction"],         [],                                        passthrough),
    ("on_platform",           "C38", ["already", "on the platform"],       [],                                        yn),
    # Decision is always 'Y': we only ever receive these screens for deals being progressed
    ("decision",              "C40", ["decision", "progress"],             [],                                        lambda v: "Y"),
]

TARGET_SHEET_NAME = "Initial Screen Result"


# ============================================================
#  POLICY CRITERIA — written into column B as constants.
#  These are your team's standard thresholds, not deal-specific.
#  Edit here if policy changes.
# ============================================================

CRITERIA = {
    "B7":  "Desired Value",                     # Column header
    "B10": "Y",                                 # First Lien Y/N
    "B11": "< 55%",                             # LTV
    "B12": "< 5.5x",                            # Leverage through tranche
    "B13": "< 1.5x",                            # Prior Ranking Debt
    "B14": "< 30%",                             # EBITDA adjustments
    "B15": "N",                                 # Cyclical
    "B16": "N",                                 # ESG issues
    "B18": "N",                                 # 3-year historical EBITDA negative
    "B19": "Y",                                 # Maintenance covenants
    "B20": "> 800bps (US), > 600bps (EU)",      # Gross Asset Yield
    "B23": "Date",                              # Initial Screen
    "B24": "Name",                              # Analyst
    "B25": "Name(s)",                           # IC Participants
    "B26": "Y/N",                               # Decision to progress
    "B27": "Preference is Lender of Record",    # Lender
    "B29": "Y",                                 # (i) Commingled Fund
    "B30": "Y",                                 # (ii) Participation allowed by Fund Docs
    "B31": "≤ 2.5x",                            # (iii) Max Fund Leverage Levels
    "B32": "Y and Date of Sign-Off",            # (iv) Legal Sign-Off by Andrew Bratt
    "B33": "< 20%",                             # (v) SSG Participation as % of Fund NAV
    "B34": "Y",                                 # (vi) Fund in investment / re-investment period
    "B37": "Country",                           # Borrower jurisdiction
    "B38": "Y",                                 # Already on the platform
    "B40": "Y/N",                               # Decision: Progress to Initial Review
}


# ============================================================
#  EUROPEAN JURISDICTION CHECK
# ============================================================
EUROPEAN_COUNTRIES = {
    "austria", "belgium", "bulgaria", "croatia", "cyprus", "czech republic", "czechia",
    "denmark", "estonia", "finland", "france", "germany", "greece", "hungary",
    "iceland", "ireland", "italy", "latvia", "liechtenstein", "lithuania",
    "luxembourg", "malta", "netherlands", "norway", "poland", "portugal", "romania",
    "slovakia", "slovenia", "spain", "sweden", "switzerland", "united kingdom",
    "uk", "u.k.", "england", "scotland", "wales", "northern ireland",
}

def is_european_country(value):
    if not value: return False
    s = str(value).strip().lower()
    return s in EUROPEAN_COUNTRIES or any(c in s for c in EUROPEAN_COUNTRIES)


# ============================================================
#  CURRENCY DETECTION
# ============================================================
# (code, symbol) keyed by lowercase country / region match.
CURRENCY_BY_COUNTRY = {
    "united states": ("USD", "$"), "us": ("USD", "$"), "u.s.": ("USD", "$"), "usa": ("USD", "$"),
    "canada": ("CAD", "C$"),
    "united kingdom": ("GBP", "£"), "uk": ("GBP", "£"), "u.k.": ("GBP", "£"),
    "england": ("GBP", "£"), "scotland": ("GBP", "£"), "wales": ("GBP", "£"),
    "switzerland": ("CHF", "CHF"),
    "australia": ("AUD", "A$"),
    "japan": ("JPY", "¥"),
    "china": ("CNY", "¥"),
    "india": ("INR", "₹"),
}
# Eurozone members → EUR
for c in ("austria", "belgium", "croatia", "cyprus", "estonia", "finland",
          "france", "germany", "greece", "ireland", "italy", "latvia",
          "lithuania", "luxembourg", "malta", "netherlands", "portugal",
          "slovakia", "slovenia", "spain"):
    CURRENCY_BY_COUNTRY[c] = ("EUR", "€")
# Non-euro EU / EEA
CURRENCY_BY_COUNTRY.update({
    "denmark": ("DKK", "kr"), "norway": ("NOK", "kr"), "sweden": ("SEK", "kr"),
    "iceland": ("ISK", "kr"), "poland": ("PLN", "zł"), "czech republic": ("CZK", "Kč"),
    "czechia": ("CZK", "Kč"), "hungary": ("HUF", "Ft"), "romania": ("RON", "lei"),
    "bulgaria": ("BGN", "лв"),
})

CURRENCY_FROM_SYMBOL = {
    "€": ("EUR", "€"), "£": ("GBP", "£"), "¥": ("JPY", "¥"), "₹": ("INR", "₹"),
    "chf": ("CHF", "CHF"), "usd": ("USD", "$"), "eur": ("EUR", "€"),
    "gbp": ("GBP", "£"), "cad": ("CAD", "C$"), "aud": ("AUD", "A$"),
    "$": ("USD", "$"),  # listed last so € / £ / ¥ win first
}

def detect_currency(raw_ws, jurisdiction):
    """Return (code, symbol). Strategy:
       1. Look at the 'Available Amount' label for an explicit currency hint.
       2. Fall back to borrower jurisdiction.
       3. Default to USD.
    """
    # 1. Inspect the available-amount label across the data sheet
    for row in raw_ws.iter_rows(values_only=True):
        for v in row:
            if not v: continue
            s = str(v).lower()
            if "available" in s and "amount" in s:
                # Check explicit codes first (chf/usd/eur), then symbols.
                for token in ("chf", "usd", "eur", "gbp", "cad", "aud", "€", "£", "¥", "₹", "$"):
                    if token in s:
                        return CURRENCY_FROM_SYMBOL[token]
    # 2. Jurisdiction-based inference
    if jurisdiction:
        j = str(jurisdiction).strip().lower()
        if j in CURRENCY_BY_COUNTRY:
            return CURRENCY_BY_COUNTRY[j]
        for country, cur in CURRENCY_BY_COUNTRY.items():
            if country in j:
                return cur
    # 3. Final fallback
    return ("USD", "$")


# ============================================================
#  RAW SHEET PARSING
# ============================================================

def find_data_sheet(wb, fields):
    """Score each sheet by how many target labels it contains; pick the highest."""
    best_score, best_ws = -1, None
    for name in wb.sheetnames:
        ws = wb[name]
        joined = " | ".join(
            str(v).lower()
            for row in ws.iter_rows(values_only=True)
            for v in row if v is not None
        )
        score = sum(
            1 for _, _, must, must_not, _ in fields
            if all(w in joined for w in must) and not any(w in joined for w in must_not)
        )
        if score > best_score:
            best_score, best_ws = score, ws
    return best_ws or wb[wb.sheetnames[0]]


def label_matches(text, must, must_not):
    if not text: return False
    s = str(text).lower()
    return all(w in s for w in must) and not any(w in s for w in must_not)


def find_value_for_label(ws, must, must_not, max_gap=3):
    """Return (label, criterion, actual). Stops scanning right after `max_gap` empties."""
    candidates = []
    for row in ws.iter_rows():
        for cell in row:
            if not label_matches(cell.value, must, must_not): continue
            values, gap = [], 0
            for c in row:
                if c.column <= cell.column: continue
                if c.value not in (None, ""):
                    values.append(c.value); gap = 0
                else:
                    gap += 1
                    if gap >= max_gap and values: break
            candidates.append((cell, values))
    if not candidates: return (None, None, None)
    candidates.sort(key=lambda x: (0 if x[1] else 1, x[0].column, x[0].row))
    cell, values = candidates[0]
    if len(values) >= 2: return (cell.value, values[0], values[1])
    if len(values) == 1: return (cell.value, None, values[0])
    return (cell.value, None, None)


def historical_ebitda_status(wb):
    """Read historical EBITDA values from the One-Page Review Memo (or any sheet)
    and classify the 3-year-negative flag:
      - any historical year < 0  -> 'Y' (yes, negative at some point)
      - all historical years >= 0 and at least one populated -> 'N'
      - nothing populated / no table -> 'N/A'
    """
    for sname in wb.sheetnames:
        ws = wb[sname]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not v: continue
                s = str(v).lower()
                if "historical" not in s or "pro forma" not in s: continue

                # Find 'Pro Forma' header on the same row to bound the historical block
                pf_col = None
                for c2 in ws[cell.row]:
                    if c2.column > cell.column and c2.value and str(c2.value).strip().lower() == "pro forma":
                        pf_col = c2.column; break

                # Year columns sit one row below
                year_row_idx = cell.row + 1
                hist_cols = []
                for yc in ws[year_row_idx]:
                    if yc.column < cell.column: continue
                    if pf_col and yc.column >= pf_col: break
                    yv = yc.value
                    try:
                        y = int(yv) if isinstance(yv, (int, float)) else int(str(yv).strip())
                        if 1990 <= y <= 2100: hist_cols.append(yc.column)
                    except (TypeError, ValueError):
                        pass
                if not hist_cols: continue

                # 'Adj. EBITDA' row sits within the next ~15 rows
                for r in range(year_row_idx + 1, year_row_idx + 16):
                    label_cell = ws.cell(r, 2).value  # column B holds the row labels
                    if not label_cell: continue
                    ls = str(label_cell).lower()
                    if "ebitda" not in ls: continue
                    if "%" in ls or "margin" in ls: continue  # skip "% Margin" sub-row
                    vals = [ws.cell(r, col).value for col in hist_cols]
                    nums = []
                    for v in vals:
                        if v in (None, ""): continue
                        try: nums.append(float(v))
                        except (TypeError, ValueError): continue
                    if not nums: return "N/A"
                    return "Y" if any(n < 0 for n in nums) else "N"
    return "N/A"


def find_ic_participants(wb):
    """Pull IC participants from anywhere in the workbook.

    Strategy: find every cell whose label contains 'IC Participants', look at the
    column immediately to the right (and the same column below as fallback), collect
    rows that look like person names, and pick the match that yields the most names.
    This way 'PD Co-Investment IC Participants' (a placeholder header on Initial
    Screen_NEW) loses to 'IC Participants' on Lookup Lists, which has the actual roster.
    """
    SKIP = {"deal team", "tbd", "team", "name", "name(s)", "names", "y/n",
            "date", "country", "value", "y", "n"}

    def looks_like_name(s):
        if not s: return False
        if s.lower() in SKIP: return False
        if not any(c.isalpha() for c in s): return False
        # Reject obviously non-name strings (numbers only, very short, contain ?).
        if "?" in s or len(s) < 3: return False
        return True

    best = []
    for sname in wb.sheetnames:
        ws = wb[sname]
        for row in ws.iter_rows():
            for cell in row:
                if not cell.value: continue
                if "ic participants" not in str(cell.value).lower(): continue
                # Try column to the right first, then same column.
                for col_offset in (1, 0):
                    names = []
                    for r in range(cell.row + 1, cell.row + 30):
                        v = ws.cell(r, cell.column + col_offset).value
                        if v is None or str(v).strip() == "":
                            if names: break  # End of contiguous block
                            else: continue
                        s = str(v).strip()
                        if looks_like_name(s):
                            names.append(s)
                        elif names:
                            break
                    if len(names) > len(best):
                        best = names

    if not best: return ""
    initials = [to_initials(n) for n in best]
    return ", ".join(i for i in initials if i)


def detect_project_name(wb, primary_ws):
    """Find the deal project name. Strategy:
       1. Look across ALL sheets for any cell starting with 'Project ' (the team's naming convention).
       2. Fall back: a standalone cell on the data sheet that isn't a known label.
    """
    # Strategy 1: explicit "Project X" cells anywhere in the workbook
    candidates = []
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if not v: continue
                s = str(v).strip()
                if s.lower().startswith("project ") and 9 <= len(s) <= 60:
                    candidates.append(s)
    if candidates:
        # Pick the most common (handles a project name appearing on several sheets)
        return max(set(candidates), key=candidates.count)

    # Strategy 2: standalone-cell heuristic on the data sheet
    GENERIC = {"initial screen", "date", "name", "value", "y/n", "y", "n",
               "country", "added", "performing", "yes", "no"}
    for row in primary_ws.iter_rows():
        non_empty = [c for c in row if c.value not in (None, "")]
        if len(non_empty) != 1: continue
        s = str(non_empty[0].value).strip()
        if not (3 <= len(s) <= 60): continue
        if s.lower() in GENERIC: continue
        if s.lower().startswith("if "): continue
        if any(label_matches(s, must, must_not) for _, _, must, must_not, _ in FIELDS): continue
        return s
    return None


# ============================================================
#  EXCEL DATE SERIAL
# ============================================================

EXCEL_EPOCH = datetime(1899, 12, 30)  # Excel's day-zero (handles its leap-year quirk)

def to_excel_serial(d):
    if isinstance(d, date) and not isinstance(d, datetime):
        d = datetime(d.year, d.month, d.day)
    delta = d - EXCEL_EPOCH
    return delta.days + delta.seconds / 86400


# ============================================================
#  SURGICAL XML EDITING
# ============================================================
# Cell encodings in OOXML:
#   <c r="C8" s="42" t="s"><v>123</v></c>          shared string (index 123)
#   <c r="C8" s="42" t="inlineStr"><is><t>X</t></is></c>   inline string
#   <c r="C8" s="42"><v>40</v></c>                  number / date serial
#   <c r="C8" s="42"/>                              empty
#
# We replace each whole <c r="REF" .../> element with a new one carrying the
# same style (s="...") but a new value.  We keep all other XML untouched.

CELL_RE = re.compile(r'<c\s+r="(?P<ref>[A-Z]+\d+)"(?P<attrs>[^/>]*)(?:/>|>(?P<inner>.*?)</c>)', re.DOTALL)

def get_cell_style(sheet_xml, ref):
    """Find the existing `s="N"` attribute for a cell, so we keep its formatting."""
    m = re.search(r'<c\s+r="' + re.escape(ref) + r'"([^/>]*)', sheet_xml)
    if not m: return None
    s = re.search(r's="(\d+)"', m.group(1))
    return s.group(1) if s else None


def build_cell_element(ref, value, style):
    """Build a new <c> element. Inline string for text, raw number for date/numeric."""
    style_attr = f' s="{style}"' if style else ""

    if value in (None, ""):
        return f'<c r="{ref}"{style_attr}/>'

    if isinstance(value, datetime):
        serial = to_excel_serial(value)
        return f'<c r="{ref}"{style_attr}><v>{serial:g}</v></c>'

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'

    text = xml_escape(str(value))
    return f'<c r="{ref}"{style_attr} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def replace_cell_in_xml(sheet_xml, ref, new_element):
    """Replace either <c r="REF" .../> (self-closing) or <c r="REF" ...>...</c>."""
    pattern_open  = re.compile(r'<c\s+r="' + re.escape(ref) + r'"[^/>]*>.*?</c>', re.DOTALL)
    pattern_close = re.compile(r'<c\s+r="' + re.escape(ref) + r'"[^/>]*/>')

    if pattern_open.search(sheet_xml):
        return pattern_open.sub(lambda m: new_element, sheet_xml, count=1)
    if pattern_close.search(sheet_xml):
        return pattern_close.sub(lambda m: new_element, sheet_xml, count=1)

    # Cell doesn't exist yet — insert it into the matching <row>.
    row_num = re.match(r'[A-Z]+(\d+)', ref).group(1)
    row_re = re.compile(r'(<row\s+r="' + row_num + r'"[^>]*>)', re.DOTALL)
    if row_re.search(sheet_xml):
        return row_re.sub(r'\1' + new_element, sheet_xml, count=1)

    return sheet_xml  # row missing — give up silently


def find_target_sheet_xml_path(zf):
    """Look up which worksheets/sheetN.xml corresponds to TARGET_SHEET_NAME."""
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    rels   = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    rid = None
    for m in re.finditer(r'<sheet\s[^>]*?/>', wb_xml):
        tag = m.group(0)
        name_m = re.search(r'\sname="([^"]+)"', tag)
        rid_m  = re.search(r'\sr:id="([^"]+)"', tag)
        if name_m and rid_m and name_m.group(1) == TARGET_SHEET_NAME:
            rid = rid_m.group(1); break
    if not rid:
        raise RuntimeError(f"Sheet '{TARGET_SHEET_NAME}' not found in template")

    # Each <Relationship> tag has attrs in arbitrary order. Don't use [^/>] —
    # the Type URL contains slashes. Match up to the closing /> only.
    target = None
    for m in re.finditer(r'<Relationship\s[^>]*?/>', rels):
        tag = m.group(0)
        if f'Id="{rid}"' in tag:
            t = re.search(r'Target="([^"]+)"', tag)
            if t:
                target = t.group(1); break
    if not target:
        raise RuntimeError(f"Could not resolve relationship {rid}")

    # Targets can be absolute ('/xl/worksheets/sheet1.xml') or relative
    # ('worksheets/sheet1.xml'). Normalize to the zip-internal path.
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target


def rewrite_zip(input_zip_path, output_zip_path, replacements):
    """Copy the zip byte-for-byte except for the parts named in `replacements` (path -> bytes)."""
    with zipfile.ZipFile(input_zip_path, "r") as zin:
        with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = replacements[item.filename] if item.filename in replacements else zin.read(item.filename)
                zout.writestr(item, data)


# ============================================================
#  MAIN
# ============================================================

def transform(raw_path, template_path, output_path):
    # 1. Read values from raw file
    raw_wb = openpyxl.load_workbook(raw_path, data_only=True, keep_vba=False, read_only=False)
    raw_ws = find_data_sheet(raw_wb, FIELDS)

    extracted = {}
    log = []

    project_name = detect_project_name(raw_wb, raw_ws)
    if project_name:
        # Keep the full name; only prepend "Project " if missing.
        full = project_name if project_name.lower().startswith("project ") else f"Project {project_name}"
        extracted["A2"] = full
        log.append(("project_name", "A2", project_name, full))

    raw_jurisdiction = None
    raw_amount = None
    for key, target_cell, must, must_not, fn in FIELDS:
        label, criterion, actual = find_value_for_label(raw_ws, must, must_not)
        raw_val = criterion if key == "gp_approved" else actual
        if key == "borrower_jurisdiction": raw_jurisdiction = raw_val
        if key == "available_amount":      raw_amount = raw_val
        new_val = fn(raw_val)
        extracted[target_cell] = new_val
        log.append((key, target_cell, raw_val, new_val))

    # IC Participants: pulled from the Lookup Lists sheet (cross-sheet lookup)
    ic_initials = find_ic_participants(raw_wb)
    if ic_initials:
        extracted["C25"] = ic_initials
        log.append(("ic_participants", "C25", "(Lookup Lists)", ic_initials))

    # Legacy decision row (C26) — also always 'Y', mirroring C40
    extracted["C26"] = "Y"
    log.append(("decision (legacy)", "C26", "(rule)", "Y"))

    # 3-year historical EBITDA negative — derived from the One-Page Review Memo
    hist_status = historical_ebitda_status(raw_wb)
    extracted["C18"] = hist_status
    log.append(("hist_ebitda_neg", "C18", "(One-Page Memo)", hist_status))

    # Currency: derived from the amount label or jurisdiction
    cur_code, cur_symbol = detect_currency(raw_ws, raw_jurisdiction)
    extracted["C8"] = fmt_amount(raw_amount, cur_symbol)
    extracted["C9"] = cur_code
    log.append(("currency", "C9", raw_jurisdiction, cur_code))
    log.append(("amount (currency)", "C8", raw_amount, extracted["C8"]))

    # Deal-type rules:
    #  - Lender of Record deals: the six Participation sub-fields don't apply → N/A
    #  - Non-European deals: borrower jurisdiction + already-on-platform don't apply → N/A
    lor_value = str(extracted.get("C27") or "").lower()
    is_lor = "lender of record" in lor_value
    if is_lor:
        for ref in ("C29", "C30", "C31", "C32", "C33", "C34"):
            extracted[ref] = "N/A"
            log.append(("(N/A: LoR)", ref, "(rule)", "N/A"))

    jurisdiction_value = extracted.get("C37")
    if not is_european_country(jurisdiction_value):
        extracted["C37"] = "N/A"
        extracted["C38"] = "N/A"
        log.append(("(N/A: non-EU)", "C37", jurisdiction_value, "N/A"))
        log.append(("(N/A: non-EU)", "C38", "(rule)", "N/A"))

    # Add the constant policy criteria into column B
    for ref, text in CRITERIA.items():
        extracted[ref] = text
        log.append(("(criterion)", ref, "(policy)", text))

    # 2. Open template, get sheet XML, apply replacements, write output
    with zipfile.ZipFile(template_path, "r") as zf:
        sheet_xml_path = find_target_sheet_xml_path(zf)
        sheet_xml = zf.read(sheet_xml_path).decode("utf-8")

    for ref, new_value in extracted.items():
        style = get_cell_style(sheet_xml, ref)
        new_element = build_cell_element(ref, new_value, style)
        sheet_xml = replace_cell_in_xml(sheet_xml, ref, new_element)

    rewrite_zip(template_path, output_path, {sheet_xml_path: sheet_xml.encode("utf-8")})
    return log


DEFAULT_TEMPLATE = "input/initial-screen-template.xlsx"

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python transform.py <raw> [<template>] [<output>]")
        sys.exit(1)

    raw = sys.argv[1]
    template = sys.argv[2] if len(sys.argv) >= 3 else DEFAULT_TEMPLATE

    if len(sys.argv) >= 4:
        out = sys.argv[3]
    else:
        os.makedirs("output", exist_ok=True)
        # Build the target name: "<YYYYMM> Co-Investment <Project>.xlsx"
        # YYYYMM comes from the initial screen date; project from A2 ("Project Foo" -> "Foo").
        peek_wb = openpyxl.load_workbook(raw, data_only=True, read_only=False)
        peek_ws = find_data_sheet(peek_wb, FIELDS)
        _, _, screen_actual = find_value_for_label(
            peek_ws, ["initial screen"], ["analyst", "ic", "decision"]
        )
        yyyymm = ""
        if isinstance(screen_actual, datetime):
            yyyymm = screen_actual.strftime("%Y%m")
        elif isinstance(screen_actual, date):
            yyyymm = screen_actual.strftime("%Y%m")
        elif screen_actual:
            for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d.%m.%Y", "%d-%m-%Y"):
                try:
                    yyyymm = datetime.strptime(str(screen_actual).strip(), fmt).strftime("%Y%m"); break
                except ValueError:
                    continue
        proj = detect_project_name(peek_wb, peek_ws) or ""
        if proj.lower().startswith("project "):
            proj = proj.split(None, 1)[1]
        proj = proj or "Project"

        parts = []
        if yyyymm: parts.append(yyyymm)
        parts.append(f"Co-Investment {proj}")
        out = f"output/{' '.join(parts)}.xlsx"

    if not os.path.exists(template):
        print(f"Template not found: {template}")
        sys.exit(1)

    log = transform(raw, template, out)
    print(f"Wrote {out}\n")
    print(f"{'FIELD':<22} {'TARGET':<7} {'RAW':<35} {'WROTE':<25}")
    print("-" * 95)
    for key, tc, raw_val, new in log:
        rv = (str(raw_val)[:33] + "..") if raw_val is not None and len(str(raw_val)) > 33 else (str(raw_val) if raw_val is not None else "")
        nv = str(new)[:23]
        print(f"{key:<22} {tc:<7} {rv:<35} {nv:<25}")
