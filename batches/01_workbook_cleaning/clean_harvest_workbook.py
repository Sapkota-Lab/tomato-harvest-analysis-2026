"""Standardize + organize Harvest 2026 workbook; write CLEANED copy.

Original file is never overwritten.

Standardizes:
  - Genotype / Sample → UKYT-YY-NNN + light commercial/NC aliases
  - Location → High tunnel | Open field | Robinson | South Farm
  - Harvest → First … Seventh
  - Date of harvest → assigned by recorded values + reasoned cadence
  - Sheet names/order → HT_ / SF_ / RC_ / Lab_ prefixes
"""

from __future__ import annotations

import os
from collections import Counter
from datetime import datetime, date
from pathlib import Path
import re

from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import Selection

BATCH_DIR = Path(__file__).resolve().parent
ROOT = BATCH_DIR.parents[1]
os.chdir(BATCH_DIR)

SRC = ROOT / "data/raw/Harvest 2026 Sapkota lab.xlsx"
OUT = ROOT / "data/processed/Harvest 2026 Sapkota lab CLEANED.xlsx"

# Sheet → default Location / Harvest when cells blank or invalid
SHEET_DEFAULTS = {
    "First": ("High tunnel", "First"),
    "Second": ("High tunnel", "Second"),
    "Third": ("High tunnel", "Third"),
    "Fourth": ("High tunnel", "Fourth"),
    "Fifth": ("High tunnel", "Fifth"),
    "High Tunnel Harvest Sixth": ("High tunnel", "Sixth"),
    "High_tunnel_Aug19_Nam": ("High tunnel", "Seventh"),
    # Lex South Farm = open-field site; use one location name everywhere
    "FirstHArvest_OpenFieldLex": ("South Farm", "First"),
    "SecondHarvestOpenField": ("South Farm", "Second"),
    "Second_Lex_Aug18_Nam": ("South Farm", "Third"),
    "Second_Lex_Aug18_Ella": ("South Farm", "Third"),
    "Second_lex_Aug19_Nam": ("South Farm", "Fourth"),
    "RobinsonCenter": ("Robinson", "First"),
    "hybrid_harvest_Lex_MS": ("South Farm", None),
    "Specific Gravity": (None, None),
}

# Date policy: (default_date | None, mode, reason)
# mode "force" = set every data row; "fill" = blanks only; "keep" = forward-fill only
DATE_POLICY = {
    "First": (
        datetime(2026, 7, 13),
        "force",
        "Recorded on sheet for all rows",
    ),
    "Second": (
        datetime(2026, 7, 20),
        "force",
        "Recorded on sheet for all rows",
    ),
    "Third": (
        datetime(2026, 7, 27),
        "fill",
        "Two-day HT harvest (7/27–7/28 recorded); blanks filled with 7/27",
    ),
    "Fourth": (
        datetime(2026, 8, 3),
        "force",
        "Inferred: weekly HT cadence after Third (Mon 7/27 → Mon 8/3)",
    ),
    "Fifth": (
        datetime(2026, 8, 10),
        "force",
        "Inferred: weekly HT cadence; SG HT samples cluster on 8/11",
    ),
    "High Tunnel Harvest Sixth": (
        datetime(2026, 8, 18),
        "force",
        "Recorded on sheet (first row); applied to all rows",
    ),
    "High_tunnel_Aug19_Nam": (
        datetime(2026, 8, 19),
        "force",
        "Recorded on sheet; seventh HT harvest",
    ),
    "FirstHArvest_OpenFieldLex": (
        datetime(2026, 7, 23),
        "force",
        "Recorded on first row; South Farm first harvest",
    ),
    "SecondHarvestOpenField": (
        None,
        "keep",
        "South Farm second harvest; mixed recorded dates 8/7 and 8/10",
    ),
    "Second_Lex_Aug18_Nam": (
        datetime(2026, 8, 18),
        "force",
        "South Farm third harvest; data entered by Nam",
    ),
    "Second_Lex_Aug18_Ella": (
        datetime(2026, 8, 18),
        "force",
        "South Farm third harvest; data entered by Ella (same day as Nam sheet)",
    ),
    "Second_lex_Aug19_Nam": (
        datetime(2026, 8, 19),
        "force",
        "South Farm fourth harvest; data entered by Nam",
    ),
    "RobinsonCenter": (
        datetime(2026, 8, 12),
        "force",
        "SG Robinson first-harvest assays all dated 2026-08-12",
    ),
    "hybrid_harvest_Lex_MS": (
        None,
        "keep",
        "Hybrid evaluation block; no harvest date recorded",
    ),
}

# Final sheet order and names (Excel max 31 chars). Keys = original names.
SHEET_RENAME_ORDER = [
    ("First", "HT_01_First_2026-07-13"),
    ("Second", "HT_02_Second_2026-07-20"),
    ("Third", "HT_03_Third_2026-07-27"),
    ("Fourth", "HT_04_Fourth_2026-08-03"),
    ("Fifth", "HT_05_Fifth_2026-08-10"),
    ("High Tunnel Harvest Sixth", "HT_06_Sixth_2026-08-18"),
    ("High_tunnel_Aug19_Nam", "HT_07_Seventh_2026-08-19_Nam"),
    ("FirstHArvest_OpenFieldLex", "SF_01_First_2026-07-23"),
    ("SecondHarvestOpenField", "SF_02_Second_2026-08-07"),
    ("Second_Lex_Aug18_Nam", "SF_03_Third_2026-08-18_Nam"),
    ("Second_Lex_Aug18_Ella", "SF_03_Third_2026-08-18_Ella"),
    ("Second_lex_Aug19_Nam", "SF_04_Fourth_2026-08-19_Nam"),
    ("hybrid_harvest_Lex_MS", "SF_Hybrid"),
    ("RobinsonCenter", "RC_01_First_2026-08-12"),
    ("Specific Gravity", "Lab_Specific_Gravity"),
    ("Carotenoid", "Lab_Carotenoid"),
    ("TA", "Lab_TA"),
    ("Anthocyanin", "Lab_Anthocyanin"),
    ("GABA", "Lab_GABA"),
    ("Polyphenol", "Lab_Polyphenol"),
]

COMMERCIAL_ALIASES = {
    "mt. fresh": "Mountain Fresh",
    "mt fresh": "Mountain Fresh",
    "m. fresh": "Mountain Fresh",
    "mountain fresh": "Mountain Fresh",
    "mt. fresh plus": "Mountain Fresh Plus",
    "mt fresh plus": "Mountain Fresh Plus",
    "mountain fresh plus": "Mountain Fresh Plus",
    "mountain freshplus": "Mountain Fresh Plus",
    "mt. gem": "Mt. Gem",
    "mt gem": "Mt. Gem",
    "mt.gem": "Mt. Gem",
    "mountain gem": "Mt. Gem",
    "mountaingem": "Mt. Gem",
    "mt. girl": "Mt. Girl",
    "mt girl": "Mt. Girl",
    "mountain girl": "Mt. Girl",
    "mountain girl -1": "Mt. Girl",
    "mt. heritage": "Mt. Heritage",
    "mt heritage": "Mt. Heritage",
    "mt.heritage": "Mt. Heritage",
    "mountain heritage": "Mt. Heritage",
    "mountain heritage -1": "Mt. Heritage",
    "mt. magic": "Mt. Magic",
    "mt magic": "Mt. Magic",
    "mountain magic": "Mt. Magic",
    "cherokee purple": "Cherokee Purple",
    "maglia r cherry": "Maglia Rosa Cherry",
    "maglia rosa ch": "Maglia Rosa Cherry",
    "maglia rosa cherry": "Maglia Rosa Cherry",
    "magnolia rosa cherry": "Maglia Rosa Cherry",
    "peace vine": "Peace Vine",
    "peace vine cherry": "Peace Vine",
    "peacevine": "Peace Vine",
    "peacevine cherry": "Peace Vine",
    "peacevine ch.": "Peace Vine",
    "peacevine ch": "Peace Vine",
    "peachvine ch.": "Peace Vine",
    "peachvine ch": "Peace Vine",
    "bhn589": "BHN589",
    "bhn 589": "BHN589",
}

LOCATION_MAP = {
    r"(?i)^high\s*tunnel$": "High tunnel",
    r"(?i)^hightunnel$": "High tunnel",
    r"(?i)^open\s*field$": "Open field",
    r"(?i)^south\s*farm$": "South Farm",
    r"(?i)^robinson$": "Robinson",
    r"(?i)^robinson\s*center$": "Robinson",
}

HARVEST_MAP = {
    r"(?i)^first(\s*harvest)?$": "First",
    r"(?i)^second(\s*harvest)?$": "Second",
    r"(?i)^third(\s*harvest)?$": "Third",
    r"(?i)^fourth(\s*harvest)?$": "Fourth",
    r"(?i)^fifth(\s*harvest)?$": "Fifth",
    r"(?i)^sixth(\s*harvest)?$": "Sixth",
    r"(?i)^seventh(\s*harvest)?$": "Seventh",
    r"(?i)^aug\.?\s*19$": "Seventh",
}

VALID_LOCATIONS = {"High tunnel", "Open field", "Robinson", "South Farm"}
VALID_HARVESTS = {
    "First",
    "Second",
    "Third",
    "Fourth",
    "Fifth",
    "Sixth",
    "Seventh",
}

# These sheets are one harvest event — always set Harvest / Location from the sheet name.
FORCE_HARVEST_SHEETS = {
    "First",
    "Second",
    "Third",
    "Fourth",
    "Fifth",
    "High Tunnel Harvest Sixth",
    "High_tunnel_Aug19_Nam",
    "FirstHArvest_OpenFieldLex",
    "SecondHarvestOpenField",
    "Second_Lex_Aug18_Nam",
    "Second_Lex_Aug18_Ella",
    "Second_lex_Aug19_Nam",
    "RobinsonCenter",
}
FORCE_LOCATION_SHEETS = set(FORCE_HARVEST_SHEETS)

# Original sheet → who entered the data (lab members), when known from sheet name
ENTERED_BY = {
    "Second_Lex_Aug18_Nam": "Nam",
    "Second_lex_Aug19_Nam": "Nam",
    "High_tunnel_Aug19_Nam": "Nam",
    "Second_Lex_Aug18_Ella": "Ella",
}

# Display token for missing text / ID fields (never for numeric measurements)
MISSING = "N/A"

HEADER_FONT = Font(name="Calibri", bold=True, size=11)
BODY_FONT = Font(name="Calibri", size=11)
HEADER_BORDER = Border(bottom=Side(style="thin", color="000000"))
NO_FILL = PatternFill(fill_type=None)
# Match HT_01 / source "First" sheet spacing
HEADER_ROW_HEIGHT = 18
BODY_ROW_HEIGHT = 25
DEFAULT_COL_WIDTH = 10.875

# Preferred column order for harvest-style sheets (audit cols always last)
HARVEST_COL_ORDER = [
    "SN",
    "Genotypes",
    "Location",
    "Date of harvest",
    "Harvest",
    "Experiment",
    "Grade 1 Marketable (weight)",
    "Grade 1 Marketable (counts)",
    "Grade 2 Marketable (weight)",
    "Grade 2 Marketable (counts)",
    "Unmarketables (weight)",
    "Unmarketables (counts)",
    "Fruit count",
    "Weight of 20 marketable (g)",
    "Weight of 20 unmarketable (g)",
    "Weight of 20 fruit (g)",
    "Total Weight (g)",
    "Total Weight (kg)",
    "Reading 1",
    "Reading 2",
    "Reading 3",
    "Average",
    "Brix",
    "Note",
]
AUDIT_COL_ORDER = ["Genotype_raw", "Sample_raw", "Replicate", "Entered_by"]

ONEDRIVE = ROOT / "data/reference"

log_rows: list[tuple] = []


def log(sheet, row, field, old, new, note=""):
    if old == new and not note:
        return
    log_rows.append((sheet, row, field, _fmt(old), _fmt(new), note))


def _fmt(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.isoformat()
    return str(v)


def extract_replicate(name):
    if name is None or (isinstance(name, float) and str(name) == "nan"):
        return name, None
    text = str(name).strip()
    # "Rep 1" / "Rep1" / "R1" at end (with or without space/hyphen)
    m = re.search(r"(?:\s+|-)?(?:Rep|R)\s*(\d+)\s*$", text, flags=re.I)
    if m:
        return text[: m.start()].strip(" -"), f"R{m.group(1)}"
    # "Second Rep" / "first rep" embedded in messy genotype cells
    ord_m = re.search(r"(?i)\b(first|second|third|fourth|fifth)\s+reps?\b", text)
    if ord_m:
        n = {
            "first": 1,
            "second": 2,
            "third": 3,
            "fourth": 4,
            "fifth": 5,
        }[ord_m.group(1).lower()]
        text = (text[: ord_m.start()] + text[ord_m.end() :]).strip(" -")
        return text, f"R{n}"
    return text, None


def clean_genotype(raw):
    """Return (cleaned_name, replicate_or_None)."""
    if raw is None:
        return None, None
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        # Excel numeric IDs like 2633 → try as YYNN / YYNNN
        if float(raw).is_integer():
            n = int(raw)
            s = str(n)
            if len(s) == 4:  # 2633 → 26-033
                raw = f"{s[:2]}-{s[2:]}"
            elif len(s) <= 3:
                raw = s
            else:
                raw = s
        else:
            return str(raw), None

    text, rep = extract_replicate(raw)
    if text is None or str(text).strip() == "":
        return None, rep
    text = re.sub(r"\s+", " ", str(text).strip())
    text = text.rstrip(".")
    # Trailing harvest tags jammed onto names: "...PLUSF3", "...PLUS F2"
    text = re.sub(r"(?i)\s*F\d+\s*$", "", text).strip()

    # Strip leading harvest/location tags: "H2 25-002", "HT3 25-004"
    tagged = re.match(r"(?i)^(?:H(?:T)?|OF|SF|RC)\s*\d+\s+(.+)$", text)
    if tagged:
        text = tagged.group(1).strip()

    # Drop trailing junk after a clean ID: "UKY-26-035 Second Rep first row..."
    leading_id = re.match(
        r"(?i)^(uky[t]?\s*-?\s*\d{2}\s*-?\s*\d{1,3}|\d{2}\s*-?\s*\d{1,3})\b(.*)$",
        text,
    )
    if leading_id and leading_id.group(2) and re.search(
        r"[A-Za-z]", leading_id.group(2)
    ):
        text = leading_id.group(1).strip()

    # trailing plot tags like " -1" / "-1" / "-2" on commercials
    text_nop = re.sub(r"\s*-\s*\d+\s*$", "", text).strip()

    key = text.lower()
    key_nop = text_nop.lower()
    if key in COMMERCIAL_ALIASES:
        return COMMERCIAL_ALIASES[key], rep
    if key_nop in COMMERCIAL_ALIASES:
        return COMMERCIAL_ALIASES[key_nop], rep

    # UKYT / UKY / UKyt variants
    uk = re.match(
        r"(?i)^uky[t]?\s*-?\s*(\d{2})\s*-?\s*(\d{1,3})$",
        text,
    )
    if uk:
        return f"UKYT-{uk.group(1)}-{int(uk.group(2)):03d}", rep

    # YY-N / YY-NN / YY-NNN (and 25-01, 26-18)
    short = re.match(r"^(\d{2})\s*[-–]\s*(\d{1,3})$", text)
    if short:
        return f"UKYT-{short.group(1)}-{int(short.group(2)):03d}", rep

    # YYNN packed without hyphen; also "2501"
    packed = re.match(r"^(\d{2})(\d{2,3})$", text)
    if packed and len(text) in (4, 5):
        return f"UKYT-{packed.group(1)}-{int(packed.group(2)):03d}", rep

    # NC / NG / NS / N accession codes — collapse spaces, drop trailing -1 plot tags
    nc = re.match(r"(?i)^(n[cgs]?)\s*[-]?\s*(\d[\d\s]*[a-z]?)\s*(?:-\s*\d+)?$", text)
    if nc:
        prefix_raw = nc.group(1).upper()
        prefix = "NG" if prefix_raw == "NG" else "NC"
        rest = re.sub(r"\s+", "", nc.group(2)).upper()
        return f"{prefix}{rest}", rep

    # Bare codes: 25218G, 264 G, 23706G → NC25218G (field convention)
    bare = re.match(r"^(\d+)\s*([A-Za-z])\s*(?:-\s*\d+)?$", text)
    if bare:
        digits, letter = bare.group(1), bare.group(2).upper()
        if letter == "G":
            return f"NC{digits}G", rep
        return f"{digits}{letter}", rep
    bare2 = re.match(r"^(\d+[A-Za-z])\s*(?:-\s*\d+)?$", text)
    if bare2 and re.search(r"[A-Za-z]", text):
        code = bare2.group(1).upper()
        if re.match(r"^\d+G$", code):
            return f"NC{code}", rep
        return code, rep

    return text_nop if text_nop != text else text, rep


def clean_location(raw, sheet_name):
    default_loc, _ = SHEET_DEFAULTS.get(sheet_name, (None, None))
    if raw is None or str(raw).strip() == "":
        return default_loc, "filled_from_sheet_default" if default_loc else ""

    text = re.sub(r"\s+", " ", str(raw).strip())
    for pat, canon in LOCATION_MAP.items():
        if re.match(pat, text):
            return canon, ""

    # Invalid tokens (e.g. "Grape") → sheet default if available
    if default_loc:
        return default_loc, f"invalid_location_replaced:{text}"
    return text, "unmapped_location"


def clean_harvest(raw, sheet_name):
    _, default_harv = SHEET_DEFAULTS.get(sheet_name, (None, None))
    if raw is None or str(raw).strip() == "":
        return default_harv, "filled_from_sheet_default" if default_harv else ""

    text = re.sub(r"\s+", " ", str(raw).strip())
    for pat, canon in HARVEST_MAP.items():
        if re.match(pat, text):
            return canon, ""
    if default_harv:
        return default_harv, f"invalid_harvest_replaced:{text}"
    return text, "unmapped_harvest"


def clean_date(raw):
    if raw is None or raw == "":
        return None
    if isinstance(raw, datetime):
        return raw.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(raw, date):
        return datetime(raw.year, raw.month, raw.day)
    try:
        return datetime.fromisoformat(str(raw)[:10])
    except Exception:
        return raw  # leave unparseable


def is_descriptive_note(text):
    """True when a 'genotype' cell is really a field note (not an ID + commentary)."""
    if text is None:
        return False
    t = str(text).strip()
    # Leading accession / commercial ID → treat as genotype (clean separately)
    if re.match(
        r"(?i)^(uky[t]?|n[cgs]|mt\.?|mountain|peace|bh|bhn|cherokee|\d{2}\s*-|\d{2}\d{2,3})",
        t,
    ):
        return False
    if len(t) >= 40:
        return True
    return bool(
        re.search(
            r"(?i)\b(fruits?|close to|mixed with|no sample|unidentified|orange fruits)\b",
            t,
        )
    )


def header_index_map(ws, header_row=1):
    """Map normalized header name → 1-based column index (first match)."""
    mapping = {}
    for col in range(1, (ws.max_column or 0) + 1):
        val = ws.cell(header_row, col).value
        if val is None:
            continue
        key = re.sub(r"\s+", " ", str(val).strip().lower())
        mapping.setdefault(key, col)
    return mapping


def find_col(hmap, *candidates):
    for c in candidates:
        if c in hmap:
            return hmap[c]
    return None


def ensure_trailing_col(ws, header, header_row=1):
    """Add header after the last occupied header cell (or reuse if present)."""
    hmap = header_index_map(ws, header_row)
    key = header.lower()
    if key in hmap:
        return hmap[key]
    max_h = 0
    for col in range(1, (ws.max_column or 0) + 1):
        if ws.cell(header_row, col).value is not None:
            max_h = col
    new_col = max_h + 1
    ws.cell(header_row, new_col).value = header
    ws.cell(header_row, new_col).font = Font(bold=True)
    return new_col


def process_table_sheet(ws, sheet_name, header_row=1, geno_keys=None, robinson_sn=False):
    geno_keys = geno_keys or ("genotypes",)
    hmap = header_index_map(ws, header_row)
    geno_col = None
    for k in geno_keys:
        geno_col = find_col(hmap, k)
        if geno_col:
            break
    loc_col = find_col(hmap, "location")
    date_col = find_col(hmap, "date of harvest", "date")
    harv_col = find_col(hmap, "harvest")
    sn_col = find_col(hmap, "sn")

    if geno_col is None and robinson_sn:
        geno_col = find_col(hmap, "genotypes") or 2
        if ws.cell(header_row, geno_col).value is None:
            ws.cell(header_row, geno_col).value = "Genotypes"

    if geno_col is None:
        return

    using_sample = find_col(hmap, "sample") == geno_col
    # Always keep Replicate: same genotype can be planted in multiple plots (R1/R2…).
    # Keep Genotype_raw on non-HT sheets (IDs are messier) and whenever a name changes.
    keep_raw = sheet_name not in {
        "First",
        "Second",
        "Third",
        "Fourth",
        "Fifth",
        "High Tunnel Harvest Sixth",
        "High_tunnel_Aug19_Nam",
    } and not sheet_name.startswith("HT_0")

    # Strip stale audit cols then re-add what we want
    for drop_header in ("Genotype_raw", "Sample_raw", "Replicate", "Entered_by"):
        hmap = header_index_map(ws, header_row)
        drop_c = find_col(hmap, drop_header.lower())
        if drop_c:
            # Only drop Entered_by if we're about to recreate; always recreate Replicate
            ws.delete_cols(drop_c)

    raw_header = "Sample_raw" if using_sample else "Genotype_raw"
    raw_col = ensure_trailing_col(ws, raw_header, header_row) if keep_raw else None
    rep_col = ensure_trailing_col(ws, "Replicate", header_row)
    entered_col = None
    if sheet_name in ENTERED_BY:
        entered_col = ensure_trailing_col(ws, "Entered_by", header_row)

    # Refresh col indexes after possible header writes / deletes
    hmap = header_index_map(ws, header_row)
    geno_col = find_col(hmap, *geno_keys) or geno_col
    if keep_raw:
        raw_col = find_col(hmap, raw_header.lower()) or raw_col
    rep_col = find_col(hmap, "replicate") or rep_col
    if entered_col:
        entered_col = find_col(hmap, "entered_by") or entered_col
    loc_col = find_col(hmap, "location")
    date_col = find_col(hmap, "date of harvest", "date")
    harv_col = find_col(hmap, "harvest")
    sn_col = find_col(hmap, "sn")
    note_col = find_col(hmap, "note", "notes", "notes:")

    last_date = None
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        values = [ws.cell(row, c).value for c in range(1, (ws.max_column or 0) + 1)]
        if all(v is None or str(v).strip() == "" for v in values):
            continue

        geno_val = ws.cell(row, geno_col).value if geno_col else None
        if robinson_sn and (geno_val is None or str(geno_val).strip() == "") and sn_col:
            sn_val = ws.cell(row, sn_col).value
            if sn_val is not None and str(sn_val).strip() != "":
                log(sheet_name, row, "Genotypes", geno_val, sn_val, "moved_from_SN")
                geno_val = sn_val
                ws.cell(row, geno_col).value = sn_val
                ws.cell(row, sn_col).value = row - header_row

        if geno_col:
            old = ws.cell(row, geno_col).value
            if isinstance(old, str) and old.startswith("="):
                pass
            elif old is not None:
                # Field notes typed into the Genotypes cell (e.g. Aug19 row 59)
                if is_descriptive_note(old):
                    if note_col:
                        existing = ws.cell(row, note_col).value
                        merged = str(old).strip()
                        if existing:
                            merged = f"{existing}; {merged}"
                        ws.cell(row, note_col).value = merged
                    log(
                        sheet_name,
                        row,
                        "Genotypes",
                        old,
                        MISSING,
                        "moved_descriptive_note_to_Note",
                    )
                    ws.cell(row, geno_col).value = MISSING
                    if raw_col:
                        ws.cell(row, raw_col).value = old
                else:
                    cleaned, rep = clean_genotype(old)
                    if raw_col:
                        ws.cell(row, raw_col).value = old
                    if cleaned != old:
                        log(
                            sheet_name,
                            row,
                            "Genotypes" if not using_sample else "Sample",
                            old,
                            cleaned,
                        )
                    ws.cell(row, geno_col).value = cleaned
                    if rep_col:
                        ws.cell(row, rep_col).value = rep
                    # Leftover commentary after ID (e.g. "first row first genotype") → Note
                    if cleaned and note_col and isinstance(old, str):
                        leftover = str(old).strip()
                        for pat in (
                            re.escape(str(cleaned)),
                            re.escape(str(old).split()[0]) if old.split() else "",
                            r"(?i)uky[t]?\s*-?\s*\d{2}\s*-?\s*\d{1,3}",
                            r"(?i)\b(?:first|second|third|fourth|fifth)\s+reps?\b",
                            r"(?i)\bR\s*\d+\b",
                            r"(?i)\bRep\s*\d+\b",
                        ):
                            if pat:
                                leftover = re.sub(pat, " ", leftover)
                        leftover = re.sub(r"\s+", " ", leftover).strip(" -;,").strip()
                        if leftover and len(leftover) > 2:
                            existing = ws.cell(row, note_col).value
                            if existing and str(existing).strip() and str(existing).strip().upper() != MISSING:
                                if leftover.lower() not in str(existing).lower():
                                    ws.cell(row, note_col).value = f"{existing}; {leftover}"
                            else:
                                ws.cell(row, note_col).value = leftover

        if entered_col and sheet_name in ENTERED_BY:
            ws.cell(row, entered_col).value = ENTERED_BY[sheet_name]

        if loc_col:
            old = ws.cell(row, loc_col).value
            if sheet_name in FORCE_LOCATION_SHEETS:
                default_loc, _ = SHEET_DEFAULTS.get(sheet_name, (None, None))
                new, note = default_loc, "forced_from_sheet_name"
                if old == new:
                    note = ""
            else:
                new, note = clean_location(old, sheet_name)
            if new != old:
                log(sheet_name, row, "Location", old, new, note)
            if new is not None:
                ws.cell(row, loc_col).value = new

        if harv_col:
            old = ws.cell(row, harv_col).value
            if sheet_name in FORCE_HARVEST_SHEETS:
                _, default_harv = SHEET_DEFAULTS.get(sheet_name, (None, None))
                new, note = default_harv, "forced_from_sheet_name"
                if old == new:
                    note = ""
            else:
                new, note = clean_harvest(old, sheet_name)
            if new != old:
                log(sheet_name, row, "Harvest", old, new, note)
            if new is not None:
                ws.cell(row, harv_col).value = new

        if date_col:
            old = ws.cell(row, date_col).value
            parsed = clean_date(old)
            # Do not forward-fill dates onto empty spacer rows (no ID, no measurements)
            has_id = geno_col and ws.cell(row, geno_col).value not in (
                None,
                "",
                MISSING,
            )
            has_measure = False
            for c in range(1, (ws.max_column or 0) + 1):
                if c in {
                    geno_col,
                    loc_col,
                    date_col,
                    harv_col,
                    sn_col,
                    raw_col,
                    rep_col,
                    entered_col,
                    note_col,
                }:
                    continue
                val = ws.cell(row, c).value
                if _is_significant_measure(ws.cell(header_row, c).value, val):
                    has_measure = True
                    break
            if not has_id and not has_measure and parsed is None:
                continue

            policy = DATE_POLICY.get(sheet_name)
            if policy:
                default_date, mode, reason = policy
                if mode == "force" and default_date is not None:
                    # Only force dates on real data rows
                    if not has_id and not has_measure:
                        continue
                    new = default_date
                    note = "date_forced:" + reason if parsed != default_date else ""
                elif mode == "fill":
                    new = parsed if parsed is not None else (
                        last_date if last_date is not None else default_date
                    )
                    note = ""
                    if parsed is None and new is not None:
                        note = "date_filled:" + reason
                else:  # keep
                    new = parsed if parsed is not None else last_date
                    note = "forward_filled" if parsed is None and new is not None else ""
            else:
                new = parsed if parsed is not None else last_date
                note = "forward_filled" if parsed is None and new is not None else ""

            if new != old and (note or (old is not None and new != parsed)):
                log(sheet_name, row, "Date of harvest", old, new, note or "normalized")
            elif note:
                log(sheet_name, row, "Date of harvest", old, new, note)
            if new is not None:
                ws.cell(row, date_col).value = new
                ws.cell(row, date_col).number_format = "yyyy-mm-dd"
                if isinstance(new, datetime):
                    last_date = new


def drop_useless_columns(ws, header_row=1):
    """Remove empty / duplicate junk columns (common on South Farm sheets)."""
    max_col = ws.max_column or 0
    max_row = ws.max_row or header_row
    seen_headers = {}
    to_delete = []

    for col in range(1, max_col + 1):
        header = ws.cell(header_row, col).value
        key = re.sub(r"\s+", " ", str(header).strip().lower()) if header else f"__empty_{col}"

        # Duplicate SN (second copy from paste-merged templates)
        if key == "sn":
            if "sn" in seen_headers:
                to_delete.append(col)
                continue
            seen_headers["sn"] = col
        elif header is not None:
            seen_headers[key] = col

        has_value = False
        has_formula = False
        for row in range(header_row + 1, max_row + 1):
            val = ws.cell(row, col).value
            if val is None or (isinstance(val, str) and val.strip() == ""):
                continue
            if isinstance(val, str) and val.startswith("="):
                has_formula = True
            else:
                has_value = True
                break

        # Drop blank header with no content, or fully empty non-core columns
        if header is None and not has_value and not has_formula:
            to_delete.append(col)
            continue
        if header is None and has_value:
            # Orphan values under a blank header — fold into Note when possible
            note_c = seen_headers.get("note") or find_col(
                header_index_map(ws, header_row), "note", "notes"
            )
            if note_c:
                for row in range(header_row + 1, max_row + 1):
                    val = ws.cell(row, col).value
                    if val is None or (isinstance(val, str) and val.strip() == ""):
                        continue
                    existing = ws.cell(row, note_c).value
                    ws.cell(row, note_c).value = (
                        f"{existing}; {val}" if existing else val
                    )
                to_delete.append(col)
                continue
        if not has_value and not has_formula:
            # Keep structural / audit columns even when mostly blank
            # (Replicate often empty later in season — still meaningful)
            if key in {
                "location",
                "date of harvest",
                "date",
                "harvest",
                "genotypes",
                "sample",
                "sn",
                "replicate",
                "entered_by",
                "genotype_raw",
                "sample_raw",
                "note",
                "notes",
            }:
                continue
            to_delete.append(col)

    # Delete from right to left so indexes stay valid
    for col in sorted(set(to_delete), reverse=True):
        header = ws.cell(header_row, col).value
        log(ws.title, header_row, "column_removed", header, None, "empty_or_duplicate")
        ws.delete_cols(col)

    # Repair Total Weight (KG) formulas if they point at a stale column
    hmap = header_index_map(ws, header_row)
    gram_c = find_col(hmap, "total weight (gram)", "total weight (gm)")
    kg_c = find_col(hmap, "total weight (kg)")
    if gram_c and kg_c:
        letter = get_column_letter(gram_c)
        for row in range(header_row + 1, (ws.max_row or header_row) + 1):
            cell = ws.cell(row, kg_c)
            if cell.value is None:
                continue
            cell.value = f"={letter}{row}/1000"


def _is_significant_measure(header, val):
    """True for real assay/yield values (ignore template placeholders)."""
    if isinstance(val, str) and val.startswith("="):
        return False  # formulas alone don't keep an empty-ID row
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    h = re.sub(r"\s+", " ", str(header or "").strip().lower())
    # Template default mass copied into blank rows
    if "exact sample mass" in h and float(val) == 1.0:
        return False
    return True


def drop_empty_data_rows(ws, header_row=1):
    """Remove unused plot slots and fully blank spacer rows."""
    hmap = header_index_map(ws, header_row)
    geno_c = find_col(hmap, "genotypes", "sample")
    ignore = {
        find_col(hmap, "sn"),
        find_col(hmap, "location"),
        find_col(hmap, "date of harvest", "date"),
        find_col(hmap, "harvest"),
        find_col(hmap, "genotype_raw", "sample_raw"),
        find_col(hmap, "replicate"),
        find_col(hmap, "entered_by"),
        find_col(hmap, "experiment"),
        find_col(hmap, "user"),
        find_col(hmap, "note", "notes"),
        geno_c,
    }
    ignore.discard(None)
    max_row = ws.max_row or header_row
    for row in range(max_row, header_row, -1):
        geno = ws.cell(row, geno_c).value if geno_c else None
        geno_missing = (
            geno is None
            or str(geno).strip() == ""
            or str(geno).strip().upper() == MISSING
        )
        has_measure = False
        for col in range(1, (ws.max_column or 0) + 1):
            if col in ignore:
                continue
            val = ws.cell(row, col).value
            if _is_significant_measure(ws.cell(header_row, col).value, val):
                has_measure = True
                break
        if geno_missing and not has_measure:
            ws.delete_rows(row)
            log(ws.title, row, "row_removed", None, None, "empty_plot_slot")


def standardize_headers(ws, header_row=1):
    """Normalize header spellings / Title Case across sheets."""
    canon = {
        "sn": "SN",
        "en": "EN",
        "genotypes": "Genotypes",
        "genotype": "Genotypes",
        "uky_id": "UKY_ID",
        "sample": "Sample",
        "location": "Location",
        "date of harvest": "Date of harvest",
        "date": "Date",
        "harvest": "Harvest",
        "grade 1 marketable (weight)": "Grade 1 Marketable (weight)",
        "grade 1 marketable (counts)": "Grade 1 Marketable (counts)",
        "grade 2 marketable (weight)": "Grade 2 Marketable (weight)",
        "grade 2 marketable (counts)": "Grade 2 Marketable (counts)",
        "unmarketables (weight)": "Unmarketables (weight)",
        "unmarketables (count)": "Unmarketables (counts)",
        "unmarketables (counts)": "Unmarketables (counts)",
        "unmarketable(count)": "Unmarketables (counts)",
        "total weight (gram)": "Total Weight (g)",
        "total weight (gm)": "Total Weight (g)",
        "total weight (kg)": "Total Weight (kg)",
        "total weight": "Total Weight (g)",
        "weight(20 samples m)": "Weight of 20 marketable (g)",
        "weight(20 samples um)": "Weight of 20 unmarketable (g)",
        "weight of 20 counts": "Weight of 20 fruit (g)",
        "weight of 20 marketable (g)": "Weight of 20 marketable (g)",
        "weight of 20 fruit (g)": "Weight of 20 fruit (g)",
        "fruit count": "Fruit count",
        "total weight (lbs)": "Total Weight (g)",
        "reading 1": "Reading 1",
        "reading 2": "Reading 2",
        "reading 3": "Reading 3",
        "average": "Average",
        "brix": "Brix",
        "note": "Note",
        "notes": "Notes",
        "notes:": "Notes",
        "genotype_raw": "Genotype_raw",
        "sample_raw": "Sample_raw",
        "replicate": "Replicate",
        "entered_by": "Entered_by",
        "experiment": "Experiment",
        "user": "User",
        "growth habit": "Growth habit",
        "fruit type / notes": "Fruit type / notes",
        "source / pedigree": "Source / pedigree",
        "other notes": "Other notes",
        "a630": "A630",
        "a750": "A750",
        "a663": "A663",
        "a645": "A645",
        "a505": "A505",
        "a453": "A453",
        "a530": "A530",
        "a657": "A657",
        "a765": "A765",
        "vol naoh (ml)": "Vol NaOH (mL)",
        "exact mass (g)": "Exact mass (g)",
        "exact sample mass (g)": "Exact sample mass (g)",
        "ta (% citric acid, g/100 g)": "TA (% citric acid, g/100 g)",
        "total anthocyanin": "Total Anthocyanin",
        "gaba (mg/g)": "GABA (mg/g)",
        "polyphenolic content (mg/g)": "Polyphenolic Content (mg/g)",
        "lycopene(mg/100ml)": "Lycopene (mg/100 mL)",
        "beta-carotene(mg/100ml)": "Beta-carotene (mg/100 mL)",
        "lycopene(mg/100g fw)": "Lycopene (mg/100 g fw)",
        "beta-carotene(mg/100g fw)": "Beta-carotene (mg/100 g fw)",
    }
    for col in range(1, (ws.max_column or 0) + 1):
        raw = ws.cell(header_row, col).value
        if raw is None:
            continue
        trimmed = re.sub(r"\s+", " ", str(raw).strip())
        key = trimmed.lower()
        if key in canon and trimmed != canon[key]:
            log(ws.title, header_row, "header", raw, canon[key], "standardized")
            ws.cell(header_row, col).value = canon[key]
        elif trimmed != str(raw):
            ws.cell(header_row, col).value = trimmed


def shift_formula_refs(formula: str, shift: int = 1, min_col: int = 1) -> str:
    """Shift A1 column letters in a formula (e.g. after inserting a column)."""
    pattern = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")

    def repl(m):
        dol1, letters, dol2, row = m.groups()
        from openpyxl.utils import column_index_from_string

        idx = column_index_from_string(letters)
        if idx < min_col:
            return m.group(0)
        return f"{dol1}{get_column_letter(idx + shift)}{dol2}{row}"

    return pattern.sub(repl, formula)


def shift_sheet_formulas(ws, header_row=1, shift=1, min_col: int = 1):
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        for col in range(1, (ws.max_column or 0) + 1):
            cell = ws.cell(row, col)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = shift_formula_refs(cell.value, shift, min_col=min_col)


def ensure_sn_column(ws, header_row=1):
    """Guarantee an SN column exists (leftmost) for row numbering."""
    hmap = header_index_map(ws, header_row)
    sn_c = find_col(hmap, "sn")
    if sn_c == 1:
        return 1
    if sn_c:
        return sn_c
    ws.insert_cols(1)
    ws.cell(header_row, 1).value = "SN"
    ws.cell(header_row, 1).font = Font(bold=True)
    shift_sheet_formulas(ws, header_row, shift=1)
    return 1


def fill_sn_sequential(ws, header_row=1):
    """Renumber SN 1…n for every non-empty data row (Entries-style EN)."""
    sn_c = ensure_sn_column(ws, header_row)
    hmap = header_index_map(ws, header_row)
    geno_c = find_col(hmap, "genotypes", "sample", "uky_id")
    n = 0
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        values = [ws.cell(row, c).value for c in range(1, (ws.max_column or 0) + 1)]
        if all(v is None or str(v).strip() == "" for v in values):
            continue
        # Skip spacer / title-only rows (no genotype/sample and no SN intent)
        if geno_c:
            g = ws.cell(row, geno_c).value
            if g is None or str(g).strip() == "" or str(g).strip().upper() == MISSING:
                # still number if other data present
                others = [
                    ws.cell(row, c).value
                    for c in range(1, (ws.max_column or 0) + 1)
                    if c != sn_c and c != geno_c
                ]
                if all(
                    v is None or str(v).strip() == "" or str(v).strip().upper() == MISSING
                    for v in others
                ):
                    continue
        n += 1
        old = ws.cell(row, sn_c).value
        if old != n:
            log(ws.title, row, "SN", old, n, "filled_sequential")
        ws.cell(row, sn_c).value = n


def fill_text_missing(ws, header_row=1):
    """Replace blank text/ID cells with N/A. Leave numeric measurement cells empty."""
    hmap = header_index_map(ws, header_row)
    text_keys = {
        "location",
        "harvest",
        "replicate",
        "entered_by",
        "genotype_raw",
        "sample_raw",
        "note",
        "notes",
        "user",
        "growth habit",
        "fruit type / notes",
        "source / pedigree",
        "other notes",
        "date of harvest",
        "date",
        "experiment",
    }
    text_cols = [col for key, col in hmap.items() if key in text_keys]
    geno_c = find_col(hmap, "genotypes", "sample", "uky_id")

    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        vals = [ws.cell(row, c).value for c in range(1, (ws.max_column or 0) + 1)]
        if all(v is None or str(v).strip() == "" for v in vals):
            continue
        # Fill N/A on any data-bearing row (including notes-only lab rows)
        for col in text_cols:
            cell = ws.cell(row, col)
            if cell.value is None or (isinstance(cell.value, str) and cell.value.strip() == ""):
                cell.value = MISSING


def title_case_user_names(ws, header_row=1):
    hmap = header_index_map(ws, header_row)
    user_c = find_col(hmap, "user", "entered_by")
    if not user_c:
        return
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        val = ws.cell(row, user_c).value
        if val is None or not isinstance(val, str):
            continue
        cleaned = val.strip()
        if not cleaned or cleaned.upper() == MISSING:
            continue
        known = {"nam": "Nam", "ella": "Ella"}
        if cleaned.lower() in known:
            new = known[cleaned.lower()]
        elif cleaned.isupper() and len(cleaned) <= 3:
            new = cleaned  # keep initials like BA
        else:
            new = cleaned.title()
        if new != val:
            ws.cell(row, user_c).value = new


def stringify_id_cells(ws, header_row=1):
    """Keep genotype/sample raw IDs as text (avoid 2633 looking like a number)."""
    hmap = header_index_map(ws, header_row)
    for key in ("genotype_raw", "sample_raw", "genotypes", "sample"):
        col = find_col(hmap, key)
        if not col:
            continue
        for row in range(header_row + 1, (ws.max_row or header_row) + 1):
            val = ws.cell(row, col).value
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                if float(val).is_integer():
                    ws.cell(row, col).value = str(int(val))
                else:
                    ws.cell(row, col).value = str(val)


def reorder_columns(ws, preferred, header_row=1):
    """Rewrite sheet columns into preferred order (Entries-like left-to-right)."""
    hmap = header_index_map(ws, header_row)
    if not hmap:
        return
    present = {}
    for key, col in hmap.items():
        header = ws.cell(header_row, col).value
        present[header] = col
    ordered_headers = [h for h in preferred if h in present]
    extras = [
        ws.cell(header_row, c).value
        for c in range(1, (ws.max_column or 0) + 1)
        if ws.cell(header_row, c).value is not None
        and ws.cell(header_row, c).value not in ordered_headers
        and ws.cell(header_row, c).value not in AUDIT_COL_ORDER
    ]
    audit = [h for h in AUDIT_COL_ORDER if h in present]
    final_headers = ordered_headers + extras + audit
    current = [
        ws.cell(header_row, c).value
        for c in range(1, (ws.max_column or 0) + 1)
        if ws.cell(header_row, c).value is not None
    ]
    if current == final_headers:
        return

    # Snapshot values only (not formats — stale date formats cause false datetimes)
    rows_out = [final_headers]
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        rows_out.append([
            ws.cell(row, present[h]).value if h in present else None
            for h in final_headers
        ])

    old_max_col = ws.max_column or len(final_headers)
    old_max_row = ws.max_row or header_row
    for row in range(header_row, old_max_row + 1):
        for col in range(1, old_max_col + 1):
            cell = ws.cell(row, col)
            cell.value = None
            cell.number_format = "General"

    for r_i, row_vals in enumerate(rows_out):
        for c_i, val in enumerate(row_vals, start=1):
            cell = ws.cell(header_row + r_i, c_i)
            cell.value = val
            cell.number_format = "General"


def drop_trailing_empty_cols(ws, header_row=1):
    max_col = ws.max_column or 0
    for col in range(max_col, 0, -1):
        header = ws.cell(header_row, col).value
        has = False
        for row in range(header_row, (ws.max_row or header_row) + 1):
            val = ws.cell(row, col).value
            if val is not None and str(val).strip() != "":
                has = True
                break
        if not has:
            ws.delete_cols(col)


def collapse_blank_header_cols(ws, header_row=1):
    """Remove columns with no header; fold stray notes into Notes when possible."""
    max_col = ws.max_column or 0
    for col in range(max_col, 0, -1):
        header = ws.cell(header_row, col).value
        if header is not None and str(header).strip() != "":
            continue

        formula_rows = 0
        text_notes = []
        for row in range(header_row + 1, (ws.max_row or header_row) + 1):
            val = ws.cell(row, col).value
            if isinstance(val, str) and val.startswith("="):
                formula_rows += 1
            elif isinstance(val, str) and val.strip():
                text_notes.append((row, val.strip()))
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                formula_rows += 1  # treat numeric assay output as keep-worthy

        if formula_rows >= 3:
            title = ws.title.lower()
            if "gaba" in title:
                ws.cell(header_row, col).value = "GABA (mg/g)"
            elif "polyphenol" in title:
                ws.cell(header_row, col).value = "Polyphenolic Content (mg/g)"
            elif "anthocyanin" in title:
                ws.cell(header_row, col).value = "Total Anthocyanin"
            elif "carotenoid" in title:
                ws.cell(header_row, col).value = "Assay result"
            else:
                ws.cell(header_row, col).value = "Calculated"
            continue

        if text_notes:
            # Re-resolve Notes column each time (indexes shift after deletes)
            hmap = header_index_map(ws, header_row)
            note_c = find_col(hmap, "note", "notes")
            if note_c is None:
                note_c = ensure_trailing_col(ws, "Notes", header_row)
            for row, text in text_notes:
                existing = ws.cell(row, note_c).value
                if existing and str(existing).strip() and str(existing).strip().upper() != MISSING:
                    ws.cell(row, note_c).value = f"{existing}; {text}"
                else:
                    ws.cell(row, note_c).value = text
        ws.delete_cols(col)



def _excel_serial_to_number(dt):
    """Convert a false datetime (from date-formatted numerics) back to Excel serial."""
    from openpyxl.utils.datetime import to_excel
    return round(float(to_excel(dt)), 4)


def apply_entries_style(ws, header_row=1):
    """Plain Entries-style formatting: bold headers, no fills/colors, freeze, filter."""
    collapse_blank_header_cols(ws, header_row)
    drop_trailing_empty_cols(ws, header_row)
    max_col = 0
    for col in range(1, (ws.max_column or 0) + 1):
        if ws.cell(header_row, col).value is not None:
            max_col = col
    if max_col == 0:
        return
    max_row = header_row
    for row in range(header_row, (ws.max_row or header_row) + 1):
        if any(
            ws.cell(row, c).value is not None and str(ws.cell(row, c).value).strip() != ""
            for c in range(1, max_col + 1)
        ):
            max_row = row

    # Drop phantom empty rows below data so Excel doesn't scroll a huge blank region
    for row in range((ws.max_row or max_row), max_row, -1):
        ws.delete_rows(row)

    date_cols = set()
    for col in range(1, max_col + 1):
        header = str(ws.cell(header_row, col).value or "").strip().lower()
        if header in {"date", "date of harvest"} or (
            "date" in header and "weight" not in header
        ):
            date_cols.add(col)

    for col in range(1, max_col + 1):
        cell = ws.cell(header_row, col)
        if cell.value is None:
            continue
        cell.font = HEADER_FONT
        cell.fill = NO_FILL
        cell.border = Border()
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.number_format = "General"

    for row in range(header_row + 1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            cell.font = BODY_FONT
            cell.fill = NO_FILL
            cell.border = Border()
            cell.alignment = Alignment(vertical="center")

            if col in date_cols:
                if isinstance(cell.value, (datetime, date)):
                    cell.number_format = "yyyy-mm-dd"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.number_format = "General"
            else:
                if isinstance(cell.value, datetime):
                    cell.value = _excel_serial_to_number(cell.value)
                elif isinstance(cell.value, date) and not isinstance(cell.value, datetime):
                    cell.value = _excel_serial_to_number(
                        datetime(cell.value.year, cell.value.month, cell.value.day)
                    )
                cell.number_format = "General"

            if cell.value == MISSING:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in range(1, max_col + 1):
        header = str(ws.cell(header_row, col).value or "")
        width = min(max(len(header) + 2, 8), 36)
        for row in range(header_row + 1, min(max_row, header_row + 40) + 1):
            val = ws.cell(row, col).value
            if val is None:
                continue
            width = max(width, min(len(str(val)) + 2, 36))
        ws.column_dimensions[get_column_letter(col)].width = width

    # Uniform row spacing like HT_01 (other sheets inherited shorter defaults)
    ws.sheet_format.defaultRowHeight = 16.5
    ws.sheet_format.defaultColWidth = DEFAULT_COL_WIDTH
    ws.sheet_format.customHeight = False
    # Clear any leftover row-dimension entries below the data
    for r in list(ws.row_dimensions.keys()):
        if isinstance(r, int) and r > max_row:
            del ws.row_dimensions[r]
    ws.row_dimensions[header_row].height = HEADER_ROW_HEIGHT
    for row in range(header_row + 1, max_row + 1):
        ws.row_dimensions[row].height = BODY_ROW_HEIGHT

    # Viewport reset (source left topLeftCell mid-sheet)
    ws.freeze_panes = None
    view = ws.sheet_view
    view.topLeftCell = "A1"
    view.selection = [Selection(activeCell="A1", sqref="A1")]
    freeze_at = f"A{header_row + 1}"
    ws.freeze_panes = freeze_at
    if view.pane is not None:
        view.pane.topLeftCell = freeze_at
    view.topLeftCell = "A1"
    view.showGridLines = True

    # Plain AutoFilter only — no Excel Table styles (those add gray/blue fills)
    for key in list(ws.tables.keys()):
        del ws.tables[key]
    if max_row >= header_row and max_col >= 1:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max_col)}{max_row}"

    # Bold header + thin underline (Entries_Lexington_field look; no fill)
    for col in range(1, max_col + 1):
        cell = ws.cell(header_row, col)
        if cell.value is None:
            continue
        cell.font = HEADER_FONT
        cell.fill = NO_FILL
        cell.border = HEADER_BORDER


def unmerge_all(ws):
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))


def polish_sheet(ws, header_row=1, harvest_style=False):
    """Final pass: headers, SN, N/A, column order, Entries-like formatting."""
    unmerge_all(ws)
    standardize_headers(ws, header_row)
    drop_trailing_empty_cols(ws, header_row)
    drop_empty_data_rows(ws, header_row)
    if harvest_style:
        reorder_columns(ws, HARVEST_COL_ORDER, header_row)
    fill_sn_sequential(ws, header_row)
    title_case_user_names(ws, header_row)
    stringify_id_cells(ws, header_row)
    fill_text_missing(ws, header_row)
    apply_entries_style(ws, header_row)


def convert_ella_lbs_to_grams(ws, header_row=1):
    """Ella sheet recorded TOTAL WEIGHT (lbs); convert to grams for site consistency."""
    hmap = header_index_map(ws, header_row)
    lbs_c = None
    for key, col in hmap.items():
        if "lbs" in key or key == "total weight (lbs)":
            lbs_c = col
            break
    # Also match before header rename
    if lbs_c is None:
        for col in range(1, (ws.max_column or 0) + 1):
            h = str(ws.cell(header_row, col).value or "")
            if "lbs" in h.lower():
                lbs_c = col
                break
    if lbs_c is None:
        return
    w20 = find_col(hmap, "weight of 20 counts", "weight of 20 fruit (g)")
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        val = ws.cell(row, lbs_c).value
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            grams = round(float(val) * 453.59237, 1)
            log(ws.title, row, "Total Weight", val, grams, "lbs_to_g")
            ws.cell(row, lbs_c).value = grams
        if w20:
            wval = ws.cell(row, w20).value
            if isinstance(wval, (int, float)) and not isinstance(wval, bool) and wval < 50:
                # Ella 20-count weights are also in lbs when < ~50
                grams20 = round(float(wval) * 453.59237, 1)
                log(ws.title, row, "Weight of 20", wval, grams20, "lbs_to_g")
                ws.cell(row, w20).value = grams20
    ws.cell(header_row, lbs_c).value = "Total Weight (g)"


def fix_date_display(wb):
    """Excel shows ###### when format/width is wrong; use Excel-safe yyyy-mm-dd."""
    for ws in wb.worksheets:
        if ws.title.startswith("_"):
            continue
        header_row = 2 if "Specific_Gravity" in ws.title or ws.title == "Lab_Specific_Gravity" else 1
        # Specific gravity title row
        if ws.title == "Lab_Specific_Gravity":
            header_row = 2
        hmap = header_index_map(ws, header_row)
        date_cols = []
        for key, col in hmap.items():
            if "date" in key:
                date_cols.append(col)
        # Also format any datetime cells on the sheet
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row or 1, max_col=ws.max_column or 1):
            for cell in row:
                if isinstance(cell.value, (datetime, date)):
                    cell.number_format = "yyyy-mm-dd"
        for col in date_cols:
            letter = get_column_letter(col)
            ws.column_dimensions[letter].width = 14


def add_ht_entry_metadata(wb):
    """Pull HT entry roster (habit/type) + pedigree from OneDrive dump if present."""
    roster_path = ONEDRIVE / "Lexington_Hightunnel/Entries_Lex_hightunnel.xlsx"
    pedigree_path = (
        ONEDRIVE / "Lexington_southFarm/Lexington_tomato_field_layout_randomization.xlsx"
    )
    if not roster_path.exists():
        print("WARNING: HT entry roster not found; skip HT_Entry_Metadata")
        return

    roster_wb = load_workbook(roster_path, data_only=True)
    # Sheet2 has correct habit/type alignment; Sheet1 headers are shifted.
    src = roster_wb["Sheet2"] if "Sheet2" in roster_wb.sheetnames else roster_wb.active

    if "HT_Entry_Metadata" in wb.sheetnames:
        del wb["HT_Entry_Metadata"]
    out = wb.create_sheet("HT_Entry_Metadata")
    out.append([
        "EN",
        "Genotypes",
        "Growth habit",
        "Fruit type / notes",
        "Source / pedigree",
    ])
    for cell in out[1]:
        cell.font = Font(bold=True)

    pedigree = {}
    if pedigree_path.exists():
        try:
            ped_wb = load_workbook(pedigree_path, data_only=True)
            if "Source_Breeding_Hybrid" in ped_wb.sheetnames:
                pws = ped_wb["Source_Breeding_Hybrid"]
                for row in pws.iter_rows(min_row=2, values_only=True):
                    if not row or row[1] is None:
                        continue
                    gid = clean_genotype(row[1])[0]
                    pedigree[gid or str(row[1])] = row[2]
        except Exception as exc:
            print(f"WARNING: pedigree join failed: {exc}")

    for row in src.iter_rows(min_row=2, values_only=True):
        if not row or row[1] is None:
            continue
        en, uky, habit, notes = row[0], row[1], row[2], row[3]
        extra = row[4] if len(row) > 4 and row[4] else None
        fruit_notes = notes
        if extra:
            fruit_notes = f"{notes}; {extra}" if notes else str(extra)
        cleaned, _ = clean_genotype(uky)
        out.append([
            en,
            cleaned or uky,
            habit,
            fruit_notes,
            pedigree.get(cleaned or str(uky)),
        ])

    for col, width in enumerate([6, 18, 16, 40, 28], start=1):
        out.column_dimensions[get_column_letter(col)].width = width
    print(f"Added HT_Entry_Metadata ({out.max_row - 1} entries)")


def _canon_experiment(raw):
    if raw is None:
        return None
    t = str(raw).strip().lower()
    if "hybrid" in t:
        return "Hybrid evaluation"
    if "varitome" in t or "wild" in t or "seed increase" in t:
        return "Varitome seed increase"
    if "advanced" in t or "breeding" in t:
        return "Advanced breeding"
    return str(raw).strip()


def build_sf_layout_maps():
    """Genotype → experiment / designed reps / pedigree from South Farm layout files."""
    layout_path = (
        ONEDRIVE / "Lexington_southFarm/Lexington_tomato_field_layout_randomization.xlsx"
    )
    entries_path = ONEDRIVE / "Lexington_southFarm/Entries_Lexington_field.xlsx"
    experiment = {}
    reps = {}
    pedigree = {}
    plants = {}
    note = {}
    en_map = {}

    if layout_path.exists():
        lwb = load_workbook(layout_path, data_only=True)
        if "Printable_Field_Tags" in lwb.sheetnames:
            tws = lwb["Printable_Field_Tags"]
            for row in tws.iter_rows(min_row=2, values_only=True):
                if not row or row[6] is None:
                    continue
                cleaned, _ = clean_genotype(row[6])
                key = cleaned or str(row[6]).strip()
                exp = _canon_experiment(row[0])
                if exp:
                    experiment[key] = exp
                if row[1] is not None and str(row[1]).strip() != "":
                    reps.setdefault(key, set()).add(int(row[1]) if str(row[1]).isdigit() else row[1])
                if row[5] is not None:
                    en_map.setdefault(key, row[5])
                if row[7]:
                    pedigree.setdefault(key, row[7])
        if "Source_Breeding_Hybrid" in lwb.sheetnames:
            for row in lwb["Source_Breeding_Hybrid"].iter_rows(min_row=2, values_only=True):
                if not row or row[1] is None:
                    continue
                cleaned, _ = clean_genotype(row[1])
                key = cleaned or str(row[1]).strip()
                if row[0] is not None:
                    en_map.setdefault(key, row[0])
                if row[2]:
                    pedigree.setdefault(key, row[2])
                if row[3]:
                    note.setdefault(key, row[3])
                if row[4] is not None:
                    plants.setdefault(key, row[4])

    if entries_path.exists():
        ewb = load_workbook(entries_path, data_only=True)
        if "breeding" in ewb.sheetnames:
            for row in ewb["breeding"].iter_rows(min_row=2, values_only=True):
                if not row or row[1] is None:
                    continue
                cleaned, _ = clean_genotype(row[1])
                key = cleaned or str(row[1]).strip()
                if row[0] is not None:
                    en_map.setdefault(key, row[0])
                if row[2]:
                    pedigree.setdefault(key, row[2])
                if row[3]:
                    note.setdefault(key, row[3])
                    if not experiment.get(key):
                        experiment[key] = _canon_experiment(row[3]) or "Advanced breeding"
                if row[4] is not None:
                    plants.setdefault(key, row[4])
        if "varitome" in ewb.sheetnames:
            for row in ewb["varitome"].iter_rows(min_row=2, values_only=True):
                # cols: EN?, Line, Line, Note, Plant number
                if not row or row[1] is None:
                    continue
                cleaned, _ = clean_genotype(row[1])
                key = cleaned or str(row[1]).strip()
                if row[0] is not None:
                    en_map.setdefault(key, row[0])
                note.setdefault(key, row[3] if len(row) > 3 else "Seed increase")
                experiment.setdefault(key, "Varitome seed increase")
                if len(row) > 4 and row[4] is not None:
                    plants.setdefault(key, row[4])

    return {
        "experiment": experiment,
        "reps": reps,
        "pedigree": pedigree,
        "plants": plants,
        "note": note,
        "en": en_map,
    }


def add_sf_entry_metadata(wb):
    """South Farm entry roster from Entries_Lexington_field + field layout tags."""
    maps = build_sf_layout_maps()
    if not maps["en"] and not maps["experiment"]:
        print("WARNING: SF entry sources not found; skip SF_Entry_Metadata")
        return maps

    if "SF_Entry_Metadata" in wb.sheetnames:
        del wb["SF_Entry_Metadata"]
    out = wb.create_sheet("SF_Entry_Metadata")
    out.append([
        "EN",
        "Genotypes",
        "Experiment",
        "Source / pedigree",
        "Note",
        "Plants per entry",
        "Designed reps",
    ])
    for cell in out[1]:
        cell.font = Font(bold=True)

    # Prefer EN-ordered unique genotypes
    rows = []
    seen = set()
    for key, en in maps["en"].items():
        if key in seen:
            continue
        seen.add(key)
        rep_set = maps["reps"].get(key) or set()
        rep_str = ", ".join(
            str(r) for r in sorted(rep_set, key=lambda x: (isinstance(x, str), x))
        ) or MISSING
        rows.append((
            en if isinstance(en, (int, float)) else en,
            key,
            maps["experiment"].get(key) or MISSING,
            maps["pedigree"].get(key) or MISSING,
            maps["note"].get(key) or MISSING,
            maps["plants"].get(key) if maps["plants"].get(key) is not None else MISSING,
            rep_str,
        ))
    # Also include genotypes that only appear in tags
    for key in maps["experiment"]:
        if key in seen:
            continue
        seen.add(key)
        rep_set = maps["reps"].get(key) or set()
        rep_str = ", ".join(
            str(r) for r in sorted(rep_set, key=lambda x: (isinstance(x, str), x))
        ) or MISSING
        rows.append((
            maps["en"].get(key) if maps["en"].get(key) is not None else MISSING,
            key,
            maps["experiment"].get(key) or MISSING,
            maps["pedigree"].get(key) or MISSING,
            maps["note"].get(key) or MISSING,
            maps["plants"].get(key) if maps["plants"].get(key) is not None else MISSING,
            rep_str,
        ))

    def sort_key(r):
        en = r[0]
        try:
            return (0, int(en), str(r[1]))
        except (TypeError, ValueError):
            return (1, 0, str(r[1]))

    for row in sorted(rows, key=sort_key):
        out.append(list(row))

    for col, width in enumerate([6, 18, 22, 28, 18, 14, 14], start=1):
        out.column_dimensions[get_column_letter(col)].width = width
    print(f"Added SF_Entry_Metadata ({out.max_row - 1} entries)")
    return maps


def add_rc_entry_metadata(wb):
    """Robinson entry roster + growth habit/type from send file."""
    entries_path = ONEDRIVE / "RobinsonCenter/Entries_Robinson.xlsx"
    traits_path = ONEDRIVE / "RobinsonCenter/To_send_Shawn_RobinsonTrial2026.xlsx"
    if not entries_path.exists():
        print("WARNING: Robinson entries not found; skip RC_Entry_Metadata")
        return

    if "RC_Entry_Metadata" in wb.sheetnames:
        del wb["RC_Entry_Metadata"]
    out = wb.create_sheet("RC_Entry_Metadata")
    out.append([
        "EN",
        "Genotypes",
        "Source / pedigree",
        "Note",
        "Plants",
        "Growth habit",
        "Fruit type",
        "Color",
        "Other notes",
    ])
    for cell in out[1]:
        cell.font = Font(bold=True)

    traits = {}
    if traits_path.exists():
        twb = load_workbook(traits_path, data_only=True)
        tws = twb.active
        # Title on row 1; headers on row 2; data from row 3
        for row in tws.iter_rows(min_row=3, values_only=True):
            if not row or row[1] is None:
                continue
            cleaned, _ = clean_genotype(row[1])
            key = cleaned or str(row[1]).strip()
            traits[key] = {
                "habit": row[4] if len(row) > 4 else None,
                "type": row[5] if len(row) > 5 else None,
                "color": row[6] if len(row) > 6 else None,
                "other": row[7] if len(row) > 7 else None,
            }

    ewb = load_workbook(entries_path, data_only=True)
    ews = ewb.active
    n = 0
    for row in ews.iter_rows(min_row=2, values_only=True):
        if not row or row[1] is None:
            continue
        cleaned, _ = clean_genotype(row[1])
        key = cleaned or str(row[1]).strip()
        t = traits.get(key, {})
        out.append([
            row[0],
            key,
            row[2],
            row[3],
            row[4],
            t.get("habit"),
            t.get("type"),
            t.get("color"),
            t.get("other"),
        ])
        n += 1

    for col, width in enumerate([6, 18, 22, 16, 10, 16, 14, 14, 40], start=1):
        out.column_dimensions[get_column_letter(col)].width = width
    print(f"Added RC_Entry_Metadata ({n} entries)")


def annotate_sf_experiments(wb, experiment_map):
    """Add Experiment column on SF harvest sheets from field-layout genotype join."""
    if not experiment_map:
        return
    sf_sheets = [n for n in wb.sheetnames if n.startswith("SF_") and "Metadata" not in n]
    for name in sf_sheets:
        ws = wb[name]
        hmap = header_index_map(ws, 1)
        geno_c = find_col(hmap, "genotypes")
        if not geno_c:
            continue
        exp_c = find_col(hmap, "experiment")
        if not exp_c:
            # Insert after Harvest when present, else after Genotypes
            harv_c = find_col(hmap, "harvest")
            insert_at = (harv_c + 1) if harv_c else (geno_c + 1)
            ws.insert_cols(insert_at)
            ws.cell(1, insert_at).value = "Experiment"
            ws.cell(1, insert_at).font = Font(bold=True)
            exp_c = insert_at
            shift_sheet_formulas(ws, header_row=1, shift=1, min_col=insert_at)

        filled = 0
        for row in range(2, (ws.max_row or 1) + 1):
            g = ws.cell(row, geno_c).value
            if g is None or str(g).strip() == "":
                continue
            cleaned, _ = clean_genotype(g)
            key = cleaned or str(g).strip()
            exp = experiment_map.get(key)
            if exp:
                ws.cell(row, exp_c).value = exp
                filled += 1
            elif ws.cell(row, exp_c).value in (None, ""):
                ws.cell(row, exp_c).value = MISSING
        print(f"  Experiment on {name}: {filled} rows matched")


def write_sheet_index(wb):
    if "_Sheet_Index" in wb.sheetnames:
        del wb["_Sheet_Index"]
    ws = wb.create_sheet("_Sheet_Index", 0)
    ws.append([
        "New sheet name",
        "Original name",
        "Location",
        "Harvest",
        "Date assigned",
        "Date reasoning",
    ])
    for cell in ws[1]:
        cell.font = Font(bold=True)

    lab_sheets = {
        "Specific Gravity",
        "Carotenoid",
        "TA",
        "Anthocyanin",
        "GABA",
        "Polyphenol",
    }
    for old, new in SHEET_RENAME_ORDER:
        loc, harv = SHEET_DEFAULTS.get(old, (None, None))
        policy = DATE_POLICY.get(old)
        if old in lab_sheets:
            loc, harv = "Lab", "—"
            dstr = "(as recorded)"
            reason = "Quality assay sheet; dates are lab run dates"
        elif policy and policy[0] is not None:
            dstr = policy[0].strftime("%Y-%m-%d")
            reason = policy[2]
        elif policy:
            dstr = "(mixed / blank)"
            reason = policy[2]
        else:
            dstr = "—"
            reason = "—"
        ws.append([new, old, loc or "—", harv or "—", dstr, reason])

    ws.append([])
    ws.append([
        "Naming key",
        "HT = High tunnel",
        "SF = South Farm / Lex field",
        "RC = Robinson Center",
        "Lab = quality assays",
        "Nam/Ella in sheet names = who entered the data",
    ])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    ws.append([])
    ws.append(["Empty cells", "Numeric blanks = not measured (left empty for Excel math)."])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    ws.append([
        "N/A",
        "Text/ID fields with no value (Replicate, notes, dates when unknown, etc.). Never means zero.",
    ])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    ws.append([
        "Replicate (R1/R2…)",
        "Same genotype planted in multiple plots; keep even if later harvests stopped logging replicas.",
    ])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    ws.append([
        "SN",
        "Sequential row number within the sheet (filled when blank), like EN in Entries_Lexington_field.",
    ])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    ws.append([
        "SF / RC metadata",
        "SF_Entry_Metadata + RC_Entry_Metadata from OneDrive entry lists; Experiment on SF harvest sheets from field layout tags.",
    ])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    for col, width in enumerate([28, 28, 14, 10, 18, 70], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def organize_sheets(wb):
    """Rename sheets and put them in a logical order."""
    for old, new in SHEET_RENAME_ORDER:
        if old in wb.sheetnames and old != new:
            wb[old].title = new

    desired = ["_Sheet_Index", "_Cleaning_Log"]
    for _, new in SHEET_RENAME_ORDER:
        desired.append(new)
        if new == "HT_07_Seventh_2026-08-19_Nam":
            desired.append("HT_Entry_Metadata")
        if new == "SF_Hybrid":
            desired.append("SF_Entry_Metadata")
        if new == "RC_01_First_2026-08-12":
            desired.append("RC_Entry_Metadata")
    for idx, name in enumerate(desired):
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=idx - wb.sheetnames.index(name))


def main():
    global log_rows
    log_rows = []
    wb = load_workbook(SRC)

    # --- Harvest-style sheets ---
    harvest_style = [
        "First",
        "Second",
        "Third",
        "Fourth",
        "Fifth",
        "High Tunnel Harvest Sixth",
        "High_tunnel_Aug19_Nam",
        "FirstHArvest_OpenFieldLex",
        "SecondHarvestOpenField",
        "Second_Lex_Aug18_Nam",
        "Second_Lex_Aug18_Ella",
        "Second_lex_Aug19_Nam",
        "hybrid_harvest_Lex_MS",
    ]
    for sheet in harvest_style:
        if sheet in wb.sheetnames:
            if sheet == "Second_Lex_Aug18_Ella":
                convert_ella_lbs_to_grams(wb[sheet], header_row=1)
            process_table_sheet(wb[sheet], sheet, header_row=1)
            drop_useless_columns(wb[sheet], header_row=1)
            drop_empty_data_rows(wb[sheet], header_row=1)
            standardize_headers(wb[sheet], header_row=1)

    if "RobinsonCenter" in wb.sheetnames:
        process_table_sheet(
            wb["RobinsonCenter"],
            "RobinsonCenter",
            header_row=1,
            robinson_sn=True,
        )
        drop_useless_columns(wb["RobinsonCenter"], header_row=1)
        drop_empty_data_rows(wb["RobinsonCenter"], header_row=1)
        standardize_headers(wb["RobinsonCenter"], header_row=1)

    # --- Specific Gravity (header row 2) ---
    if "Specific Gravity" in wb.sheetnames:
        ws = wb["Specific Gravity"]
        process_table_sheet(ws, "Specific Gravity", header_row=2)
        hmap = header_index_map(ws, 2)
        geno_c = find_col(hmap, "genotypes")
        r2_c = find_col(hmap, "reading 2")
        if geno_c and r2_c:
            for row in range(3, (ws.max_row or 2) + 1):
                g = ws.cell(row, geno_c).value
                r2 = ws.cell(row, r2_c).value
                if g == "NC25218G" and isinstance(r2, (int, float)) and r2 > 1.1:
                    log("Specific Gravity", row, "Reading 2", r2, 1.022, "typo_fix")
                    ws.cell(row, r2_c).value = 1.022

    # --- Lab Sample sheets ---
    for sheet in ["Carotenoid", "TA", "Anthocyanin", "GABA", "Polyphenol"]:
        if sheet in wb.sheetnames:
            process_table_sheet(
                wb[sheet],
                sheet,
                header_row=1,
                geno_keys=("sample", "genotypes"),
            )

    # --- Cleaning log ---
    if "_Cleaning_Log" in wb.sheetnames:
        del wb["_Cleaning_Log"]
    log_ws = wb.create_sheet("_Cleaning_Log", 0)
    log_ws.append(["Sheet", "Row", "Field", "Old", "New", "Note"])
    for cell in log_ws[1]:
        cell.font = Font(bold=True)
    for row in log_rows:
        log_ws.append(list(row))

    start = len(log_rows) + 4
    log_ws.cell(start, 1, "SUMMARY").font = Font(bold=True)
    field_counts = Counter(r[2] for r in log_rows)
    log_ws.cell(start + 1, 1, "Field")
    log_ws.cell(start + 1, 2, "Changes")
    for i, (f, n) in enumerate(field_counts.most_common(), start=start + 2):
        log_ws.cell(i, 1, f)
        log_ws.cell(i, 2, n)
    note_counts = Counter(r[5] for r in log_rows if r[5])
    base = start + 2 + len(field_counts) + 2
    log_ws.cell(base, 1, "Notes").font = Font(bold=True)
    for i, (n, c) in enumerate(note_counts.most_common(), start=base + 1):
        log_ws.cell(i, 1, n)
        log_ws.cell(i, 2, c)

    for col in range(1, 7):
        log_ws.column_dimensions[get_column_letter(col)].width = 22

    write_sheet_index(wb)
    add_ht_entry_metadata(wb)
    sf_maps = add_sf_entry_metadata(wb)
    add_rc_entry_metadata(wb)
    organize_sheets(wb)
    annotate_sf_experiments(wb, sf_maps.get("experiment") if sf_maps else {})
    fix_date_display(wb)

    # --- Final polish: SN fill, N/A, Title Case headers, Entries-like formatting ---
    harvest_polished = {
        "HT_01_First_2026-07-13",
        "HT_02_Second_2026-07-20",
        "HT_03_Third_2026-07-27",
        "HT_04_Fourth_2026-08-03",
        "HT_05_Fifth_2026-08-10",
        "HT_06_Sixth_2026-08-18",
        "HT_07_Seventh_2026-08-19_Nam",
        "SF_01_First_2026-07-23",
        "SF_02_Second_2026-08-07",
        "SF_03_Third_2026-08-18_Nam",
        "SF_03_Third_2026-08-18_Ella",
        "SF_04_Fourth_2026-08-19_Nam",
        "SF_Hybrid",
        "RC_01_First_2026-08-12",
    }
    meta_sheets = {"HT_Entry_Metadata", "SF_Entry_Metadata", "RC_Entry_Metadata"}
    for name in list(wb.sheetnames):
        if name.startswith("_"):
            apply_entries_style(wb[name], header_row=1)
            continue
        ws = wb[name]
        if name == "Lab_Specific_Gravity":
            polish_sheet(ws, header_row=2, harvest_style=True)
            if ws.cell(1, 1).value:
                ws.cell(1, 1).font = Font(name="Calibri", bold=True, size=14)
        elif name.startswith("Lab_"):
            polish_sheet(ws, header_row=1, harvest_style=False)
        elif name in meta_sheets:
            polish_sheet(ws, header_row=1, harvest_style=False)
        elif name in harvest_polished:
            polish_sheet(ws, header_row=1, harvest_style=True)
        else:
            polish_sheet(ws, header_row=1, harvest_style=False)

    wb.save(OUT)
    print(f"Wrote {OUT}")
    print(f"Logged changes: {len(log_rows)}")
    print("By field:", dict(field_counts))
    print("By note:", dict(note_counts))
    print("Sheet order:")
    for name in wb.sheetnames:
        print(f"  {name}")


if __name__ == "__main__":
    main()
