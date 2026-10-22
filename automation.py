import customtkinter as ctk
from tkinter import messagebox, filedialog
import win32com.client
import win32gui
import win32con
import ctypes
import pythoncom
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import shutil
import tempfile
from datetime import datetime
import time as _time
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")
C = {
    # ── Blue theme: cool light canvas, white cards, refined blue accent ──
    "bg": "#eef1f6",
    "card": "#ffffff",
    "card2": "#f4f7fb",
    "header": "#1e293b",
    "accent": "#2f6fed",
    "accent_hover": "#2559c9",
    "accent_light": "#e6eefc",
    "accent_mid": "#86a9ee",
    "accent_line": "#cfe0fa",
    "success": "#15803d",
    "success_bg": "#e6f4ec",
    "success_mid": "#a9dcbd",
    "warning": "#b45309",
    "warning_bg": "#fbf1de",
    "danger": "#dc2626",
    "danger_bg": "#fbeae7",
    "text": "#14203a",
    "text_dim": "#475569",
    "text_muted": "#64748b",
    "border": "#dce3ee",
    "border_light": "#e8edf5",
    "row_hover": "#e9f0fd",
    "log_bg": "#0f172a",
    "log_text": "#c7d7ec",
    "log_hdr": "#0b1120",
    "exist_bg": "#eef3fb",
    "exist_hdr": "#1e293b",
    "new_bg": "#e7f4ed",
    "new_hdr": "#15803d",
    "new_row1": "#f1f9f4",
    "new_row2": "#dff0e6",
    "gold": "#a16207",
    "gold_bg": "#faf4e3",
}

PARAMS = [
    ("LENGTH",          "Case Length"),
    ("DEPTH",           "Case Depth"),
    ("HEIGHT",          "Case Height"),
    ("TOEKICK_HEIGHT",  "Toekick Height"),
    ("TOP_DEPTH",       "Top Depth"),
    ("DIE_BOARD_HEIGHT","DieBoard Height"),
]

PARAM_SHORT = {
    "LENGTH":           "Length",
    "DEPTH":            "Depth",
    "HEIGHT":           "Height",
    "TOEKICK_HEIGHT":   "Toekick",
    "TOP_DEPTH":        "Top Depth",
    "DIE_BOARD_HEIGHT": "DieBoard",
}

EQUATION_TO_PROP = {
    "LENGTH":           "UnitLength",
    "DEPTH":            "UnitDepth",
    "HEIGHT":           "UnitHeight",
    "TOEKICK_HEIGHT":   "UnitBaseHeight",
    "TOP_DEPTH":        "UnitTopDepth",
    # DIE_BOARD_HEIGHT deliberately has NO custom-property mapping: its equation
    # and the D1 dimension are updated, but no custom property is written.
}

# Master BOM column positions (0-indexed for values_only tuples from openpyxl).
# FRONT_PANEL and RACEWAY columns have been REMOVED from the BOM — variant info
# (front-panel type + raceway) now comes from parsing the description string.
MASTER_BOM_COLS = {
    "PART_NUMBER":    1,   # col B
    "DESCRIPTION":    2,   # col C
    "MATERIAL":       3,   # col D
    "LENGTH":         4,   # col E (was 6)
    "DEPTH":          5,   # col F (was 7)
    "HEIGHT":         6,   # col G (was 8)
    "TOEKICK_HEIGHT": 7,   # col H (was 9)
    "TOP_DEPTH":      8,   # col I (was 10)
    "DIE_BOARD_HEIGHT": 9,   # col J — matched in find_in_master_bom via app_names
}

# ─────────────────────────────────────────
# Product-line configuration
# ─────────────────────────────────────────
# Detected at startup from the open assembly's Description custom property:
#   starts with "PF-"  -> PF product line
#   starts with "PT-"  -> PT product line
# Anything else -> error, refuse to run.
PRODUCT_LINES = {
    "PF": {
        "bom_filename":          "PF - Master List of Parts & Assemblies.xlsm",
        "guideline_sheet":       "PF-R Detailed",
        "assy_sheet":            "Master BOM - PF -Sub Assemblies",
        "part_sheet":            "Master BOM - Part Level",
        "ui_label":              "(PF Product Line)",
        "custom_prop_value":     "PF",
        "equation_aliases":      {},
        "design_ref_check_dp":     True,
        "design_ref_check_siltp":  True,
    },
    "PT": {
        "bom_filename":          "PT - Master List of Parts & Assemblies.xlsm",
        "guideline_sheet":       "PT-R Detailed",
        "assy_sheet":            "Master BOM - PT -Sub Assemblies",
        "part_sheet":            "Master BOM - Part Level",
        "ui_label":              "(PT Product Line)",
        "custom_prop_value":     "PT",
        "equation_aliases":      {"OVERALL_DEPTH": "DEPTH", "OVERALL_HEIGHT": "HEIGHT"},
        "design_ref_check_dp":     True,
        "design_ref_check_siltp":  False,
    },
    "PFI": {
        # NOTE: bom_filename and guideline_sheet must match the real file/sheet
        # EXACTLY (openpyxl is case- and space-sensitive). assy_sheet/part_sheet
        # are kept identical to PF per "same sheet names".
        "bom_filename":          "PFI - Master List of Parts & Assemblies.xlsm",
        "guideline_sheet":       "PFI-R Detailed",
        "assy_sheet":            "Master BOM - PFI -Sub Assemblies",
        "part_sheet":            "Master BOM - Part Level",
        "ui_label":              "(PFI Product Line)",
        "custom_prop_value":     "PFI",
        "equation_aliases":      {},      # same equation names as PF
        "design_ref_check_dp":     False, # PFI: skip Deck Pan design-ref lookup
        "design_ref_check_siltp":  False, # PFI: skip Case Top design-ref lookup
    },                                    # -> BOM match, else equation update
}

# Runtime: holds the selected product line and resolved paths. Populated by
# detect_product_line() at startup.
ACTIVE_LINE = {
    "key":           None,   # "PF" or "PT"
    "cfg":           None,   # the PRODUCT_LINES[key] dict
    "bom_path":      None,   # resolved absolute path to the BOM
}

def _tool_dir():
    """Return the directory containing the EXE (when frozen) or this script."""
    import sys as _sys
    if getattr(_sys, "frozen", False):
        return os.path.dirname(_sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_active_cfg():
    """Return the active product-line config dict. Returns None if not set."""
    return ACTIVE_LINE.get("cfg")

def get_active_bom_path():
    """Return the resolved Master BOM path for the active product line."""
    return ACTIVE_LINE.get("bom_path")

MACRO_PATH     = r"\\172.16.208.35\Hillphoenix_Specialty\Barker Specialty\PART NUMBER GENERATION TOOL\New Part Numbering Scheme\NEW PART NUMBER TOOL.swp"
MACRO_MODULE   = "GeneratePartNumber1"
MACRO_FUNCTION = "GenerateNewPartNumber"
QC_MACRO_FILENAME = "QC tools 1.swp"
QC_MACRO_MODULE = "QC_tools"
QC_MACRO_FUNCTION = "WriteExceptionsToSheet"

# ─────────────────────────────────────────
# Drawing date stamp
# ─────────────────────────────────────────
# When a drawing is renamed/renumbered, stamp this custom property with today's
# date. The format matches the DD-MM-YYYY style shown in the drawing's Summary
# Info — change DRAWING_DATE_FORMAT if a different format is wanted, or
# DRAWING_DATE_PROP if the property has a different name.
DRAWING_DATE_PROP   = "DrawnDate"
DRAWING_DATE_FORMAT = "%d-%m-%Y"   # e.g. 11-06-2026

def get_user_display_name():
    """Best-effort Windows display name for the greeting, normalised to
    'First Last' (Windows often returns 'Last, First'). Numeric-only values
    (employee IDs) are treated as 'no name'. Returns '' if nothing usable.
    """
    def _norm(n):
        n = (n or "").strip()
        if not n or n.isdigit():
            return ""                      # blank or an employee id -> not a name
        if "," in n:                       # "Last, First" -> "First Last"
            a = [p.strip() for p in n.split(",", 1)]
            if len(a) == 2 and a[0] and a[1]:
                return f"{a[1]} {a[0]}"
        return n
    # 1. Full display name via the Win32 extended API (NameDisplay = 3)
    try:
        import win32api
        r = _norm(win32api.GetUserNameEx(3))
        if r:
            return r
    except Exception:
        pass
    # 2. full_name from the local user database
    try:
        import win32net
        u = os.environ.get("USERNAME") or ""
        if u:
            info = win32net.NetUserGetInfo(None, u, 2)
            r = _norm((info or {}).get("full_name"))
            if r:
                return r
    except Exception:
        pass
    # 3. login name as a last resort (dropped by _norm if it's all digits)
    try:
        import getpass
        r = _norm(getpass.getuser())
        if r:
            return r
    except Exception:
        pass
    return ""

def get_employee_id():
    """The Windows login / employee id (e.g. '11046403') for the footer."""
    try:
        v = os.getlogin()
        if v and v.strip():
            return v.strip()
    except Exception:
        pass
    return (os.environ.get("USERNAME") or "").strip()

# ─────────────────────────────────────────
# Material + gauge-table updates
# ─────────────────────────────────────────
# Material database and gauge table directory (network shares).
MATERIAL_DB_PATH = r"\\Hp-keo-ecm2\solidworks admin\SolidWorks Admin\Materials\HPx.sldmat"
GAUGE_DIR        = r"\\HP-KEO-ECM2\Solidworks Admin\SolidWorks Admin\Sheet Metal Gauge Tables"

# Map Guideline "Abbreviation" (exact match, case-insensitive) -> list of
# (radio_label, material_name) options the user can choose from. The material
# name is BOTH the material DB entry name AND the gauge file name (without
# extension). To add more part categories later, just add entries here.
MATERIAL_CHOICES = {
    "Front Panel": [
        ("Painted",    "16NEAL"),
        ("Brushed SS", "430-16BRSS"),
    ],
    "Case Top": [
        ("Painted",    "22NEAL"),
        ("Brushed SS", "430-22BRSS"),
    ],
    "Outside Back": [
        ("Painted",    "22NEAL"),
        ("Brushed SS", "430-22BRSS"),
    ],
    "Front Toekick": [
        ("Painted",    "22NEAL"),
        ("Brushed SS", "430-22BRSS"),
    ],
    "End Toekick Cover": [
        ("Painted",    "22NEAL"),
        ("Brushed SS", "430-22BRSS"),
    ],
    "Rear Toekick": [
        ("Painted",    "22NEAL"),
        ("Brushed SS", "430-22BRSS"),
    ],
    "End Panel LH": [
        ("Painted",    "22NEAL"),
        ("Brushed SS", "430-22BRSS"),
    ],
    "End Panel RH": [
        ("Painted",    "22NEAL"),
        ("Brushed SS", "430-22BRSS"),
    ],
    "End Toekick Filler": [
        ("Painted",    "22NEAL"),
        ("Brushed SS", "430-22BRSS"),
    ],
    "Interior": [
        ("Painted",    "18NIZE"),
        ("Brushed SS", "430-16BRSS"),
    ],
}

# Some categories are a SINGLE selector that covers many Guideline abbreviations
# which DON'T all share one gauge. "Interior" is the first: one radio (Painted /
# Brushed), but each member part has its own (painted, brushed) codes below. The
# radio chooses the finish KIND; the member's pair decides the exact material.
# Keys are normalised (lower-case) to match normalize_abbrev().
MATERIAL_GROUP_CODES = {
    "Interior": {
        # 12-gauge shelves
        "shelf standard lh,rh":          ("12HSLAS", "430-12BRSS"),
        "shelf standard center":         ("12HSLAS", "430-12BRSS"),
        "shelf standard center support": ("12HSLAS", "430-12BRSS"),
        # 16/18-gauge deck pan + product stop
        "deck pan":                      ("18NIZE",  "430-16BRSS"),
        "product stop":                  ("18NIZE",  "430-16BRSS"),
        # everything else interior — 22-gauge
        "frame kit dx":                  ("22NEAL",  "430-22BRSS"),
        "frame kit glycol":              ("22NEAL",  "430-22BRSS"),
        "deck pan support":              ("22NEAL",  "430-22BRSS"),
        "baffle assembly":               ("22NEAL",  "430-22BRSS"),
        "shelf standard baffel lh,rh":   ("22NEAL",  "430-22BRSS"),
        "inside back":                   ("22NEAL",  "430-22BRSS"),
        "baffle discharge":              ("22NEAL",  "430-22BRSS"),
        "honeycomb retainer front":      ("22NEAL",  "430-22BRSS"),
    },
    # End Panel LH / RH — single-member groups. The selector lives in
    # MATERIAL_CHOICES (so the radio shows AND the End Panel assembly's own
    # abbreviation 'End Panel LH/RH' is recognised, which is what makes the
    # subtree get descended on a material-only run). Only the OUTSIDE-SKIN
    # abbreviation is a member here, so the skin takes the finish while
    # skin-inside / core / capping (different abbreviations) are left untouched.
    "End Panel LH": {"end panel skin outside lh": ("22NEAL", "430-22BRSS")},
    "End Panel RH": {"end panel skin outside rh": ("22NEAL", "430-22BRSS")},
    "End Toekick Filler": {"toekick filler": ("22NEAL", "430-22BRSS")},
    # Trims follow another panel's finish choice (no card of their own): the Front
    # Panel Trim uses the Front Panel radio's Painted/Brushed kind (22-ga codes),
    # the Case Top Trim uses the Case Top radio's kind (22-ga codes).
    "Front Panel": {"front panel trim": ("22NEAL", "430-22BRSS")},
    "Case Top": {"case top trim": ("22NEAL", "430-22BRSS")},
    # Trims follow an existing selector's finish KIND (no separate UI card):
    # Front Panel Trim follows the Front Panel radio, Case Top Trim follows the
    # Case Top radio. Both map to 22-gauge codes regardless of the parent's gauge.
    "Front Panel": {"front panel trim": ("22NEAL", "430-22BRSS")},
    "Case Top":    {"case top trim":    ("22NEAL", "430-22BRSS")},
}

def gauge_file_path(material_name):
    """Resolve the gauge table .xls path for a given material name."""
    return os.path.join(GAUGE_DIR, material_name + ".xls")

def normalize_abbrev(s):
    """Normalize an abbreviation string for case-insensitive exact match."""
    return (s or "").strip().lower()

# Runtime: holds the user's material radio selection per abbreviation.
# Populated from the UI radio button group before scan/Phase 2.
# Example: {"Front Panel": "16NEAL"}
SELECTED_MATERIALS = {}

# ─────────────────────────────────────────
# Sub-assembly configuration activation
# ─────────────────────────────────────────
# Some sub-assemblies carry configurations that differ only by a refrigerant
# token (DX / GLYCOL) and, for the base frame, a leveling-pad token (LP). We never
# hardcode the full config strings — they vary per assembly. Instead we read the
# component's current referenced config, flip just the relevant token axes per the
# user's selection, and re-point it to whichever config name has the matching
# token SET.
#
# Each "family" is identified by the leading tokens of its Description and declares
# which axes apply:
#   • Base Frame (BF...)        -> refrigerant AND leveling pad
#       e.g. BF,FLT,FP,BAL,BX,DX  /  BF,FLT,FP,BAL,BX,LP,GLYCOL
#   • Frame Kit  (FRM,KIT...)    -> refrigerant ONLY (no LP variants)
#       e.g. FRM,KIT,BAL,BX,DX    /  FRM,KIT,BAL,BX,GLYCOL
# Add another family by appending one row below — no other code changes needed.
REFRIGERANT_TOKENS = ("DX", "GLYCOL")   # mutually-exclusive refrigerant tokens
LP_TOKEN           = "LP"               # leveling-pad token

CONFIG_FAMILIES = [
    {"name": "Base Frame", "lead": ["BF"],         "use_lp": True},
    {"name": "Frame Kit",  "lead": ["FRM", "KIT"], "use_lp": False},
]

# Runtime: the user's configuration selection for THIS run. None on an axis means
# "leave that axis as the component's current config already has it".
#   leveling_pad : True (add LP) / False (remove LP) / None (leave as-is)
#   refrigerant  : "DX" / "GLYCOL" / None (leave as-is)
BASE_FRAME_CFG = {"leveling_pad": None, "refrigerant": None}

# Path -> family dict, for every configurable sub-assembly seen during the scan
# (populated by the walker). Mapped to FINAL paths at apply time, so an assembly
# that was renumbered or replaced still gets its config re-pointed.
CONFIG_TARGETS = {}

def _cfg_tokens(name):
    """Split a config name / description into upper-case, comma-delimited tokens."""
    return [t.strip().upper() for t in (name or "").split(",") if t.strip()]

def config_family_for_desc(desc):
    """Return the CONFIG_FAMILIES entry whose leading tokens match this
    description, or None. Matches on the leading tokens so e.g. BF,FLT,FP,RR,RCWY
    and BF,ANGLD,... both map to Base Frame, and FRM,KIT,... maps to Frame Kit."""
    toks = _cfg_tokens(desc)
    for fam in CONFIG_FAMILIES:
        lead = [t.upper() for t in fam["lead"]]
        if toks[:len(lead)] == lead:
            return fam
    return None

def base_frame_cfg_selected():
    """True only if the user actively chose a configuration this run (at least one
    axis is not None). When False nothing is re-pointed and behaviour is identical
    to before this feature existed."""
    return (BASE_FRAME_CFG.get("leveling_pad") is not None
            or BASE_FRAME_CFG.get("refrigerant") is not None)

# Tracks file paths the tool itself opens during a run (scan + Phase 2), so the
# end-of-run cleanup can close ONLY these and leave unrelated user models open.
# Cleared at the start of each scan. Paths are stored lowercased.
DOCS_OPENED_IN_RUN = set()

def _track_opened_doc(path):
    """Record a path the tool opened, for selective cleanup later."""
    if path:
        try: DOCS_OPENED_IN_RUN.add(path.lower())
        except Exception: pass

def _close_allow_set(entries):
    """Paths the end-of-run cleanup is allowed to close: everything the tool
    opened, PLUS every old and new path of the components it processed. When a
    file is renamed to its new number, it stays open under the NEW path — which
    is not in DOCS_OPENED_IN_RUN (that holds the path it was opened under), so
    without adding the new paths the renamed assemblies would be left open.
    Unrelated models the user already had open are still untouched."""
    s = set(DOCS_OPENED_IN_RUN)
    for e in (entries or []):
        for k in ("old_path", "new_path"):
            p = e.get(k)
            if p:
                try: s.add(p.lower())
                except Exception: pass
    return s

def get_selected_material(abbrev):
    """Return the target material name for an abbreviation, or None.

    Two ways an abbreviation can carry a finish:
      • it IS a selector key (Front Panel, the toekicks, Case Top, Interior) —
        return the code chosen on that radio directly; or
      • it's a MEMBER of a group selector (e.g. any Interior part) — the group's
        radio chose the finish KIND, and the member's own (painted, brushed)
        codes decide the exact gauge."""
    if not abbrev: return None
    norm = normalize_abbrev(abbrev)
    for k in MATERIAL_CHOICES.keys():
        if normalize_abbrev(k) == norm:
            return SELECTED_MATERIALS.get(k)
    for gname, members in MATERIAL_GROUP_CODES.items():
        if norm in members:
            kind = _finish_kind(SELECTED_MATERIALS.get(gname))
            painted, brushed = members[norm]
            if kind == "brushed": return brushed
            if kind == "painted": return painted
            return None
    return None

def selected_fp_material():
    """The material code the user selected for the FRONT PANEL this run (e.g.
    '16NEAL' for Painted, '430-16BRSS' for Brushed SS), or None. Other
    material-bearing parts (toekicks, end cover) are handled per-part and must
    NOT influence front-panel assembly matching, so this returns the Front
    Panel selection specifically rather than 'any selected material'."""
    for k, v in SELECTED_MATERIALS.items():
        if normalize_abbrev(k) == "front panel" and v:
            return v
    return None

def material_for_description(lookup, description):
    """The selected finish for a description.
      • 'SIL,TP…' -> the Case Top selector (its Guideline abbreviation is
        'Case Top Trim', which differs from the selector name, so match by prefix);
      • everything else — including the Deck Pan AND every other Interior part —
        resolves by its Guideline abbreviation, which then routes through the
        Interior group. This is important because 'DP…' is NOT all one gauge:
        'DP,96.5,42…' is the Deck Pan (18NIZE) but 'DP,SP,CVR' is the Deck Pan
        Support (22NEAL), and only the abbreviation tells them apart.
    Returns a material code (e.g. '22NEAL') or None when no finish applies."""
    du = (description or "").strip().upper()
    if du.startswith("SIL,TP"):
        return SELECTED_MATERIALS.get("Case Top")
    return get_selected_material(get_abbreviation(lookup, description))

def _interior_assembly_finish_word(lookup, description):
    """For an Interior-group ASSEMBLY (Frame Kit, Baffle Assembly, Deck Pan, …)
    return the finish WORD the Master BOM should carry — 'Painted' or 'BRSS' —
    because the assembly spans several gauges and a single gauge code would be
    wrong on its row. Returns '' when the description isn't an Interior assembly
    (those keep the normal front-panel-style code)."""
    if not lookup:
        return ""
    ab = normalize_abbrev(get_abbreviation(lookup, description))
    if ab in MATERIAL_GROUP_CODES.get("Interior", {}):
        kind = _finish_kind(SELECTED_MATERIALS.get("Interior"))
        if kind == "painted": return "Painted"
        if kind == "brushed": return "BRSS"
    return ""

def _finish_kind(s):
    """Classify a finish/material string as 'painted', 'brushed', or '' (blank /
    unknown). Used to match the Design Reference column H, which stores finish
    WORDS that vary per sheet — 'Painted', 'Brushed stainless steel', 'BRSS' —
    rather than material codes. Works for codes too (e.g. '430-22BRSS' -> brushed,
    '22NEAL'/'18NIZE' -> painted), so the same selection compares cleanly against
    either a label or a code."""
    u = (s or "").strip().upper()
    if not u:
        return ""
    if "BRUSH" in u or "BRSS" in u or u.startswith("430"):
        return "brushed"
    if "PAINT" in u or "NEAL" in u or "NIZE" in u:
        return "painted"
    return ""

# ─────────────────────────────────────────
# Design Reference Library (Deck Pan + Case Top configurations)
# ─────────────────────────────────────────
DESIGN_REF_PATH = os.path.join(
    os.environ.get("USERPROFILE", r"C:\Users\Default"),
    "OneDrive - Dover Corporation",
    "HPX-SD TEAM - General",
    "Kaizen Ideas",
    "Design Reference Library.xlsx"
)

# Deck Pan Configurations sheet — for descriptions starting with "DP"
DECKPAN_SHEET = "Deck pan Configurations"
DECKPAN_HEADER_ROW = 7   # data starts at row 8
DECKPAN_COLS = {
    "ASSEMBLY": 3,   # Excel col C
    "DESCRIPTION": 4,
    "LENGTH": 5,
    "DEPTH": 6,
    "MATERIAL": 8,   # Excel col H
}

# Case Top Configurations sheet — for descriptions starting with "SIL,TP"
CASETOP_SHEET = "Case Top Configurations"
CASETOP_HEADER_ROW = 2   # data starts at row 3
CASETOP_COLS = {
    "ASSEMBLY": 1,    # Excel col A
    "DESCRIPTION": 3,
    "LENGTH": 4,
    "TOP_DEPTH": 5,   # "Case Top Depth"
    "MATERIAL": 8,    # Excel col H
}

# ─────────────────────────────────────────
# Description-based variant detection
# ─────────────────────────────────────────
# Front-panel and raceway type are parsed from the description string itself.
# Strict matching of these exact tokens (no whitespace between comma and word):
#   ANGLD,FP or FP,ANGLD -> Angled Front Panel
#   FLT,FP   or FP,FLT   -> Flat (Straight) Front Panel
#   RR,RCWY  or RCWY,RR  -> Rear Raceway
#   FR,RCWY  or RCWY,FR  -> Front Raceway
#   BALLAST,BOX          -> Ballast Box raceway
# A part with no front-panel tag is treated as "fits any front panel".
# A part with no raceway tag is treated as "fits any raceway".
# Shared parts use "/" between two variant descriptions, e.g.:
#   "FR,RCWY/BALLAST,BOX"  means part works in both Front Raceway and Ballast Box.

# Front-panel tokens (canonical -> set of token spellings to look for)
_FP_TOKENS = {
    "ANGLED":   ("ANGLD,FP",  "FP,ANGLD",  "ANGLED,FP", "FP,ANGLED"),
    "STRAIGHT": ("FLT,FP",    "FP,FLT",    "FLAT,FP",   "FP,FLAT"),
}
# Raceway tokens (canonical code -> set of token spellings)
_RCW_TOKENS = {
    "FR": ("FR,RCWY", "RCWY,FR"),
    "RR": ("RR,RCWY", "RCWY,RR"),
    "BB": ("BALLAST,BOX", "BAL,BX"),
}

def _desc_pieces(description):
    """Split a description on '/' to handle shared-variant descriptions.
    Returns a list of upper-cased piece strings (each with whitespace stripped)."""
    if not description: return []
    return [p.strip().upper() for p in str(description).split("/") if p.strip()]

# ─────────────────────────────────────────────────────────────────────────────
# Descriptions that ENCODE dimensions and/or finish
#
# A few families carry their case dimensions AND their finish inside the
# Description itself, e.g.
#     DP,48.5,42,2IN_PS,PAINTED,W-O PRICE TAG MOLDING   (deck pan assembly)
#     DP,48.5,42                                        (deck pan part)
#     PS,48.5,2IN                                       (product stop part)
#     SIL,TP,NEAL,SOLAR DIGITAL THERMOMETER             (case top assembly)
# When a run changes those dimensions or that finish, the Description has to be
# rewritten to match. This matters for more than cosmetics: the Master BOM is
# searched BY DESCRIPTION, so a stale description would look up the wrong row
# and could never find the correct existing part.
#
# Two rules, deliberately conservative:
#   1. Only PURE-NUMBER tokens are treated as dimensions. That is what protects
#      "2IN" and "2IN_PS" — they contain letters, so they are never touched.
#      A number is swapped only when it equals one of the ORIGINAL dimension
#      values, and it is replaced by that same dimension's NEW value.
#   2. The family's finish word is swapped both ways (painted <-> brushed).
#      The word differs per family: deck pans say PAINTED, case tops say NEAL,
#      and both say BRSS when brushed.
# Any description not in a listed family is returned untouched.
# ─────────────────────────────────────────────────────────────────────────────
# (prefix, painted word, brushed word, dimensions the description encodes)
# The dimension list mirrors find_in_design_ref: deck pans carry Length+Depth,
# case tops carry Length+Top Depth. Scoping it per family stops a number being
# matched against the wrong parameter when two dimensions share a value.
DESC_MATERIAL_FAMILIES = [
    ("SIL,TP", "NEAL",    "BRSS", ("LENGTH", "TOP_DEPTH")),
    ("DP",     "PAINTED", "BRSS", ("LENGTH", "DEPTH")),
    ("PS",     "PAINTED", "BRSS", ("LENGTH",)),
]

def _material_is_brushed(mat_code):
    """Brushed codes are all 430-xxBRSS. Painted codes (22NEAL, 16NEAL, 18NIZE,
    12HSLAS, ...) never contain BRSS, so this one test covers every group."""
    return "BRSS" in (mat_code or "").upper()

def _desc_family(description):
    """Return (prefix, painted_word, brushed_word, dim_params) if this description
    belongs to a family that encodes dims/finish, else None."""
    du = (description or "").strip().upper()
    if not du: return None
    for fam in DESC_MATERIAL_FAMILIES:
        prefix = fam[0]
        if du == prefix or du.startswith(prefix + ","):
            return fam
    return None

def _fmt_desc_dim(v):
    """Format a dimension the way descriptions carry it: 48.5 -> '48.5',
    42.0 -> '42' (no trailing zeros)."""
    try:    f = float(v)
    except Exception: return str(v)
    return str(int(round(f))) if abs(f - round(f)) < 1e-9 else ("%g" % f)

def build_target_description(description, old_vals, new_vals, target_material=None):
    """The Description this component SHOULD carry after the run.

    old_vals / new_vals : {"LENGTH": 48.5, "DEPTH": 42, ...} before and after.
    target_material     : material code being applied (None if not changing).

    Returns the original description unchanged when the family doesn't encode
    anything, when nothing actually changed, or on any error.
    """
    fam = _desc_family(description)
    if not fam or not description:
        return description
    _prefix, painted_w, brushed_w, dim_params = fam
    old_vals = old_vals or {}; new_vals = new_vals or {}
    want_brushed = _material_is_brushed(target_material) if target_material else None

    out = []
    for tok in str(description).split(","):
        t = tok.strip()
        # (1) dimension token — pure numbers only, so 2IN / 2IN_PS are immune
        try:
            tval = float(t)
            is_num = True
        except Exception:
            is_num = False
        if is_num:
            swapped = None
            for pname in dim_params:
                ov = old_vals.get(pname); nv = new_vals.get(pname)
                if ov is None or nv is None: continue
                try:
                    if abs(float(ov) - tval) < 1e-6:
                        swapped = _fmt_desc_dim(nv)
                        break
                except Exception:
                    continue
            out.append(swapped if swapped is not None else tok)
            continue
        # (2) finish word — swap both ways
        if want_brushed is not None and t.upper() in (painted_w, brushed_w):
            out.append(brushed_w if want_brushed else painted_w)
            continue
        out.append(tok)   # untouched (keeps original spacing)
    return ",".join(out)

def _piece_has_token(piece, token):
    """Check if `token` appears in `piece` with a non-alphanumeric boundary
    (or string edge) before and after. Prevents false matches like 'FR,RCWY'
    matching inside 'OTHERFR,RCWY' while still accepting tokens that follow
    a space, dash, or any other separator — not only commas. The token
    itself can contain commas (e.g. 'ANGLD,FP'); those are part of the token,
    not boundaries.
    """
    if not piece or not token: return False
    start = 0
    while True:
        idx = piece.find(token, start)
        if idx < 0: return False
        # Character just before the match must be non-alphanumeric (or edge)
        before_ok = (idx == 0) or (not piece[idx - 1].isalnum())
        end_idx = idx + len(token)
        # Character just after must be non-alphanumeric (or edge)
        after_ok = (end_idx >= len(piece)) or (not piece[end_idx].isalnum())
        if before_ok and after_ok:
            return True
        start = idx + 1

def extract_front_panel(description):
    """Return 'ANGLED', 'STRAIGHT', or None if the description has no FP tag.
    A part may carry multiple FP tags across slash-separated pieces. If pieces
    disagree, return the union as a set; if only one is found, return the string;
    if none, return None."""
    if not description: return None
    found = set()
    for piece in _desc_pieces(description):
        for canonical, tokens in _FP_TOKENS.items():
            for tok in tokens:
                if _piece_has_token(piece, tok):
                    found.add(canonical)
                    break
    if not found:
        return None
    if len(found) == 1:
        return next(iter(found))
    return found  # rare: a single description encoding both FP variants

def extract_raceways(description):
    """Return a set of canonical raceway codes ({'FR','RR','BB'} subset) found
    in the description. Empty set means 'no raceway tag found' — fits any."""
    found = set()
    if not description: return found
    for piece in _desc_pieces(description):
        for canonical, tokens in _RCW_TOKENS.items():
            for tok in tokens:
                if _piece_has_token(piece, tok):
                    found.add(canonical)
                    break
    return found

def detect_variant_from_subassemblies(model, sw_app, log=None):
    """Detect the model's FP + raceway variant from its TWO designated
    DIRECT-CHILD assemblies, using ONLY the leading token of each one's
    description (NOT tokens that appear anywhere in the string):

        Front-panel assembly description starts with
            FP,FLT    -> Flat front panel        (returned as 'STRAIGHT')
            FP,ANGLD  -> Angled front panel      (returned as 'ANGLED')
        Raceway assembly description starts with
            RR,RCWY   -> Rear raceway            (returned as {'RR'})
            FR,RCWY   -> Front raceway           (returned as {'FR'})
        No direct-child assembly starts with a raceway prefix
            -> Ballast Box                       (returned as {'BB'})

    Only the START of the correct assembly matters. A token appearing LATER in
    a description does NOT count: a flat front-panel assembly that is shared
    across rear-raceway and ballast-box builds reads 'FP,FLT,RR,RCWY/BAL,BX' —
    the 'RR,RCWY' inside it is NOT a raceway assembly and must not make the
    model a rear-raceway build.

    Performance: only DIRECT children are examined (GetComponents(True)), read
    through get_doc so LIGHTWEIGHT components read WITHOUT resolving. Suppressed
    direct children are resolved only if a needed assembly is still missing,
    then restored to their original suppression state. We stop as soon as both
    the front-panel and raceway assemblies are identified.

    Returns (front_panel, raceways): front_panel in {'STRAIGHT','ANGLED',None};
    raceways is a one-element set, {'RR'} | {'FR'} | {'BB'}.
    """
    FP_PREFIX  = (("FP,FLT", "STRAIGHT"), ("FP,ANGLD", "ANGLED"))
    RCW_PREFIX = (("RR,RCWY", "RR"),      ("FR,RCWY", "FR"))

    def _leads_with(desc, prefix):
        # Anchored at the very start, with a non-alphanumeric boundary after the
        # prefix (or end of string) so 'FP,FLT' can't match 'FP,FLTX'.
        if not desc.startswith(prefix):
            return False
        nxt = desc[len(prefix):len(prefix) + 1]
        return (nxt == "") or (not nxt.isalnum())

    def _classify(desc):
        d = (desc or "").strip().upper()
        fp  = next((canon for pre, canon in FP_PREFIX  if _leads_with(d, pre)), None)
        rcw = next((canon for pre, canon in RCW_PREFIX if _leads_with(d, pre)), None)
        return fp, rcw

    # DIRECT children only (True = top-level; False would flatten the whole tree)
    try:
        comps = model.GetComponents(True)
        children = [] if comps is None else list(comps)
    except Exception:
        children = []

    fp_result = None
    rcw_result = None
    pending = []

    def _consume(doc, resolved=False):
        nonlocal fp_result, rcw_result
        try:
            if int(doc.GetType()) != 2:   # only assemblies carry these prefixes
                return
        except Exception:
            pass
        desc = get_description(doc) or ""
        fp, rcw = _classify(desc)
        # (per-direct-child detection lines intentionally not logged — only the
        # final "Assembly variant: ..." conclusion is shown.)
        if fp and not fp_result:   fp_result = fp
        if rcw and not rcw_result: rcw_result = rcw

    # Pass 1 — already-loaded / lightweight direct children (no resolve).
    for comp in children:
        doc, _ = get_doc(sw_app, comp)
        if doc is None:
            pending.append(comp)      # suppressed / not loaded — resolve only if needed
            continue
        _consume(doc)
        if fp_result and rcw_result:
            break

    # Pass 2 — resolve suppressed direct children only if one is still missing,
    # restoring each one's original suppression state afterwards.
    if not (fp_result and rcw_result):
        for comp in pending:
            orig = get_suppress_state(comp)
            try:    comp.SetSuppression2(2)      # 2 = Resolved
            except Exception: pass
            doc, _ = get_doc(sw_app, comp)
            if doc is not None:
                _consume(doc, resolved=True)
            if orig is not None and orig != 2:
                try:    comp.SetSuppression2(orig)   # restore exact prior state
                except Exception: pass
            if fp_result and rcw_result:
                break

    raceways = {rcw_result} if rcw_result else {"BB"}
    if log and rcw_result is None:
        log("    · no raceway assembly (starts with RR,RCWY / FR,RCWY) among "
            "direct children → Ballast Box\n")
    return fp_result, raceways

def variants_match(row_fp, row_rcw, assy_fp, assy_rcw):
    """Decide whether a BOM row's variant tags match the assembly's variant.
    row_fp:  string ('ANGLED'/'STRAIGHT'), a set of those, or None (universal)
    row_rcw: a set of raceway codes (empty = universal)
    assy_fp: string ('ANGLED'/'STRAIGHT') or None (assembly itself has no tag)
    assy_rcw: a set of raceway codes (assembly's raceways; empty = no tag)

    Rules:
      - Front panel:
          row has no FP tag (None)        -> matches any
          row's FP equals assy's FP        -> matches
          row carries a SET containing assy FP -> matches (rare)
          else                              -> reject
      - Raceway:
          row has no raceway tag (empty)   -> matches any
          row's set intersects assy's set  -> matches
          else                              -> reject
    """
    # Front panel
    if row_fp is not None:
        if assy_fp is None:
            # Row demands a specific FP but the assembly has no FP tag.
            # Be conservative: reject.
            return False
        # Normalize both sides to sets for a uniform comparison
        row_set  = row_fp  if isinstance(row_fp,  set) else {row_fp}
        assy_set = assy_fp if isinstance(assy_fp, set) else {assy_fp}
        if not (row_set & assy_set):
            return False
    # Raceway
    if row_rcw:
        if not assy_rcw:
            return False
        if not (row_rcw & assy_rcw):
            return False
    return True


def _bom_match_score(row, description, assy_config=None, required_material=None):
    """Return a confidence score for a candidate BOM row. Higher is a better
    match when several rows satisfy the same description/dimension filter.

    Exact full-string matches outrank shared-variant-piece matches, and a final
    material/variant tie-break keeps the best candidate deterministic even when a
    partial shared-description row is otherwise ambiguous.
    """
    row_desc_raw = row[MASTER_BOM_COLS["DESCRIPTION"]]
    desc_upper = (description or "").strip().upper()
    row_full = str(row_desc_raw or "").strip().upper()
    row_pieces = set(_desc_pieces(row_desc_raw))
    part_pieces = set(_desc_pieces(description))

    score = 0
    if desc_upper and row_full and desc_upper == row_full:
        score += 100
    elif desc_upper and row_pieces and desc_upper in row_pieces:
        score += 80
    elif row_full and part_pieces and row_full in part_pieces:
        score += 70

    if assy_config:
        assy_fp = assy_config.get("front_panel")
        assy_rcw = assy_config.get("raceways") or set()
        if assy_fp or assy_rcw:
            row_fp = extract_front_panel(row_desc_raw)
            row_rcw = extract_raceways(row_desc_raw)
            if variants_match(row_fp, row_rcw, assy_fp, assy_rcw):
                score += 20

    if required_material:
        row_mat = row[MASTER_BOM_COLS["MATERIAL"]]
        row_mat_norm = str(row_mat).strip().upper() if row_mat else ""
        req_mat_norm = (required_material or "").strip().upper()
        if row_mat_norm:
            if _finish_kind(row_mat_norm) == _finish_kind(req_mat_norm):
                score += 10
            else:
                score -= 999

    return score


# ─────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────
def format_value(value):
    try:
        f = float(value)
        return str(int(f)) if f == int(f) else str(f)
    except Exception:
        return str(value)

def get_part_number_from_path(path):
    return os.path.splitext(os.path.basename(path))[0].upper()

def snapshot_folder(folder, ext):
    try:
        return set(f.lower() for f in os.listdir(folder)
                   if f.lower().endswith(ext.lower()) and not f.startswith("~$"))
    except Exception:
        return set()

def find_new_file(folder, before_set, ext):
    new_files = snapshot_folder(folder, ext) - before_set
    return os.path.join(folder, list(new_files)[0]) if new_files else ""

# ─────────────────────────────────────────
# Guideline Excel
# ─────────────────────────────────────────
def _open_for_read(path, **kwargs):
    """Open a workbook for READING even when the file is currently open in Excel.

    Excel keeps an exclusive *write* lock on an open file, which makes a normal
    openpyxl open fail with PermissionError. But it still allows other programs
    to *read/copy* the bytes. So we copy the file to a throwaway temp file first
    and open that copy — the original's lock no longer matters.

    Returns (workbook, temp_path). Pass both to _close_read() when done so the
    temp file gets deleted. Falls back to opening the original directly if the
    copy step itself fails (e.g. a truly exclusive lock), preserving the old
    behaviour in that case.

    NOTE: the copy reflects the file as last *saved*. If someone has unsaved
    edits open in Excel, those aren't seen until they save — fine for the
    reference/BOM workbooks, which aren't edited mid-run.
    """
    tmp = None
    try:
        suffix = os.path.splitext(path)[1] or ".xlsx"
        fd, tmp = tempfile.mkstemp(suffix=suffix)
        os.close(fd)
        shutil.copy2(path, tmp)
        wb = openpyxl.load_workbook(tmp, **kwargs)
        return wb, tmp
    except Exception:
        if tmp and os.path.exists(tmp):
            try: os.remove(tmp)
            except Exception: pass
        # Last resort: open the original directly (will still fail if locked).
        wb = openpyxl.load_workbook(path, **kwargs)
        return wb, None

def _close_read(wb, tmp):
    """Close a workbook opened via _open_for_read and delete its temp copy."""
    try: wb.close()
    except Exception: pass
    if tmp:
        try: os.remove(tmp)
        except Exception: pass

def build_qc_questions():
    """Build the ordered free-text QC prompts from the app's supported fields."""
    questions = [
        (name, f"Profile Attribute - {label}")
        for name, label in PARAMS
    ]
    questions.extend(
        (f"finish:{category}", f"Finish - {category}")
        for category in MATERIAL_CHOICES
    )
    return questions

def save_qc_record(answers, model_title):
    """Write completed QC prompts and answers to a unique Excel record."""
    missing = [
        prompt for key, prompt in build_qc_questions()
        if not str(answers.get(key, "")).strip()
    ]
    if missing:
        raise ValueError("Missing QC answers: " + ", ".join(missing))

    safe_title = "".join(
        char if char.isalnum() or char in " _-" else "_"
        for char in os.path.splitext(model_title or "model")[0]
    ).strip(" ._-")[:80] or "model"
    stem = f"QC_Status_{safe_title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    output_path = os.path.join(_tool_dir(), stem + ".xlsx")
    suffix = 1
    while os.path.exists(output_path):
        output_path = os.path.join(_tool_dir(), f"{stem}_{suffix}.xlsx")
        suffix += 1

    wb = openpyxl.Workbook()
    try:
        ws = wb.active
        ws.title = "QC Record"
        ws.append(("Question", "Answer"))
        for key, prompt in build_qc_questions():
            ws.append((prompt, str(answers[key]).strip()))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 60
        ws["A1"].font = Font(bold=True)
        ws["B1"].font = Font(bold=True)
        wb.save(output_path)
    finally:
        wb.close()
    return output_path

def _workbook_is_open(path):
    """Best-effort: is this workbook currently open in Excel (which would block
    a save)?  Excel keeps a write lock on an open file, so trying to open it for
    writing raises PermissionError/OSError. We open 'r+b' (read+write, no
    truncation) and never write — so this is non-destructive. Returns False if
    the file doesn't exist (nothing to collide with)."""
    if not path or not os.path.exists(path):
        return False
    try:
        with open(path, "r+b"):
            return False
    except (PermissionError, OSError):
        return True
    except Exception:
        return False

# ─────────────────────────────────────────
def load_excel_lookup(excel_path):
    lookup = {}
    cfg = get_active_cfg() or {}
    sheet_name = cfg.get("guideline_sheet", "PF-R Detailed")
    try:
        wb, _tmp = _open_for_read(excel_path, read_only=True, data_only=True)
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[2]: continue
            key = str(row[2]).strip().upper()
            # row[3] is the "Abbreviation" column (col D). We store it on the
            # SAME dict under a key that can't collide with PARAM names.
            abbrev = str(row[3]).strip() if row[3] else ""
            new_entry = {
                "LENGTH":         str(row[4]).strip().upper() == "YES" if row[4] else False,
                "HEIGHT":         str(row[5]).strip().upper() == "YES" if row[5] else False,
                "DEPTH":          str(row[6]).strip().upper() == "YES" if row[6] else False,
                "TOEKICK_HEIGHT": str(row[7]).strip().upper() == "YES" if row[7] else False,
                "TOP_DEPTH":      str(row[8]).strip().upper() == "YES" if row[8] else False,
                "DIE_BOARD_HEIGHT": str(row[9]).strip().upper() == "YES" if len(row) > 9 and row[9] else False,
                "_ABBREV":        abbrev,
            }
            # The same description can appear on multiple rows (e.g. a green
            # section-header row with no abbreviation, plus a detail row that
            # carries "Front Panel"). Don't let a later row with an EMPTY
            # abbreviation clobber an abbreviation we already captured. Merge:
            # keep the first non-empty abbreviation we see.
            if key in lookup:
                existing = lookup[key]
                if not new_entry["_ABBREV"] and existing.get("_ABBREV"):
                    new_entry["_ABBREV"] = existing["_ABBREV"]
                # Also OR the dimension flags so a "Yes" on any row sticks.
                for p in ("LENGTH","HEIGHT","DEPTH","TOEKICK_HEIGHT","TOP_DEPTH","DIE_BOARD_HEIGHT"):
                    new_entry[p] = bool(new_entry[p]) or bool(existing.get(p))
            lookup[key] = new_entry
        _close_read(wb, _tmp)
        return lookup, None
    except Exception as e:
        return {}, str(e)

def _guideline_row(lookup, description):
    """Look up a description's Guideline row.

    Exact match first. If that misses, DP, SIL,TP and PS fall back to their
    GENERIC Guideline entry: the model carries variable suffixes on these (e.g.
    'DP,96.5,42,2IN_PS,PAINTED,W-O PRICE TAG MOLDING', 'SIL,TP,SOLAR DIGITAL
    THERMOMETER', 'PS,96.5,...'), while the Guideline lists them only as 'DP',
    'SIL,TP' and 'PS'. Without this the exact description never matches, the
    part is treated as 'not affected', and its equations/material never update.

    Returns the row dict, or None if nothing matches.
    """
    if not description or description == "N/A":
        return None
    key = description.strip().upper()
    row = lookup.get(key)
    if row is None:
        if key.startswith("SIL,TP"):
            row = lookup.get("SIL,TP")
        elif key.startswith("DP"):
            row = lookup.get("DP")
        elif key.startswith("PS,"):
            # Product Stop: the Guideline lists it generically as 'PS', but the
            # model carries a variable suffix (e.g. 'PS,96.5,...'). A comma-
            # anchored prefix falls back to the generic 'PS' row — same idea as
            # DP / SIL,TP. (A bare 'PS' already matches exactly above.)
            row = lookup.get("PS")
    return row

def get_applicable_params(lookup, description, selected_names, selected_values):
    row = _guideline_row(lookup, description)
    if row is None:
        return [], []
    names, values = [], []
    for n, v in zip(selected_names, selected_values):
        if row.get(n, False):
            names.append(n)
            values.append(v)
    return names, values

# ─── WHY-LOG (diagnostic, removable) ──────────────────────────────────────
# Flip WHY_LOG to False to silence, or delete this block + the two call sites
# (grep "WHY-LOG" and "_why_affected"). It prints, for each part/assembly, the
# exact Guideline row that decided affected/not-affected and which columns were
# YES — so the unsuppress behaviour is never guesswork.
WHY_LOG = True

def _why_affected(lookup, desc, sel_names):
    key = (desc or "").strip().upper()
    row = _guideline_row(lookup, desc)
    if row is None:
        return f"[why] '{desc}': not in Guideline -> not affected (left alone)"
    if lookup.get(key) is not None:
        matched_key, how = key, "exact"
    elif key.startswith("SIL,TP"):
        matched_key, how = "SIL,TP", "prefix-fallback"
    elif key.startswith("DP"):
        matched_key, how = "DP", "prefix-fallback"
    else:
        matched_key, how = key, "?"
    flags = " ".join(f"{p}={'Y' if row.get(p) else 'N'}"
                     for p in ("LENGTH", "DEPTH", "HEIGHT", "TOEKICK_HEIGHT", "TOP_DEPTH", "DIE_BOARD_HEIGHT"))
    matched = [n for n in sel_names if row.get(n, False)]
    verdict = ("AFFECTED by " + ",".join(matched)) if matched else "NOT affected (no changed column is YES)"
    return (f"[why] '{desc}' -> Guideline row '{matched_key}' ({how}): {flags} "
            f"| you are changing: {','.join(sel_names) or '(none)'} -> {verdict}")
# ──────────────────────────────────────────────────────────────────────────

def get_abbreviation(lookup, description):
    """Return the Guideline 'Abbreviation' string for a description, or ''."""
    row = _guideline_row(lookup, description)
    if row is None:
        return ""
    return row.get("_ABBREV", "") or ""

def get_referenced_config(child):
    """The configuration name a component references, read WITHOUT resolving it.
    Used as a cheap description proxy for a SUPPRESSED component (whose model
    doc is not loaded, so get_description can't read it). Returns '' if
    unavailable."""
    try:
        v = child.ReferencedConfiguration
    except Exception:
        v = None
    if callable(v):
        try:    v = v()
        except Exception: v = None
    try:
        return str(v) if v else ""
    except Exception:
        return ""

def should_skip_unaffected_part(lookup, desc, sel_names, sel_values, reliable):
    """Decide — WITHOUT resolving the component — whether a PART is guaranteed
    NOT to be affected by this run, so we can avoid unsuppressing it just to
    read it.

    A part is affected iff (a) a selected dimension applies to it per the
    Guideline, or (b) it is a material-option part whose material is being
    changed this run. Both depend ONLY on the description, so both can be
    evaluated from a cheap description read.

    reliable=True  -> `desc` is the real Description, read from a loaded/
                      lightweight doc; trust it as-is.
    reliable=False -> `desc` is the referenced-configuration proxy for a
                      suppressed part. To be safe we additionally require it to
                      be a RECOGNISED Guideline key, so a misread or blank
                      config name can never cause us to skip a part that
                      actually needed changing. In any uncertain case we return
                      False and the caller resolves exactly as before.

    Returns True only when it is SAFE to skip resolving this part.
    """
    if not desc or desc == "N/A":
        return False
    key = desc.strip().upper()
    if not reliable and key not in lookup:
        return False                       # cheap proxy not recognised — resolve
    app, _ = get_applicable_params(lookup, desc, sel_names, sel_values)
    if app:
        return False                       # a selected dimension applies
    if get_selected_material(get_abbreviation(lookup, desc)) is not None:
        return False                       # material-option part this run (incl. End Panel skin via group)
    return True                            # no dimension, no material — safe to skip

# ─────────────────────────────────────────
# Master BOM
# ─────────────────────────────────────────
def load_master_bom():
    cfg = get_active_cfg()
    bom_path = get_active_bom_path()
    if not cfg or not bom_path:
        return [], [], "Product line not selected"
    try:
        wb, _tmp = _open_for_read(bom_path, read_only=True, data_only=True)
        parts_rows = [r for r in wb[cfg["part_sheet"]].iter_rows(min_row=2, values_only=True) if r[1]]
        assy_rows  = [r for r in wb[cfg["assy_sheet"]].iter_rows(min_row=2, values_only=True)  if r[1]]
        _close_read(wb, _tmp)
        return parts_rows, assy_rows, None
    except Exception as e:
        return [], [], str(e)

def find_in_master_bom(description, app_names, app_values, is_assembly,
                       parts_rows, assy_rows, assy_config=None,
                       required_material=None):
    """Look up a part/assembly in the Master BOM by description + dimensions.

    description : the part's Description property (single — no slashes expected
                  on the part side; slashes appear on the BOM row side for
                  shared-variant entries).
    app_names/app_values : the dimensions to match.
    is_assembly : True searches assy_rows, False searches parts_rows.
    assy_config : optional dict carrying the TOP assembly's variant tags:
                   {"front_panel": "ANGLED"/"STRAIGHT"/None,
                    "raceways":    set of {"FR","RR","BB"} (may be empty)}
                  Used as a tiebreaker when multiple rows match description+dims.
    required_material : optional string. When provided (e.g., for Front Panel
                  parts where the user has chosen a material), only rows whose
                  Material column (case-insensitive) equals this value qualify.

    Returns the matched PART_NUMBER string, or None if no row qualified.
    """
    rows = assy_rows if is_assembly else parts_rows
    desc_upper = (description or "").strip().upper()
    req_mat_norm = (required_material or "").strip().upper() if required_material else None
    # Slash-aware matching. A match exists if ANY of these are true:
    #   (a) part's full description == row's full description (trivial)
    #   (b) part's full description equals one of the row's pieces
    #       (the part is a single-variant, the BOM row is a shared row)
    #   (c) row's full description equals one of the part's pieces
    #       (the part is shared, the BOM row is a single-variant row)
    # IMPORTANT: matching only on a shared sub-piece (e.g. "BAL,BX") is too
    # loose — two unrelated shared parts could each contain "BAL,BX" and
    # falsely match each other. The full description must appear on one side.
    part_pieces = set(_desc_pieces(description))
    matches = []
    for row in rows:
        row_desc_raw = row[MASTER_BOM_COLS["DESCRIPTION"]]
        if not row_desc_raw: continue
        row_pieces = set(_desc_pieces(row_desc_raw))
        row_full = str(row_desc_raw).strip().upper()
        # (a) (b) (c)
        if not (
            (desc_upper and row_full  and desc_upper == row_full) or
            (desc_upper and row_pieces and desc_upper in row_pieces) or
            (row_full  and part_pieces and row_full  in part_pieces)
        ):
            continue
        # Material filter (only applies if caller passed required_material).
        # A BLANK Material cell is a wildcard — it matches any requested
        # material, so leaving the column empty in the BOM does not force a new
        # number. A NON-BLANK cell must match exactly, so a Brushed-SS row is
        # rejected for a Painted request (and vice-versa).
        if req_mat_norm is not None:
            row_mat = row[MASTER_BOM_COLS["MATERIAL"]]
            row_mat_norm = str(row_mat).strip().upper() if row_mat else ""
            if row_mat_norm:
                if is_assembly:
                    # Assembly rows hold a finish WORD ('Painted'/'BRSS') for
                    # Interior assemblies, or a code for FP/toekick assemblies —
                    # so compare by finish KIND. Either form matches, and a
                    # Painted request still rejects a Brushed row (and vice-versa).
                    if _finish_kind(row_mat_norm) != _finish_kind(req_mat_norm):
                        continue
                else:
                    # Parts carry a specific gauge code — exact (gauge matters).
                    if row_mat_norm != req_mat_norm:
                        continue
        # Dimension match
        all_match = True
        for name, value in zip(app_names, app_values):
            col = MASTER_BOM_COLS.get(name)
            if col is None: continue
            row_val = row[col]
            if row_val is None: all_match = False; break
            try:
                if abs(float(row_val) - float(value)) > 0.0001:
                    all_match = False; break
            except Exception:
                all_match = False; break
        if all_match:
            matches.append(row)
    if not matches:
        return None
    if len(matches) == 1:
        return str(matches[0][MASTER_BOM_COLS["PART_NUMBER"]]).strip()

    # Multiple matches — use the variant tags from both sides as a filtering
    # tie-breaker, and fall back to a deterministic confidence score when more
    # than one row still qualifies. This keeps an ambiguous shared-description
    # match from silently returning an arbitrary first row.
    if assy_config:
        assy_fp  = assy_config.get("front_panel")  # 'ANGLED'/'STRAIGHT'/None
        assy_rcw = assy_config.get("raceways") or set()
        # If the assy has no FP or no raceway, treat those as "no constraint".
        if assy_fp or assy_rcw:
            filtered = []
            for r in matches:
                rd = r[MASTER_BOM_COLS["DESCRIPTION"]]
                row_fp  = extract_front_panel(rd)
                row_rcw = extract_raceways(rd)
                if variants_match(row_fp, row_rcw, assy_fp, assy_rcw):
                    filtered.append(r)
            if filtered:
                matches = filtered

    if len(matches) > 1:
        matches = sorted(
            matches,
            key=lambda r: _bom_match_score(r, description, assy_config=assy_config,
                                          required_material=required_material),
            reverse=True,
        )
    return str(matches[0][MASTER_BOM_COLS["PART_NUMBER"]]).strip()

# ─────────────────────────────────────────
# Design Reference Library lookup
# ─────────────────────────────────────────
_DESIGN_REF_CACHE = {}   # sheet_name -> list of value-only row tuples (or None)

def clear_design_ref_cache():
    """Drop the cached Design Ref rows. Call once at the start of each run so a
    long-lived app picks up an edited Design Ref file."""
    _DESIGN_REF_CACHE.clear()

def _design_ref_sheet_rows(sheet_name, log=None):
    """Return all value-only rows for a Design Ref sheet, loading + caching the
    workbook ONCE per run. Subsequent lookups (e.g. DP and SIL,TP in the same
    run) reuse the cache instead of re-opening the file. Returns None on error."""
    if sheet_name in _DESIGN_REF_CACHE:
        return _DESIGN_REF_CACHE[sheet_name]
    rows = None
    _tmp = None
    try:
        wb, _tmp = _open_for_read(DESIGN_REF_PATH, read_only=True, data_only=True)
        try:
            ws = wb[sheet_name]
            rows = [r for r in ws.iter_rows(values_only=True)]
        except Exception as e:
            if log:
                try: log(f"     Design Ref sheet '{sheet_name}' not found: {e}\n")
                except Exception: pass
        finally:
            _close_read(wb, _tmp)
    except Exception as e:
        if _tmp and os.path.exists(_tmp):
            try: os.remove(_tmp)
            except Exception: pass
        if log:
            try: log(f"     Design Ref open error: {e}\n")
            except Exception: pass
    _DESIGN_REF_CACHE[sheet_name] = rows
    return rows

def find_in_design_ref(description, length_val, depth_val, log=None, required_material=None):
    """
    Look up an assembly in the Design Reference Library based on its description prefix.
    - If description starts with "DP"     -> Deck pan Configurations (match Length + Depth)
    - If description starts with "SIL,TP" -> Case Top Configurations (match Length + Top Depth)

    length_val / depth_val are the target values to match (already the user's new value
    for selected params, or current SW value for unselected — caller decides).
    Either can be None, meaning "don't filter on this dimension".

    Returns the matched assembly number (str) or None if not found / not applicable.
    """
    desc_up = (description or "").strip().upper()

    # Decide which sheet applies
    if desc_up.startswith("DP"):
        sheet_name = DECKPAN_SHEET
        cols = DECKPAN_COLS
        header_row = DECKPAN_HEADER_ROW
        dim_col_len = cols["LENGTH"]
        dim_col_dep = cols["DEPTH"]
    elif desc_up.startswith("SIL,TP"):
        sheet_name = CASETOP_SHEET
        cols = CASETOP_COLS
        header_row = CASETOP_HEADER_ROW
        dim_col_len = cols["LENGTH"]
        dim_col_dep = cols["TOP_DEPTH"]
    else:
        return None  # not a DP or SIL,TP assembly — not handled here

    def _logmsg(m):
        if log:
            try: log(m)
            except Exception: pass

    rows = _design_ref_sheet_rows(sheet_name, log)
    if not rows:
        return None

    # 0-indexed positions for values_only tuples
    i_assy = cols["ASSEMBLY"] - 1
    i_len  = dim_col_len - 1
    i_dep  = dim_col_dep - 1
    i_mat  = (cols["MATERIAL"] - 1) if cols.get("MATERIAL") else None
    req_kind = _finish_kind(required_material) if required_material else ""

    last_assy = None  # forward-fill merged assembly numbers
    result = None
    try:
        for row in rows[header_row:]:          # data starts AFTER the header row
            # Forward-fill the assembly number across merged rows
            cell_assy = row[i_assy] if i_assy < len(row) else None
            if cell_assy:
                last_assy = str(cell_assy).strip()
            if not last_assy:
                continue

            row_len = row[i_len] if i_len < len(row) else None
            row_dep = row[i_dep] if i_dep < len(row) else None

            # Match Length if a target length was provided
            if length_val is not None:
                if row_len is None:
                    continue
                try:
                    if abs(float(row_len) - float(length_val)) > 0.0001:
                        continue
                except Exception:
                    continue

            # Match Depth if a target depth was provided
            if depth_val is not None:
                if row_dep is None:
                    continue
                try:
                    if abs(float(row_dep) - float(depth_val)) > 0.0001:
                        continue
                except Exception:
                    continue

            # Finish filter (column H). The Design Ref stores finish WORDS that
            # vary per sheet ('Painted' / 'Brushed stainless steel' / 'BRSS'),
            # so compare by finish KIND (painted vs brushed), not exact text.
            # A blank cell is a wildcard.
            if req_kind and i_mat is not None:
                row_kind = _finish_kind(row[i_mat] if i_mat < len(row) else None)
                if row_kind and row_kind != req_kind:
                    continue

            # First matching row's assembly wins
            result = last_assy
            break
    except Exception as e:
        _logmsg(f"     Design Ref read error: {e}\n")

    return result

def append_to_master_bom(entry, is_assembly):
    cfg = get_active_cfg()
    bom_path = get_active_bom_path()
    if not cfg or not bom_path:
        return False, "Product line not selected"
    try:
        sheet_name = cfg["assy_sheet"] if is_assembly else cfg["part_sheet"]
        material = ""
        old_length = old_depth = old_height = old_toekick = old_top_depth = old_dieboard = ""
        sl_no = 1
        try:
            wb_ro = openpyxl.load_workbook(bom_path, read_only=True, data_only=True)
            ws_ro = wb_ro[sheet_name]
            row_count = 0
            for row in ws_ro.iter_rows(min_row=2, values_only=True):
                if not row[1]: continue
                row_count += 1
                if str(row[1]).strip().upper() == entry["old_number"].upper():
                    material      = row[MASTER_BOM_COLS["MATERIAL"]]       or "NA"
                    old_length    = row[MASTER_BOM_COLS["LENGTH"]]          or ""
                    old_depth     = row[MASTER_BOM_COLS["DEPTH"]]           or ""
                    old_height    = row[MASTER_BOM_COLS["HEIGHT"]]          or ""
                    old_toekick   = row[MASTER_BOM_COLS["TOEKICK_HEIGHT"]]  or ""
                    old_top_depth = row[MASTER_BOM_COLS["TOP_DEPTH"]]       or ""
                    old_dieboard  = (row[MASTER_BOM_COLS["DIE_BOARD_HEIGHT"]]
                                     if len(row) > MASTER_BOM_COLS["DIE_BOARD_HEIGHT"] else "") or ""
            sl_no = row_count + 1
            wb_ro.close()
        except Exception as e:
            return False, f"Read error: {e}"
        try:
            wb = openpyxl.load_workbook(bom_path, keep_vba=True)
        except Exception:
            wb = openpyxl.load_workbook(bom_path)
        ws = wb[sheet_name]
        pv = entry.get("param_values", {})
        # Material for the appended row, in priority order:
        #   1. target_material — when this new number came from a MATERIAL change
        #      (the file already had it written in Phase 2).
        #   2. current_material — the part's ACTUAL material read during the scan
        #      (fills the column for equation-driven new parts too).
        #   3. the material copied from the old BOM row above (last-ditch fallback).
        # Assemblies have no SW material, so they keep the old-row value.
        if not is_assembly:
            if entry.get("apply_material") and entry.get("target_material"):
                material = entry["target_material"]
            elif entry.get("current_material"):
                material = entry["current_material"]
        elif entry.get("fp_material"):
            # Front Panel assembly: show the material of its Front Panel part.
            material = entry["fp_material"]
        # Column layout (no FRONT_PANEL/RACEWAY): A=SL, B=PartNo, C=Desc,
        # D=Material, E=Length, F=Depth, G=Height, H=Toekick, I=TopDepth,
        # J=Dieboard (sub-assemblies only).
        new_row = [
            sl_no,
            entry["new_number"],
            entry["desc"],
            material,
            pv.get("LENGTH",         old_length),
            pv.get("DEPTH",          old_depth),
            pv.get("HEIGHT",         old_height),
            pv.get("TOEKICK_HEIGHT", old_toekick),
            pv.get("TOP_DEPTH",      old_top_depth),
        ]
        if is_assembly:
            new_row.append("")  # DIEBOARD column for assemblies
        else:
            new_row.append(pv.get("DIE_BOARD_HEIGHT", old_dieboard))
        first_empty = None
        for row_idx in range(2, ws.max_row + 2):
            cell_val = ws.cell(row=row_idx, column=2).value
            if not cell_val:
                first_empty = row_idx
                break
        if first_empty is None:
            first_empty = ws.max_row + 1
        for col_idx, val in enumerate(new_row, start=1):
            ws.cell(row=first_empty, column=col_idx, value=val)
        ws.cell(row=first_empty, column=1, value=first_empty - 1)
        try:
            wb.save(bom_path)
            wb.close()
            return True, None
        except Exception as e1:
            xlsx_path = bom_path.replace(".xlsm", "_updated.xlsx")
            try:
                wb.save(xlsx_path)
                wb.close()
                return True, f"Saved as xlsx: {xlsx_path}"
            except Exception as e2:
                wb.close()
                return False, f"Save failed: {e1} | xlsx also failed: {e2}"
    except Exception as e:
        return False, str(e)

def append_all_to_master_bom(new_entries, lookup=None):
    """Append MANY new rows to the Master BOM in a SINGLE open + save.

    new_entries: list of (entry, is_assembly) pairs. Parts and assemblies are
    written to their respective sheets. Returns (ok, err, n_written).

    Replaces calling append_to_master_bom once per item — that re-opened the
    workbook read-only, re-opened it with VBA, and re-saved the whole macro file
    for EVERY new number (2 loads + 1 save each). Here it's one read pass for the
    old-row data and one keep_vba save for all rows. Row-building logic mirrors
    append_to_master_bom exactly.
    """
    cfg = get_active_cfg(); bom_path = get_active_bom_path()
    if not cfg or not bom_path:
        return False, "Product line not selected", 0
    if not new_entries:
        return True, None, 0
    try:
        # ── One read pass: per sheet, map old PartNo -> existing row tuple ──
        old_by_sheet = {}
        try:
            wb_ro = openpyxl.load_workbook(bom_path, read_only=True, data_only=True)
            for _e, _is_assy in new_entries:
                sheet = cfg["assy_sheet"] if _is_assy else cfg["part_sheet"]
                if sheet not in old_by_sheet:
                    m = {}
                    for row in wb_ro[sheet].iter_rows(min_row=2, values_only=True):
                        if not row or not row[1]:
                            continue
                        m[str(row[1]).strip().upper()] = row
                    old_by_sheet[sheet] = m
            wb_ro.close()
        except Exception as e:
            return False, f"Read error: {e}", 0
        # ── One write open ──
        try:
            wb = openpyxl.load_workbook(bom_path, keep_vba=True)
        except Exception:
            wb = openpyxl.load_workbook(bom_path)
        next_row = {}   # sheet -> next empty row index (so appends stack)
        written = 0
        for entry, is_assembly in new_entries:
            sheet = cfg["assy_sheet"] if is_assembly else cfg["part_sheet"]
            ws = wb[sheet]
            if sheet not in next_row:
                fe = None
                for ridx in range(2, ws.max_row + 2):
                    if not ws.cell(row=ridx, column=2).value:
                        fe = ridx; break
                next_row[sheet] = fe if fe is not None else ws.max_row + 1
            old_row = old_by_sheet.get(sheet, {}).get((entry["old_number"] or "").upper())
            if old_row:
                material      = old_row[MASTER_BOM_COLS["MATERIAL"]]      or "NA"
                old_length    = old_row[MASTER_BOM_COLS["LENGTH"]]         or ""
                old_depth     = old_row[MASTER_BOM_COLS["DEPTH"]]          or ""
                old_height    = old_row[MASTER_BOM_COLS["HEIGHT"]]         or ""
                old_toekick   = old_row[MASTER_BOM_COLS["TOEKICK_HEIGHT"]] or ""
                old_top_depth = old_row[MASTER_BOM_COLS["TOP_DEPTH"]]      or ""
                old_dieboard  = (old_row[MASTER_BOM_COLS["DIE_BOARD_HEIGHT"]]
                                 if len(old_row) > MASTER_BOM_COLS["DIE_BOARD_HEIGHT"] else "") or ""
            else:
                material = ""
                old_length = old_depth = old_height = old_toekick = old_top_depth = old_dieboard = ""
            # Material priority — identical to append_to_master_bom.
            if not is_assembly:
                if entry.get("apply_material") and entry.get("target_material"):
                    material = entry["target_material"]
                elif entry.get("current_material"):
                    material = entry["current_material"]
            else:
                # EVERY assembly records the FINISH WORD ('Painted' / 'BRSS') —
                # never a gauge code — since one assembly can span several gauges.
                # Interior assemblies derive it from the Interior selector; all
                # others from the material that changed inside them (fp_material,
                # or the propagated _mat_changed_name), falling back to the finish
                # kind of whatever material the BOM row already carried.
                _word = _interior_assembly_finish_word(lookup, entry.get("desc"))
                if not _word:
                    _src = entry.get("fp_material") or entry.get("_mat_changed_name") or material
                    _k = _finish_kind(_src)
                    _word = "Painted" if _k == "painted" else ("BRSS" if _k == "brushed" else "")
                if _word:
                    material = _word
                elif entry.get("fp_material"):
                    material = entry["fp_material"]
            pv = entry.get("param_values", {})
            r = next_row[sheet]
            new_row = [
                r - 1,
                entry["new_number"],
                entry["desc"],
                material,
                pv.get("LENGTH",         old_length),
                pv.get("DEPTH",          old_depth),
                pv.get("HEIGHT",         old_height),
                pv.get("TOEKICK_HEIGHT", old_toekick),
                pv.get("TOP_DEPTH",      old_top_depth),
            ]
            # DieBoard Height column (J): assemblies stay blank (they don't carry
            # it); parts write their DieBoard value so future BOM lookups can match
            # on it — the updated value if it changed, else the old one.
            if is_assembly:
                new_row.append("")
            else:
                new_row.append(pv.get("DIE_BOARD_HEIGHT", old_dieboard))
            for col_idx, val in enumerate(new_row, start=1):
                ws.cell(row=r, column=col_idx, value=val)
            next_row[sheet] = r + 1
            written += 1
        # ── One save ──
        try:
            wb.save(bom_path); wb.close()
            return True, None, written
        except Exception as e1:
            xlsx_path = bom_path.replace(".xlsm", "_updated.xlsx")
            try:
                wb.save(xlsx_path); wb.close()
                return True, f"Saved as xlsx: {xlsx_path}", written
            except Exception as e2:
                wb.close()
                return False, f"Save failed: {e1} | xlsx also failed: {e2}", 0
    except Exception as e:
        return False, str(e), 0

# ─────────────────────────────────────────
# Custom properties
# ─────────────────────────────────────────
def _is_flat_pattern_cfg(name):
    """True for a sheet-metal flat-pattern configuration (e.g. 'DefaultSM-FLAT-PATTERN'
    or '<config>SM-FLAT-PATTERN')."""
    return "FLAT-PATTERN" in (name or "").upper()

def _delete_all_config_props(mgr, cfg_name, log):
    """Delete every custom property under one configuration's property manager.
    GetNames is exposed as a PROPERTY on this binding (returns a tuple), not a
    method — calling it as GetNames() raises 'tuple object is not callable', which
    is why nothing was deleting. Read it as a property, fall back to method."""
    try:
        names = mgr.GetNames
        if callable(names):        # some builds DO expose it as a method
            names = names()
    except Exception as e:
        log(f"    ! could not read props to delete: {e}\n"); return
    names = list(names) if names else []
    if not names:
        log(f"  Flat pattern '{cfg_name}': no custom properties to delete\n"); return
    log(f"  Flat pattern '{cfg_name}': deleting {len(names)} custom propert"
        f"{'y' if len(names)==1 else 'ies'}\n")
    for nm in names:
        done = False
        for meth in ("Delete2", "Delete"):
            try: getattr(mgr, meth)(nm); done = True; break
            except Exception: continue
        log(f"    - deleted {nm}\n" if done else f"    ! could not delete {nm}\n")

def write_description_only(model, new_description, old_description, log):
    """Rewrite ONLY the Description property. Needed when a component's finish
    changed but no equations did (e.g. a SIL,TP assembly whose description carries
    NEAL/BRSS but no dimensions) — in that case update_custom_properties never runs.
    Only touches the config whose Description is exactly the one we read, so other
    configurations are never clobbered."""
    if not (new_description and old_description) or new_description == old_description:
        return
    try:
        try:   config_names = list(model.GetConfigurationNames())
        except TypeError: config_names = list(model.GetConfigurationNames)
    except Exception as e:
        log(f"  ! Description update skipped: {e}\n"); return
    for config_name in config_names:
        try:
            if _is_flat_pattern_cfg(config_name): continue
            cfg = model.GetConfigurationByName(config_name)
            if not cfg: continue
            mgr = cfg.CustomPropertyManager
            if not mgr: continue
            try:   cur = model.CustomInfo2(config_name, "Description")
            except Exception: cur = None
            if cur and str(cur).strip().upper() == str(old_description).strip().upper():
                mgr.Set2("Description", new_description)
                log(f"  Description -> {new_description}  (config '{config_name}')\n")
        except Exception as e:
            log(f"  ! Description write ('{config_name}'): {e}\n")

def update_custom_properties(model, updated_names, updated_values, log,
                             new_description=None, old_description=None):
    try:
        try:   config_names = list(model.GetConfigurationNames())
        except TypeError: config_names = list(model.GetConfigurationNames)
        except Exception as e: log(f"  Config names error: {e}\n"); return
        part_no = '$PRP:"SW-File Name"'
        for config_name in config_names:
            try:
                config = model.GetConfigurationByName(config_name)
                if not config: continue
                mgr = config.CustomPropertyManager
                if not mgr: continue
                # Flat-pattern configs: never write to them — strip them clean instead.
                if _is_flat_pattern_cfg(config_name):
                    _delete_all_config_props(mgr, config_name, log)
                    continue
                log(f"  Config: '{config_name}'\n")
                def set_prop(name, value):
                    try:   mgr.Set2(name, format_value(value)); log(f"    + {name} = {format_value(value)}\n")
                    except Exception as e: log(f"    ! {name}: {e}\n")
                for eq_name, value in zip(updated_names, updated_values):
                    prop = EQUATION_TO_PROP.get(eq_name)
                    if prop: set_prop(prop, value)
                # DP / PS / SIL,TP descriptions encode the case dims and the finish,
                # so they must be rewritten in step with the update (the Master BOM
                # is searched BY description). Only rewrite the config whose current
                # Description is exactly the one we read and matched on — never
                # clobber a different config that legitimately says something else.
                if new_description and old_description and new_description != old_description:
                    try:    _cur_d = model.CustomInfo2(config_name, "Description")
                    except Exception: _cur_d = None
                    if _cur_d and str(_cur_d).strip().upper() == str(old_description).strip().upper():
                        set_prop("Description", new_description)
                set_prop("Revision", "1.1")
                _cfg = get_active_cfg() or {}
                set_prop("ProductLine", _cfg.get("custom_prop_value", "PF"))
                try:   mgr.Set2("PartNo", part_no); log(f"    + PartNo = {part_no}\n")
                except Exception as e: log(f"    ! PartNo: {e}\n")
            except Exception as e: log(f"  Config error: {e}\n")
    except Exception as e: log(f"  update_custom_properties error: {e}\n")

# ─────────────────────────────────────────
# SolidWorks helpers
# ─────────────────────────────────────────
def _byref_int():
    return win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

def activate_doc(sw_app, path):
    """Activate an open document by path, robust to the ActivateDoc3 overload
    quirk (its errors arg accepts a VARIANT on some builds, a plain int on
    others — the VARIANT form can throw 'Type mismatch'). Returns the activated
    ModelDoc2 or None, so callers can verify instead of failing silently."""
    for errs in (_byref_int(), 0):
        try:
            return sw_app.ActivateDoc3(path, False, 0, errs)
        except Exception:
            continue
    try:
        return sw_app.ActivateDoc2(path, False, 0)
    except Exception:
        return None

def open_doc(sw_app, path, doctype):
    """Open a document (or return it if already open). Returns (doc, opened_here).
    OpenDoc6's byref errors arg throws 'Type mismatch' on some pywin32 bindings,
    so we fall back to the simpler OpenDoc(name, type) which has no byref args."""
    try:
        d = sw_app.GetOpenDocument(path)
        if d is not None:
            return d, False
    except Exception:
        pass
    # try OpenDoc6 (byref VARIANTs), then the simpler OpenDoc as fallback
    try:
        d = sw_app.OpenDoc6(path, doctype, 1, "", _byref_int(), _byref_int())
        if d is not None:
            return d, True
    except Exception:
        pass
    try:
        d = sw_app.OpenDoc(path, doctype)
        if d is not None:
            return d, True
    except Exception:
        pass
    return None, False

def _save(model):
    for method, args in [("Save", ()), ("Save2", (False,)), ("Save3", (1, _byref_int(), _byref_int()))]:
        try:   getattr(model, method)(*args); return True
        except Exception: continue
    return False

def _save_as_job(sw_app, model, cur_path, job, log):
    """Save the open top assembly As a job-numbered file in the SAME folder, and
    switch the working document to it (references are preserved — only the top
    assembly is renamed). Returns the new path on success, else "" (run continues
    on the current file). Blank job number is a no-op.

    Robustness: the 3DEXPERIENCE SaveAs sometimes RAISES a COM 'Type mismatch'
    even though it actually saved. So we DON'T trust the return value / exception —
    we check the ground truth afterwards: is the active doc now pointing at the job
    file, or does the job file now exist on disk? If so, it's a success.
    """
    job = (job or "").strip()
    if not job:
        return ""
    if not cur_path:
        log("  Job number given but the model has no saved path — skipping Save As\n")
        return ""
    folder = os.path.dirname(cur_path)
    ext    = os.path.splitext(cur_path)[1] or ".SLDASM"
    new_path = os.path.join(folder, job + ext)
    _np = os.path.normcase(os.path.abspath(new_path))
    if _np == os.path.normcase(os.path.abspath(cur_path)):
        log(f"  Working file is already {os.path.basename(new_path)} — no Save As needed\n")
        return ""
    log(f"  Saving as job: {os.path.basename(cur_path)} -> {os.path.basename(new_path)}\n")
    # Try both call forms; ignore their thrown errors — we verify by result below.
    for form, call in (
        ("Extension.SaveAs", lambda: model.Extension.SaveAs(new_path, 0, 1, None, _byref_int(), _byref_int())),
        ("SaveAs3",          lambda: model.SaveAs3(new_path, 0, 1)),
    ):
        try:
            call()
            break
        except Exception as e:
            log(f"  (SaveAs via {form} raised: {e} — verifying whether it saved anyway)\n")
    # Ground-truth check 1: what path is the active doc on now?
    now = ""
    try:
        adoc = sw_app.ActiveDoc
        if adoc is not None:
            try:    now = adoc.GetPathName()
            except TypeError: now = adoc.GetPathName
    except Exception:
        now = ""
    if now and os.path.normcase(os.path.abspath(now)) == _np:
        log(f"  Working file is now {os.path.basename(new_path)}\n")
        return new_path
    # Ground-truth check 2: was the job file WRITTEN just now? (Not merely
    # present — on a rerun the file can exist from a previous run while this
    # SaveAs actually failed. Requiring a fresh mtime avoids adopting a stale
    # file while sw_model still points at the original document.)
    try:
        on_disk = os.path.exists(new_path) and (_time.time() - os.path.getmtime(new_path) < 120)
    except Exception:
        on_disk = False
    if on_disk:
        log(f"  Job file saved on disk ({os.path.basename(new_path)}); adopting it as the working file\n")
        return new_path
    log("  ! SaveAs did not produce the job file — staying on the current file\n")
    return ""

def safe_close(sw_app, doc, main_path, log, label):
    try:
        try:   doc_path = doc.GetPathName()
        except TypeError: doc_path = doc.GetPathName
        if doc_path.lower() == main_path.lower(): return
        try:   title = doc.GetTitle()
        except TypeError: title = doc.GetTitle
        sw_app.CloseDoc(title); log(f"  Closed: {label}\n")
    except Exception as e: log(f"  Close warning ({label}): {e}\n")

def get_description(model_doc):
    # Get the list of actual configuration names for this document
    try:
        try:   config_names = list(model_doc.GetConfigurationNames())
        except TypeError: config_names = list(model_doc.GetConfigurationNames)
    except Exception:
        config_names = []

    # Default behaviour: read the Description from the "Default" configuration.
    # Only fall back to the first configuration when there is NO "Default" config.
    has_default = any(c.strip().lower() == "default" for c in config_names)
    if has_default:
        try:
            val = model_doc.CustomInfo2("Default", "Description")
            if val and val.strip():
                return val.strip()
        except Exception:
            pass
    else:
        # No Default config -> use the FIRST configuration in the list
        if config_names:
            try:
                val = model_doc.CustomInfo2(config_names[0], "Description")
                if val and val.strip():
                    return val.strip()
            except Exception:
                pass

    # Last resort: file-level / configuration-independent property (empty config name)
    try:
        val = model_doc.CustomInfo2("", "Description")
        if val and val.strip():
            return val.strip()
    except Exception:
        pass
    return "N/A"

# ─────────────────────────────────────────
# Material + gauge helpers
# ─────────────────────────────────────────
def _material_differs(current_mat, target_mat):
    """Decide whether the current material genuinely differs from the target.

    Handles the messy reality of reading the 'Material' custom property:
      - If we couldn't read anything (empty), we can't prove a difference, so
        return False. This avoids BOTH the 'same material says updated' bug and
        spurious new numbers when the read momentarily fails.
      - If the read returned a FORMULA string (e.g. '"SW-Material@@Default@..."')
        instead of a resolved name, we also can't prove a difference → False.
      - Otherwise compare case-insensitively.
    """
    if not target_mat:
        return False
    cur = (current_mat or "").strip()
    if not cur:
        return False
    # Unresolved formula — not a real material name we can compare against.
    if cur.startswith('"SW-Material') or cur.upper().startswith("SW-MATERIAL"):
        return False
    return cur.upper() != target_mat.strip().upper()

def get_current_material(model_doc, log=None, indent=""):
    """Read the part's current material name for its active configuration.

    Strategy, most reliable first:
      1. GetMaterialPropertyName2 — the DIRECT SolidWorks material API and the
         exact counterpart to SetMaterialPropertyName2 used to WRITE material.
         Returns the applied material name (e.g. '16NEAL') directly, with no
         formula to resolve. Because this is the same thing we change, it is the
         most reliable to read — it works even when the part has no 'Material'
         custom property at all.
      2. The 'Material' custom property (Get5/Get4/Get3 resolved value) — for
         parts that only expose the material as a config-specific custom property
         whose stored value is a formula like "SW-Material@@<config>@<file>".
      3. CustomInfo2 raw text (skipped if it is an unresolved formula).

    Returns the resolved string (e.g. '430-16BRSS') or '' if nothing is found.
    Pass `log` to record which methods were tried when the read comes back empty,
    so an empty read can be diagnosed from the output log.
    """
    if model_doc is None:
        return ""

    def _dbg(msg):
        if log is not None:
            try: log(f"{indent}     [material-read] {msg}\n")
            except Exception: pass

    # Determine the active config name (material is config-specific).
    # pywin32 sometimes needs () and sometimes not — try both shapes.
    active_cfg_name = ""
    try:
        try:    active_cfg = model_doc.GetActiveConfiguration()
        except TypeError: active_cfg = model_doc.GetActiveConfiguration
        if active_cfg:
            try:    active_cfg_name = active_cfg.Name
            except Exception: active_cfg_name = ""
    except Exception:
        active_cfg_name = ""

    # ── 1. Direct SolidWorks material API (preferred; symmetric with write) ──
    # GetMaterialPropertyName2(ConfigName, [out] Database) returns the material
    # NAME. The out Database param is NOT optional, and passing a plain "" fails
    # with a type mismatch on this pywin32 build, so the reliable call passes a
    # VT_BYREF|VT_BSTR VARIANT for it. (Confirmed via material_debug.py, which
    # returned '16NEAL' only with the VARIANT style.) We still try the simpler
    # shapes as fallbacks for other environments, across the active config, the
    # file-level (""), and every named config. Anything that looks like a
    # .sldmat database name is discarded.
    def _looks_like_db(s):
        s2 = (s or "").strip().lower()
        return (not s2) or s2.endswith(".sldmat") or "\\" in s2 or "/" in s2

    def _call_variant(cfg):
        db = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")
        return model_doc.GetMaterialPropertyName2(cfg, db)   # returns the name
    def _call_one(cfg):
        return model_doc.GetMaterialPropertyName2(cfg)
    def _call_two(cfg):
        return model_doc.GetMaterialPropertyName2(cfg, "")

    # Config candidates: active first, then file-level, then every named config.
    cfg_candidates = [active_cfg_name, ""]
    try:
        try:    _all_cfgs = list(model_doc.GetConfigurationNames())
        except TypeError: _all_cfgs = list(model_doc.GetConfigurationNames)
        cfg_candidates += _all_cfgs
    except Exception:
        pass
    seen = set(); cfgs_to_try = []
    for c in cfg_candidates:
        key = c or ""
        if key not in seen:
            seen.add(key); cfgs_to_try.append(key)

    for cfg in cfgs_to_try:
        for caller in (_call_variant, _call_one, _call_two):
            try:
                res = caller(cfg)
            except Exception:
                continue
            cand = ""
            if isinstance(res, (tuple, list)):
                for item in res:
                    if isinstance(item, str) and item.strip() and not _looks_like_db(item):
                        cand = item.strip(); break
            elif isinstance(res, str) and res.strip() and not _looks_like_db(res):
                cand = res.strip()
            if cand:
                return cand
    _dbg(f"GetMaterialPropertyName2 empty (cfgs={cfgs_to_try!r})")

    # ── 2. 'Material' custom property (resolved value) ──
    mgrs_to_try = []
    try:
        ext = model_doc.Extension
        if ext is not None:
            if active_cfg_name:
                try:
                    m = ext.CustomPropertyManager(active_cfg_name)
                    if m is not None: mgrs_to_try.append(m)
                except Exception: pass
            # Also the file-level (empty string) as a fallback
            try:
                m = ext.CustomPropertyManager("")
                if m is not None: mgrs_to_try.append(m)
            except Exception: pass
    except Exception:
        pass

    for mgr in mgrs_to_try:
        # Get5/Get4/Get3 all return (retval, valOut, resolvedValOut[, wasResolved]).
        # Prefer the resolved value (res[2]); only use the raw value (res[1]) if
        # it is not an unresolved "SW-Material@@..." formula.
        for getter in ("Get5", "Get4", "Get3"):
            try:
                fn = getattr(mgr, getter, None)
                if fn is None: continue
                res = fn("Material", False)
                if isinstance(res, tuple):
                    if len(res) >= 3 and isinstance(res[2], str) and res[2].strip():
                        return res[2].strip()
                    if len(res) >= 2 and isinstance(res[1], str) and res[1].strip():
                        val = res[1].strip()
                        if not val.startswith('"SW-Material') and not val.upper().startswith("SW-MATERIAL"):
                            return val
            except Exception: pass
    _dbg("'Material' custom property empty or unresolved")

    # ── 3. Last-ditch: CustomInfo2 raw text (skip unresolved formulas) ──
    try:
        raw = model_doc.CustomInfo2(active_cfg_name or "", "Material")
        if raw and not raw.strip().startswith('"SW-Material'):
            return raw.strip()
    except Exception: pass

    _dbg("no material found via any method")
    return ""

def apply_material_and_gauge(model_doc, material_name, log, indent=""):
    """Set material on every configuration and update the Sheet-Metal feature's
    gauge table to the file matching `material_name`. Returns True on success.

    Mirrors the standalone Material script's approach. Mistakes are logged but
    do not raise — the caller decides whether to continue.
    """
    if model_doc is None:
        log(f"{indent}  ! Material: model is None\n"); return False
    gauge_path = gauge_file_path(material_name)
    if not os.path.isfile(gauge_path):
        log(f"{indent}  ! Gauge table file not found: {gauge_path}\n")
        return False
    # 1. Material — apply to every config so the part is consistent
    try:
        try:   config_names = list(model_doc.GetConfigurationNames())
        except TypeError: config_names = list(model_doc.GetConfigurationNames)
    except Exception:
        config_names = []
    if not config_names: config_names = [""]   # fall back to file-level
    material_ok = True
    for cfg in config_names:
        try:
            result = model_doc.SetMaterialPropertyName2(cfg, MATERIAL_DB_PATH, material_name)
            if result is False:
                log(f"{indent}  ! SetMaterial returned False for cfg '{cfg}'\n")
                material_ok = False
        except Exception as e:
            log(f"{indent}  ! SetMaterial '{cfg}': {e}\n")
            material_ok = False
    # 2. Gauge table — only updates the Sheet-Metal feature
    gauge_ok = True
    try:
        sw_feat = model_doc.FeatureByName("Sheet-Metal")
        if sw_feat is None:
            log(f"{indent}  ! No Sheet-Metal feature on this part — gauge skipped\n")
            gauge_ok = False
        else:
            # Match the standalone script's pattern: property access, not call
            try:   sw_data = sw_feat.GetDefinition
            except TypeError: sw_data = sw_feat.GetDefinition()
            try:
                sw_data.SetUseGaugeTable(True, gauge_path)
            except Exception as e:
                log(f"{indent}  ! SetUseGaugeTable: {e}\n")
                gauge_ok = False
            try:
                sw_feat.ModifyDefinition(sw_data, model_doc, pythoncom.Nothing)
            except Exception as e:
                log(f"{indent}  ! ModifyDefinition: {e}\n")
                gauge_ok = False
    except Exception as e:
        log(f"{indent}  ! Sheet-Metal feature lookup: {e}\n")
        gauge_ok = False
    if material_ok and gauge_ok:
        log(f"{indent}  + Material={material_name}, gauge applied\n")
        return True
    if material_ok:
        log(f"{indent}  ~ Material={material_name} set; gauge update had issues\n")
        return True   # material at least got applied
    return False

def _match_config_by_tokens(anchor_name, names, log, indent="  ", apply_lp=True):
    """Given an anchor config name and the list of available config names, return
    the config name whose token SET matches the user's selection. Copies the
    anchor's tokens, swaps the refrigerant token, and (only when apply_lp) toggles
    the LP token per BASE_FRAME_CFG, then finds the name with the matching set.
    Returns None if the configs aren't switchable (no DX/GLYCOL) or nothing matches."""
    if not names:
        return None
    # Only switchable if its configs carry a refrigerant token — guards against
    # acting on the wrong assembly.
    if not any(any(t in _cfg_tokens(n) for t in REFRIGERANT_TOKENS) for n in names):
        log(f"{indent}No DX/GLYCOL configs — not switchable\n")
        return None
    if anchor_name not in names:
        anchor_name = names[0]
    target = _cfg_tokens(anchor_name)
    # Refrigerant axis (applies to every family)
    want_ref = BASE_FRAME_CFG.get("refrigerant")
    if want_ref:
        target = [t for t in target if t not in REFRIGERANT_TOKENS]
        target.append(want_ref.upper())
    # Leveling-pad axis (only for families whose configs have LP variants)
    if apply_lp:
        want_lp = BASE_FRAME_CFG.get("leveling_pad")
        if want_lp is True and LP_TOKEN not in target:
            target.append(LP_TOKEN)
        elif want_lp is False:
            target = [t for t in target if t != LP_TOKEN]
    target_set = set(target)
    match = next((n for n in names if set(_cfg_tokens(n)) == target_set), None)
    if match is None:
        log(f"{indent}No config matches {sorted(target_set)} (have: {names})\n")
        return None
    return match

def _config_names_for_component(sw_app, comp, path):
    """Configuration names of a component's document. GetModelDoc2 is unreliable
    on some components (it raises 'Member not found' / returns None when the doc
    isn't fully loaded as a dispatchable model), so fall back to reading the
    config list straight from the FILE — the proven approach from the probe.
    Returns [] only if nothing works."""
    # 1) try the loaded model doc (cheap when it works)
    try:    md = comp.GetModelDoc2()
    except Exception: md = None
    if md is not None:
        try:
            try:    n = list(md.GetConfigurationNames())
            except TypeError: n = list(md.GetConfigurationNames)
            if n: return n
        except Exception: pass
    # 2) fall back to the file: reuse the already-open doc, else open it read-only
    if not path:
        return []
    try:
        doc = sw_app.GetOpenDocument(path)
    except Exception:
        doc = None
    if doc is None:
        try:
            dtype = 2 if path.lower().endswith(".sldasm") else 1   # assembly vs part
            doc = sw_app.OpenDoc6(path, dtype, 1, "", _byref_int(), _byref_int())
            _track_opened_doc(path)
        except Exception:
            doc = None
    if doc is not None:
        try:
            try:    return list(doc.GetConfigurationNames())
            except TypeError: return list(doc.GetConfigurationNames)
        except Exception: pass
    return []

def apply_subassembly_configs(sw_app, main_path, sw_model, all_entries, log):
    """Re-point each configurable sub-assembly COMPONENT in the main assembly to
    the user's chosen configuration, by setting its ReferencedConfiguration. This
    is what actually makes the MAIN assembly use the variant.

    Handles every CONFIG_FAMILIES entry: the base frame switches on refrigerant +
    leveling pad, the frame kit on refrigerant only (use_lp=False). The main
    references these by NAMED config, so changing a file's active config does
    nothing — the component's referenced config is the control that matters.
    Setting it only re-points to an already-existing config, so it can't corrupt
    stored suppression states. Run as the LAST step, after re-suppress. No-op
    unless the user chose a configuration.
    """
    if not base_frame_cfg_selected():
        return
    if not CONFIG_TARGETS:
        log("  Sub-assembly config: no configurable sub-assembly found in the tree\n")
        return
    # old_path -> final path (renumber/replace set new_path; "" = unchanged).
    final_of = {e["old_path"].lower(): (e.get("new_path") or e["old_path"])
                for e in all_entries if e.get("old_path")}
    # final-path(lower) -> family dict
    targets = {}
    for p, fam in CONFIG_TARGETS.items():
        fp = final_of.get(p.lower(), p)
        if fp:
            targets[fp.lower()] = fam
    log("=" * 52 + "\n SUB-ASSEMBLY CONFIGURATION\n" + "=" * 52 + "\n")
    sel = []
    if BASE_FRAME_CFG.get("leveling_pad") is True:    sel.append("With Leveling Pad")
    elif BASE_FRAME_CFG.get("leveling_pad") is False: sel.append("Without Leveling Pad")
    if BASE_FRAME_CFG.get("refrigerant"): sel.append(BASE_FRAME_CFG["refrigerant"])
    log(f"  Selection: {', '.join(sel) if sel else '(none)'}\n")
    # Re-acquire the LIVE main document from the app before touching it. After a
    # long run + saves (and a job Save-As), the sw_model handle passed in can be
    # stale — reading it throws "object disconnected", which previously bubbled up
    # into cleanup and closed the assembly. We activate the main by path, take the
    # doc the app hands back, and if we can't get a live handle we skip this step
    # cleanly (no changes, no crash, assembly stays open).
    live = None
    try:    live = activate_doc(sw_app, main_path)
    except Exception: pass
    if live is None:
        try:    live = sw_app.ActiveDoc
        except Exception: live = None
    if live is None:
        log("  ! Main assembly is not available to re-point configs — skipping this step (no changes made)\n" + "=" * 52 + "\n")
        return
    try:
        comps = live.GetComponents(False)
    except Exception as e:
        log(f"  ! Could not read components — skipping sub-assembly config (no changes made): {e}\n" + "=" * 52 + "\n"); return
    comps = [] if comps is None else list(comps)
    any_changed = False
    matched = 0
    for c in comps:
        try:
            try:    cpath = c.GetPathName()
            except TypeError: cpath = c.GetPathName
        except Exception:
            cpath = ""
        fam = targets.get((cpath or "").lower())
        if fam is None:
            continue
        matched += 1
        base = os.path.basename(cpath)
        # Config list straight from the file (GetModelDoc2 is unreliable here),
        # plus the component's current referenced config as the anchor.
        names = _config_names_for_component(sw_app, c, cpath)
        if not names:
            log(f"  ! [{fam['name']}] {base}: could not read its configurations — skipping\n")
            continue
        cur_ref = get_referenced_config(c) or names[0]
        # Frame Kit switches on refrigerant only (use_lp=False); base frame on both.
        target = _match_config_by_tokens(cur_ref, names, log, indent="    ",
                                         apply_lp=fam.get("use_lp", True))
        if not target:
            continue
        if target == cur_ref:
            log(f"  [{fam['name']}] {base}: already references '{cur_ref}' — no change\n")
            continue
        try:
            c.ReferencedConfiguration = target
            any_changed = True
            log(f"  [{fam['name']}] {base}: referenced config '{cur_ref}' -> '{target}'\n")
        except Exception as e:
            log(f"  ! [{fam['name']}] {base}: failed to set referenced config: {e}\n")
    if matched == 0:
        log("  ! No configurable sub-assembly found among the main assembly's components\n")
    if any_changed:
        # EditRebuild3 first (what worked in the probe), ForceRebuild3 as backup,
        # so the graphics + tree pick up the new referenced config.
        try:
            live.EditRebuild3()
        except Exception:
            try: live.ForceRebuild3(True)
            except Exception: pass
        if _save(live): log("  Saved main (sub-assembly config applied)\n")
        else:           log("  ! Save main failed\n")
    elif matched:
        log("  No sub-assembly component needed re-pointing\n")
    log("=" * 52 + "\n")

def get_eq_count(eq_mgr):
    try:   return int(eq_mgr.GetCount())
    except TypeError: return int(eq_mgr.GetCount)

def _active_aliases():
    """Return the equation-name alias map for the active product line.
    PF returns {}. PT returns {'OVERALL_DEPTH':'DEPTH','OVERALL_HEIGHT':'HEIGHT'}."""
    cfg = get_active_cfg() or {}
    return cfg.get("equation_aliases", {}) or {}

def build_eq_index(eq_mgr):
    index = {}
    aliases = _active_aliases()
    for i in range(get_eq_count(eq_mgr)):
        try:   eq = eq_mgr.Equation(i)
        except TypeError: eq = eq_mgr.Equation
        if eq:
            try:
                raw_name = eq.split('"')[1].upper()
                # Index the raw name as it appears in the equation
                index[raw_name] = i
                # Also index the alias target (e.g. OVERALL_DEPTH -> DEPTH) so
                # callers asking for "DEPTH" find the OVERALL_DEPTH equation.
                aliased = aliases.get(raw_name)
                if aliased and aliased not in index:
                    index[aliased] = i
            except IndexError: pass
    return index

def get_current_equation_values(model):
    current = {}
    aliases = _active_aliases()
    try:
        eq_mgr = model.GetEquationMgr
        for i in range(get_eq_count(eq_mgr)):
            try:   eq = eq_mgr.Equation(i)
            except TypeError: eq = eq_mgr.Equation
            if eq and '"' in eq and '=' in eq:
                try:
                    parts = eq.split('=')
                    name  = parts[0].strip().strip('"').upper()
                    raw   = parts[1].strip().split('"')[0].strip()
                    # Translate alias to canonical (e.g. OVERALL_DEPTH -> DEPTH)
                    name = aliases.get(name, name)
                    # `raw` is the equation TEXT. For a driven equation (e.g.
                    # DIE_BOARD_HEIGHT = ( TOEKICK_HEIGHT + 15.3125 )) the text
                    # isn't a number — parsing it yields junk like "(". When the
                    # text isn't a plain number, read the EVALUATED value instead.
                    def _is_number(t):
                        try: float(t); return True
                        except Exception: return False
                    if not _is_number(raw):
                        try:    ev = eq_mgr.Value(i)
                        except TypeError: ev = eq_mgr.Value
                        except Exception: ev = None
                        if ev is not None:
                            try:    raw = format_value(float(ev))
                            except Exception: pass
                    for p_name, _ in PARAMS:
                        if name == p_name.upper(): current[p_name] = raw
                except Exception: pass
    except Exception: pass
    return current

def get_doc(sw_app, child):
    try:   path = child.GetPathName()
    except TypeError: path = child.GetPathName
    except Exception: path = ""
    if not path: return None, ""
    doc = None
    try:   doc = sw_app.GetOpenDocument(path)
    except Exception: pass
    if doc is None:
        try:   doc = child.GetModelDoc2()
        except Exception: pass
    return doc, path

def get_children_from_model(model):
    # DIRECT children only — GetComponents(True). The walk descends the tree by
    # RECURSING into each affected sub-assembly, so it must see ONE LEVEL at a
    # time. GetComponents(False) returns the WHOLE flattened tree, which means a
    # pruned sub-assembly's children still appear here as top-level entries and
    # get resolved/unsuppressed anyway — defeating the subtree prune entirely.
    # With (True), a pruned or replaced-as-unit sub-assembly's descendants are
    # never enumerated, hence never unsuppressed. The recursion (for needs_update
    # / FP-material assemblies) is what reaches deeper levels.
    try:
        comps = model.GetComponents(True)
        return [] if comps is None else list(comps)
    except Exception: return []

def _prop_call(obj, name, *args):
    """Access obj.<name> whether this binding exposes it as a method or property.
    (Same defensive pattern used across the tool for GetType/GetPathName/etc.)"""
    a = getattr(obj, name)
    try:    return a(*args)
    except TypeError: return a
    except Exception:
        try: return a
        except Exception: return None

def shift_d1_for_dieboard(model, old_db, new_db, log):
    """When DieBoard Height changes, the D1@END VIEW sketch dimension must move by
    the SAME delta (however much DieBoard moves, D1 moves the same — verified live).
    Reads the current D1 at runtime and shifts it; no hard-coded offset. Returns
    True if a D1 was changed. Only touches 'D1@END VIEW' (NOT 'D1@END VIEW (STD
    BASE)' or other D1s). Never raises."""
    try:
        delta = float(new_db) - float(old_db)
    except Exception:
        return False
    if abs(delta) < 1e-9:
        return False
    # Resolve the exact dimension. Parameter() wants the config-qualified name;
    # try the plain name first, then discover the real full name by walking
    # display dimensions (some builds need 'D1@END VIEW@<file>.Part').
    target = "D1@END VIEW"
    dim = None
    try:    dim = _prop_call(model, "Parameter", target)
    except Exception: dim = None
    resolved_name = target
    if dim is None:
        # walk features -> display dims to find 'D1@END VIEW' (ignoring STD BASE)
        try:
            feat = _prop_call(model, "FirstFeature")
            while feat and dim is None:
                dd = None
                try: dd = _prop_call(feat, "GetFirstDisplayDimension")
                except Exception: dd = None
                while dd:
                    try:
                        d = _prop_call(dd, "GetDimension2", 0) or _prop_call(dd, "GetDimension")
                        fn = _prop_call(d, "FullName") or _prop_call(d, "Name") or ""
                        fu = str(fn).upper()
                        if "D1@END VIEW" in fu and "STD BASE" not in fu:
                            dim = _prop_call(model, "Parameter", str(fn)) or d
                            resolved_name = str(fn)
                            break
                    except Exception: pass
                    try: dd = _prop_call(dd, "GetNext")
                    except Exception: dd = None
                feat = _prop_call(feat, "GetNextFeature")
        except Exception:
            pass
    if dim is None:
        log(f"    (D1@END VIEW not found — D1 not shifted)\n")
        return False
    try:
        cur_m = _prop_call(dim, "SystemValue")       # meters
        cur_in = cur_m / 0.0254
        new_in = cur_in + delta
        dim.SystemValue = new_in * 0.0254
        log(f"    D1 shifted: {resolved_name} {cur_in:.4f} -> {new_in:.4f} in (Δ{delta:+.4f})\n")
        return True
    except Exception as e:
        log(f"    (could not set D1@END VIEW: {e})\n")
        return False

def update_equations_on_doc(sw_app, doc, path, names, values, log, main_path,
                             close_after=False, target_material=None,
                             apply_material=False,
                             new_description=None, old_description=None):
    """Update equations on a model. Optionally also apply target_material and
    gauge table on the same open instance (one save, one close).

    target_material : material name (also the gauge file basename) to apply.
    apply_material  : if True and target_material is provided, the material and
                       gauge table are applied to the active doc before saving.

    Returns the active_doc reference (or None if the part could not be opened).
    """
    try:   activate_doc(sw_app, path)
    except Exception as e: log(f"  Activate warning: {e}\n")
    try:   active_doc = sw_app.GetOpenDocument(path) or doc
    except Exception: active_doc = doc
    eq_mgr = active_doc.GetEquationMgr
    eq_count = get_eq_count(eq_mgr)

    upd_names, upd_values = [], []
    _db_old_val = None    # DieBoard's value BEFORE we overwrite it (for the D1 shift)
    _db_new_val = None
    if eq_count > 0 and names:
        eq_index = build_eq_index(eq_mgr)
        for name, value in zip(names, values):
            idx = eq_index.get(name.upper())
            if idx is None: log(f"  '{name}' not found\n"); continue
            # Preserve the equation's existing name (canonical or an alias like
            # OVERALL_DEPTH). Reading the current equation tells us which variant
            # to write back so we don't accidentally rename it.
            try:   existing_eq = eq_mgr.Equation(idx)
            except TypeError: existing_eq = eq_mgr.Equation
            except Exception: existing_eq = None
            actual_name = name
            if existing_eq and '"' in existing_eq:
                try: actual_name = existing_eq.split('"')[1]
                except IndexError: pass
            # DieBoard Height is a DRIVEN equation (= TOEKICK_HEIGHT + 15.3125).
            # Before we overwrite it with a literal, capture its current evaluated
            # value so we can shift D1@END VIEW by the exact same delta.
            if name.upper() == "DIE_BOARD_HEIGHT":
                try:    _db_old_val = float(eq_mgr.Value(idx))
                except TypeError:
                    try: _db_old_val = float(eq_mgr.Value)
                    except Exception: _db_old_val = None
                except Exception: _db_old_val = None
                try:    _db_new_val = float(value)
                except Exception: _db_new_val = None
            try:
                eq_mgr.Equation(idx, f'"{actual_name}"= {format_value(value)}')
                log(f"  Eq '{actual_name}'= {format_value(value)}\n")
                upd_names.append(name); upd_values.append(value)
            except Exception as e: log(f"  Failed '{name}': {e}\n")
    elif eq_count == 0 and names:
        log(f"  No equations on this part — equation updates skipped\n")

    eq_applied = bool(upd_names)
    mat_applied = False

    # DieBoard changed -> shift the D1@END VIEW dimension by the same delta.
    if _db_old_val is not None and _db_new_val is not None:
        shift_d1_for_dieboard(active_doc, _db_old_val, _db_new_val, log)

    if eq_applied:
        try:   eq_mgr.EvaluateAll()
        except Exception: pass
        try:   active_doc.ForceRebuild3(True)
        except Exception: pass
        log(f"  Updating custom properties...\n")
        update_custom_properties(active_doc, upd_names, upd_values, log,
                                 new_description=new_description,
                                 old_description=old_description)

    # Apply material + gauge if requested. Works regardless of whether equations
    # were updated — supports material-only updates with no dimension change.
    if apply_material and target_material:
        log(f"  Applying material: {target_material}\n")
        # CRITICAL: open the part FILE as a standalone document for the material
        # operations. ActivateDoc3 on a component within an assembly only brings
        # the component to the foreground — the resulting doc reference lacks
        # FeatureByName / ModifyDefinition. OpenDoc6 returns a proper PartDoc
        # interface. If the part is already open as a standalone doc, OpenDoc6
        # returns the existing instance instead of re-opening.
        part_doc = None
        try:
            errors   = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            # swDocPART = 1, swOpenDocOptions_Silent = 1
            part_doc = sw_app.OpenDoc6(path, 1, 1, "", errors, warnings)
            _track_opened_doc(path)
        except Exception as e:
            log(f"  ! OpenDoc6 failed: {e}\n")
            part_doc = None
        if part_doc is None:
            log(f"  ! Could not open part file for material update — skipping\n")
        else:
            # Now that the part is properly open, re-read its current material.
            # This read is reliable (proper PartDoc interface). If it already
            # matches the target (and we can prove it), skip the write — a final
            # safety net so we never rewrite an already-correct material.
            current_here = get_current_material(part_doc, log, indent="  ")
            if current_here and not _material_differs(current_here, target_material):
                log(f"  Material already {target_material} — no change needed\n")
                mat_applied = False
            else:
                if current_here:
                    log(f"  Material currently {current_here} → changing to {target_material}\n")
                else:
                    log(f"  Could not re-read current material — applying {target_material} anyway\n")
                mat_applied = apply_material_and_gauge(part_doc, target_material, log, indent="  ")
                if mat_applied:
                    try:   part_doc.ForceRebuild3(True)
                    except Exception: pass
                    # Save the part doc itself (separate from the assembly-context save below)
                    try:    _save(part_doc)
                    except Exception as e: log(f"  ! Save part: {e}\n")
            # Use part_doc as the doc reference for the rest of this function so
            # we don't accidentally save/close stale references.
            active_doc = part_doc

    # Description-only rewrite: the finish word changed but no equations did, so
    # update_custom_properties never ran. Write the description on its own.
    desc_only = False
    if (not eq_applied) and new_description and old_description and new_description != old_description:
        write_description_only(active_doc, new_description, old_description, log)
        desc_only = True

    if not eq_applied and not mat_applied and not desc_only:
        # Nothing to save. Close if requested.
        if not names: log(f"  No equation params and no material update — nothing to do\n")
        if close_after:
            safe_close(sw_app, active_doc, main_path, log, path.split("\\")[-1])
        return active_doc

    saved = _save(active_doc); log(f"  {'Saved' if saved else 'Save failed'}\n")
    if close_after and saved:
        safe_close(sw_app, active_doc, main_path, log, path.split("\\")[-1])
    return active_doc

def replace_in_assembly(sw_app, assy_path, old_path, new_path, main_path, log):
    errors = _byref_int()
    try:
        assy_doc = sw_app.GetOpenDocument(assy_path)
        if assy_doc is None:
            assy_doc = sw_app.OpenDoc6(assy_path, 2, 1, "", errors, _byref_int())
        if assy_doc is None: log(f"  Could not open assembly\n"); return False
        activate_doc(sw_app, assy_path)
        try:   comps = list(assy_doc.GetComponents(True))
        except TypeError: comps = list(assy_doc.GetComponents)
        assy_doc.ClearSelection2(True)
        matched = False
        for comp in comps:
            try:
                try:   comp_path = comp.GetPathName()
                except TypeError: comp_path = comp.GetPathName
                if comp_path.lower() == old_path.lower():
                    log(f"  Found: {comp.Name2}\n")
                    try:   sel = comp.Select2(True, 0)
                    except TypeError: sel = comp.Select2
                    log(f"  Select2: {sel}\n")
                    try:   result = assy_doc.ReplaceComponents2(new_path, "", True, 0, True)
                    except Exception as e: log(f"  Replace error: {e}\n"); result = False
                    log(f"  Replace result: {result}\n")
                    matched = result; break
            except Exception as e: log(f"  Comp error: {e}\n")
        if not matched: log(f"  Component not found or failed\n")
        _save(assy_doc)
        # Close the newly-inserted component (e.g. Deck Pan / Case Top from the
        # Design Reference Library) so it doesn't stay open while the user
        # downloads parts from Enovia. ReplaceComponents2 loads it as a child;
        # closing it by its own title unloads it the same way needs_update
        # assemblies are closed after their equation update.
        if matched and new_path and new_path.lower() != main_path.lower():
            try:
                new_doc = sw_app.GetOpenDocument(new_path)
                if new_doc is not None:
                    try:   ntitle = new_doc.GetTitle()
                    except TypeError: ntitle = new_doc.GetTitle
                    sw_app.CloseDoc(ntitle)
                    log(f"  Closed inserted: {os.path.basename(new_path)}\n")
            except Exception as e:
                log(f"  Close inserted warning: {e}\n")
        if assy_path.lower() != main_path.lower():
            try:
                try:   title = assy_doc.GetTitle()
                except TypeError: title = assy_doc.GetTitle
                sw_app.CloseDoc(title)
            except Exception: pass
        return matched
    except Exception as e: log(f"  Replace error: {e}\n"); return False

# ─────────────────────────────────────────────────────────────────────────────
# Reset "Title to Filename" — fixes the stale 3DEXPERIENCE tree-node title that a
# renumber (SaveAs) leaves behind. The File Title can't be set via the classic
# API, only via the platform's Rename Title dialog — so we drive that dialog by
# window messages (no mouse, no menu, no coordinates): fire the command, read the
# title vs the filename, and click Reset+OK ONLY if they differ. Command id 33185
# is the "Rename Title" menu command (verified stable across restarts).
# ─────────────────────────────────────────────────────────────────────────────
RENAME_TITLE_CMD_ID = 33185
_RENAME_DIALOG_TITLE = "Rename Title"

def _tf_find_sw_main():
    hits = []
    def cb(h, _):
        try:
            if "SOLIDWORKS Premium" in win32gui.GetWindowText(h):
                hits.append(h)
        except Exception:
            pass
        return True
    try: win32gui.EnumWindows(cb, None)
    except Exception: pass
    return hits[0] if hits else None

def _tf_dialog_hwnd():
    try:    return win32gui.FindWindow(None, _RENAME_DIALOG_TITLE) or None
    except Exception: return None

def _tf_wm_gettext(hwnd):
    try:
        length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
        if length <= 0: length = 260
        buf = ctypes.create_unicode_buffer(length + 1)
        n = ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_GETTEXT, length + 1, ctypes.byref(buf))
        if n and buf.value: return buf.value
        t = win32gui.GetWindowText(hwnd)
        return t if t else None
    except Exception:
        try:    return win32gui.GetWindowText(hwnd)
        except Exception: return None

def _tf_children(hwnd):
    out = []
    def cb(h, _):
        try:
            out.append((h, win32gui.GetClassName(h),
                        _tf_wm_gettext(h) or win32gui.GetWindowText(h) or ""))
        except Exception:
            pass
        return True
    try: win32gui.EnumChildWindows(hwnd, cb, None)
    except Exception: pass
    return out

def _tf_click_button(dlg_hwnd, want_text):
    for h, cls, txt in _tf_children(dlg_hwnd):
        if cls == "Button" and txt.strip().lower() == want_text.strip().lower():
            try:    win32gui.SendMessage(h, win32con.BM_CLICK, 0, 0); return True
            except Exception: return False
    return False

def _tf_read_title(dlg_hwnd):
    for h, cls, _ in _tf_children(dlg_hwnd):
        if cls == "Edit":
            v = _tf_wm_gettext(h)
            if v not in (None, ""): return v
    return None

def _tf_read_filename(dlg_hwnd):
    for h, cls, txt in _tf_children(dlg_hwnd):
        if cls == "Static" and txt.startswith("Filename:"):
            return txt
    return ""

def reset_title_if_mismatch(sw_app, log):
    """Fire Rename Title on the ACTIVE doc; reset the File Title only if it differs
    from the filename. Returns 'fixed' / 'ok' / 'skip'. Never raises — a failure is
    logged and the run continues."""
    try:
        sw_hwnd = _tf_find_sw_main()
        if not sw_hwnd:
            log("    (title check: SolidWorks window not found — skipped)\n"); return "skip"
        if _tf_dialog_hwnd():
            log("    (title check: a Rename dialog was already open — skipped)\n"); return "skip"

        win32gui.PostMessage(sw_hwnd, win32con.WM_COMMAND, RENAME_TITLE_CMD_ID, 0)

        h = None
        for _ in range(120):
            h = _tf_dialog_hwnd()
            if h: break
            _time.sleep(0.03)
        if not h:
            log("    (title check: Rename dialog did not open — skipped)\n"); return "skip"

        before = None; fname = ""
        for _ in range(40):
            before = _tf_read_title(h); fname = _tf_read_filename(h)
            if fname and before is not None: break
            _time.sleep(0.02)

        stem = ""
        if fname.startswith("Filename:"):
            stem = os.path.splitext(fname.split(":", 1)[1].strip())[0].strip()

        def _cancel():
            if not _tf_click_button(h, "Cancel"):
                try: win32gui.PostMessage(h, win32con.WM_COMMAND, 2, 0)
                except Exception: pass
            for _ in range(60):
                if not _tf_dialog_hwnd(): break
                _time.sleep(0.02)

        if before is None or not stem:
            log("    (title check: could not read title/filename — no change)\n"); _cancel(); return "skip"

        if before.strip().upper() == stem.upper():
            log(f"    Title ok ({stem})\n"); _cancel(); return "ok"

        # mismatch -> reset
        if not _tf_click_button(h, "Reset Title to Filename"):
            log("    (title check: Reset button not found — no change)\n"); _cancel(); return "skip"
        _time.sleep(0.03)
        if not _tf_click_button(h, "OK"):
            try: win32gui.PostMessage(h, win32con.WM_COMMAND, 1, 0)
            except Exception: pass
        for _ in range(60):
            if not _tf_dialog_hwnd(): break
            _time.sleep(0.02)
        log(f"    Title fixed: {before} -> {stem}\n")
        return "fixed"
    except Exception as e:
        log(f"    (title check error, continuing: {e})\n")
        try:
            if _tf_dialog_hwnd():
                win32gui.PostMessage(_tf_dialog_hwnd(), win32con.WM_COMMAND, 2, 0)
        except Exception: pass
        return "skip"

def _drawing_referenced_part(sw_app, draw_path, log):
    """Return the path of the model a drawing references (for fixing its title too).
    Uses document dependencies (reliable across interfaces)."""
    try:
        deps = sw_app.GetDocumentDependencies2(draw_path, True, False, False)
        if deps:
            deps = list(deps)
            for i in range(1, len(deps), 2):
                p = deps[i]
                if isinstance(p, str) and p.lower().endswith(".sldprt"):
                    return p
            for i in range(1, len(deps), 2):
                p = deps[i]
                if isinstance(p, str) and p.lower().endswith(".sldasm"):
                    return p
    except Exception as e:
        log(f"    (could not read drawing dependencies: {e})\n")
    return ""

def fix_titles_after_renumber(sw_app, new_path, is_drawing, log, counts=None):
    """After a renumber: fix the renumbered doc's title, and if it's a drawing,
    also fix its referenced part (open->fix->close). Each gated by the mismatch
    check. Fully guarded — never raises into the run."""
    def _bump(r):
        if counts is not None and r in ("fixed", "ok"):
            counts[r] = counts.get(r, 0) + 1
    def _active_path():
        try:
            ad = sw_app.ActiveDoc
            if ad is None: return "", None
            try:    p = ad.GetPathName()
            except TypeError: p = ad.GetPathName
            return (p or ""), ad
        except Exception:
            return "", None

    # 1) the renumbered document. Renaming a referenced sub can pull focus to the
    # PARENT, so DON'T trust whatever happens to be active — explicitly activate
    # the renumbered file and verify it before checking its title.
    activate_doc(sw_app, new_path)
    _time.sleep(0.15)
    ap, ad = _active_path()
    if not (ap and os.path.normcase(ap) == os.path.normcase(new_path)):
        log(f"    (title check: {os.path.basename(new_path)} is not active — skipped)\n")
    else:
        r = reset_title_if_mismatch(sw_app, log); _bump(r)
        if r == "fixed":
            try:
                if ad is not None: ad.Save3(1, _byref_int(), _byref_int())
            except Exception: pass

    # 2) if a drawing, also fix its referenced model
    if is_drawing and new_path:
        part_path = _drawing_referenced_part(sw_app, new_path, log)
        if not part_path:
            return
        log(f"    drawing references {os.path.basename(part_path)} — checking its title\n")
        pdoc, opened_here = open_doc(sw_app, part_path, 1)
        if pdoc is None:
            log("    (could not open referenced part — skipped)\n")
            try: activate_doc(sw_app, new_path)
            except Exception: pass
            return
        activate_doc(sw_app, part_path)
        _time.sleep(0.3)
        pap, pad = _active_path()
        if pap and os.path.normcase(pap) == os.path.normcase(part_path):
            r = reset_title_if_mismatch(sw_app, log); _bump(r)
            try:
                if pad is not None: pad.Save3(1, _byref_int(), _byref_int())
            except Exception: pass
        else:
            log("    (referenced part not active after open — skipped)\n")
        # close the part if WE opened it (keep the session as we found it), then
        # return to the drawing so the caller's close logic still works.
        if opened_here:
            try:
                try:   pt = pdoc.GetTitle()
                except TypeError: pt = pdoc.GetTitle
                sw_app.CloseDoc(pt)
            except Exception: pass
        try: activate_doc(sw_app, new_path)
        except Exception: pass

def run_macro_on_file(sw_app, file_path, is_drawing, main_path, log, title_counts=None):
    folder = os.path.dirname(file_path)
    ext    = ".slddrw" if is_drawing else ".sldasm"
    errors = _byref_int()
    before = snapshot_folder(folder, ext)
    try:
        file_type = 3 if is_drawing else 2
        doc = sw_app.GetOpenDocument(file_path)
        if doc is None:
            doc = sw_app.OpenDoc6(file_path, file_type, 1, "", errors, _byref_int())
        if doc is None: log(f"  Could not open: {file_path}\n"); return ""
    except Exception as e: log(f"  Open failed: {e}\n"); return ""
    try:   activate_doc(sw_app, file_path)
    except Exception: pass
    try:
        result = sw_app.RunMacro2(MACRO_PATH, MACRO_MODULE, MACRO_FUNCTION, 0, errors)
        log(f"  Macro result: {result}\n")
    except Exception as e: log(f"  Macro failed: {e}\n"); return ""
    new_path = find_new_file(folder, before, ext)
    log(f"  {'Detected: ' + os.path.basename(new_path) if new_path else 'No new file detected'}\n")
    # Reset the 3DEXPERIENCE tree title if the renumber left it stale. Only when a
    # NEW file was actually produced — no point opening the Rename dialog (and, for
    # drawings, the referenced part) on an unchanged document. The renumbered doc
    # is active right now, before we close it below. Fully guarded — never blocks.
    if new_path and new_path.lower() != file_path.lower():
        try:
            fix_titles_after_renumber(sw_app, new_path, is_drawing, log, counts=title_counts)
        except Exception as e:
            log(f"  (title fix skipped: {e})\n")
    if file_path.lower() != main_path.lower():
        try:
            try:   title = doc.GetTitle()
            except TypeError: title = doc.GetTitle
            sw_app.CloseDoc(title); log(f"  Closed: {os.path.basename(file_path)}\n")
        except Exception as e: log(f"  Close warning: {e}\n")
    return new_path

def resolve_qc_macro_path(parent=None):
    """Find the QC macro by its expected name or a similar filename.

    If the app folder contains several plausible candidates, ask the user to
    select the intended macro instead of risking execution of the wrong one.
    """
    import difflib

    folder = _tool_dir()
    expected = os.path.join(folder, QC_MACRO_FILENAME)
    if os.path.isfile(expected):
        return expected

    candidates = [
        os.path.join(folder, name)
        for name in os.listdir(folder)
        if name.lower().endswith(".swp") and "qc" in name.lower()
    ]
    expected_stem = os.path.splitext(QC_MACRO_FILENAME)[0].lower()

    def _normalise(value):
        return "".join(ch for ch in value.lower() if ch.isalnum())

    ranked = sorted(
        (
            difflib.SequenceMatcher(
                None, _normalise(expected_stem),
                _normalise(os.path.splitext(os.path.basename(path))[0]),
            ).ratio(),
            path,
        )
        for path in candidates
    )
    if ranked and ranked[-1][0] >= 0.62:
        best_score, best_path = ranked[-1]
        second_score = ranked[-2][0] if len(ranked) > 1 else 0.0
        if best_score - second_score >= 0.08:
            return best_path

    selected = filedialog.askopenfilename(
        parent=parent,
        title="Select the QC Tools SolidWorks macro",
        initialdir=folder,
        filetypes=[("SolidWorks macro", "*.swp"), ("All files", "*.*")],
    )
    if not selected:
        raise FileNotFoundError(
            f"Could not uniquely find '{QC_MACRO_FILENAME}' beside the app. "
            "Select the QC macro to continue."
        )
    return selected

def run_qc_tools_macro(sw_app, main_path, model_title,
                       module_name=QC_MACRO_MODULE,
                       procedure_name=QC_MACRO_FUNCTION,
                       macro_path=None):
    """Run the selected QC macro and find the workbook it creates or updates."""
    macro_path = macro_path or resolve_qc_macro_path()
    if not os.path.isfile(macro_path):
        raise FileNotFoundError(f"QC tools macro not found:\n{macro_path}")
    if not module_name.strip() or not procedure_name.strip():
        raise ValueError("Enter both the QC macro module and procedure names.")

    job_name = os.path.splitext(os.path.basename(model_title or ""))[0].strip()
    if not job_name:
        raise ValueError("Cannot determine the job name for the QC report.")
    report_dir = _tool_dir()
    before_reports = {}
    try:
        for name in os.listdir(report_dir):
            if name.lower().endswith(".xlsm"):
                path = os.path.join(report_dir, name)
                stat = os.stat(path)
                before_reports[os.path.normcase(path)] = (
                    stat.st_mtime_ns, stat.st_size
                )
    except OSError as e:
        raise OSError(f"Could not inspect the QC report folder:\n{e}") from e

    if main_path and activate_doc(sw_app, main_path) is None:
        raise RuntimeError("Could not activate the top-level assembly for QC.")

    errors = _byref_int()
    try:
        macro_result = sw_app.RunMacro2(
            macro_path, module_name.strip(), procedure_name.strip(), 0, errors
        )
    except Exception as e:
        raise RuntimeError(f"Could not run the QC tools macro: {e}") from e
    changed_reports = []
    try:
        for name in os.listdir(report_dir):
            if not name.lower().endswith(".xlsm"):
                continue
            path = os.path.join(report_dir, name)
            stat = os.stat(path)
            current = (stat.st_mtime_ns, stat.st_size)
            if before_reports.get(os.path.normcase(path)) != current:
                changed_reports.append((stat.st_mtime_ns, path))
    except OSError as e:
        raise OSError(f"QC macro returned, but the report folder could not be checked:\n{e}") from e

    if not changed_reports:
        try:
            error_detail = errors.value
        except Exception:
            error_detail = "unavailable"
        raise RuntimeError(
            "SolidWorks did not create or update an .xlsm report in the app folder. "
            f"RunMacro2 returned {macro_result!r}; error code: {error_detail}. "
            f"Macro: {macro_path}; module: {module_name.strip()}; "
            f"procedure: {procedure_name.strip()}."
        )

    expected_name = "".join(
        char for char in job_name.lower() if char.isalnum()
    )
    changed_reports.sort(
        key=lambda item: (
            expected_name in "".join(
                char for char in os.path.splitext(os.path.basename(item[1]))[0].lower()
                if char.isalnum()
            ),
            item[0],
        ),
        reverse=True,
    )
    return changed_reports[0][1]

def stamp_drawing_date(sw_app, drw_path, log, indent="  "):
    """Stamp DRAWING_DATE_PROP on the (renamed) drawing with today's date.

    Opens the drawing if it isn't already open, writes the property (preserving
    its Date type via swCustomPropertyDeleteAndAdd), saves, and closes it again
    only if we opened it here. The value is written in DRAWING_DATE_FORMAT.
    Read-back of custom properties is unreliable on this pywin32 build, so we
    write-and-log rather than verify. Returns True if the write succeeded.
    """
    today = datetime.now().strftime(DRAWING_DATE_FORMAT)
    drw = None; opened_here = False
    try:
        drw = sw_app.GetOpenDocument(drw_path)
    except Exception:
        drw = None
    if drw is None:
        try:
            errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            warns  = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            drw = sw_app.OpenDoc6(drw_path, 3, 1, "", errors, warns)   # 3 = drawing, silent
            opened_here = True
        except Exception as e:
            log(f"{indent}Date stamp: could not open drawing ({e})\n"); return False
    if drw is None:
        log(f"{indent}Date stamp: drawing not available\n"); return False

    ok = False
    try:
        mgr = drw.Extension.CustomPropertyManager("")   # file-level ('Custom' tab)
        # swCustomPropertyDeleteAndAdd = 1 ; field types: 64 = Date, 30 = Text
        for ftype, tname in ((64, "Date"), (30, "Text")):
            try:
                mgr.Add3(DRAWING_DATE_PROP, ftype, today, 1)
                ok = True
                log(f"{indent}Date stamp: {DRAWING_DATE_PROP} = {today} (type {tname})\n")
                break
            except Exception:
                continue
        if not ok:
            log(f"{indent}Date stamp: could not set {DRAWING_DATE_PROP}\n")
        else:
            try:   drw.ForceRebuild3(True)
            except Exception: pass
            try:   _save(drw)
            except Exception as e: log(f"{indent}Date stamp: save warning ({e})\n")
    except Exception as e:
        log(f"{indent}Date stamp: error ({e})\n")

    if opened_here and drw is not None:
        try:
            try:   t = drw.GetTitle()
            except TypeError: t = drw.GetTitle
            sw_app.CloseDoc(t)
        except Exception: pass
    return ok

# ─────────────────────────────────────────
# Suppression helpers
# ─────────────────────────────────────────
# swComponentSuppressionState_e: 0 = Suppressed, 2 = Resolved
def get_suppress_state(child):
    """Return the component's suppression state as an int, or None if unreadable.
    swComponentSuppressionState_e: 0=Suppressed, 1=Lightweight,
    2=Resolved, 3=FullyResolved.

    NOTE: this INTEGER read (GetSuppression2 / GetSuppressState) has proven
    unreliable in this environment — it reported non-0 values for genuinely
    suppressed components. Prefer IsSuppressed() (a Boolean) for decisions; this
    is kept only as a secondary hint (e.g. to spot lightweight)."""
    try:
        return int(child.GetSuppressState())
    except Exception:
        try:
            return int(child.GetSuppression2())
        except Exception:
            return None

def is_component_suppressed(child):
    """True iff the component is currently suppressed. Uses IsSuppressed() (a
    reliable Boolean API), falling back to the integer enum only if that call
    is unavailable."""
    try:
        return bool(child.IsSuppressed())
    except Exception:
        return get_suppress_state(child) == 0

def capture_restore_target(child):
    """Decide what suppression state to RESTORE a component to after we resolve
    it for editing. MUST be called BEFORE resolving (IsSuppressed flips once
    resolved). Returns:
        0    -> was suppressed; re-suppress it at the end
        1    -> was lightweight; set lightweight at the end
        None -> was already resolved; leave it resolved (nothing to restore)

    The primary signal is IsSuppressed() because the integer suppression enum
    has been unreliable here (it caused the end-of-run restore to skip every
    genuinely-suppressed part and leave it resolved). IsSuppressed answers
    'was it suppressed?' directly; the enum is consulted only to recognise the
    lightweight case."""
    try:
        if bool(child.IsSuppressed()):
            return 0
    except Exception:
        pass
    st = get_suppress_state(child)
    if st == 0:
        return 0
    if st == 1:
        return 1
    return None

def resolve_component(sw_app, child, log, indent=""):
    """Resolve (unsuppress) a component so its model doc loads.
    Returns (doc, path) just like get_doc. Returns (None, path) on failure."""
    try:
        child.SetSuppression2(2)   # 2 = Resolved
    except Exception as e:
        try: log(f"{indent}  Resolve failed: {e}\n")
        except Exception: pass
    doc, path = get_doc(sw_app, child)
    return doc, path

def resuppress_component(child, log, indent=""):
    """Restore a component to the Suppressed state."""
    try:
        child.SetSuppression2(0)   # 0 = Suppressed
    except Exception as e:
        try: log(f"{indent}  Re-suppress failed: {e}\n")
        except Exception: pass

def close_all_except_main(sw_app, main_path, log, only_close_paths=None):
    """Close open SolidWorks documents except the main assembly.

    only_close_paths : optional set of lowercased file paths. When provided,
        ONLY documents whose path is in this set are closed — unrelated models
        the user had open are left untouched. When None, every non-main doc is
        closed (legacy behaviour).

    Walks the open documents via GetFirstDocument / GetNext, comparing each
    path to main_path. Documents that are loaded children of the still-open
    main assembly cannot be closed independently — those silently no-op.
    """
    closed = 0
    skipped = 0
    failed = 0
    docs_to_consider = []
    # Try a few enumeration approaches — different SolidWorks versions expose
    # this differently. If all fail, log details and bail gracefully.
    enumerated = False
    err_msgs = []

    # Approach 1: GetFirstDocument / GetNext walk
    try:
        try:   doc = sw_app.GetFirstDocument()
        except TypeError: doc = sw_app.GetFirstDocument
        while doc is not None:
            try:
                try:   p = doc.GetPathName()
                except TypeError: p = doc.GetPathName
                except Exception: p = ""
                try:   t = doc.GetTitle()
                except TypeError: t = doc.GetTitle
                except Exception: t = ""
            except Exception:
                p, t = "", ""
            docs_to_consider.append((doc, p or "", t or ""))
            try:
                try:   doc = doc.GetNext()
                except TypeError: doc = doc.GetNext
            except Exception:
                doc = None
        enumerated = True
    except Exception as e:
        err_msgs.append(f"GetFirstDocument: {e}")

    # Approach 2: enumerate via sw_app.GetDocuments (some versions)
    if not enumerated:
        try:
            docs = sw_app.GetDocuments
            try: docs = list(docs)
            except Exception: docs = []
            for d in docs:
                try:
                    try:   p = d.GetPathName()
                    except TypeError: p = d.GetPathName
                    except Exception: p = ""
                    try:   t = d.GetTitle()
                    except TypeError: t = d.GetTitle
                    except Exception: t = ""
                except Exception:
                    p, t = "", ""
                docs_to_consider.append((d, p or "", t or ""))
            enumerated = True
        except Exception as e:
            err_msgs.append(f"GetDocuments: {e}")

    if not enumerated:
        log(f"  Cleanup: could not enumerate open docs ({'; '.join(err_msgs)})\n")
        return

    main_lower = (main_path or "").lower()
    for d, p, t in docs_to_consider:
        if not t:
            skipped += 1; continue
        if p and p.lower() == main_lower:
            skipped += 1; continue  # never close main
        # If an allow-list was provided, only close docs that the tool itself
        # opened during this run. Anything else (unrelated models the user had
        # open) is left alone.
        if only_close_paths is not None and (not p or p.lower() not in only_close_paths):
            skipped += 1; continue
        try:
            sw_app.CloseDoc(t)
            closed += 1
        except Exception:
            failed += 1
    log(f"  Cleanup: closed {closed} doc(s), skipped {skipped}, failed {failed}\n")


def resuppress_all_resolved(sw_app, main_path, sw_model, resolved_comps_all,
                            all_entries, log):
    """Re-suppress every component we resolved during the scan.

    Components may have been renamed or replaced during Phase 2, so the COM refs
    stored at scan time are stale. We re-find each component by path inside its
    parent assembly (opening the parent if needed) and suppress the fresh ref.

    resolved_comps_all entries are tuples: (parent_path, old_path, name, restore_target)
    where restore_target is what we must restore the component to AFTER resolving
    it (0=re-suppress, 1=set lightweight, None=leave resolved). It was captured
    BEFORE we resolved the component, using IsSuppressed() so a genuinely-
    suppressed component is reliably re-suppressed (the integer suppression enum
    proved unreliable and previously caused every part to be left resolved).
    """
    if not resolved_comps_all:
        return

    # Build a map from old_path -> final_path using all_entries. If a component
    # was renamed/replaced in Phase 2, new_path holds where it lives now.
    path_map = {}
    for e in all_entries:
        old_p = (e.get("old_path") or "").lower()
        new_p = e.get("new_path") or ""
        if old_p:
            path_map[old_p] = new_p if new_p else e.get("old_path")

    # ── DIAGNOSTIC ───────────────────────────────────────────────
    log("  [debug] path_map entries with rename:\n")
    rename_count = 0
    for k, v in path_map.items():
        if v and v.lower() != k:
            log(f"    [debug]   {os.path.basename(k)} -> {os.path.basename(v)}\n")
            rename_count += 1
    log(f"  [debug] {rename_count} renames in path_map (out of {len(path_map)} total)\n")
    log(f"  [debug] resolved_comps_all has {len(resolved_comps_all)} items:\n")
    _rt_lbl = {0: "re-suppress", 1: "lightweight", None: "leave-resolved"}
    _n_suppress = sum(1 for *_x, rt in resolved_comps_all if rt == 0)
    for parent_path, old_path, name, _rt in resolved_comps_all:
        translated_parent = path_map.get(parent_path.lower(), parent_path)
        translated_child  = path_map.get(old_path.lower(), old_path)
        log(f"    [debug]   {name}: restore={_rt_lbl.get(_rt, _rt)} | parent {os.path.basename(parent_path)} "
            f"-> {os.path.basename(translated_parent)}, child "
            f"-> {os.path.basename(translated_child)}\n")
    log(f"  [debug] {_n_suppress} of {len(resolved_comps_all)} will be re-suppressed "
        f"(restore target 0)\n")
    # ── END DIAGNOSTIC ───────────────────────────────────────────

    # Group by parent so we open each parent assembly only once.
    # IMPORTANT: translate each parent_path through path_map too — the parent
    # assembly itself may have been renamed in Phase 2 (old A8000867.SLDASM
    # becomes new A8003206.SLDASM), and we need to re-suppress inside the
    # renamed version because that's what the main assembly now references.
    by_parent = {}
    for parent_path, old_path, name, restore_target in resolved_comps_all:
        current_parent = path_map.get(parent_path.lower(), parent_path)
        if not current_parent:
            current_parent = parent_path
        by_parent.setdefault(current_parent, []).append((old_path, name, parent_path, restore_target))

    log(f"  Restoring suppression on {len(resolved_comps_all)} component(s) "
        f"across {len(by_parent)} parent(s)...\n")

    for parent_path, items in by_parent.items():
        is_main = parent_path.lower() == main_path.lower()
        parent_doc = None
        opened_here = False
        try:
            if is_main:
                parent_doc = sw_model
            else:
                # Try to find it already open; otherwise open it
                try: parent_doc = sw_app.GetOpenDocument(parent_path)
                except Exception: parent_doc = None
                if parent_doc is None:
                    try:
                        parent_doc = sw_app.OpenDoc6(
                            parent_path, 2, 0, "", _byref_int(), _byref_int())
                        opened_here = True
                        _track_opened_doc(parent_path)  # ensure end-of-run cleanup closes it
                    except Exception as e:
                        log(f"    Could not open parent {os.path.basename(parent_path)}: {e}\n")
                        continue
            if parent_doc is None:
                log(f"    Parent doc unavailable: {os.path.basename(parent_path)}\n")
                continue

            # Activate so SetSuppression2 affects this doc
            try: activate_doc(sw_app, parent_path)
            except Exception: pass

            # Walk this parent's components looking for our tracked items
            try:   comps = parent_doc.GetComponents(False)
            except Exception: comps = None
            comps = [] if comps is None else list(comps)

            for old_path, name, original_parent, restore_target in items:
                # Reapply the state captured before the scan.
                # 0=re-suppress, 1=set lightweight. None means it was already
                # resolved before we touched it, so leave it resolved.
                if restore_target is None:
                    continue
                restore_state = restore_target
                final_path = path_map.get(old_path.lower(), old_path)
                target_lower = (final_path or "").lower()
                # Also try to find the entry's new_number so we can fall back to
                # matching any current child whose path contains that number —
                # protects against rename-but-reference-not-updated edge cases.
                new_number = None
                for e in all_entries:
                    if (e.get("old_path") or "").lower() == old_path.lower():
                        new_number = (e.get("new_number") or "").strip()
                        break
                found = False
                for c in comps:
                    try:
                        try:   cpath = c.GetPathName()
                        except TypeError: cpath = c.GetPathName
                        except Exception: cpath = ""
                    except Exception:
                        cpath = ""
                    cpath_lower = (cpath or "").lower()
                    matches = False
                    if cpath_lower:
                        if cpath_lower == target_lower:
                            matches = True
                        elif new_number and new_number.lower() in cpath_lower:
                            matches = True
                        elif cpath_lower == old_path.lower():
                            # Component still references the old path — happens
                            # if the macro renamed the file but didn't repoint
                            # the assembly. Restore it anyway.
                            matches = True
                    if matches:
                        try:    c.SetSuppression2(restore_state)   # 0=suppress, 1=lightweight
                        except Exception as e:
                            try: log(f"      Restore-state failed: {e}\n")
                            except Exception: pass
                        _sl = {0: "suppressed", 1: "lightweight"}.get(restore_state, str(restore_state))
                        log(f"    Restored to {_sl}: {name} ({os.path.basename(cpath)})\n")
                        found = True
                        # Do NOT break: pattern copies (e.g. X<2>, X<3>) are
                        # separate components that reference the SAME file, so we
                        # must keep scanning and restore every matching instance —
                        # otherwise the seed gets restored but its pattern
                        # instances stay behind (showing as unresolved).
                if not found:
                    log(f"    Could not locate {name} in {os.path.basename(parent_path)}\n")

            # Save the parent so the suppression persists on disk
            try:
                _save(parent_doc)
                log(f"    Saved parent: {os.path.basename(parent_path)}\n")
            except Exception as e:
                log(f"    Save warning ({os.path.basename(parent_path)}): {e}\n")
        finally:
            if opened_here and not is_main and parent_doc is not None:
                try:
                    try:   title = parent_doc.GetTitle()
                    except TypeError: title = parent_doc.GetTitle
                    sw_app.CloseDoc(title)
                except Exception: pass

# ─────────────────────────────────────────
# Assembly tree walker
# ─────────────────────────────────────────
def walk_assembly_tree(sw_app, model, sel_names, sel_values,
                       lookup, parts_rows, assy_rows,
                       visited, log, status_cb, root_tk,
                       folder, assy_config=None, depth=0, cur_eq_vals=None,
                       resolved_comps=None, node_info=None):
    if cur_eq_vals is None: cur_eq_vals = {}
    if resolved_comps is None: resolved_comps = []
    # node_info maps every visited component path (lowercased) ->
    # {parent, desc, depth, old_number, type}. It records ALL components, even
    # unaffected intermediate assemblies, so the parent chain can be walked
    # upward later to cascade new numbers to ancestor assemblies.
    if node_info is None: node_info = {}
    indent   = "  " * depth
    children = get_children_from_model(model)
    entries  = []
    if not children: return entries
    try:   current_path = model.GetPathName()
    except TypeError: current_path = model.GetPathName
    assy_children, part_children = [], []
    for child in children:
        try:
            doc, path = get_doc(sw_app, child)
            # RELIABLE 'was suppressed' signal: a suppressed / not-loaded
            # component returns NO doc from get_doc. Lightweight and resolved
            # components DO return a doc (GetOpenDocument hands back lightweight
            # docs). The integer suppression enum (GetSuppression2/GetSuppressState)
            # and IsSuppressed() both proved unreliable in this environment —
            # they reported 'not suppressed' for components that were actually
            # suppressed, so the end-of-run restore skipped all of them. We key
            # the restore decision off whether the model actually loaded:
            #   doc is None -> it was suppressed -> re-suppress (0) at the end
            #   doc present -> leave it (None); it was lightweight/resolved
            restore_target = 0 if doc is None else None
            pre_resolved = False
            # Suppressed components often return empty path because the file is
            # not loaded. Try resolving once so we can route them to the right
            # list (asm/prt). Mark them so the per-loop logic re-suppresses later.
            if not path:
                try:
                    child.SetSuppression2(2)  # 2 = Resolved
                    doc, path = get_doc(sw_app, child)
                    if path:
                        pre_resolved = True
                        restore_target = 0   # we resolved a not-loaded component -> must re-suppress
                except Exception:
                    pass
            if not path:
                try: log(f"{indent}  [skip: no path] {child.Name2}\n")
                except Exception: pass
                continue
            path_key = path.lower()
            if path_key in visited:
                try: log(f"{indent}  [skip: already visited] {os.path.basename(path)}\n")
                except Exception: pass
                continue
            if path_key.endswith(".sldasm"): assy_children.append((child, doc, path, pre_resolved, restore_target))
            elif path_key.endswith(".sldprt"): part_children.append((child, doc, path, pre_resolved, restore_target))
            else:
                try: log(f"{indent}  [skip: not asm/prt] {os.path.basename(path)}\n")
                except Exception: pass
        except Exception as e:
            try: log(f"{indent}  [skip: error {e}]\n")
            except Exception: pass
            continue

    # Subassemblies first
    for child, doc, path, pre_resolved, restore_target in assy_children:
        was_suppressed = pre_resolved
        try:
            path_key = path.lower()
            if path_key in visited: continue
            log("\n")  # blank line before each component block (readability)
            name = child.Name2
            if pre_resolved:
                log(f"{indent}  Resolved (was suppressed): {name}\n")
            # Resolve if the component was suppressed/lightweight OR its doc is
            # None (not loaded). We need it fully loaded to read + update
            # equations across all configurations. Its restore target was
            # captured in restore_target and is reapplied at the end of the run.
            elif (restore_target in (0, 1)) or doc is None:
                was_suppressed = True
                state_lbl = {0: "suppressed", 1: "lightweight"}.get(restore_target, "not loaded")
                log(f"{indent}  {state_lbl.capitalize()} (resolving to update): {name}\n")
                doc, path = resolve_component(sw_app, child, log, indent)
                if doc is None:
                    log(f"{indent}  Could not resolve {name} — skipping\n")
                    continue
            visited.add(path_key)
            _track_opened_doc(path)
            desc = get_description(doc) if doc else "N/A"
            node_info[path_key] = {"parent": current_path, "desc": desc, "depth": depth,
                                   "old_number": get_part_number_from_path(path), "type": "assembly"}
            # Remember configurable sub-assemblies (base frame / frame kit, by the
            # leading tokens of their Description) so the chosen config can be
            # re-pointed on the final version of each at the end of the run.
            _fam = config_family_for_desc(desc)
            if _fam is not None:
                CONFIG_TARGETS[path] = _fam
                log(f"{indent}  ↳ {_fam['name']} detected: {name}\n")

            # ─────────────────────────────────────────
            # SPECIAL CASE: Design Reference Library
            # DP -> Deck Pan (match Length + Depth)
            # SIL,TP -> Case Top (match Length + Top Depth)
            # Only triggers if the relevant parameter was selected by the user.
            # ─────────────────────────────────────────
            desc_up = (desc or "").strip().upper()
            _cfg = get_active_cfg() or {}
            is_dp     = desc_up.startswith("DP")     and _cfg.get("design_ref_check_dp",    True)
            is_siltp  = desc_up.startswith("SIL,TP") and _cfg.get("design_ref_check_siltp", True)
            if is_dp or is_siltp:
                if is_dp:
                    relevant = [p for p in ("LENGTH", "DEPTH") if p in sel_names]
                    len_target = (sel_values[sel_names.index("LENGTH")] if "LENGTH" in sel_names
                                  else cur_eq_vals.get("LENGTH"))
                    dep_target = (sel_values[sel_names.index("DEPTH")] if "DEPTH" in sel_names
                                  else cur_eq_vals.get("DEPTH"))
                    kind = "Deck Pan"
                else:
                    relevant = [p for p in ("LENGTH", "TOP_DEPTH") if p in sel_names]
                    len_target = (sel_values[sel_names.index("LENGTH")] if "LENGTH" in sel_names
                                  else cur_eq_vals.get("LENGTH"))
                    dep_target = (sel_values[sel_names.index("TOP_DEPTH")] if "TOP_DEPTH" in sel_names
                                  else cur_eq_vals.get("TOP_DEPTH"))
                    kind = "Case Top"

                _dref_mat = material_for_description(lookup, desc)
                # DP / SIL,TP descriptions carry their dims and finish inline
                # (e.g. "DP,48.5,42,2IN_PS,PAINTED,..."), so work out what this
                # description SHOULD read after the run. find_in_design_ref matches
                # on the dimension columns (it only reads the prefix), so the lookup
                # itself is unaffected — this is what gets WRITTEN back to the model.
                _newv_d = dict(cur_eq_vals)
                for _i, _n in enumerate(sel_names):
                    _newv_d[_n] = sel_values[_i]
                tgt_desc = build_target_description(desc, cur_eq_vals, _newv_d, _dref_mat)
                if relevant or _dref_mat:  # a relevant dimension OR a finish applies
                    log(f"{indent}  >> {name} [{kind}] [{desc}]\n")
                    log(f"{indent}     Design Ref check: L={len_target} D={dep_target} MAT={_dref_mat or '-'}\n")
                    status_cb(f"Checking {name}..."); root_tk.update()
                    dref_assy = find_in_design_ref(desc, len_target, dep_target, log=log, required_material=_dref_mat)
                    if dref_assy:
                        pv = {n: format_value(v) for n, v in zip(sel_names, sel_values)}
                        for p_name in cur_eq_vals:
                            if p_name not in pv: pv[p_name] = format_value(cur_eq_vals[p_name])
                        if dref_assy.upper() == get_part_number_from_path(path).upper():
                            log(f"{indent}     → Same number already (Design Ref) — skip\n")
                        else:
                            ext_path = os.path.join(folder, dref_assy + ".SLDASM")
                            on_disk  = os.path.exists(ext_path)
                            entry_status = "existing" if on_disk else "file_missing"
                            log(f"{indent}     → DESIGN REF FOUND: {dref_assy} {'✓' if on_disk else '✗ File Missing'}\n")
                            entries.append({
                                "old_path":       path,
                                "old_number":     get_part_number_from_path(path),
                                "desc":           desc,
                                "target_desc":    tgt_desc,
                                "parent_path":    current_path,
                        "depth":          depth,
                                "type":           "assembly",
                                "app_names":      relevant,
                                "app_values":     [len_target if p=="LENGTH" else dep_target for p in relevant],
                                "param_values":   pv,
                                "existing_number": dref_assy,
                                "new_number":     dref_assy,
                                "new_path":       "",
                                "status":         entry_status,
                                "_status_lbl":    None,
                                "_design_ref":    True,
                            })
                            def mark_visited_recursive(mdl, vis, depth_limit=10):
                                if depth_limit <= 0: return
                                try:
                                    ch_list = mdl.GetComponents(False)
                                    if not ch_list: return
                                    for ch in list(ch_list):
                                        try:
                                            try:   cp = ch.GetPathName()
                                            except TypeError: cp = ch.GetPathName
                                            if cp and cp.lower() not in vis:
                                                vis.add(cp.lower())
                                                ch_doc = sw_app.GetOpenDocument(cp)
                                                if ch_doc and cp.lower().endswith(".sldasm"):
                                                    mark_visited_recursive(ch_doc, vis, depth_limit-1)
                                        except Exception: pass
                                except Exception: pass
                            if doc:
                                mark_visited_recursive(doc, visited)
                                log(f"{indent}     → All children marked visited (replaced as unit)\n")
                        if doc: safe_close(sw_app, doc, "", log, name)
                        continue
                    else:
                        log(f"{indent}     → Not in Design Ref → falling back to Master BOM\n")

            app_names, app_values = get_applicable_params(lookup, desc, sel_names, sel_values)
            if WHY_LOG: log(f"{indent}     {_why_affected(lookup, desc, sel_names)}\n")
            # Finish for material-bearing assemblies (toekicks via abbreviation,
            # Case Top / Interior via prefix). Lets a material-only change be
            # processed even when no dimension applies. None for everything else,
            # so non-material assemblies keep their exact dimension-gated behavior.
            _amat = material_for_description(lookup, desc)
            # Assembly-level finish gate. ONLY Case Top / Interior (prefix-based)
            # may open this gate on a material-only basis, because their finish
            # lives on the ASSEMBLY (description / Design Ref / BOM row). Toekicks
            # and the front panel carry their finish on a PART, so a material-only
            # change for them is handled at the part level (re-material the part,
            # then cascade the assembly) — exactly like the front panel — via the
            # pruning exception below, NOT by matching the assembly here. Without
            # this an UNCHANGED toekick (e.g. Brushed staying Brushed) would open
            # the gate, find no exact-material BOM row, and get a bogus new number.
            _amat_gate = _amat if desc_up.startswith(("DP", "SIL,TP")) else None
            if app_names or _amat_gate:
                log(f"{indent}  >> {name} [Assembly] [{desc}]\n")
                log(f"{indent}     Params: {','.join(app_names)}\n")
                status_cb(f"Checking {name}..."); root_tk.update()
                # pv = all selected params new values + current SW values for unselected
                # Used for Master BOM append so all columns are correctly populated
                pv = {n: format_value(v) for n, v in zip(sel_names, sel_values)}
                for p_name in cur_eq_vals:
                    if p_name not in pv: pv[p_name] = format_value(cur_eq_vals[p_name])
                # Build FULL bom_names/bom_values: for params applicable to this part
                # but NOT selected by user, use current SolidWorks equation value
                # so BOM matching is always done against all relevant dimensions
                guideline_row = _guideline_row(lookup, desc) or {}
                bom_names, bom_values = [], []
                for p_name, _ in PARAMS:
                    if not guideline_row.get(p_name, False): continue
                    if p_name in sel_names:
                        idx = sel_names.index(p_name)
                        bom_names.append(p_name); bom_values.append(sel_values[idx])
                    elif p_name in cur_eq_vals:
                        bom_names.append(p_name); bom_values.append(cur_eq_vals[p_name])
                log(f"{indent}     BOM check: {list(zip(bom_names,bom_values))}\n")
                # The front-panel assembly's BOM row carries a Material (e.g.
                # 430-16BRSS for Brushed SS). Require the SELECTED front-panel
                # material so we don't replace it with a wrong-material row
                # (matching a Brushed-SS assembly when Painted was chosen). Other
                # assemblies carry no material, so they stay unconstrained.
                _assy_req_mat = None
                _adesc_u = (desc or "").strip().upper()
                if _adesc_u.startswith(("FP,FLT", "FP,ANGLD")):
                    _assy_req_mat = selected_fp_material()
                else:
                    _assy_req_mat = _amat
                _newv_a = dict(cur_eq_vals)
                for _i, _n in enumerate(sel_names):
                    _newv_a[_n] = sel_values[_i]
                tgt_desc = build_target_description(desc, cur_eq_vals, _newv_a, _assy_req_mat)
                if tgt_desc != desc:
                    log(f"{indent}     Description will become: {tgt_desc}\n")
                existing = find_in_master_bom(tgt_desc, bom_names, bom_values, True, parts_rows, assy_rows, assy_config, required_material=_assy_req_mat)
                if existing:
                    if existing.upper() == get_part_number_from_path(path).upper():
                        log(f"{indent}     → Same number already in BOM — skip\n")
                        existing = None; entry_status = "skip"
                    else:
                        ext_path = os.path.join(folder, existing + ".SLDASM")
                        on_disk  = os.path.exists(ext_path)
                        entry_status = "existing" if on_disk else "file_missing"
                        log(f"{indent}     → BOM FOUND: {existing} {'✓' if on_disk else '✗ File Missing'}\n")
                        def mark_visited_recursive(mdl, vis, depth_limit=10):
                            if depth_limit <= 0: return
                            try:
                                ch_list = mdl.GetComponents(False)
                                if not ch_list: return
                                for ch in list(ch_list):
                                    try:
                                        try:   cp = ch.GetPathName()
                                        except TypeError: cp = ch.GetPathName
                                        if cp and cp.lower() not in vis:
                                            vis.add(cp.lower())
                                            ch_doc = sw_app.GetOpenDocument(cp)
                                            if ch_doc and cp.lower().endswith(".sldasm"):
                                                mark_visited_recursive(ch_doc, vis, depth_limit-1)
                                    except Exception: pass
                            except Exception: pass
                        if doc:
                            mark_visited_recursive(doc, visited)
                            log(f"{indent}     → All children marked as visited (will be skipped)\n")
                else:
                    existing = None; entry_status = "needs_update"
                    log(f"{indent}     → Not in BOM → needs new number\n")
                if entry_status != "skip":
                    entries.append({
                        "old_path":       path,
                        "old_number":     get_part_number_from_path(path),
                        "desc":           desc,
                        "target_desc":    tgt_desc,
                        "parent_path":    current_path,
                        "depth":          depth,
                        "type":           "assembly",
                        "app_names":      app_names,
                        "app_values":     app_values,
                        "param_values":   pv,
                        "existing_number": existing,
                        "new_number":     existing or "Pending...",
                        "new_path":       "",
                        "status":         entry_status,
                        "_status_lbl":    None,
                        # Non-FP material-bearing assemblies (toekicks / end cover)
                        # carry their selected finish so the BOM row + cascade
                        # re-check filter by it. FP assemblies get it via propagation.
                        "fp_material":    (_assy_req_mat if not _adesc_u.startswith(("FP,FLT","FP,ANGLD")) else None),
                    })
            else:
                log(f"{indent}  -- {name} [{desc}] (not affected)\n")
            comp_status = entry_status if (app_names or _amat_gate) else "no_params"
            if doc and comp_status == "needs_update":
                sub_entries = walk_assembly_tree(
                    sw_app, doc, sel_names, sel_values,
                    lookup, parts_rows, assy_rows,
                    visited, log, status_cb, root_tk,
                    folder, assy_config=assy_config, depth=depth+1, cur_eq_vals=cur_eq_vals, resolved_comps=resolved_comps, node_info=node_info)
                entries.extend(sub_entries)
                safe_close(sw_app, doc, "", log, name)
            elif doc and comp_status in ("existing", "file_missing"):
                log(f"{indent}  → Skipping children (assembly replaced as unit)\n")
                safe_close(sw_app, doc, "", log, name)
            elif doc:
                # comp_status is "no_params" here: NO selected parameter applies
                # to this subassembly. By the Guideline's hierarchical design, no
                # part under it is affected either, so PRUNE the whole subtree —
                # its children are never resolved / unsuppressed or modified.
                #
                # This prunes a subassembly that is NOT in the Guideline too:
                # per the product rule, anything outside the Guideline is not
                # ours to touch and cannot contain an affected part. (Previously
                # unrecognised assemblies were still walked "to be safe", which
                # unsuppressed their parts for nothing.)
                #
                # EXCEPTION: during a material change, descend anyway when a
                # material-bearing part could live inside this subtree — material
                # is NOT a Guideline dimension, so the dimension-based pruning
                # can't be trusted for it. This is the same mechanism the front
                # panel has always used; here it's generalised to every material
                # part type so toekicks / end cover / Case Top / Interior update
                # exactly like the front panel.
                _key = (desc or "").strip().upper()
                _mat_active = any(SELECTED_MATERIALS.values())
                # Front panel + Case Top + Interior are recognised by prefix.
                _is_mat_assy = _key.startswith(("FP,FLT", "FP,ANGLD", "SIL,TP", "DP"))
                # Toekicks / end cover: the assembly description IS a material
                # part description (optionally with a suffix), e.g. 'KP,MTL,FR…'.
                if not _is_mat_assy and _mat_active:
                    for _gk, _grow in (lookup or {}).items():
                        if get_selected_material(_grow.get("_ABBREV", "")) is None:
                            continue
                        if _key == _gk or _key.startswith(_gk + ","):
                            _is_mat_assy = True
                            break
                if _mat_active and _is_mat_assy:
                    log(f"{indent}  → Material selected — descending to reach material part(s)\n")
                    sub_entries = walk_assembly_tree(
                        sw_app, doc, sel_names, sel_values,
                        lookup, parts_rows, assy_rows,
                        visited, log, status_cb, root_tk,
                        folder, assy_config=assy_config, depth=depth+1, cur_eq_vals=cur_eq_vals, resolved_comps=resolved_comps, node_info=node_info)
                    entries.extend(sub_entries)
                    safe_close(sw_app, doc, "", log, name)
                else:
                    _why = ("not affected" if _guideline_row(lookup, desc) is not None
                            else "not in Guideline — not touched")
                    log(f"{indent}  → Subtree {_why} — skipping its children "
                        f"(not resolved / unsuppressed)\n")
                    safe_close(sw_app, doc, "", log, name)
        except Exception as e:
            log(f"{indent}  Error (assy): {e}\n")
        finally:
            # Defer re-suppression until after the apply phase saves. Collect the
            # component identity (by path) so we can re-find it later, since the
            # COM reference may go stale if the component is renamed/replaced.
            if was_suppressed:
                resolved_comps.append((current_path, path, name, restore_target))

    # Then parts
    for child, doc, path, pre_resolved, restore_target in part_children:
        was_suppressed = pre_resolved
        status = None
        try:
            path_key = path.lower()
            if path_key in visited: continue
            log("\n")  # blank line before each component block (readability)
            name = child.Name2

            # ── Skip-resolve pre-filter ─────────────────────────────────────
            # Before unsuppressing a part just to read it, check (cheaply,
            # WITHOUT resolving) whether it can be affected at all. If neither a
            # selected dimension nor a material change can apply, skip it — no
            # reason to unsuppress. For loaded/lightweight parts we read the real
            # Description; for suppressed parts we use the referenced config name
            # and only skip when it's a recognised Guideline entry, so a misread
            # name can never drop a part that actually needed changing.
            if not pre_resolved:
                if doc is not None:
                    _peek, _reliable = (get_description(doc) or ""), True
                else:
                    _peek, _reliable = get_referenced_config(child), False
                if should_skip_unaffected_part(lookup, _peek, sel_names, sel_values, _reliable):
                    log(f"{indent}  -- {name} [{_peek or '?'}] (not affected — not unsuppressed)\n")
                    continue
            # ────────────────────────────────────────────────────────────────

            if pre_resolved:
                log(f"{indent}  Resolved (was suppressed): {name}\n")
            # Resolve if the component was suppressed/lightweight OR its doc is
            # None (not loaded). Restore target captured in restore_target,
            # reapplied at the end of the run.
            elif (restore_target in (0, 1)) or doc is None:
                was_suppressed = True
                state_lbl = {0: "suppressed", 1: "lightweight"}.get(restore_target, "not loaded")
                log(f"{indent}  {state_lbl.capitalize()} (resolving to update): {name}\n")
                doc, path = resolve_component(sw_app, child, log, indent)
                if doc is None:
                    log(f"{indent}  Could not resolve {name} — skipping\n")
                    continue
            visited.add(path_key)
            _track_opened_doc(path)
            desc = get_description(doc) if doc else "N/A"
            node_info[path_key] = {"parent": current_path, "desc": desc, "depth": depth,
                                   "old_number": get_part_number_from_path(path), "type": "part"}
            app_names, app_values = get_applicable_params(lookup, desc, sel_names, sel_values)
            if WHY_LOG: log(f"{indent}     {_why_affected(lookup, desc, sel_names)}\n")

            # ─── Material check (Front Panel, etc.) ───
            abbrev      = get_abbreviation(lookup, desc)
            target_mat  = get_selected_material(abbrev)   # None unless abbrev is in MATERIAL_CHOICES / a group
            current_mat = ""
            material_needs_update = False
            if abbrev and not target_mat:
                # Abbreviation found but it's not one we have material options for.
                # Helpful when debugging why a part didn't get a material change.
                log(f"{indent}     [material] abbrev='{abbrev}' (no material options configured)\n")
            if target_mat:
                current_mat = get_current_material(doc, log, indent) if doc else ""
                # Diagnostic: surface exactly what was read so material issues are
                # easy to debug from the log.
                log(f"{indent}     [material] abbrev='{abbrev}' read='{current_mat or '(empty)'}' target='{target_mat}'\n")
                # Flag material-needs-update only when we have a usable current
                # material reading AND it differs from the target. _material_differs
                # guards against unreadable / formula-string reads so selecting the
                # SAME material the part already has does not get flagged.
                if _material_differs(current_mat, target_mat):
                    material_needs_update = True
            elif app_names and doc:
                # Not a material-option part, but it may get a new number if it
                # isn't in the BOM — record its ACTUAL material so the Master BOM
                # row is filled. Quiet read (no diagnostic spam for parts that
                # legitimately have no material).
                current_mat = get_current_material(doc)

            # Process this part if EITHER dimensions matter (app_names) OR material needs update
            if app_names or material_needs_update:
                log(f"{indent}  >> {name} [Part] [{desc}]\n")
                if app_names:
                    log(f"{indent}     Params: {','.join(app_names)}\n")
                if target_mat:
                    mat_lbl = f"current={current_mat or '?'} target={target_mat}"
                    if material_needs_update:
                        log(f"{indent}     Material check: {mat_lbl}  ⇒ needs update\n")
                    else:
                        log(f"{indent}     Material check: {mat_lbl}  ⇒ ok\n")
                status_cb(f"Checking {name}..."); root_tk.update()
                # pv = all selected params new values + current SW values for unselected
                pv = {n: format_value(v) for n, v in zip(sel_names, sel_values)}
                for p_name in cur_eq_vals:
                    if p_name not in pv: pv[p_name] = format_value(cur_eq_vals[p_name])
                # Build FULL bom_names/bom_values: for params applicable to this part
                # but NOT selected by user, use current SolidWorks equation value
                guideline_row = _guideline_row(lookup, desc) or {}
                bom_names, bom_values = [], []
                for p_name, _ in PARAMS:
                    if not guideline_row.get(p_name, False): continue
                    if p_name in sel_names:
                        idx = sel_names.index(p_name)
                        bom_names.append(p_name); bom_values.append(sel_values[idx])
                    elif p_name in cur_eq_vals:
                        bom_names.append(p_name); bom_values.append(cur_eq_vals[p_name])
                log(f"{indent}     BOM check: {list(zip(bom_names,bom_values))}\n")
                # Always filter by the selected material for material-bearing
                # parts. target_mat is None for parts with no material option, so
                # those stay unconstrained. Blank Material cells are treated as
                # wildcards by find_in_master_bom, so this no longer forces
                # unnecessary new numbers — but it DOES reject a wrong-material
                # row (e.g. a Brushed-SS row when Painted is selected) even when
                # the part's own material isn't changing.
                bom_mat_filter = target_mat
                # DP / PS / SIL,TP descriptions encode the case dims and the finish.
                # The BOM is searched BY DESCRIPTION, so search on what the part WILL
                # be after the run, not what it currently says.
                _newv = dict(cur_eq_vals)
                for _i, _n in enumerate(sel_names):
                    _newv[_n] = sel_values[_i]
                tgt_desc = build_target_description(desc, cur_eq_vals, _newv, target_mat)
                if tgt_desc != desc:
                    log(f"{indent}     Description will become: {tgt_desc}\n")
                existing = find_in_master_bom(tgt_desc, bom_names, bom_values, False,
                                              parts_rows, assy_rows, assy_config,
                                              required_material=bom_mat_filter)
                if existing:
                    if existing.upper() == get_part_number_from_path(path).upper():
                        # BOM says this exact number matches desc+dims+material. If
                        # we still think material needs update there is a BOM
                        # inconsistency — force a needs_update so the macro path
                        # generates a fresh number and the new file gets target
                        # material applied. Otherwise everything is in sync → skip.
                        if material_needs_update:
                            log(f"{indent}     ! BOM lists this number but material is changing to {target_mat}. Treating as needs new number.\n")
                            existing = None; status = "needs_update"
                        else:
                            log(f"{indent}     → Same number already in BOM — skip\n")
                            existing = None; status = "skip"
                    else:
                        ext_path = os.path.join(folder, existing + ".SLDPRT")
                        on_disk  = os.path.exists(ext_path)
                        status   = "existing" if on_disk else "file_missing"
                        log(f"{indent}     → BOM FOUND: {existing} {'✓' if on_disk else '✗ File Missing'}\n")
                else:
                    existing = None; status = "needs_update"
                    log(f"{indent}     → Not in BOM → needs new number\n")
                if status != "skip":
                    # apply_material: True if we plan to overwrite material on this
                    # entry's part. Only relevant when status == 'needs_update' —
                    # for 'existing' / 'file_missing' the replacement file from BOM
                    # is assumed to carry the correct material already.
                    apply_material = (status == "needs_update" and target_mat is not None and material_needs_update)
                    entries.append({
                        "old_path":        path,
                        "old_number":      get_part_number_from_path(path),
                        "desc":            desc,
                        "target_desc":     tgt_desc,
                        "parent_path":     current_path,
                        "depth":           depth,
                        "type":            "part",
                        "app_names":       app_names,
                        "app_values":      app_values,
                        "param_values":    pv,
                        "existing_number": existing,
                        "new_number":      existing or "Pending...",
                        "new_path":        "",
                        "status":          status,
                        "_status_lbl":     None,
                        "abbrev":          abbrev,
                        "target_material": target_mat,
                        "current_material": current_mat,
                        "apply_material":  apply_material,
                    })
            else:
                log(f"{indent}  -- {name} [{desc}] (not affected)\n")
        except Exception as e:
            log(f"{indent}  Error (part): {e}\n")
        finally:
            # Defer re-suppression until after the apply phase saves, so the
            # part's equations get updated while it stays resolved. Store identity
            # by path (COM ref would go stale if the part is renamed/replaced).
            if was_suppressed:
                resolved_comps.append((current_path, path, name, restore_target))
    return entries

# ─────────────────────────────────────────
# Export Excel
# ─────────────────────────────────────────
def export_to_excel(summary_data, save_path):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Update Summary"
    navy="1A4F8A"; cream="F5F0E8"; green="D4EDDA"; amber="FFF3CD"; red="FADADD"; white="FFFFFF"; lblue="CDDDF5"
    HF = Font(name="Calibri", bold=True,  color="FFFFFF", size=12)
    SF = Font(name="Calibri", bold=True,  color=navy,     size=12)
    BF = Font(name="Calibri",              size=11)
    DF = Font(name="Calibri", bold=True,   size=11)
    CA = Alignment(horizontal="center", vertical="center")
    LA = Alignment(horizontal="left",   vertical="center")
    thin = Side(style="thin", color="C8BFB2")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    def hc(r,c,v,span=1):
        x=ws.cell(row=r,column=c,value=v); x.font=HF; x.fill=PatternFill("solid",fgColor=navy); x.alignment=CA; x.border=bdr
        if span>1: ws.merge_cells(start_row=r,start_column=c,end_row=r,end_column=c+span-1)
    def sc(r,c,v,bg=cream,span=1):
        x=ws.cell(row=r,column=c,value=v); x.font=SF; x.fill=PatternFill("solid",fgColor=bg); x.alignment=LA; x.border=bdr
        if span>1: ws.merge_cells(start_row=r,start_column=c,end_row=r,end_column=c+span-1)
    def bc(r,c,v,bg=white,bold=False,center=False):
        x=ws.cell(row=r,column=c,value=v); x.font=DF if bold else BF; x.fill=PatternFill("solid",fgColor=bg); x.alignment=CA if center else LA; x.border=bdr
    r=1; hc(r,1,"SOLIDWORKS CONFIGURATOR — SUMMARY REPORT",span=5); ws.row_dimensions[r].height=22; r+=1
    bc(r,1,"Job Number:",bg=cream,bold=True)
    c2=ws.cell(row=r,column=2,value=os.path.splitext(summary_data["assembly"] or "")[0]); c2.font=BF; c2.fill=PatternFill("solid",fgColor=white); c2.alignment=LA; c2.border=bdr
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5); r+=1
    if summary_data.get("reference"): bc(r,1,"Reference Used:",bg=cream,bold=True); bc(r,2,summary_data["reference"]); r+=1
    if summary_data.get("user"): bc(r,1,"Generated by:",bg=cream,bold=True); bc(r,2,summary_data["user"]); r+=1
    bc(r,1,"Timestamp:",bg=cream,bold=True); bc(r,2,summary_data["timestamp"]); r+=1
    if summary_data.get("time_taken"): bc(r,1,"Time Taken:",bg=cream,bold=True); bc(r,2,summary_data["time_taken"])
    r+=2
    sc(r,1,"DIMENSION CHANGES",span=5); r+=1
    for col,txt in enumerate(["Parameter","Original Value","New Value"],1): hc(r,col,txt)
    r+=1
    for dim in summary_data["dimensions"]:
        bc(r,1,dim["name"],bg=cream,bold=True); bc(r,2,dim["original"],center=True); bc(r,3,dim["new"],bg=lblue,center=True); r+=1
    r+=1
    if summary_data.get("materials"):
        sc(r,1,"MATERIAL CHANGES",span=5); r+=1
        for col,txt in enumerate(["Material Type","Finish Applied","Parts Changed"],1): hc(r,col,txt)
        r+=1
        for m in summary_data["materials"]:
            bc(r,1,m["name"],bg=cream,bold=True); bc(r,2,m["finish"],bg=lblue,center=True); bc(r,3,m["count"],center=True); r+=1
        r+=1
    sc(r,1,"SUMMARY COUNTS",span=5); r+=1
    counts=summary_data["counts"]
    bc(r,1,"Parts Updated:",bg=cream,bold=True); bc(r,2,f"{counts['parts_new']} new numbers",bg=green); bc(r,3,f"Existing replaced: {counts['parts_existing']}",bg=amber); bc(r,4,f"Missing: {counts['parts_missing']}",bg=amber); r+=1
    bc(r,1,"Subassemblies Updated:",bg=cream,bold=True); bc(r,2,f"{counts['assy_new']} new numbers",bg=green); bc(r,3,f"Existing replaced: {counts['assy_existing']}",bg=amber); bc(r,4,f"Failed: {counts['assy_failed']}",bg=red); r+=2
    def tbl(entries, title):
        nonlocal r
        if not entries: return
        sc(r,1,f"{title} ({len(entries)})",span=5); r+=1
        for col,txt in enumerate(["Old Number","Description","Updates","New/Existing Number","Status"],1): hc(r,col,txt)
        r+=1
        for e in entries:
            st=e.get("status_display","")
            bg=green if st in ("New Number","Replaced") else (amber if st in ("File Missing","Drawing Missing") else red)
            bc(r,1,e["old_number"],bold=True); bc(r,2,e["desc"]); bc(r,3,e.get("updates_label","—"),center=True)
            bc(r,4,e["new_number"],bg=bg,bold=True,center=True); bc(r,5,st,bg=bg,center=True); r+=1
        r+=1
    tbl(summary_data["parts"],"PARTS"); tbl(summary_data["assemblies"],"SUBASSEMBLIES")
    ws.column_dimensions["A"].width=22; ws.column_dimensions["B"].width=38
    ws.column_dimensions["C"].width=20; ws.column_dimensions["D"].width=22; ws.column_dimensions["E"].width=16
    wb.save(save_path)

# ─────────────────────────────────────────
# Export PDF
# ─────────────────────────────────────────
def export_to_pdf(summary_data, save_path):
    doc=SimpleDocTemplate(save_path,pagesize=landscape(letter),leftMargin=0.5*inch,rightMargin=0.5*inch,topMargin=0.5*inch,bottomMargin=0.5*inch)
    story=[]; nc=colors.HexColor; navy_c=nc("#1A4F8A"); cream_c=nc("#F5F0E8"); green_c=nc("#D4EDDA"); amber_c=nc("#FFF3CD"); red_c=nc("#FADADD"); blue_c=nc("#CDDDF5")
    ts=ParagraphStyle("t",fontSize=16,fontName="Helvetica-Bold",textColor=colors.white)
    h2=ParagraphStyle("h2",fontSize=11,fontName="Helvetica-Bold",textColor=navy_c,spaceAfter=4,spaceBefore=10)
    tt=Table([[Paragraph("SolidWorks Configurator — Summary Report",ts)]],colWidths=[10*inch])
    tt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),navy_c),("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),("LEFTPADDING",(0,0),(-1,-1),12)]))
    story.append(tt); story.append(Spacer(1,8))
    _mi_l=ParagraphStyle("mi_l",fontName="Helvetica-Bold",fontSize=9,leading=11)
    _mi_v=ParagraphStyle("mi_v",fontName="Helvetica",fontSize=9,leading=11)
    def _MI(txt,lbl=False):
        s=str(txt if txt is not None else "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        return Paragraph(s,_mi_l if lbl else _mi_v)
    _mi_rows=[["Job Number:",os.path.splitext(summary_data["assembly"] or "")[0]]]
    if summary_data.get("reference"): _mi_rows.append(["Reference Used:",summary_data["reference"]])
    if summary_data.get("user"): _mi_rows.append(["Generated by:",summary_data["user"]])
    _mi_rows.append(["Timestamp:",summary_data["timestamp"]])
    if summary_data.get("time_taken"): _mi_rows.append(["Time Taken:",summary_data["time_taken"]])
    mt_data=[[_MI(a,lbl=True),_MI(b)] for a,b in _mi_rows]
    mt=Table(mt_data,colWidths=[1.7*inch,8.3*inch])
    mt.setStyle(TableStyle([("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),("FONTNAME",(1,0),(1,-1),"Helvetica"),("FONTSIZE",(0,0),(-1,-1),9),("BACKGROUND",(0,0),(0,-1),cream_c),("GRID",(0,0),(-1,-1),0.5,nc("#C8BFB2")),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
    story.append(mt); story.append(Spacer(1,8))
    counts=summary_data["counts"]
    cd=[["Parts:",f"{counts['parts_new']} new numbers",f"Existing: {counts['parts_existing']}",f"Missing: {counts['parts_missing']}"],
        ["Subassemblies:",f"{counts['assy_new']} new numbers",f"Existing: {counts['assy_existing']}",f"Failed: {counts['assy_failed']}"]]
    ct=Table(cd,colWidths=[1.8*inch,2.2*inch,2*inch,2*inch])
    ct.setStyle(TableStyle([("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),("FONTNAME",(1,0),(-1,-1),"Helvetica"),("FONTSIZE",(0,0),(-1,-1),9),("BACKGROUND",(0,0),(0,-1),cream_c),("BACKGROUND",(1,0),(1,-1),green_c),("GRID",(0,0),(-1,-1),0.5,nc("#C8BFB2")),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
    story.append(Paragraph("SUMMARY COUNTS",h2)); story.append(ct); story.append(Spacer(1,8))
    if summary_data["dimensions"]:
        story.append(Paragraph("DIMENSION CHANGES",h2))
        dd=[["Parameter","Original","New Value"]]+[[d["name"],d["original"],d["new"]] for d in summary_data["dimensions"]]
        dt=Table(dd,colWidths=[2*inch,2*inch,2*inch])
        dt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),navy_c),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTNAME",(0,1),(-1,-1),"Helvetica"),("FONTSIZE",(0,0),(-1,-1),9),("BACKGROUND",(2,1),(2,-1),blue_c),("GRID",(0,0),(-1,-1),0.5,nc("#C8BFB2")),("ALIGN",(1,0),(-1,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
        story.append(dt); story.append(Spacer(1,8))
    if summary_data.get("materials"):
        story.append(Paragraph("MATERIAL CHANGES",h2))
        md=[["Material Type","Finish Applied","Parts Changed"]]+[[m["name"],m["finish"],str(m["count"])] for m in summary_data["materials"]]
        mtb=Table(md,colWidths=[2.4*inch,2*inch,1.6*inch])
        mtb.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),navy_c),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTNAME",(0,1),(-1,-1),"Helvetica"),("FONTSIZE",(0,0),(-1,-1),9),("BACKGROUND",(1,1),(1,-1),blue_c),("GRID",(0,0),(-1,-1),0.5,nc("#C8BFB2")),("ALIGN",(1,0),(-1,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
        story.append(mtb); story.append(Spacer(1,8))
    def mktbl(title,entries):
        if not entries: return
        story.append(Paragraph(f"{title} ({len(entries)})",h2))
        # Wrap the two long-text columns (Description, Updates) in Paragraphs so
        # reportlab wraps them inside the column instead of overflowing (e.g.
        # "Equation + Material (430-22BRSS)" spilling past the cell).
        _cs_l=ParagraphStyle("cell_l",fontName="Helvetica",fontSize=8,leading=9)
        _cs_c=ParagraphStyle("cell_c",fontName="Helvetica",fontSize=8,leading=9,alignment=1)
        def _P(txt,ctr=False):
            s=str(txt if txt is not None else "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
            return Paragraph(s,_cs_c if ctr else _cs_l)
        td=[["Old Number","Description","Updates","New/Existing Number","Status"]]+[
            [e["old_number"],_P(e["desc"]),_P(e.get("updates_label","—"),ctr=True),e["new_number"],e.get("status_display","")]
            for e in entries]
        t=Table(td,colWidths=[1.4*inch,3.1*inch,2.2*inch,1.8*inch,1.2*inch])
        rs=[]
        for i,e in enumerate(entries,1):
            st=e.get("status_display","")
            if st=="New Number": bg_st=green_c
            elif st=="Replaced": bg_st=blue_c
            elif st in ("File Missing","Drawing Missing"): bg_st=amber_c
            else: bg_st=red_c
            rs.append(("BACKGROUND",(3,i),(4,i),bg_st))
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),navy_c),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTNAME",(0,1),(-1,-1),"Helvetica"),("FONTSIZE",(0,0),(-1,-1),9),("GRID",(0,0),(-1,-1),0.5,nc("#C8BFB2")),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("ALIGN",(2,0),(-1,-1),"CENTER")]+rs))
        story.append(t); story.append(Spacer(1,8))
    mktbl("PARTS",summary_data["parts"]); mktbl("SUBASSEMBLIES",summary_data["assemblies"])
    doc.build(story)

# ─────────────────────────────────────────
# GUI
# ─────────────────────────────────────────
def build_app():
    pythoncom.CoInitialize()
    sw_app=sw_model=None; assy_title="Not connected"; main_path=""; top_desc=""
    try:
        sw_app=win32com.client.Dispatch("SldWorks.Application"); doc=sw_app.ActiveDoc
        if doc is None: assy_title="No active document"
        else:
            try:   doc_type=int(doc.GetType())
            except TypeError: doc_type=int(doc.GetType)
            if doc_type!=2: assy_title="Active doc is not an assembly"
            else:
                sw_model=doc
                try:   assy_title=doc.GetTitle()
                except TypeError: assy_title=doc.GetTitle
                try:   main_path=doc.GetPathName()
                except TypeError: main_path=doc.GetPathName
    except Exception as e: assy_title=f"Connection failed: {e}"

    # ── Product-line detection ──────────────────────────────────────
    # Read the open assembly's Description custom property; the prefix
    # determines which product line's BOM/sheets/etc. the tool uses.
    if sw_model is not None:
        try:
            top_desc = get_description(sw_model) or ""
        except Exception:
            top_desc = ""
        top_desc_up = top_desc.strip().upper()
        if top_desc_up.startswith("PFI-"):
            picked = "PFI"
        elif top_desc_up.startswith("PF-"):
            picked = "PF"
        elif top_desc_up.startswith("PT-"):
            picked = "PT"
        else:
            picked = None
        if picked is None:
            messagebox.showerror(
                "Product Line Not Detected",
                "Could not determine the product line from the open assembly.\n\n"
                "The Description custom property must start with 'PF-', 'PFI-' or 'PT-'.\n"
                f"Got: '{top_desc}'\n\n"
                "Open the correct assembly and relaunch the tool."
            )
            return
        cfg = PRODUCT_LINES[picked]
        ACTIVE_LINE["key"] = picked
        ACTIVE_LINE["cfg"] = cfg
        # Resolve Master BOM path (next to the EXE / script)
        ACTIVE_LINE["bom_path"] = os.path.join(_tool_dir(), cfg["bom_filename"])
    # If no model is open we still let the UI come up (the tool will show a
    # "Not connected" status); product-line-dependent operations will be
    # blocked anyway because there's nothing to scan.

    import sys
    if getattr(sys, "frozen", False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path=os.path.join(script_dir,"Guideline Document for the Parts Impacted by Dimensional Parameters.xlsx")
    excel_ok=os.path.exists(excel_path)
    current_vals=get_current_equation_values(sw_model) if sw_model else {}

    # Mutable connection state so the "Connect to Model" button can refresh it
    # without restarting the app. Index [0] holds the live value.
    conn_state = {
        "sw_model": sw_model,
        "main_path": main_path,
        "assy_title": assy_title,
    }

    all_entries      = []
    orig_vals_snap   = {}
    new_vals_snap    = {}
    resolved_comps_all = []   # components resolved from suppressed; re-suppress after apply
    update_done      = [False]
    gen_done         = [False]
    automation_started = [False]
    summary_generated = [False]
    qc_answers       = {}
    qc_measure_started = set()
    qc_macro_complete = [False]
    qc_report_path   = [""]
    qc_record_path   = [""]
    qc_completed     = [False]
    qc_question_index = [0]
    qc_dialog        = [None]
    # Re-entrancy guard. True while a long operation (scan / phase 2 / generate)
    # is actively running. Action handlers refuse to start a second operation
    # while one is in flight — important because long ops pump the event loop
    # with root.update(), which can otherwise deliver a stray click (e.g. on
    # "Connect to Model") mid-run and corrupt state. False = idle / waiting for
    # the user's next action.
    _busy            = [False]
    update_timestamp = [""]
    total_time_str   = [""]
    job_reference    = [""]   # the number the model was originally open as, before a job Save-As
    stored_sel       = [[], []]
    last_results     = ["replace"]
    timer_running    = [False]
    timer_elapsed    = [0.0]
    timer_start      = [0.0]
    timer_job        = [None]
    # Per-run cache of the loaded lookup tables, so PHASE 2 reuses what the scan
    # already parsed instead of re-reading both workbooks. Set in run_update,
    # consumed in do_phase2, cleared on a fresh run / reconnect.
    cached_lookup     = [None]
    cached_parts_rows = [None]
    cached_assy_rows  = [None]

    root=ctk.CTk(); root.title("SolidWorks Configurator")

    def _close_qc_dialog(dialog):
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()
        qc_dialog[0] = None

    def _show_qc_step():
        dialog = qc_dialog[0]
        if dialog is None or not dialog.winfo_exists():
            return
        for child in dialog.winfo_children():
            child.destroy()

        questions = build_qc_questions()
        index = qc_question_index[0]
        key, prompt = questions[index]
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            dialog,
            text=f"QC Checkpoint {index + 1} of {len(questions)}",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=C["text"],
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(22, 6))
        card = ctk.CTkFrame(
            dialog, fg_color=C["card"], corner_radius=14,
            border_width=1, border_color=C["border_light"],
        )
        card.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            card, text=prompt, font=ctk.CTkFont(size=15, weight="bold"),
            text_color=C["text"], anchor="w", wraplength=470, justify="left",
        ).grid(row=0, column=0, sticky="w", padx=18, pady=(22, 12))
        answer_var = ctk.StringVar(value=qc_answers.get(key, ""))
        is_profile_attribute = key in dict(PARAMS)
        answer_entry = ctk.CTkEntry(
            card, textvariable=answer_var, height=42,
            placeholder_text="Type QC result",
            font=ctk.CTkFont(size=14),
            state="normal" if not is_profile_attribute or key in qc_measure_started else "disabled",
        )
        if is_profile_attribute:
            measure_status = ctk.CTkLabel(
                card,
                text=("Measure tool opened. Measure this attribute in SolidWorks, then enter the result."
                      if key in qc_measure_started else
                      "Open SolidWorks Measure before entering this result."),
                font=ctk.CTkFont(size=11), text_color=C["text_dim"],
                anchor="w", wraplength=470, justify="left",
            )
            measure_status.grid(row=1, column=0, sticky="w", padx=18, pady=(0, 8))

            def open_measure():
                try:
                    if sw_app is None or sw_model is None:
                        raise RuntimeError("No active SolidWorks model is connected.")
                    if main_path:
                        if activate_doc(sw_app, main_path) is None:
                            raise RuntimeError("Could not activate the assembly being checked.")
                    elif sw_app.ActiveDoc is None:
                        raise RuntimeError("No active SolidWorks document is available.")
                    launched = sw_app.RunCommand(92, "")
                    if launched is False:
                        raise RuntimeError("SolidWorks did not start the Measure command.")
                except Exception as e:
                    messagebox.showerror(
                        "Could not open Measure",
                        f"Measure could not be started in SolidWorks:\n{e}",
                        parent=dialog,
                    )
                    return
                qc_measure_started.add(key)
                answer_entry.configure(state="normal")
                measure_status.configure(
                    text="Measure is open in SolidWorks. Measure this attribute, then return here to enter the result."
                )

            ctk.CTkButton(
                card, text="Open Measure in SolidWorks", height=34,
                fg_color=C["accent_light"], hover_color=C["accent_line"],
                text_color=C["accent"], command=open_measure,
            ).grid(row=2, column=0, sticky="w", padx=18, pady=(0, 12))
            answer_entry.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 22))
        else:
            answer_entry.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 22))
            answer_entry.focus_set()

        buttons = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 18))
        buttons.grid_columnconfigure(1, weight=1)

        def go_back():
            if answer_var.get().strip():
                qc_answers[key] = answer_var.get().strip()
            if qc_question_index[0] > 0:
                qc_question_index[0] -= 1
                _show_qc_step()

        def save_answer():
            answer = answer_var.get().strip()
            if not answer:
                messagebox.showwarning(
                    "QC answer required",
                    f"Enter an answer for {prompt} before continuing.",
                    parent=dialog,
                )
                answer_entry.focus_set()
                return
            qc_answers[key] = answer
            if qc_question_index[0] < len(questions) - 1:
                qc_question_index[0] += 1
                _show_qc_step()
                return
            try:
                qc_record_path[0] = save_qc_record(qc_answers, assy_title)
            except Exception as e:
                messagebox.showerror(
                    "QC answers could not be saved",
                    f"The answers were not saved, so the app will remain open.\n\n{e}",
                    parent=dialog,
                )
                return
            _show_qc_completion()

        if index > 0:
            ctk.CTkButton(
                buttons, text="← Back", width=100, height=36,
                command=go_back,
            ).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            buttons, text="Cancel", width=100, height=36,
            fg_color=C["card"], hover_color=C["card2"],
            text_color=C["text_dim"], border_width=1,
            border_color=C["border"],
            command=lambda: _close_qc_dialog(dialog),
        ).grid(row=0, column=1, sticky="e", padx=(0, 8))
        ctk.CTkButton(
            buttons,
            text="Save QC Record & Exit" if index == len(questions) - 1 else "Next →",
            width=190, height=36, fg_color=C["accent"],
            hover_color=C["accent_hover"], text_color="#ffffff",
            command=save_answer,
        ).grid(row=0, column=2, sticky="e")
        answer_entry.bind("<Return>", lambda _event: save_answer())

    def start_qc_wizard():
        if qc_dialog[0] is not None and qc_dialog[0].winfo_exists():
            qc_dialog[0].lift()
            return
        dialog = ctk.CTkToplevel(root)
        qc_dialog[0] = dialog
        dialog.title("Quality Control")
        dialog.geometry("560x360")
        dialog.resizable(False, False)
        dialog.transient(root)
        dialog.protocol("WM_DELETE_WINDOW", lambda: _close_qc_dialog(dialog))
        _show_qc_step()

    def _show_qc_completion():
        dialog = qc_dialog[0]
        if dialog is None or not dialog.winfo_exists():
            return
        for child in dialog.winfo_children():
            child.destroy()
        dialog.title("Run QC Tools Macro")
        dialog.geometry("650x390")
        dialog.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            dialog, text="QC answers saved",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=C["success"],
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(22, 8))
        ctk.CTkLabel(
            dialog,
            text=f"QC answers:\n{qc_record_path[0]}\n\n"
                 "Run the SolidWorks QC Tools macro to generate the job report.",
            font=ctk.CTkFont(size=11), text_color=C["text"],
            anchor="w", justify="left", wraplength=560,
        ).grid(row=1, column=0, sticky="ew", padx=24, pady=10)

        fields = ctk.CTkFrame(
            dialog, fg_color=C["card"], corner_radius=12,
            border_width=1, border_color=C["border_light"],
        )
        fields.grid(row=2, column=0, sticky="ew", padx=20, pady=8)
        fields.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            fields, text="Macro module", font=ctk.CTkFont(size=11, weight="bold"),
            text_color=C["text_dim"],
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 4))
        module_var = ctk.StringVar(value=QC_MACRO_MODULE)
        ctk.CTkEntry(fields, textvariable=module_var, height=32).grid(
            row=0, column=1, sticky="ew", padx=14, pady=(12, 4)
        )
        ctk.CTkLabel(
            fields, text="Procedure", font=ctk.CTkFont(size=11, weight="bold"),
            text_color=C["text_dim"],
        ).grid(row=1, column=0, sticky="w", padx=14, pady=(4, 12))
        procedure_var = ctk.StringVar(value=QC_MACRO_FUNCTION)
        ctk.CTkEntry(fields, textvariable=procedure_var, height=32).grid(
            row=1, column=1, sticky="ew", padx=14, pady=(4, 12)
        )
        macro_status = ctk.CTkLabel(
            dialog,
            text=(f"Report ready: {qc_report_path[0]}" if qc_macro_complete[0]
                  else "Macro has not run yet."),
            font=ctk.CTkFont(size=11),
            text_color=C["success"] if qc_macro_complete[0] else C["text_dim"],
            anchor="w", wraplength=600, justify="left",
        )
        macro_status.grid(row=3, column=0, sticky="ew", padx=24, pady=6)

        buttons = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons.grid(row=4, column=0, sticky="ew", padx=20, pady=(8, 18))
        buttons.grid_columnconfigure(0, weight=1)

        def open_report(path, title):
            try:
                os.startfile(path)
            except OSError as e:
                messagebox.showerror(
                    title, f"Could not open the report:\n{e}", parent=dialog
                )

        def run_macro():
            macro_button.configure(state="disabled", text="Running macro...")
            dialog.update_idletasks()
            try:
                macro_path = resolve_qc_macro_path(parent=dialog)
                qc_report_path[0] = run_qc_tools_macro(
                    sw_app, main_path, assy_title,
                    module_name=module_var.get(),
                    procedure_name=procedure_var.get(),
                    macro_path=macro_path,
                )
            except Exception as e:
                macro_status.configure(
                    text=f"Macro failed: {e}", text_color=C["danger"]
                )
                messagebox.showerror(
                    "QC Tools Macro Failed",
                    f"The macro did not complete. Correct the macro details or select "
                    f"the right file, then try again.\n\n{e}",
                    parent=dialog,
                )
                macro_button.configure(state="normal", text="Run QC Tools Macro")
                return
            qc_macro_complete[0] = True
            qc_completed[0] = True
            macro_status.configure(
                text=f"Report ready: {qc_report_path[0]}",
                text_color=C["success"],
            )
            open_report_button.configure(state="normal")
            close_button.configure(state="normal")
            macro_button.configure(state="normal", text="Run Macro Again")

        macro_button = ctk.CTkButton(
            buttons, text="Run QC Tools Macro", width=190, height=38,
            fg_color=C["accent"], hover_color=C["accent_hover"],
            text_color="#ffffff", command=run_macro,
        )
        macro_button.grid(row=0, column=0, sticky="w", padx=(0, 8))
        open_report_button = ctk.CTkButton(
            buttons, text="Open QC Report", width=140, height=38,
            state="normal" if qc_macro_complete[0] else "disabled",
            command=lambda: open_report(
                qc_report_path[0], "Could not open QC report"
            ),
        )
        open_report_button.grid(row=0, column=1, sticky="e", padx=(0, 8))
        ctk.CTkButton(
            buttons, text="Open QC Answers", width=130, height=38,
            command=lambda: open_report(
                qc_record_path[0], "Could not open QC answers"
            ),
        ).grid(row=0, column=2, sticky="e", padx=(0, 8))
        close_button = ctk.CTkButton(
            buttons, text="Close App", width=100, height=38,
            fg_color=C["success"], hover_color=C["success"],
            text_color="#ffffff", state="normal" if qc_macro_complete[0] else "disabled",
            command=root.destroy,
        )
        close_button.grid(row=0, column=3, sticky="e")

    def request_app_exit():
        if _busy[0]:
            messagebox.showwarning(
                "Operation in progress",
                "Wait for the current operation to finish before exiting.",
                parent=root,
            )
            return
        if not automation_started[0]:
            root.destroy()
            return
        if qc_completed[0]:
            root.destroy()
            return
        if not summary_generated[0]:
            messagebox.showwarning(
                "Automation must be completed",
                "The automation process has started. Complete the update and generate its summary "
                "before exiting; the QC checkpoints will follow.",
                parent=root,
            )
            return
        start_qc_wizard()

    root.protocol("WM_DELETE_WINDOW", request_app_exit)

    # Window icon (top-left of the window / Alt-Tab). With --onefile, bundled
    # data is extracted to sys._MEIPASS at runtime — NOT next to the exe — so we
    # look there first, then fall back to a loose assets/ folder for dev runs.
    try:
        import sys as _sys
        _icon_bases = []
        if getattr(_sys, "_MEIPASS", None):        # PyInstaller onefile temp dir
            _icon_bases.append(_sys._MEIPASS)
        _icon_bases.append(os.path.dirname(os.path.abspath(__file__)))
        if getattr(_sys, "frozen", False):
            _icon_bases.append(os.path.dirname(_sys.executable))
        for _b in _icon_bases:
            _ico = os.path.join(_b, "assets", "app.ico")
            if os.path.exists(_ico):
                try:    root.iconbitmap(_ico)          # Windows .ico
                except Exception:
                    try: root.wm_iconbitmap(_ico)
                    except Exception: pass
                break
    except Exception:
        pass
    root.geometry("1250x780"); root.minsize(1000,680); root.resizable(True,True); root.configure(fg_color=C["bg"])
    # Single-column shell: MAIN fills the window and stacks a slim TOP BAR
    # (row 0: Connect-to-Model + Log), a clickable STEPPER (row 1), and the
    # swappable CONTENT (row 2). Navigation is the stepper — there is no sidebar.
    root.grid_columnconfigure(0,weight=1); root.grid_rowconfigure(0,weight=1)

    cur_panel_key=["input"]
    nav_enabled={"input":True,"log":True,"replace":False,"newnum":False,"summary":False}
    nav_done   ={"input":False,"log":False,"replace":False,"newnum":False,"summary":False}

    sw_bc=C["success"] if sw_model else C["danger"]; sw_bb=C["success_bg"] if sw_model else C["danger_bg"]
    xc=C["success"] if excel_ok else C["danger"]; xb=C["success_bg"] if excel_ok else C["danger_bg"]

    main=ctk.CTkFrame(root,fg_color=C["bg"],corner_radius=0)
    main.grid(row=0,column=0,sticky="nsew"); main.grid_columnconfigure(0,weight=1); main.grid_rowconfigure(2,weight=1)

    # ── TOP BAR (row 0) ──
    topbar=ctk.CTkFrame(main,fg_color=C["card"],corner_radius=0,height=56)
    topbar.grid(row=0,column=0,sticky="ew"); topbar.grid_propagate(False); topbar.grid_columnconfigure(1,weight=1)
    topbar.grid_rowconfigure(0,weight=1)
    ctk.CTkFrame(topbar,fg_color=C["border_light"],height=1,corner_radius=0).place(relx=0,rely=1.0,anchor="sw",relwidth=1.0)
    brand=ctk.CTkFrame(topbar,fg_color="transparent"); brand.grid(row=0,column=0,sticky="w",padx=18)
    ctk.CTkLabel(brand,text="SW",width=26,height=26,corner_radius=10,
                 fg_color=C["accent"],text_color="#ffffff",font=ctk.CTkFont(size=13,weight="bold")).pack(side="left")
    ctk.CTkLabel(brand,text="SolidWorks Configurator",font=ctk.CTkFont(size=15,weight="bold"),
                 text_color=C["text"]).pack(side="left",padx=(11,0))
    tb_right=ctk.CTkFrame(topbar,fg_color="transparent"); tb_right.grid(row=0,column=2,sticky="e",padx=14)
    connect_btn=ctk.CTkButton(tb_right,text="\u27f3 Connect to Model",font=ctk.CTkFont(size=12,weight="bold"),
                              fg_color=C["accent_light"],hover_color=C["accent_line"],text_color=C["accent"],
                              border_width=1,border_color=C["accent_line"],
                              corner_radius=10,height=34,width=164)
    connect_btn.grid(row=0,column=0,padx=4)

    # ── STEPPER (row 1) — clickable navigation ──
    stepbar=ctk.CTkFrame(main,fg_color=C["card"],corner_radius=0)
    stepbar.grid(row=1,column=0,sticky="ew")
    import tkinter as _tk
    stepbar_in=ctk.CTkFrame(stepbar,fg_color="transparent"); stepbar_in.pack(pady=(10,8))
    STEPS=[("input","1","User Input"),("replace","2","Replace Parts"),
           ("newnum","3","New Numbers"),("summary","4","Summary")]
    CONN_LABELS=["scanning tree","updating equations & materials","generating new numbers"]
    _CONN_W=64
    step_widgets={}; connectors=[]
    for si,(key,num,label) in enumerate(STEPS):
        if si>0:
            cf=ctk.CTkFrame(stepbar_in,fg_color="transparent"); cf.grid(row=0,column=si*2-1,padx=1)
            cvs=_tk.Canvas(cf,width=_CONN_W,height=10,bg=C["card"],highlightthickness=0,bd=0)
            cvs.grid(row=0,column=0,pady=(2,3))
            cap=ctk.CTkLabel(cf,text=CONN_LABELS[si-1],font=ctk.CTkFont(size=9),
                             text_color=C["text_muted"],wraplength=_CONN_W,justify="center")
            cap.grid(row=1,column=0)
            connectors.append(dict(canvas=cvs,cap=cap,frac=0.0,state="idle"))
        cell=ctk.CTkFrame(stepbar_in,fg_color="transparent",corner_radius=16)
        cell.grid(row=0,column=si*2)
        circ=ctk.CTkLabel(cell,text=num,width=24,height=24,corner_radius=16,
                          font=ctk.CTkFont(size=11,weight="bold"),fg_color=C["card2"],text_color=C["text_muted"])
        circ.grid(row=0,column=0,padx=(10,7),pady=4)
        txt=ctk.CTkLabel(cell,text=label,font=ctk.CTkFont(size=12,weight="bold"),text_color=C["text_muted"])
        txt.grid(row=0,column=1,padx=(0,12),pady=5)
        for w in (cell,circ,txt):
            w.bind("<Button-1>", lambda e,k=key: nav_go(k))
        step_widgets[key]=dict(cell=cell,circ=circ,txt=txt)
    ctk.CTkFrame(stepbar,fg_color=C["border"],height=1,corner_radius=0).pack(fill="x",side="bottom")

    # ── CONTENT (row 2; holds the swappable pages) ──
    # body: a PERSISTENT left context column (col 0) | divider (col 1) | swappable CONTENT (col 2).
    body=ctk.CTkFrame(main,fg_color=C["bg"],corner_radius=0)
    body.grid(row=2,column=0,sticky="nsew"); body.grid_rowconfigure(0,weight=1)
    body.grid_columnconfigure(0,weight=0,minsize=300); body.grid_columnconfigure(1,weight=0); body.grid_columnconfigure(2,weight=1)
    _divider=ctk.CTkFrame(body,fg_color=C["border"],width=1,corner_radius=0)
    _divider.grid(row=0,column=1,sticky="ns",pady=10); _divider.grid_propagate(False)
    content=ctk.CTkFrame(body,fg_color=C["bg"],corner_radius=0)
    content.grid(row=0,column=2,sticky="nsew"); content.grid_rowconfigure(0,weight=1); content.grid_columnconfigure(0,weight=1)

    # ── INPUT PAGE: just the form now (the context panel is persistent, built below) ──
    input_panel=ctk.CTkFrame(content,fg_color=C["bg"],corner_radius=0)
    input_panel.grid_rowconfigure(0,weight=1); input_panel.grid_columnconfigure(0,weight=1)

    def _section(parent,text,row,colspan=2,padx=6,pady=(2,8)):
        hf=ctk.CTkFrame(parent,fg_color="transparent")
        hf.grid(row=row,column=0,columnspan=colspan,sticky="ew",padx=padx,pady=pady)
        b=ctk.CTkFrame(hf,fg_color=C["accent"],width=4,height=18,corner_radius=2)
        b.grid(row=0,column=0,sticky="ns",padx=(0,8)); b.grid_propagate(False)
        ctk.CTkLabel(hf,text=text,font=ctk.CTkFont(size=12,weight="bold"),
                     text_color=C["text"],anchor="w").grid(row=0,column=1,sticky="w")
        return hf

    # ===== LEFT CONTEXT PANEL (persistent: body col 0, visible on every tab) =====
    # Scrollable left context column — on smaller screens the cards scroll
    # instead of clipping at the bottom.
    ctx=ctk.CTkScrollableFrame(body,fg_color=C["bg"],corner_radius=0,width=262)
    ctx.grid(row=0,column=0,sticky="nsew",padx=(12,4),pady=12); ctx.grid_columnconfigure(0,weight=1)

    # ── JOB NUMBER — entered first. On Run, the open model is saved as this job
    # number (same folder) and all updates then run on that job copy. Blank = run
    # on the current file, exactly as before.
    jobcard=ctk.CTkFrame(ctx,fg_color=C["card"],corner_radius=16,border_width=1,border_color=C["border_light"])
    jobcard.grid(row=0,column=0,sticky="ew",pady=(0,12)); jobcard.grid_columnconfigure(0,weight=1)
    _section(jobcard,"JOB NUMBER",0,colspan=1,padx=16,pady=(14,6))
    job_var=ctk.StringVar(value="")
    job_entry=ctk.CTkEntry(jobcard,textvariable=job_var,
                           placeholder_text="blank = use current file",
                           font=ctk.CTkFont(size=14,weight="bold"),height=36,
                           fg_color=C["card2"],border_color=C["border"],text_color=C["text"])
    job_entry.grid(row=1,column=0,sticky="ew",padx=16,pady=(0,6))
    ctk.CTkLabel(jobcard,text="Enter the new job number as per the job text",
                 font=ctk.CTkFont(size=10),text_color=C["text_dim"],anchor="w",
                 justify="left",wraplength=228).grid(row=2,column=0,sticky="w",padx=16,pady=(0,12))

    csum=ctk.CTkFrame(ctx,fg_color=C["card"],corner_radius=16,border_width=1,border_color=C["border_light"])
    csum.grid(row=1,column=0,sticky="ew"); csum.grid_columnconfigure(0,weight=1)
    _section(csum,"CASE SUMMARY",0,colspan=1,padx=16,pady=(14,6))
    title_lbl=ctk.CTkLabel(csum,text=assy_title,font=ctk.CTkFont(size=15,weight="bold"),
                           text_color=C["text"],anchor="w",justify="left",wraplength=228)
    title_lbl.grid(row=1,column=0,sticky="w",padx=16,pady=(0,2))
    _desc_show=top_desc if len(top_desc)<=72 else top_desc[:70]+"..."
    desc_lbl=ctk.CTkLabel(csum,text=_desc_show,font=ctk.CTkFont(size=11,weight="bold"),
                          text_color=C["text_dim"],anchor="w",justify="left",wraplength=228)
    desc_lbl.grid(row=2,column=0,sticky="w",padx=16,pady=(0,10))
    pillbox=ctk.CTkFrame(csum,fg_color="transparent"); pillbox.grid(row=3,column=0,sticky="w",padx=14,pady=(0,14))
    conn_pill=ctk.CTkLabel(pillbox,text="\u25cf Connected" if sw_model else "\u25cf Not Connected",
                           font=ctk.CTkFont(size=11,weight="bold"),fg_color=sw_bb,text_color=sw_bc,corner_radius=16,padx=10,pady=4)
    conn_pill.grid(row=0,column=0,padx=2,pady=2,sticky="w")
    excel_pill=ctk.CTkLabel(pillbox,text="\u2713 Excel OK" if excel_ok else "\u2717 Excel",
                            font=ctk.CTkFont(size=11,weight="bold"),fg_color=xb,text_color=xc,corner_radius=16,padx=10,pady=4)
    excel_pill.grid(row=0,column=1,padx=2,pady=2,sticky="w")

    status_var=ctk.StringVar(value="Ready to update" if sw_model else "Not connected")
    sline=ctk.CTkFrame(ctx,fg_color="transparent"); sline.grid(row=2,column=0,sticky="ew",pady=(10,0))
    sline.grid_columnconfigure(1,weight=1)
    sdot=ctk.CTkLabel(sline,text="\u25cf",font=ctk.CTkFont(size=12),text_color=C["success"] if sw_model else C["danger"])
    sdot.grid(row=0,column=0,padx=(4,6))
    ctk.CTkLabel(sline,textvariable=status_var,font=ctk.CTkFont(size=12),text_color=C["text_dim"],anchor="w").grid(row=0,column=1,sticky="w")
    ctk.CTkButton(sline,text="Log",font=ctk.CTkFont(size=11,weight="bold"),
                  fg_color="transparent",hover_color=C["border_light"],text_color=C["accent"],
                  border_width=1,border_color=C["border"],corner_radius=10,height=26,width=58,
                  command=lambda:show_panel(log_panel)).grid(row=0,column=2,sticky="e",padx=(4,2))
    def set_status(text,color=C["success"]):
        try: status_var.set(text); sdot.configure(text_color=color)
        except Exception: pass

    # Live "remaining" readout, shown under the status line during the
    # equation/material and new-number phases (blank otherwise).
    eq_remaining_var=ctk.StringVar(value="")
    eq_remaining_lbl=ctk.CTkLabel(sline,textvariable=eq_remaining_var,
                                  font=ctk.CTkFont(size=11,weight="bold"),
                                  text_color=C["accent"],anchor="w",justify="left",
                                  wraplength=250)   # wrap instead of clipping
    eq_remaining_lbl.grid(row=1,column=0,columnspan=3,sticky="ew",padx=(22,6),pady=(3,0))
    eq_remaining_lbl.grid_remove()   # hidden until a run populates it
    def set_eq_remaining(parts=0,assemblies=0):
        """Show 'N parts and M assemblies remaining'. Call with 0,0 to hide."""
        try:
            parts=max(0,int(parts or 0)); assemblies=max(0,int(assemblies or 0))
            if parts<=0 and assemblies<=0:
                eq_remaining_var.set(""); eq_remaining_lbl.grid_remove(); return
            bits=[]
            if parts>0:      bits.append(f"{parts} part{'s' if parts!=1 else ''}")
            if assemblies>0: bits.append(f"{assemblies} assembl{'ies' if assemblies!=1 else 'y'}")
            eq_remaining_var.set("\u23f3  " + " and ".join(bits) + " remaining")
            eq_remaining_lbl.grid()
        except Exception: pass

    tpcard=ctk.CTkFrame(ctx,fg_color=C["card"],corner_radius=16,border_width=1,border_color=C["border_light"])
    tpcard.grid(row=3,column=0,sticky="ew",pady=(12,0)); tpcard.grid_columnconfigure(0,weight=1)
    timer_var=ctk.StringVar(value="00:00")
    timer_lbl=ctk.CTkLabel(tpcard,textvariable=timer_var,font=ctk.CTkFont(size=36,weight="bold"),text_color=C["text"])
    timer_lbl.grid(row=0,column=0,sticky="w",padx=16,pady=(14,0))
    timer_status=ctk.CTkLabel(tpcard,text="",font=ctk.CTkFont(size=11,weight="bold"),text_color=C["text_dim"],anchor="w")
    timer_status.grid(row=1,column=0,sticky="w",padx=16,pady=(0,8))
    ctk.CTkLabel(tpcard,text="PROGRESS",font=ctk.CTkFont(size=10,weight="bold"),text_color=C["text_muted"],anchor="w").grid(row=2,column=0,sticky="w",padx=16,pady=(2,2))
    import tkinter as _tk
    prog_canvas=_tk.Canvas(tpcard,height=22,bg=C["card"],highlightthickness=0,bd=0)
    prog_canvas.grid(row=3,column=0,sticky="ew",padx=16,pady=(0,16))
    _prog_pct=[0]
    def _draw_progress(pct):
        _prog_pct[0]=pct
        prog_canvas.delete("all")
        w=prog_canvas.winfo_width() or 240
        h=22; r=11
        prog_canvas.create_rounded_rect=lambda x1,y1,x2,y2,radius,**kw: prog_canvas.create_polygon(
            x1+radius,y1, x2-radius,y1, x2,y1, x2,y1+radius,
            x2,y2-radius, x2,y2, x2-radius,y2, x1+radius,y2,
            x1,y2, x1,y2-radius, x1,y1+radius, x1,y1, smooth=True,**kw)
        prog_canvas.create_rounded_rect(0,0,w,h,r,fill=C["border"],outline="")
        fill_w=int(w*(pct/100))
        if fill_w>r*2: prog_canvas.create_rounded_rect(0,0,fill_w,h,r,fill=C["accent"],outline="")
        elif fill_w>0: prog_canvas.create_rectangle(0,0,fill_w,h,fill=C["accent"],outline="")
        txt=f"{int(pct)}%"
        prog_canvas.create_text(w//2,h//2,text=txt,fill="#ffffff" if pct>15 else C["text_dim"],font=("Calibri",10,"bold"))
    prog_canvas.bind("<Configure>",lambda e:_draw_progress(_prog_pct[0]))
    _draw_progress(0)
    def set_progress(pct, label="", animate=True):
        try: pct=max(0.0,min(100.0,float(pct)))
        except Exception: return
        start=float(_prog_pct[0]); diff=pct-start
        # Snap for resets/decreases/tiny steps, or when animation is suppressed.
        # Otherwise slide smoothly so phase jumps (e.g. 15->30, 60->65) don't pop.
        # Step count scales with the jump size, so frequent small loop updates stay
        # instant and only big transitions animate — total animation time is bounded
        # by the percentage covered, not by the number of calls. update_idletasks()
        # repaints without processing input events, so there's no re-entrancy risk.
        if (not animate) or diff<=1.0:
            try: _draw_progress(pct); prog_canvas.update_idletasks()
            except Exception: pass
            try: _pulse(pct)
            except Exception: pass
            return
        steps=max(1,min(24,int(diff/1.2)))
        for i in range(1,steps+1):
            try: _draw_progress(start+diff*(i/steps)); prog_canvas.update_idletasks()
            except Exception: pass
            try: _time.sleep(0.010)
            except Exception: pass
        try: _draw_progress(pct); _pulse(pct)
        except Exception: pass

    def timer_tick():
        if timer_running[0]:
            elapsed=timer_elapsed[0]+(_time.time()-timer_start[0])
            mins=int(elapsed//60); secs=int(elapsed%60)
            timer_var.set(f"{mins:02d}:{secs:02d}")
            timer_job[0]=root.after(500,timer_tick)
    def timer_start_fn(label="Running..."):
        timer_start[0]=_time.time(); timer_running[0]=True
        timer_status.configure(text=label); timer_tick()
    def timer_pause(label="Waiting..."):
        if timer_running[0]:
            timer_elapsed[0]+=_time.time()-timer_start[0]
            timer_running[0]=False; timer_status.configure(text=label)
            if timer_job[0]: root.after_cancel(timer_job[0])
    def timer_resume(label="Running..."):
        if timer_running[0]:
            timer_status.configure(text=label); return
        timer_start[0]=_time.time(); timer_running[0]=True
        timer_status.configure(text=label); timer_tick()
    def timer_stop():
        if timer_running[0]:
            timer_elapsed[0]+=_time.time()-timer_start[0]
            timer_running[0]=False
        elapsed=timer_elapsed[0]; mins=int(elapsed//60); secs=int(elapsed%60)
        timer_var.set(f"{mins:02d}:{secs:02d}"); timer_status.configure(text="Complete")
        if timer_job[0]: root.after_cancel(timer_job[0])
        return f"{mins:02d}:{secs:02d}"
    def timer_reset():
        timer_running[0]=False; timer_elapsed[0]=0.0; timer_start[0]=0.0
        timer_var.set("00:00"); timer_status.configure(text="")
        if timer_job[0]: root.after_cancel(timer_job[0])
        set_progress(0,"")

    def _flash_busy():
        # Non-destructive hint shown when a tab is clicked mid-run (tabs are locked
        # while processing). Briefly overrides the timer subtitle, then restores it.
        try:
            prev=timer_status.cget("text")
            timer_status.configure(text="Please wait — still processing this step")
            root.after(1300,lambda:timer_status.configure(text=prev))
        except Exception: pass


    # ===== FORM (col 1) =====
    formwrap=ctk.CTkFrame(input_panel,fg_color=C["bg"],corner_radius=0)
    formwrap.grid(row=0,column=0,sticky="nsew",padx=14,pady=(12,0))
    formwrap.grid_columnconfigure(0,weight=1); formwrap.grid_rowconfigure(0,weight=1)
    form=ctk.CTkScrollableFrame(formwrap,fg_color=C["bg"],corner_radius=0)
    form.grid(row=0,column=0,sticky="nsew")
    form.grid_columnconfigure(0,weight=1,uniform="cols"); form.grid_columnconfigure(1,weight=1,uniform="cols")

    # ── CASE DIMENSIONS — container card on the LEFT (single column, narrower boxes) ──
    dims_card=ctk.CTkFrame(form,fg_color=C["card"],corner_radius=16,border_width=1,border_color=C["border_light"])
    dims_card.grid(row=0,column=0,sticky="new",padx=(6,8),pady=(2,10))
    dims_card.grid_columnconfigure(0,weight=1)
    _section(dims_card,"CASE DIMENSIONS",0,colspan=1,padx=18,pady=(16,8))

    check_vars={}; entry_vars={}; entry_widgets={}; row_frames={}; check_widgets={}; radio_widgets=[]
    def toggle(name):
        on=check_vars[name].get()
        entry_widgets[name].configure(
            state="normal" if on else "disabled",
            text_color=C["text"] if on else C["text_muted"],
            border_color=C["accent"] if on else C["border"])
        row_frames[name].configure(
            fg_color=C["accent_light"] if on else C["card2"],border_width=2 if on else 1,
            border_color=C["accent"] if on else C["border"])
        if not on: entry_vars[name].set("")

    for idx,(name,label) in enumerate(PARAMS):
        cv=ctk.BooleanVar(); ev=ctk.StringVar(); check_vars[name]=cv; entry_vars[name]=ev
        # Single-line row: [checkbox + name] .......... [value entry][in]
        card=ctk.CTkFrame(dims_card,fg_color=C["card2"],corner_radius=10,border_width=1,border_color=C["border"])
        card.grid(row=1+idx,column=0,sticky="ew",padx=16,pady=6); card.grid_columnconfigure(1,weight=1); row_frames[name]=card
        _chk=ctk.CTkCheckBox(card,variable=cv,text="",width=24,checkbox_width=24,checkbox_height=24,
                        fg_color=C["accent"],hover_color=C["accent_hover"],
                        checkmark_color="#ffffff",border_color=C["accent_mid"],border_width=2,
                        command=lambda n=name:toggle(n))
        _chk.grid(row=0,column=0,padx=(14,8),pady=14,sticky="w"); check_widgets[name]=_chk
        ctk.CTkLabel(card,text=label,font=ctk.CTkFont(size=13,weight="bold"),
                     text_color=C["text"],anchor="w").grid(row=0,column=1,sticky="w",pady=14)
        ent=ctk.CTkEntry(card,textvariable=ev,width=120,height=34,font=ctk.CTkFont(size=14,weight="bold"),
                         fg_color=C["card"],text_color=C["text_muted"],border_color=C["border"],border_width=1,
                         corner_radius=10,state="disabled",justify="right",placeholder_text="0.00")
        ent.grid(row=0,column=2,sticky="e",padx=(8,6),pady=10); entry_widgets[name]=ent
        ctk.CTkLabel(card,text="in",font=ctk.CTkFont(size=11,weight="bold"),
                     text_color=C["text_muted"]).grid(row=0,column=3,sticky="e",padx=(0,14))
    ctk.CTkFrame(dims_card,fg_color="transparent",height=6).grid(row=1+len(PARAMS),column=0)

    # ── Config selector variables + segmented builder (defined BEFORE Materials
    # so Materials can use the same segmented style) ──
    bf_lp_var =ctk.StringVar(value="nolp")
    bf_ref_var=ctk.StringVar(value="DX")
    cfg_refreshers=[]

    def _build_cfg_options(parent,var,options,start_row=1,cols=1,pad=(18,18),col=0,colspan=None,pady=5):
        """Segmented control bound to `var`. Registers its buttons in
        radio_widgets (so they lock with the rest of the inputs) and its refresh
        in cfg_refreshers (so reset() re-syncs the visuals). Returns the refresh
        function so callers (e.g. Materials) can register it elsewhere too."""
        wrap=ctk.CTkFrame(parent,fg_color=C["card2"],corner_radius=10,
                          border_width=1,border_color=C["border"])
        span=colspan if colspan else (2 if cols==2 else 1)
        wrap.grid(row=start_row,column=col,columnspan=span,sticky="ew",
                  padx=(pad[0],pad[1]),pady=pady)
        btns={}
        def _refresh():
            sel=var.get()
            for val,b in btns.items():
                if val==sel: b.configure(fg_color=C["accent"],text_color="#ffffff")
                else:        b.configure(fg_color="transparent",text_color=C["text_dim"])
        for lbl,val in options:
            b=ctk.CTkButton(wrap,text=lbl,font=ctk.CTkFont(size=12,weight="bold"),
                            height=32,corner_radius=8,fg_color="transparent",
                            text_color=C["text_dim"],hover=False,cursor="hand2",
                            command=lambda v=val:(var.set(v),_refresh()))
            b.pack(side="left",expand=True,fill="x",padx=3,pady=3)
            radio_widgets.append(b); btns[val]=b
        _refresh(); cfg_refreshers.append(_refresh)
        return _refresh

    def _refresh_cfg_selectors():
        for _r in cfg_refreshers:
            try: _r()
            except Exception: pass

    # ── MATERIALS — full-width card on row 1: each group is a label + segmented
    # pill row, two groups side by side (mockup arrangement) ──
    mats_card=ctk.CTkFrame(form,fg_color=C["card"],corner_radius=16,border_width=1,border_color=C["border_light"])
    mats_card.grid(row=1,column=0,columnspan=2,sticky="new",padx=6,pady=(2,10))
    mats_card.grid_columnconfigure(0,weight=1,uniform="mcol"); mats_card.grid_columnconfigure(1,weight=1,uniform="mcol")
    _section(mats_card,"MATERIALS",0,colspan=2,padx=18,pady=(16,8))
    material_radio_vars={}; material_refreshers={}
    _groups=[(ab,opts) for ab,opts in MATERIAL_CHOICES.items() if opts]
    for gi,(mat_abbrev,options) in enumerate(_groups):
        rv=ctk.StringVar(value=options[0][1]); material_radio_vars[mat_abbrev]=rv
        _r=1+(gi//2)*2; _c=gi%2
        _lp,_rp=((18,9) if _c==0 else (9,18))
        ctk.CTkLabel(mats_card,text=mat_abbrev,font=ctk.CTkFont(size=12,weight="bold"),
                     text_color=C["text"],anchor="w").grid(row=_r,column=_c,sticky="w",padx=(_lp,_rp),pady=(10,2))
        material_refreshers[mat_abbrev]=_build_cfg_options(
            mats_card,rv,[(lbl,val) for lbl,val in options],
            start_row=_r+1,col=_c,colspan=1,pad=(_lp,_rp),pady=(0,6))
    ctk.CTkFrame(mats_card,fg_color="transparent",height=10).grid(
        row=1+((len(_groups)+1)//2)*2,column=0,columnspan=2)

    # ── BASE FRAME CONSTRUCTION + REFRIGERATION — one card, top-right (row 0,
    # col 1), beside Case Dimensions. Both drive which base-frame / frame-kit
    # configuration is activated at the end of the run. ──
    bfc_card=ctk.CTkFrame(form,fg_color=C["card"],corner_radius=16,border_width=1,border_color=C["border_light"])
    bfc_card.grid(row=0,column=1,sticky="new",padx=(8,6),pady=(2,10))
    bfc_card.grid_columnconfigure(0,weight=1)
    _section(bfc_card,"BASE FRAME CONSTRUCTION",0,colspan=1,padx=18,pady=(16,8))
    _build_cfg_options(bfc_card,bf_lp_var,
                       [("Without Leveling Pad","nolp"),("With Leveling Pad","lp")],start_row=1,cols=1)
    _section(bfc_card,"REFRIGERATION",2,colspan=1,padx=18,pady=(16,8))
    _build_cfg_options(bfc_card,bf_ref_var,
                       [("DX/CO2","DX"),("Glycol","GLYCOL")],start_row=3,cols=1)
    ctk.CTkLabel(bfc_card,text="Applies to the base frame and frame kit configurations",
                 font=ctk.CTkFont(size=10),text_color=C["text_muted"],anchor="w",
                 justify="left",wraplength=340).grid(row=4,column=0,sticky="w",padx=18,pady=(2,14))

    # current/new value strips — in the persistent left panel (2-col compact grid).
    # The phase-2 handlers still grid these at rows 5/6 of `vbars`.
    vbars=ctk.CTkFrame(ctx,fg_color="transparent"); vbars.grid(row=4,column=0,sticky="ew",pady=(12,0))
    vbars.grid_columnconfigure(0,weight=1); vbars.grid_columnconfigure(1,weight=1)
    orig_bar=ctk.CTkFrame(vbars,fg_color=C["card2"],corner_radius=10,border_width=1,border_color=C["border_light"])
    orig_bar.grid(row=5,column=0,sticky="new",padx=(0,4))
    orig_bar.grid_columnconfigure(0,weight=1)
    ctk.CTkLabel(orig_bar,text="ORIGINAL",font=ctk.CTkFont(size=10,weight="bold"),
                 text_color=C["text_muted"]).grid(row=0,column=0,sticky="w",padx=10,pady=(8,2))
    orig_bar_labels={}
    for i,(name,_) in enumerate(PARAMS):
        val=format_value(current_vals.get(name,"\u2014 ")); display=f"{val} in" if val!="\u2014 " else "\u2014 "
        _olbl=ctk.CTkLabel(orig_bar,text=f"{PARAM_SHORT[name]}\n{display}",font=ctk.CTkFont(size=11,weight="bold"),
                     text_color=C["accent"],fg_color=C["accent_light"],corner_radius=6,padx=6,pady=3)
        _olbl.grid(row=1+i,column=0,padx=6,pady=3,sticky="ew"); orig_bar_labels[name]=_olbl
    new_bar=ctk.CTkFrame(vbars,fg_color=C["success_bg"],corner_radius=10,border_width=1,border_color=C["success_mid"])
    new_bar.grid_columnconfigure(0,weight=1)
    new_bar.grid_remove()
    ctk.CTkLabel(new_bar,text="NEW",font=ctk.CTkFont(size=10,weight="bold"),
                 text_color=C["success"]).grid(row=0,column=0,sticky="w",padx=10,pady=(8,2))
    new_bar_labels={}
    for i,(name,_) in enumerate(PARAMS):
        lbl=ctk.CTkLabel(new_bar,text=f"{PARAM_SHORT[name]}\n\u2014 in",font=ctk.CTkFont(size=11,weight="bold"),
                         text_color=C["success"],fg_color=C["success_mid"],corner_radius=6,padx=6,pady=3)
        lbl.grid(row=1+i,column=0,padx=6,pady=3,sticky="ew"); new_bar_labels[name]=lbl

    def _paint_new_bar(vals):
        """Fill the NEW strip and visually flag any value that differs from
        CURRENT, so it's obvious at a glance which dimension(s) changed. A
        changed value gets a solid chip + white text + leading dot; unchanged
        values stay in the subdued style."""
        for nm, lbl in new_bar_labels.items():
            raw = vals.get(nm, "\u2014 ")
            nv = format_value(raw) if raw not in ("\u2014 ", "") else "\u2014 "
            ov = format_value(current_vals.get(nm, "\u2014 "))
            changed = (nv not in ("\u2014 ", "")) and (nv != ov)
            disp = f"{nv} in" if nv != "\u2014 " else "\u2014 "
            if changed:
                lbl.configure(text=f"\u25cf {PARAM_SHORT[nm]}\n{disp}",
                              text_color="#ffffff", fg_color=C["success"])
            else:
                lbl.configure(text=f"{PARAM_SHORT[nm]}\n{disp}",
                              text_color=C["success"], fg_color=C["success_mid"])

    inputs_locked=[False]
    def set_inputs_locked(locked):
        inputs_locked[0]=locked
        st="disabled" if locked else "normal"
        try: job_entry.configure(state=st)
        except Exception: pass
        for n in check_widgets:
            try: check_widgets[n].configure(state=st)
            except Exception: pass
        for rb in radio_widgets:
            try: rb.configure(state=st)
            except Exception: pass
        for n in entry_widgets:
            try:
                if locked:
                    entry_widgets[n].configure(state="disabled")
                else:
                    on=check_vars[n].get()
                    entry_widgets[n].configure(state="normal" if on else "disabled")
            except Exception: pass

    # ── Button bar (Reset left, Run Update right) — pinned to the bottom of the page ──
    bf=ctk.CTkFrame(input_panel,fg_color=C["bg"],corner_radius=0,height=64)
    bf.grid(row=1,column=0,sticky="ew",padx=14,pady=(6,12)); bf.grid_propagate(False)
    bf.grid_columnconfigure(0,weight=1)

    def collect():
        names,values=[],[]
        fresh=get_current_equation_values(sw_model) if sw_model else {}
        for n,lbl in PARAMS:
            if not check_vars[n].get(): continue
            raw=entry_vars[n].get().strip()
            if not raw: messagebox.showwarning("Missing Value",f"Enter a value for {lbl}."); return None,None
            try: new_val=float(raw)
            except ValueError: messagebox.showerror("Invalid Value",f"'{raw}' is not valid."); return None,None
            try:
                cur=float(fresh.get(n,""))
                if abs(new_val-cur)<0.0001:
                    if not messagebox.askyesno("Same Value",f"\u26a0 {lbl} is already {format_value(cur)} in.\nProceed?"): return None,None
            except (ValueError,TypeError): pass
            values.append(new_val); names.append(n)
        if not names:
            if not messagebox.askyesno(
                "No dimension parameters selected",
                "No dimension parameters are checked.\n\n"
                "The scan will only look for material changes. Continue?"):
                return None, None
        return names,values

    # ── Swappable pages (all live in the content area, shown one at a time) ──
    log_panel     =ctk.CTkFrame(content,fg_color=C["log_bg"],corner_radius=0)
    results_panel =ctk.CTkFrame(content,fg_color=C["bg"],corner_radius=0)
    summary_panel =ctk.CTkFrame(content,fg_color=C["bg"],corner_radius=0)
    for p in [log_panel,results_panel,summary_panel]:
        p.grid_columnconfigure(0,weight=1); p.grid_rowconfigure(1,weight=1)

    cur_panel=[None]; log_return=[None]
    def show_panel(panel):
        # Remember the page we're leaving when opening the log, so the log's
        # Back button returns there instead of guessing.
        if panel is log_panel and cur_panel[0] not in (None, log_panel):
            log_return[0]=cur_panel[0]
        for p in [input_panel,log_panel,results_panel,summary_panel]: p.grid_remove()
        panel.grid(row=0,column=0,sticky="nsew")
        cur_panel[0]=panel
        # Reflect the active page in the sidebar.
        if   panel is input_panel:   cur_panel_key[0]="input"
        elif panel is log_panel:     cur_panel_key[0]="log"
        elif panel is summary_panel: cur_panel_key[0]="summary"
        elif panel is results_panel: cur_panel_key[0]="newnum" if last_results[0]=="new_numbers" else "replace"
        try: refresh_nav()
        except Exception: pass

    def copy_cb(text): root.clipboard_clear(); root.clipboard_append(text); root.update()

    # Log panel
    lhdr=ctk.CTkFrame(log_panel,fg_color=C["log_hdr"],corner_radius=0,height=44)
    lhdr.grid(row=0,column=0,sticky="ew"); lhdr.grid_propagate(False)
    ctk.CTkButton(lhdr,text="← Back",font=ctk.CTkFont(size=11,weight="bold"),fg_color="transparent",hover_color="#26303f",
                  text_color="#7fb0e8",width=64,height=28,corner_radius=10,
                  command=lambda:show_panel(log_return[0] or input_panel)).place(x=4,rely=0.5,anchor="w")
    ctk.CTkLabel(lhdr,text="OUTPUT LOG",font=ctk.CTkFont(size=11,weight="bold"),text_color="#7fb0e8").place(x=80,rely=0.5,anchor="w")
    view_results_btn=ctk.CTkButton(lhdr,text="View Results →",font=ctk.CTkFont(size=11),
                                   fg_color=C["accent"],hover_color=C["accent_hover"],
                                   text_color="#ffffff",width=110,height=28,corner_radius=10,command=lambda:None)
    ctk.CTkButton(lhdr,text="Clear",font=ctk.CTkFont(size=11),fg_color="transparent",hover_color="#26303f",
                  text_color="#8590a3",width=52,height=26,corner_radius=10,
                  command=lambda:(output.configure(state="normal"),output.delete("1.0","end"),output.configure(state="disabled"))
                  ).place(relx=1.0,x=-10,rely=0.5,anchor="e")
    ctk.CTkFrame(log_panel,fg_color=C["accent_mid"],height=2,corner_radius=0).grid(row=0,column=0,sticky="sew")
    output=ctk.CTkTextbox(log_panel,font=ctk.CTkFont(size=11),fg_color=C["log_bg"],text_color=C["log_text"],
                          scrollbar_button_color="#2b3444",scrollbar_button_hover_color=C["accent_mid"],
                          corner_radius=0,border_width=0,wrap="none",state="disabled")
    output.grid(row=1,column=0,sticky="nsew")

    def log(msg):
        try: output.configure(state="normal"); output.insert("end",msg); output.see("end"); output.configure(state="disabled"); root.update()
        except Exception: pass

    def _show_results():
        try:
            if last_results[0]=="new_numbers": build_new_numbers_panel()
            else: build_replace_panel()
            show_panel(results_panel)
        except Exception as e:
            import traceback
            log(f"\n View Results error: {e}\n{traceback.format_exc()}\n")
            messagebox.showerror("Error",f"View Results failed:\n{e}")
    view_results_btn.configure(command=_show_results)

    # ============ STEPPER NAV LOGIC (steps were built in the top area) ============
    def refresh_nav():
        for key,num,label in STEPS:
            sw=step_widgets[key]; active=(cur_panel_key[0]==key); enabled=nav_enabled[key]; done=nav_done[key]
            if done:
                # a completed step always shows its green check, regardless of the active tab
                sw["circ"].configure(fg_color=C["success"],text_color="#ffffff",text="\u2713")
                sw["txt"].configure(text_color=C["accent"] if active else C["text"])
            elif active:
                sw["circ"].configure(fg_color=C["accent"],text_color="#ffffff",text=num)
                sw["txt"].configure(text_color=C["accent"])
            elif enabled:
                sw["circ"].configure(fg_color=C["accent_light"],text_color=C["accent"],text=num)
                sw["txt"].configure(text_color=C["text"])
            else:  # locked — dimmed number (a lock emoji renders as a box on many Windows fonts)
                sw["circ"].configure(fg_color=C["card2"],text_color=C["text_muted"],text=num)
                sw["txt"].configure(text_color=C["text_muted"])

    def nav_go(key):
        # While a run is in progress, ignore tab clicks. Navigating would rebuild
        # the panel (destroying the buttons/labels the running job still holds),
        # which silently kills the workflow. Tabs become active again when idle.
        if _busy[0]:
            try: _flash_busy()
            except Exception: pass
            return
        if not nav_enabled.get(key): return
        if   key=="input":   show_panel(input_panel)
        elif key=="log":     show_panel(log_panel)
        elif key=="replace": build_replace_panel(); show_panel(results_panel)
        elif key=="newnum":  build_new_numbers_panel(); show_panel(results_panel)
        elif key=="summary": build_summary_panel(); show_panel(summary_panel)

    def unlock_nav(key):
        nav_enabled[key]=True
        try: refresh_nav()
        except Exception: pass

    refresh_nav()

    # ── Stepper connector animation + button throbbers ─────────────────
    # The green "route" line between two steps fills in lockstep with the
    # real progress of the phase that bridges them (scan / equations / macro),
    # like a train advancing along a track. Button throbbers are advanced from
    # the same progress callback so they move when actual work moves.
    _GREEN=C["success"]; _TRACK=C["border"]
    def _draw_conn(i):
        try:
            c=connectors[i]; cv=c["canvas"]; cv.delete("all")
            w=_CONN_W; y=5; x0=3; x1=w-3
            cv.create_line(x0,y,x1,y,fill=_TRACK,width=3,capstyle="round")
            frac=1.0 if c["state"]=="done" else max(0.0,min(1.0,c["frac"]))
            if frac>0:
                xf=x0+(x1-x0)*frac
                cv.create_line(x0,y,xf,y,fill=_GREEN,width=3,capstyle="round")
                if c["state"]=="active" and frac<0.999:
                    cv.create_oval(xf-3,y-3,xf+3,y+3,fill=_GREEN,outline="")
        except Exception: pass
    def _set_conn(i,state,frac=None):
        if i<0 or i>=len(connectors): return
        c=connectors[i]; c["state"]=state
        if state=="done": c["frac"]=1.0
        elif state=="idle": c["frac"]=0.0
        elif frac is not None: c["frac"]=frac
        try: c["cap"].configure(text_color=(C["success"] if state in ("active","done") else C["text_muted"]))
        except Exception: pass
        _draw_conn(i)
    active_conn=[None]            # (idx, lo_pct, hi_pct)
    def _begin_conn(idx,lo,hi):
        active_conn[0]=(idx,lo,hi); _set_conn(idx,"active",0.0)
    def _finish_conn(idx):
        active_conn[0]=None; _set_conn(idx,"done")
    def _reset_conns():
        active_conn[0]=None
        for i in range(len(connectors)): _set_conn(i,"idle")

    SPIN="\u280b\u2819\u2839\u2838\u283c\u2834\u2826\u2827"   # braille throbber
    _spin_i=[0]; active_btn=[None]   # active_btn = (button, base_label)
    def _set_btn_busy(btn,label):
        active_btn[0]=(btn,label); _spin_i[0]=0
        try: btn.configure(text=f"{SPIN[0]}  {label}")
        except Exception: pass
    def _clear_btn_busy():
        active_btn[0]=None
    def _pulse(pct):
        ac=active_conn[0]
        if ac is not None:
            idx,lo,hi=ac
            f=(pct-lo)/(hi-lo) if hi>lo else 1.0
            connectors[idx]["frac"]=max(0.0,min(1.0,f)); _draw_conn(idx)
        ab=active_btn[0]
        if ab is not None:
            _spin_i[0]=(_spin_i[0]+1)%len(SPIN)
            try: ab[0].configure(text=f"{SPIN[_spin_i[0]]}  {ab[1]}")
            except Exception: pass
    for _ci in range(len(connectors)): _draw_conn(_ci)


    def panel_header(panel,title,btns):
        # Light 2026 header: white bar, accent tick + dark title, thin bottom
        # border. Buttons: colored = primary (filled), "transparent" = ghost.
        h=ctk.CTkFrame(panel,fg_color=C["card"],corner_radius=0,height=52)
        h.grid(row=0,column=0,sticky="ew"); h.grid_propagate(False)
        ctk.CTkFrame(h,fg_color=C["border_light"],height=1,corner_radius=0).place(relx=0,rely=1.0,anchor="sw",relwidth=1.0)
        ctk.CTkFrame(h,fg_color=C["accent"],width=4,height=20,corner_radius=2).place(x=16,rely=0.5,anchor="w")
        ctk.CTkLabel(h,text=title,font=ctk.CTkFont(size=14,weight="bold"),text_color=C["text"]).place(x=28,rely=0.5,anchor="w")
        bf2=ctk.CTkFrame(h,fg_color="transparent"); bf2.place(relx=1.0,x=-12,rely=0.5,anchor="e")
        for txt,clr,cmd in btns:
            is_ghost=clr=="transparent"
            _hov={C["accent"]:C["accent_hover"],C["success"]:"#116936",
                  C["accent_mid"]:C["accent"],C["danger"]:"#b91c1c"}.get(clr,C["accent_hover"])
            ctk.CTkButton(bf2,text=txt,font=ctk.CTkFont(size=11,weight="bold"),
                          fg_color=C["card"] if is_ghost else clr,
                          hover_color=C["accent_light"] if is_ghost else _hov,
                          text_color=C["text_dim"] if is_ghost else "#ffffff",
                          width=104,height=30,corner_radius=10,
                          border_width=1 if is_ghost else 0,
                          border_color=C["border"],
                          command=cmd).pack(side="left",padx=(0,6))

    def make_col_hdr(body,ri,cols,hdr_bg=None):
        ch=ctk.CTkFrame(body,fg_color=C["accent_light"],corner_radius=6,height=28)
        ch.grid(row=ri[0],column=0,sticky="ew",padx=10)
        # Columns: Old | Description (grow) | Action | New/Existing | Status
        weights=[0,1,0,0,0]; minsizes=[145,0,110,130,110]
        # If caller passes only 4 columns (legacy), drop the Action column
        if len(cols) == 4:
            weights=[0,1,0,0]; minsizes=[145,0,130,110]
        for i,(w,ms) in enumerate(zip(weights,minsizes)): ch.grid_columnconfigure(i,weight=w,minsize=ms)
        for col,txt in enumerate(cols):
            ctk.CTkLabel(ch,text=txt,font=ctk.CTkFont(size=10,weight="bold"),text_color=C["accent"],anchor="center",justify="center").grid(row=0,column=col,sticky="ew",padx=6,pady=4)
        ri[0]+=1

    def _entry_action_label(entry):
        """Return a short label describing what will be applied to this entry.
        When material is being applied, include the target material name so
        the user can see exactly what material will be set. For parts that are
        being direct-replaced from the BOM (no edits applied to the part file)
        the label is '—' since nothing is updated on the original part."""
        # Direct-replacement / existing-BOM-part statuses mean we didn't modify
        # the original part file — we simply swapped it for the BOM-listed one.
        # Nothing was "updated" on the part itself.
        status = (entry.get("status") or "").lower()
        status_disp = (entry.get("status_display") or "").lower()
        if status in ("existing", "replaced", "file_missing") or \
           status_disp in ("replaced", "file missing"):
            return "—"
        has_eq  = bool(entry.get("app_names"))
        applied   = bool(entry.get("apply_material"))   # material applied to THIS part
        inherited = entry.get("_mat_changed_name")       # assembly inherits a CHANGED material
        has_mat   = applied or bool(inherited)
        mat_name  = (entry.get("target_material") if applied else inherited) or ""
        # An ASSEMBLY can span several gauges, so its row names the finish change
        # generically ("Equation + Material" / "Material") without a code; a PART
        # is a single gauge, so it keeps the specific code.
        is_assy = (entry.get("type") == "assembly")
        if has_eq and has_mat:
            if is_assy or not mat_name:
                return "Equation + Material"
            return f"Equation + Material ({mat_name})"
        if has_mat:
            if is_assy or not mat_name:
                return "Material"
            return f"Material ({mat_name})"
        if has_eq:
            return "Equation"
        return "—"

    def make_entry_row(body,ri,bg,entry,labels_dict,show_new_number=False,show_action=True):
        fr=ctk.CTkFrame(body,fg_color=bg,corner_radius=0)
        fr.grid(row=ri[0],column=0,sticky="ew",padx=10)
        if show_action:
            weights=[0,1,0,0,0]; minsizes=[145,0,110,130,110]
        else:
            weights=[0,1,0,0];   minsizes=[145,0,130,110]
        for i,(w,ms) in enumerate(zip(weights,minsizes)): fr.grid_columnconfigure(i,weight=w,minsize=ms)
        ri[0]+=1
        ctk.CTkFrame(body,fg_color=C["border"],height=1,corner_radius=0).grid(row=ri[0],column=0,sticky="ew",padx=10); ri[0]+=1
        of=ctk.CTkFrame(fr,fg_color="transparent"); of.grid(row=0,column=0,pady=8,padx=6)
        ctk.CTkLabel(of,text=entry["old_number"],font=ctk.CTkFont(size=11,weight="bold"),text_color=C["text"]).pack(side="left")
        ctk.CTkButton(of,text="\u29c9",font=ctk.CTkFont(size=11),fg_color="transparent",hover_color=C["border"],
                      text_color=C["accent"],width=24,height=24,corner_radius=6,
                      command=lambda n=entry["old_number"]:(copy_cb(n),messagebox.showinfo("Copied",f"Copied: {n}"))
                      ).pack(side="left",padx=(4,0))
        ctk.CTkLabel(fr,text=entry["desc"],font=ctk.CTkFont(size=11),text_color=C["text_dim"],anchor="center",justify="center").grid(row=0,column=1,sticky="ew",padx=10,pady=8)
        # Action column (col 2) — when shown
        if show_action:
            act_lbl_text = _entry_action_label(entry)
            act_color = C["text_muted"] if act_lbl_text == "—" else C["text"]
            ctk.CTkLabel(fr, text=act_lbl_text,
                         font=ctk.CTkFont(size=10, weight="bold"),
                         text_color=act_color, anchor="center", justify="center"
                         ).grid(row=0, column=2, sticky="ew", padx=6, pady=8)
            new_col_idx = 3; status_col_idx = 4
        else:
            new_col_idx = 2; status_col_idx = 3
        num_to_show=entry.get("existing_number","— ") if not show_new_number else entry.get("new_number","Pending...")
        num_color=C["accent"] if entry.get("existing_number") and not show_new_number else C["text_muted"]
        nf=ctk.CTkFrame(fr,fg_color="transparent"); nf.grid(row=0,column=new_col_idx,pady=8,padx=6)
        num_lbl=ctk.CTkLabel(nf,text=num_to_show,font=ctk.CTkFont(size=11,weight="bold"),text_color=num_color); num_lbl.pack(side="left")
        if num_to_show and num_to_show not in ("— ","Pending...","No change","Drawing Missing","Failed"):
            ctk.CTkButton(nf,text="\u29c9",font=ctk.CTkFont(size=10),fg_color="transparent",hover_color=C["border"],
                          text_color=num_color,width=22,height=22,corner_radius=6,
                          command=lambda n=num_to_show:(copy_cb(n),messagebox.showinfo("Copied",f"Copied: {n}"))
                          ).pack(side="left",padx=(3,0))
        labels_dict[entry["old_path"]]=num_lbl
        raw_st=entry.get("status","")
        if   raw_st=="existing":       st_txt,st_c,st_bg="Part Found",   C["success"],C["success_bg"]
        elif raw_st=="file_missing":   st_txt,st_c,st_bg="File Missing",  C["warning"],C["warning_bg"]
        elif raw_st=="Replaced":       st_txt,st_c,st_bg="Replaced",      C["success"],C["success_bg"]
        elif raw_st=="New Number":     st_txt,st_c,st_bg="New Number",    C["success"],C["success_bg"]
        elif raw_st=="Drawing Missing":st_txt,st_c,st_bg="Drawing Missing",C["warning"],C["warning_bg"]
        elif raw_st=="Failed":         st_txt,st_c,st_bg="Failed",        C["danger"], C["danger_bg"]
        else:                          st_txt,st_c,st_bg="Pending...",    C["text_muted"],C["card2"]
        st_lbl=ctk.CTkLabel(fr,text=st_txt,font=ctk.CTkFont(size=10,weight="bold"),fg_color=st_bg,text_color=st_c,corner_radius=6,padx=6,anchor="center",justify="center")
        st_lbl.grid(row=0,column=status_col_idx,sticky="ew",padx=6,pady=8)
        entry["_status_lbl"]=st_lbl

    def _confirm_action(action):
        """Yes/No gate for a destructive step. Returns True only if the user
        clicks Yes. Cancel (No / closing the dialog) returns False and the caller
        does nothing — the interface is left exactly as it was."""
        try:
            return bool(messagebox.askyesno("Please Confirm",
                        f"Are you sure you want to {action}?", icon="warning"))
        except Exception:
            return True

    # ─────────────────────────────────────────
    # Results panel — Phase 1 (replace existing)
    # ─────────────────────────────────────────
    def build_replace_panel():
        nav_done["input"]=True; unlock_nav("replace")
        for w in results_panel.winfo_children(): w.destroy()
        results_panel.grid_rowconfigure(1,weight=1)
        existing_entries=[e for e in all_entries if e["status"] in ("existing","file_missing","Replaced")]
        result_btns=[]
        if gen_done[0]:
            result_btns.append(("View Summary",C["success"],lambda:(build_summary_panel(),show_panel(summary_panel))))
        panel_header(results_panel,f"PHASE 1 — EXISTING PARTS — Direct Replace ({len(existing_entries)})",result_btns)
        body=ctk.CTkScrollableFrame(results_panel,fg_color=C["bg"],corner_radius=0)
        body.grid(row=1,column=0,sticky="nsew"); body.grid_columnconfigure(0,weight=1)
        ri=[0]; exist_num_labels={}
        folder=os.path.dirname(main_path) if main_path else ""

        def sec_hdr(title,count,bg=C["card2"]):
            tint=C["success_bg"] if bg==C["new_hdr"] else (C["exist_bg"] if bg==C["exist_hdr"] else C["card2"])
            acc=C["success"] if bg==C["new_hdr"] else C["accent"]
            sh=ctk.CTkFrame(body,fg_color=tint,corner_radius=10)
            sh.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=(10,2)); sh.grid_columnconfigure(1,weight=1)
            _bb=ctk.CTkFrame(sh,fg_color=acc,width=4,height=16,corner_radius=2); _bb.grid(row=0,column=0,sticky="ns",padx=(12,8),pady=8); _bb.grid_propagate(False)
            ctk.CTkLabel(sh,text=f"{title} ({count})",font=ctk.CTkFont(size=12,weight="bold"),text_color=acc).grid(row=0,column=1,sticky="w",pady=6); ri[0]+=1

        def sub_lbl(txt):
            ctk.CTkLabel(body,text=f"  {txt}",font=ctk.CTkFont(size=10,weight="bold"),text_color=C["text_muted"]).grid(row=ri[0],column=0,sticky="w",padx=14,pady=(4,0)); ri[0]+=1

        if existing_entries:
            exist_parts=[e for e in existing_entries if e["type"]=="part"]
            exist_assy =[e for e in existing_entries if e["type"]=="assembly"]
            if exist_parts:
                sub_lbl("PARTS")
                make_col_hdr(body,ri,["Old Number","Description","Existing Number","Status"],hdr_bg=C["exist_hdr"])
                for i,e in enumerate(exist_parts):
                    make_entry_row(body,ri,C["exist_bg"] if i%2==0 else C["accent_light"],e,exist_num_labels,show_new_number=False,show_action=False)
            if exist_assy:
                sub_lbl("SUBASSEMBLIES")
                make_col_hdr(body,ri,["Old Number","Description","Existing Number","Status"],hdr_bg=C["exist_hdr"])
                for i,e in enumerate(exist_assy):
                    make_entry_row(body,ri,C["exist_bg"] if i%2==0 else C["accent_light"],e,exist_num_labels,show_new_number=False,show_action=False)
            missing=[e for e in existing_entries if e["status"]=="file_missing"]
            if missing:
                mf=ctk.CTkFrame(body,fg_color=C["warning_bg"],corner_radius=10)
                mf.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=(6,2)); ri[0]+=1
                ctk.CTkLabel(mf,text=f"⚠ {len(missing)} file(s) not found on disk.\nPlease download from Enovia. Then click 'Confirm & Replace'.",
                             font=ctk.CTkFont(size=11),text_color=C["warning"],wraplength=560,justify="left").pack(padx=14,pady=(8,4))
                missing_nums=[e["existing_number"] for e in missing if e.get("existing_number")]
                ctk.CTkButton(mf,text=f"\u29c9  Copy {len(missing_nums)} Missing File Number(s)",
                              font=ctk.CTkFont(size=11),fg_color=C["warning"],hover_color=C["warning"],
                              text_color="#ffffff",corner_radius=10,height=30,
                              command=lambda nums=missing_nums:(copy_cb("\n".join(nums)),messagebox.showinfo("Copied",f"Copied {len(nums)} number(s):\n"+"\n".join(nums)))
                              ).pack(padx=14,pady=(0,8))
            cbf=ctk.CTkFrame(body,fg_color="transparent"); cbf.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=8); ri[0]+=1
            confirm_btn=ctk.CTkButton(cbf,text="✓ Confirm & Replace Existing Parts",
                                      font=ctk.CTkFont(size=13,weight="bold"),
                                      fg_color=C["accent"],hover_color=C["accent_hover"],
                                      text_color="#ffffff",corner_radius=10,height=46)
            if nav_done["replace"]:
                # Phase 1 already ran this session — show the done state so revisiting
                # this tab can't re-trigger the replace.
                _rn=sum(1 for e in existing_entries if e.get("status")=="Replaced")
                confirm_btn.configure(state="disabled",fg_color=C["accent_light"],text_color=C["accent"],
                                      text=f"\u2713 {_rn} Parts Replaced")
            else:
                confirm_btn.configure(command=lambda:(do_confirm_replace(existing_entries,confirm_btn) if _confirm_action("replace parts") else None))
            confirm_btn.pack(fill="x")
            ctk.CTkFrame(body,fg_color=C["accent_mid"],height=4,corner_radius=0).grid(row=ri[0],column=0,sticky="ew",padx=0,pady=12); ri[0]+=1

        if not existing_entries:
            skip_fr=ctk.CTkFrame(body,fg_color=C["success_bg"],corner_radius=10)
            skip_fr.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=(10,4)); ri[0]+=1
            ctk.CTkLabel(skip_fr,text="No existing parts found in Master BOM.\nAll components need new numbers.",
                         font=ctk.CTkFont(size=12),text_color=C["success"],wraplength=560,justify="center").pack(padx=14,pady=8)
            proceed_btn=ctk.CTkButton(body,text="→ Proceed to Update Equations & Materials",
                          font=ctk.CTkFont(size=12,weight="bold"),fg_color=C["success"],hover_color=C["success"],
                          text_color="#ffffff",corner_radius=10,height=46)
            def _go_phase2(b=proceed_btn):
                if _busy[0]: return
                if not _confirm_action("update equations"): return
                b.configure(state="disabled",fg_color=C["accent_mid"],text_color="#ffffff")
                _set_btn_busy(b,"Updating equations & materials\u2026"); root.update(); do_phase2()
            proceed_btn.configure(command=_go_phase2)
            proceed_btn.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=8); ri[0]+=1

    # ─────────────────────────────────────────
    # Results panel — Phase 2 (new numbers)
    # ─────────────────────────────────────────
    def build_new_numbers_panel():
        nav_done["replace"]=True; unlock_nav("newnum")
        for w in results_panel.winfo_children(): w.destroy()
        results_panel.grid_rowconfigure(1,weight=1)
        result_btns=[]
        if gen_done[0]:
            result_btns.append(("View Summary",C["success"],lambda:(build_summary_panel(),show_panel(summary_panel))))
        panel_header(results_panel,"NEW NUMBERS — Phase 2",result_btns)
        body=ctk.CTkScrollableFrame(results_panel,fg_color=C["bg"],corner_radius=0)
        body.grid(row=1,column=0,sticky="nsew"); body.grid_columnconfigure(0,weight=1)
        ri=[0]; new_num_labels={}

        def sec_hdr(title,count,bg=C["card2"]):
            tint=C["success_bg"] if bg==C["new_hdr"] else (C["exist_bg"] if bg==C["exist_hdr"] else C["card2"])
            acc=C["success"] if bg==C["new_hdr"] else C["accent"]
            sh=ctk.CTkFrame(body,fg_color=tint,corner_radius=10)
            sh.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=(12,2)); sh.grid_columnconfigure(1,weight=1)
            _bb=ctk.CTkFrame(sh,fg_color=acc,width=4,height=16,corner_radius=2); _bb.grid(row=0,column=0,sticky="ns",padx=(12,8),pady=10); _bb.grid_propagate(False)
            ctk.CTkLabel(sh,text=f"{title}",font=ctk.CTkFont(size=12,weight="bold"),text_color=acc).grid(row=0,column=1,sticky="w",pady=8)
            ctk.CTkLabel(sh,text=f"{count}",font=ctk.CTkFont(size=12,weight="bold"),
                         fg_color=acc,text_color="#ffffff",
                         corner_radius=6,padx=8,pady=1).grid(row=0,column=2,sticky="e",padx=12,pady=8)
            ri[0]+=1

        def sub_lbl(txt):
            ctk.CTkLabel(body,text=f"  {txt}",font=ctk.CTkFont(size=10,weight="bold"),text_color=C["text_muted"]).grid(row=ri[0],column=0,sticky="w",padx=14,pady=(4,0)); ri[0]+=1

        # Include entries that STILL need a number (before generation) AND those
        # that already got one (after generation), so returning here from the
        # Summary re-shows the generated results instead of the empty message.
        _NN=("needs_update","New Number","No change","Drawing Missing")
        new_parts_list=[e for e in all_entries if e["type"]=="part"     and e["status"] in _NN]
        new_assy_list =sorted([e for e in all_entries if e["type"]=="assembly" and e["status"] in _NN],
                              key=lambda e: e.get("depth",0), reverse=True)
        replaced_entries=[e for e in all_entries if e.get("status") in ("Replaced","file_missing","Failed")]
        if replaced_entries:
            sec_hdr(f"✓ PHASE 1 COMPLETE — {sum(1 for e in replaced_entries if e['status']=='Replaced')} Replaced",len(replaced_entries),bg=C["accent"])
            ri[0]+=1

        if new_parts_list or new_assy_list:
            sec_hdr("NEW NUMBERS NEEDED — Generate via Macro",len(new_parts_list)+len(new_assy_list),bg=C["new_hdr"])
            if new_parts_list:
                sub_lbl("PARTS")
                make_col_hdr(body,ri,["Old Number","Description","Updates","New Number","Status"],hdr_bg=C["new_hdr"])
                for i,e in enumerate(new_parts_list):
                    make_entry_row(body,ri,C["new_row1"] if i%2==0 else C["new_row2"],e,new_num_labels,show_new_number=True,show_action=True)
                pb=ctk.CTkFrame(body,fg_color="transparent"); pb.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=6); ri[0]+=1
                ctk.CTkButton(pb,text="\u29c9  Copy All Old Part Numbers",font=ctk.CTkFont(size=11),
                              fg_color=C["accent_light"],hover_color=C["accent"],text_color=C["accent"],corner_radius=10,height=30,
                              command=lambda:copy_cb("\n".join(e["old_number"] for e in new_parts_list))).pack(side="left",padx=(0,8))
                if gen_done[0]:
                    ctk.CTkButton(pb,text="\u29c9  Copy All New Part Numbers",font=ctk.CTkFont(size=11),
                                  fg_color=C["success_bg"],hover_color=C["success"],text_color=C["success"],corner_radius=10,height=30,
                                  command=lambda:copy_cb("\n".join(e["new_number"] for e in new_parts_list if e.get("status")=="New Number"))).pack(side="left")
            if new_assy_list:
                sub_lbl("SUBASSEMBLIES")
                make_col_hdr(body,ri,["Old Number","Description","Updates","New Number","Status"],hdr_bg=C["new_hdr"])
                for i,e in enumerate(new_assy_list):
                    make_entry_row(body,ri,C["new_row1"] if i%2==0 else C["new_row2"],e,new_num_labels,show_new_number=True,show_action=True)
                ab=ctk.CTkFrame(body,fg_color="transparent"); ab.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=6); ri[0]+=1
                ctk.CTkButton(ab,text="\u29c9  Copy All Old Assembly Numbers",font=ctk.CTkFont(size=11),
                              fg_color=C["accent_light"],hover_color=C["accent"],text_color=C["accent"],corner_radius=10,height=30,
                              command=lambda:copy_cb("\n".join(e["old_number"] for e in new_assy_list))).pack(side="left",padx=(0,8))
                if gen_done[0]:
                    ctk.CTkButton(ab,text="\u29c9  Copy All New Assembly Numbers",font=ctk.CTkFont(size=11),
                                  fg_color=C["success_bg"],hover_color=C["success"],text_color=C["success"],corner_radius=10,height=30,
                                  command=lambda:copy_cb("\n".join(e["new_number"] for e in new_assy_list if e.get("status")=="New Number"))).pack(side="left")
            if not gen_done[0]:
                folder=os.path.dirname(main_path) if main_path else ""
                parts_missing_drw=[]
                for e in new_parts_list:
                    drw=os.path.join(os.path.dirname(e["old_path"]),e["old_number"]+".SLDDRW")
                    if not os.path.exists(drw):
                        parts_missing_drw.append(e)
                        if e.get("_status_lbl"): e["_status_lbl"].configure(text="Drawing Missing",text_color=C["warning"])
                        if e["old_path"] in new_num_labels: new_num_labels[e["old_path"]].configure(text="Drawing Missing",text_color=C["warning"])
                if parts_missing_drw:
                    mf=ctk.CTkFrame(body,fg_color=C["warning"],corner_radius=10)
                    mf.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=(10,4)); ri[0]+=1
                    ctk.CTkLabel(mf,text=f"📥 {len(parts_missing_drw)} drawing(s) missing — status shown above.\nDownload from Enovia. Then click 'Generate New Numbers'.",
                                 font=ctk.CTkFont(size=11),text_color="#ffffff",wraplength=560,justify="left").pack(padx=14,pady=(10,4))
                    missing_nums=[e["old_number"] for e in parts_missing_drw]
                    ctk.CTkButton(mf,text=f"\u29c9  Copy {len(missing_nums)} Missing Drawing Number(s)",
                                  font=ctk.CTkFont(size=11),fg_color="transparent",hover_color="#9c4a08",
                                  text_color="#ffffff",corner_radius=10,height=30,
                                  command=lambda nums=missing_nums:(copy_cb("\n".join(nums)),messagebox.showinfo("Copied",f"Copied {len(nums)} number(s):\n"+"\n".join(nums)))
                                  ).pack(padx=14,pady=(0,10))
                elif new_parts_list:
                    af=ctk.CTkFrame(body,fg_color=C["success_bg"],corner_radius=10)
                    af.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=(8,4)); ri[0]+=1
                    ctk.CTkLabel(af,text="✅ All drawings found — ready to generate!",font=ctk.CTkFont(size=11),text_color=C["success"]).pack(padx=14,pady=8)
                gf=ctk.CTkFrame(body,fg_color="transparent"); gf.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=8); ri[0]+=1
                gen_btn=ctk.CTkButton(gf,text="⚙ Generate New Numbers",
                                      font=ctk.CTkFont(size=13,weight="bold"),
                                      fg_color=C["success"],hover_color=C["success"],
                                      text_color="#ffffff",corner_radius=10,height=46,
                                      command=lambda:(do_generate_new_numbers(new_parts_list,new_assy_list,new_num_labels,gen_btn) if _confirm_action("generate new numbers") else None))
                gen_btn.pack(fill="x")
        else:
            mf=ctk.CTkFrame(body,fg_color=C["success_bg"],corner_radius=10)
            mf.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=20); ri[0]+=1
            ctk.CTkLabel(mf,text="✅ All components replaced with existing parts.\nNo new numbers needed!",
                         font=ctk.CTkFont(size=12,weight="bold"),text_color=C["success"],wraplength=560,justify="center").pack(padx=20,pady=20)

    # ─────────────────────────────────────────
    # Action: Confirm & Replace existing
    # ─────────────────────────────────────────
    def do_confirm_replace(entries, btn):
        if _busy[0]: return
        _busy[0]=True
        btn.configure(state="disabled",fg_color=C["accent_mid"],text_color="#ffffff")
        _set_btn_busy(btn,"Replacing parts\u2026"); _begin_conn(1,15,60)
        timer_resume("Replacing parts...")
        set_progress(15,"Replacing parts...")
        set_status("Replacing existing parts...",C["warning"])
        log("\n"+"="*52+"\n [ PHASE 1: REPLACING EXISTING PARTS ]\n"+"="*52+"\n\n")
        folder=os.path.dirname(main_path)
        rep_total=len(entries); rep_idx=0
        for entry in entries:
            old_path=entry["old_path"]; old_num=entry["old_number"]
            exist_num=entry["existing_number"]; is_assy=entry["type"]=="assembly"
            ext=".SLDASM" if is_assy else ".SLDPRT"
            exist_path=os.path.join(folder,exist_num+ext)
            rep_idx+=1
            pct=15+(rep_idx/max(rep_total,1))*5
            set_progress(pct,f"Replacing {rep_idx}/{rep_total}...")
            log(f"  {old_num} → {exist_num}\n")
            if exist_num.upper()==old_num.upper():
                log(f"  Same number — no replacement needed\n")
                entry["status"]="Replaced"; entry["status_display"]="Replaced"
                entry["new_number"]=exist_num; entry["new_path"]=exist_path
                if entry.get("_status_lbl"): entry["_status_lbl"].configure(text="Replaced",text_color=C["accent"])
                continue
            if not os.path.exists(exist_path):
                log(f"  File not found: {exist_path}\n")
                entry["status"]="file_missing"; entry["new_number"]=exist_num
                if entry.get("_status_lbl"): entry["_status_lbl"].configure(text="File Missing",text_color=C["warning"])
                continue
            replace_target=entry.get("parent_path",main_path) or main_path
            log(f"  Replacing in: {os.path.basename(replace_target)}\n")
            ok=replace_in_assembly(sw_app,replace_target,old_path,exist_path,main_path,log)
            if not ok and replace_target.lower()!=main_path.lower():
                log(f"  Fallback: main assembly\n")
                ok=replace_in_assembly(sw_app,main_path,old_path,exist_path,main_path,log)
            entry["new_number"]=exist_num; entry["new_path"]=exist_path
            entry["status"]="Replaced" if ok else "Failed"
            entry["status_display"]="Replaced" if ok else "Failed"
            if entry.get("_status_lbl"): entry["_status_lbl"].configure(
                text="Replaced" if ok else "Failed",
                text_color=C["accent"] if ok else C["danger"])
            log(f"  {'Replaced ✓' if ok else 'Failed ✗'}\n")
        replaced_count=sum(1 for e in entries if e.get("status")=="Replaced")
        failed_count  =sum(1 for e in entries if e.get("status")=="Failed")
        log("\n  Rebuild after replacement...\n")
        try: activate_doc(sw_app, main_path); sw_model.ForceRebuild3(True); log("  Done\n")
        except Exception as e: log(f"  Warning: {e}\n")
        _save(sw_model); log("  Saved\n")
        _clear_btn_busy()
        if failed_count>0:
            btn.configure(state="disabled",fg_color=C["warning_bg"],text_color=C["warning"],text=f"⚠ {replaced_count} Replaced | {failed_count} Failed")
        else:
            btn.configure(state="disabled",fg_color=C["accent_light"],text_color=C["accent"],text=f"✓ {replaced_count} Parts Replaced Successfully")
        set_progress(20,"Replacement complete ✓")
        log("\n"+"="*52+"\n PHASE 1 complete — starting PHASE 2\n"+"="*52+"\n\n")
        root.update()
        timer_pause("Starting Phase 2...")
        do_phase2()

    # ─────────────────────────────────────────
    # Phase 2
    # ─────────────────────────────────────────
    def do_phase2():
        _busy[0]=True
        if not (active_conn[0] and active_conn[0][0]==1): _begin_conn(1,15,60)
        timer_resume("Updating equations & materials...")
        set_status("Phase 2: Updating equations & materials...",C["warning"]); root.update()
        log("  PHASE 2: Updating equations & materials on flagged parts...\n")
        # Reuse the tables the scan already parsed (set in run_update). Only
        # re-read if the cache is somehow empty (e.g. Phase 2 entered directly).
        if cached_lookup[0] is not None:
            lookup = cached_lookup[0]
        else:
            lookup,err=load_excel_lookup(excel_path)
            if err: log(f"  Excel error: {err}\n"); _busy[0]=False; return
        if cached_parts_rows[0] is not None:
            parts_rows, assy_rows = cached_parts_rows[0], cached_assy_rows[0]
        else:
            parts_rows,assy_rows,bom_err=load_master_bom()
            if bom_err: parts_rows=[]; assy_rows=[]
        sel_names=stored_sel[0]; sel_values=stored_sel[1]
        visited2=set()
        if main_path: visited2.add(main_path.lower())
        for e in all_entries:
            if e.get("status") in ("Replaced","file_missing","existing"):
                visited2.add(e["old_path"].lower())
        needs_update=[e for e in all_entries if e["status"]=="needs_update"]
        log(f"  needs_update count: {len(needs_update)}\n")
        log(f"  all_entries statuses: {[e['status'] for e in all_entries]}\n\n")
        if not needs_update:
            log("  No needs_update entries — updating main assembly equations + custom props\n")
            # Update main assembly equations and custom properties even when
            # all components were replaced (no needs_update children)
            set_progress(30,"Updating main assembly equations...")
            log("="*52+"\n MAIN ASSEMBLY EQUATIONS + CUSTOM PROPERTIES\n"+"="*52+"\n")
            update_equations_on_doc(sw_app,sw_model,main_path,sel_names,sel_values,log,main_path=main_path,close_after=False)
            log("\n")
            # Final rebuild + save
            log("  Final rebuild + save...\n")
            try: activate_doc(sw_app, main_path); sw_model.ForceRebuild3(True); log("  Rebuild done\n")
            except Exception as e: log(f"  Rebuild warning: {e}\n")
            _save(sw_model); log("  Saved\n")
            # This branch ends the workflow (no Phase 2). Restore suppression on
            # any components we resolved during the scan.
            if resolved_comps_all:
                resuppress_all_resolved(sw_app, main_path, sw_model,
                                        resolved_comps_all, all_entries, log)
                try: activate_doc(sw_app, main_path); sw_model.ForceRebuild3(True)
                except Exception: pass
                _save(sw_model); log("  Saved main (after re-suppress)\n")
            # Base-frame configuration LAST (after re-suppress): re-suppress restores
            # each resolved component in the config that was active during the scan;
            # activating the chosen config now loads that config's own stored
            # suppression states, so they can't be clobbered or corrupted.
            apply_subassembly_configs(sw_app, main_path, sw_model, all_entries, log)
            # Final cleanup: close every other open document
            close_all_except_main(sw_app, main_path, log, only_close_paths=_close_allow_set(all_entries))
            # Populate new_vals_snap from SW readback AFTER equations updated
            new_snap = get_current_equation_values(sw_model); new_vals_snap.update(new_snap)
            # Override with user-entered values for selected params (most accurate)
            for n, v in zip(sel_names, sel_values):
                new_vals_snap[n] = format_value(v)
            _paint_new_bar(new_vals_snap)
            orig_bar.grid(row=5,column=0,sticky="new",padx=(0,4))
            new_bar.grid(row=5,column=1,sticky="new",padx=(4,0))
            total_time=timer_stop(); total_time_str[0]=total_time
            set_progress(100,"BOM updated ✓ Complete!")
            set_status("Complete — all replaced ✓",C["success"])
            _busy[0]=False
            try: build_summary_panel(); show_panel(summary_panel)
            except Exception as e:
                log(f"\n Summary error: {e}\n")
                import traceback; log(traceback.format_exc())
                messagebox.showerror("Summary Error",str(e))
            return
        set_progress(30,"Updating equations & materials...")
        log("="*52+"\n MAIN ASSEMBLY EQUATIONS + CUSTOM PROPERTIES\n"+"="*52+"\n")
        update_equations_on_doc(sw_app,sw_model,main_path,sel_names,sel_values,log,main_path=main_path,close_after=False)
        log("\n")
        eq_progress_start=30; eq_progress_end=60; eq_total=len(needs_update)
        for eq_idx,entry in enumerate(needs_update):
            _rem=needs_update[eq_idx:]   # this one + everything after
            set_eq_remaining(parts=sum(1 for e in _rem if e["type"]!="assembly"),
                             assemblies=sum(1 for e in _rem if e["type"]=="assembly"))
            path=entry["old_path"]; is_assy=entry["type"]=="assembly"
            doc=sw_app.GetOpenDocument(path)
            if doc is None:
                try: doc=sw_app.OpenDoc6(path,2 if is_assy else 1,1,"",_byref_int(),_byref_int())
                except Exception as e: log(f"  Could not open {entry['old_number']}: {e}\n"); continue
            if doc is None: continue
            log(f"\n  >> {entry['old_number']} [{entry['desc']}]\n")
            log(f"     Params: {','.join(entry['app_names'])}\n")
            pct=eq_progress_start+((eq_idx+1)/max(eq_total,1))*(eq_progress_end-eq_progress_start)
            set_progress(pct,f"Updating {eq_idx+1}/{eq_total}...")
            _ud=(entry.get('desc') or '').strip()
            _ud=(_ud[:34]+'\u2026') if len(_ud)>34 else _ud
            set_status((f"Updating {entry['old_number']} \u2014 {_ud}..." if _ud
                        else f"Updating {entry['old_number']}..."),C["warning"]); root.update()
            # Close BOTH parts and assemblies after equation update so documents
            # don't pile up in SolidWorks while user downloads from Enovia.
            # They will be reopened in Phase 3 (generate new numbers) as needed.
            update_equations_on_doc(
                sw_app, doc, path,
                entry["app_names"], entry["app_values"],
                log, main_path=main_path, close_after=True,
                target_material=entry.get("target_material"),
                apply_material=entry.get("apply_material", False),
                new_description=entry.get("target_desc"),
                old_description=entry.get("desc"),
            )
        log("\n  Final rebuild + save...\n")
        set_eq_remaining(0,0)   # equations done — hide the remaining readout
        try: activate_doc(sw_app, main_path); sw_model.ForceRebuild3(True); log("  Rebuild done\n")
        except Exception as e: log(f"  Rebuild warning: {e}\n")
        _save(sw_model); log("  Saved\n")
        # NOTE: do NOT re-suppress here. Components stay resolved through Phase 2
        # (new-number generation + replacement), which needs them selectable.
        # Re-suppression happens at the end of do_generate_new_numbers.

        # Update new values bar — show BELOW original values, flag what changed
        new_snap=get_current_equation_values(sw_model); new_vals_snap.update(new_snap)
        _paint_new_bar(new_snap)

        # Show the NEW VALUES strip beneath CURRENT VALUES (both live in `vbars`)
        orig_bar.grid(row=5,column=0,sticky="new",padx=(0,4))
        new_bar.grid(row=5,column=1,sticky="new",padx=(4,0))

        set_progress(60,"Equations & materials updated ✓")
        log("="*52+"\n PHASE 2 complete\n"+"="*52+"\n")
        set_status("Phase 2 done — generate new numbers →",C["success"])
        timer_pause("Waiting for generate...")
        _clear_btn_busy(); _finish_conn(1)
        _busy[0]=False
        last_results[0]="new_numbers"; build_new_numbers_panel(); show_panel(results_panel)

    # ─────────────────────────────────────────
    # Action: Generate new numbers
    # ─────────────────────────────────────────
    def do_generate_new_numbers(new_parts_list,new_assy_list,num_labels,gen_btn):
        if _busy[0]: return
        # Gate: new numbers must be written to the Master BOM, which we can't do
        # while it's open in Excel. Check BEFORE doing anything and ask the user
        # to close it, rather than generating numbers we can't record.
        if _workbook_is_open(get_active_bom_path()):
            messagebox.showwarning("Master BOM is open in Excel",
                "The Master BOM is open in Excel, so new part numbers can't be written to it.\n\n"
                "Please close the Master BOM in Excel, then click \"Generate New Numbers\" again.")
            return
        _busy[0]=True
        gen_btn.configure(state="disabled",fg_color=C["accent_mid"],text_color="#ffffff")
        _set_btn_busy(gen_btn,"Running macro\u2026"); _begin_conn(2,65,95)
        timer_resume("Generating numbers...")
        set_progress(65,"Running macro...")
        set_status("Generating new numbers...",C["warning"]); root.update()
        gen_total=len(new_assy_list)+len(new_parts_list); gen_idx=[0]
        bom_pending=[]   # (entry, is_assembly) rows to write to the BOM in ONE save
        title_counts={"fixed":0,"ok":0}   # tree-title reset tally for the summary
        log("\n"+"="*52+"\n [ GENERATING NEW NUMBERS ]\n"+"="*52+"\n\n")
        _gp_total=len(new_parts_list); _ga_total=len(new_assy_list)
        for _gpi,entry in enumerate(new_parts_list):
            old_path=entry["old_path"]; old_num=entry["old_number"]
            gen_idx[0]+=1
            pct=65+(gen_idx[0]/max(gen_total,1))*25
            set_progress(pct,f"Macro {gen_idx[0]}/{gen_total}...")
            set_eq_remaining(parts=_gp_total-_gpi, assemblies=_ga_total)
            folder=os.path.dirname(old_path); drw_path=os.path.join(folder,old_num+".SLDDRW")
            log(f"  Part: {old_num}\n")
            if not os.path.exists(drw_path):
                log(f"  Drawing not found\n"); entry["status"]="Drawing Missing"; entry["status_display"]="Drawing Missing"
                if old_path in num_labels: num_labels[old_path].configure(text="Drawing Missing",text_color=C["warning"])
                if entry.get("_status_lbl"): entry["_status_lbl"].configure(text="Drawing Missing",text_color=C["warning"])
                continue
            new_drw=run_macro_on_file(sw_app,drw_path,is_drawing=True,main_path=main_path,log=log,title_counts=title_counts)
            if new_drw and new_drw.lower()!=drw_path.lower():
                new_num=get_part_number_from_path(new_drw)
                # Derive the renamed PART path from the renamed drawing path.
                # The macro renames drawing + part together, same name different ext.
                new_part_path = os.path.splitext(new_drw)[0] + ".SLDPRT"
                if not os.path.exists(new_part_path):
                    # Try .sldprt lowercase as a fallback
                    alt = os.path.splitext(new_drw)[0] + ".sldprt"
                    if os.path.exists(alt): new_part_path = alt
                entry["new_path"]=new_part_path; entry["new_number"]=new_num; entry["status"]="New Number"; entry["status_display"]="New Number"
                log(f"  New number: {new_num}\n")
                # Stamp the renamed drawing's date custom property with today.
                stamp_drawing_date(sw_app, new_drw, log, indent="  ")
                # Close the renamed drawing and (if open) the renamed part so
                # they don't pile up while later parts are processed.
                for dp,label in [(new_drw,"new drw"),(new_part_path,"new prt"),(drw_path,"old drw")]:
                    try:
                        d=sw_app.GetOpenDocument(dp)
                        if d:
                            try:   t=d.GetTitle()
                            except TypeError: t=d.GetTitle
                            sw_app.CloseDoc(t); log(f"  Closed {label}: {t}\n")
                    except Exception: pass
                if old_path in num_labels: num_labels[old_path].configure(text=new_num,text_color=C["success"])
                if entry.get("_status_lbl"): entry["_status_lbl"].configure(text="New Number",text_color=C["success"])
                bom_pending.append((entry, False))   # batched write at end of run
            else:
                log(f"  No new file\n"); entry["status"]="No change"; entry["status_display"]="No change"
                if old_path in num_labels: num_labels[old_path].configure(text="No change",text_color=C["text_muted"])
                if entry.get("_status_lbl"): entry["_status_lbl"].configure(text="No change",text_color=C["text_muted"])
            log("\n" + "-"*40 + "\n")   # separator so each part's log block is distinct
        for _gai,entry in enumerate(new_assy_list):
            old_path=entry["old_path"]; old_num=entry["old_number"]
            gen_idx[0]+=1
            pct=65+(gen_idx[0]/max(gen_total,1))*25
            set_progress(pct,f"Macro {gen_idx[0]}/{gen_total}...")
            set_eq_remaining(parts=0, assemblies=_ga_total-_gai)
            log(f"  Subassembly: {old_num}\n")
            try:
                errors=_byref_int(); d=sw_app.GetOpenDocument(old_path)
                if d is None: d=sw_app.OpenDoc6(old_path,2,1,"",errors,_byref_int())
                if d: activate_doc(sw_app, old_path)
            except Exception as e: log(f"  Activate: {e}\n")
            new_path=run_macro_on_file(sw_app,old_path,is_drawing=False,main_path=main_path,log=log,title_counts=title_counts)
            if new_path and new_path.lower()!=old_path.lower():
                new_num=get_part_number_from_path(new_path)
                entry["new_path"]=new_path; entry["new_number"]=new_num
                for ce in all_entries:
                    if ce["parent_path"].lower()==old_path.lower():
                        ce["parent_path"]=new_path; log(f"  Child parent updated: {ce['old_number']}\n")
                for dp,label in [(old_path,"old"),(new_path,"new")]:
                    try:
                        d=sw_app.GetOpenDocument(dp)
                        if d:
                            try:   t=d.GetTitle()
                            except TypeError: t=d.GetTitle
                            sw_app.CloseDoc(t); log(f"  Closed {label}: {t}\n")
                    except Exception as e: log(f"  Close {label} warning: {e}\n")
                replace_target=entry.get("parent_path",main_path) or main_path
                try: activate_doc(sw_app, replace_target)
                except Exception: pass
                log(f"  Replacing in: {os.path.basename(replace_target)}\n")
                ok=replace_in_assembly(sw_app,replace_target,old_path,new_path,main_path,log)
                if not ok and replace_target.lower()!=main_path.lower():
                    ok=replace_in_assembly(sw_app,main_path,old_path,new_path,main_path,log)
                entry["status"]="New Number" if ok else "Failed"
                entry["status_display"]="New Number" if ok else "Failed"
                log(f"  New number: {new_num}\n")
                if old_path in num_labels: num_labels[old_path].configure(text=new_num,text_color=C["success"] if ok else C["danger"])
                if entry.get("_status_lbl"): entry["_status_lbl"].configure(text="New Number" if ok else "Failed",text_color=C["success"] if ok else C["danger"])
                bom_pending.append((entry, True))   # batched write at end of run
            else:
                log(f"  No new file\n"); entry["status"]="No change"; entry["status_display"]="No change"
                if old_path in num_labels: num_labels[old_path].configure(text="No change",text_color=C["text_muted"])
                if entry.get("_status_lbl"): entry["_status_lbl"].configure(text="No change",text_color=C["text_muted"])
            log("\n" + "-"*40 + "\n")   # separator so each assembly's log is distinct
        # ── Tree-title reset summary ──
        set_eq_remaining(0,0)   # generation done — hide the remaining readout
        _tf_fixed=title_counts.get("fixed",0); _tf_ok=title_counts.get("ok",0)
        if _tf_fixed or _tf_ok:
            log(f"\n  Tree titles: {_tf_fixed} reset to filename, {_tf_ok} already correct\n")
        # ── Write ALL new rows to the Master BOM in a single open + save ──
        if bom_pending:
            set_status("Updating Master BOM...",C["warning"]); root.update()
            ok_bom,bom_err,n_written=append_all_to_master_bom(bom_pending, lookup=cached_lookup[0])
            for _ent,_ in bom_pending:
                _ent["bom_appended"]=ok_bom
            if ok_bom:
                log(f"  Appended {n_written} row(s) to Master BOM"
                    f"{' ('+bom_err+')' if bom_err else ''}\n")
            else:
                log(f"  Master BOM append failed: {bom_err}\n")
        log("\n  Final rebuild...\n")
        try: activate_doc(sw_app, main_path); sw_model.ForceRebuild3(True); log("  Done\n")
        except Exception as e: log(f"  Warning: {e}\n")
        _save(sw_model); log("  Saved\n")
        # Workflow complete — restore suppression on any components we resolved
        # during the scan, now that equations are updated and new numbers/
        # replacements are done. Re-find each component fresh (its COM ref may
        # be stale after rename/replace) and suppress it in its parent.
        if resolved_comps_all:
            resuppress_all_resolved(sw_app, main_path, sw_model,
                                    resolved_comps_all, all_entries, log)
            try: activate_doc(sw_app, main_path); sw_model.ForceRebuild3(True)
            except Exception: pass
            _save(sw_model); log("  Saved main (after re-suppress)\n")
        # Base-frame configuration LAST: re-suppress (above) restores each resolved
        # component to its pre-scan state in whatever config was active during the
        # scan. Activating the chosen config now loads THAT config's own stored
        # suppression states, so re-suppress can't clobber them and the base
        # frame's configurations aren't corrupted. Must run after re-suppress.
        apply_subassembly_configs(sw_app, main_path, sw_model, all_entries, log)
        # Final cleanup: close every other open document so the user has only
        # the main assembly open after the workflow completes.
        close_all_except_main(sw_app, main_path, log, only_close_paths=_close_allow_set(all_entries))
        log("="*52+"\n ALL DONE\n")
        set_progress(90,"New numbers complete ✓")
        gen_done[0]=True; set_status("New numbers generated — opening summary",C["success"])
        _clear_btn_busy(); _finish_conn(2)
        gen_btn.configure(state="disabled",fg_color=C["success_bg"],text_color=C["success"],text="✓ Generation Complete")
        root.update()
        total_time=timer_stop(); total_time_str[0]=total_time
        set_progress(100,"BOM updated ✓ Complete!")
        _busy[0]=False
        try: build_summary_panel(); show_panel(summary_panel)
        except Exception as e:
            log(f"\n Summary error: {e}\n")
            import traceback; log(traceback.format_exc())
            messagebox.showerror("Summary Error",str(e))

    # ─────────────────────────────────────────
    # Summary panel
    # ─────────────────────────────────────────
    def build_summary_panel():
        summary_generated[0] = True
        nav_done["replace"]=True; nav_done["newnum"]=True; nav_done["summary"]=True; unlock_nav("summary")
        for w in summary_panel.winfo_children(): w.destroy()
        summary_panel.grid_rowconfigure(1,weight=1)
        panel_header(summary_panel,"SUMMARY REPORT",[
            ("Export Excel",C["success"],do_export_excel),
            ("Export PDF",C["accent"],do_export_pdf),
        ])
        body=ctk.CTkScrollableFrame(summary_panel,fg_color=C["bg"],corner_radius=0)
        body.grid(row=1,column=0,sticky="nsew"); body.grid_columnconfigure(0,weight=1)
        ri=[0]
        def sec(title):
            f=ctk.CTkFrame(body,fg_color=C["card2"],corner_radius=10)
            f.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=(10,2)); f.grid_columnconfigure(0,weight=1)
            ctk.CTkLabel(f,text=f"  {title}",font=ctk.CTkFont(size=12,weight="bold"),text_color=C["text"]).grid(row=0,column=0,sticky="w",padx=10,pady=6); ri[0]+=1
        def info_row(label,value):
            f=ctk.CTkFrame(body,fg_color=C["card"],corner_radius=0)
            f.grid(row=ri[0],column=0,sticky="ew",padx=10); f.grid_columnconfigure(1,weight=1)
            ctk.CTkLabel(f,text=label,font=ctk.CTkFont(size=11,weight="bold"),text_color=C["text_dim"],width=180,anchor="w").grid(row=0,column=0,padx=12,pady=5)
            ctk.CTkLabel(f,text=value,font=ctk.CTkFont(size=11),text_color=C["text"],anchor="w").grid(row=0,column=1,sticky="w",padx=6,pady=5); ri[0]+=1
        def div():
            ctk.CTkFrame(body,fg_color=C["border"],height=1,corner_radius=0).grid(row=ri[0],column=0,sticky="ew",padx=10); ri[0]+=1
        counts=build_counts()
        sec("REPORT INFORMATION")
        info_row("Job Number:",os.path.splitext(assy_title or "")[0])
        if job_reference[0]: info_row("Reference Used:",job_reference[0])
        info_row("Timestamp:",update_timestamp[0])
        if total_time_str[0]: info_row("Time Taken:",total_time_str[0])
        info_row("Parts Updated:",f"{counts['parts_total']} ({counts['parts_new']} new numbers | Existing replaced: {counts['parts_existing']} | Missing: {counts['parts_missing']})")
        info_row("Subassemblies Updated:",f"{counts['assy_total']} ({counts['assy_new']} new numbers | Existing replaced: {counts['assy_existing']} | Failed: {counts['assy_failed']})")
        ri[0]+=1
        changed_dims=[(pl,orig_vals_snap.get(pn),new_vals_snap.get(pn)) for pn,pl in PARAMS
                      if orig_vals_snap.get(pn) and new_vals_snap.get(pn) and orig_vals_snap.get(pn)!=new_vals_snap.get(pn)]
        if changed_dims:
            sec("DIMENSION CHANGES")
            dh=ctk.CTkFrame(body,fg_color=C["accent"],corner_radius=0,height=26)
            dh.grid(row=ri[0],column=0,sticky="ew",padx=10); ri[0]+=1
            for i in range(3): dh.grid_columnconfigure(i,weight=1)
            for col,txt in enumerate(["Parameter","Original","New Value"]):
                ctk.CTkLabel(dh,text=txt,font=ctk.CTkFont(size=10,weight="bold"),text_color="#ffffff",anchor="center",justify="center").grid(row=0,column=col,sticky="ew",padx=8,pady=4)
            for i,(plabel,o,n) in enumerate(changed_dims):
                bg=C["card"] if i%2==0 else C["card2"]
                dr=ctk.CTkFrame(body,fg_color=bg,corner_radius=0)
                dr.grid(row=ri[0],column=0,sticky="ew",padx=10); ri[0]+=1
                for j in range(3): dr.grid_columnconfigure(j,weight=1)
                ctk.CTkLabel(dr,text=plabel,font=ctk.CTkFont(size=11,weight="bold"),text_color=C["text"],anchor="center",justify="center").grid(row=0,column=0,sticky="ew",padx=8,pady=6)
                ctk.CTkLabel(dr,text=f"{format_value(o)} in",font=ctk.CTkFont(size=11),text_color=C["text_dim"],anchor="center",justify="center").grid(row=0,column=1,sticky="ew",padx=8,pady=6)
                ctk.CTkLabel(dr,text=f"{format_value(n)} in",font=ctk.CTkFont(size=11,weight="bold"),fg_color=C["accent_light"],text_color=C["accent"],corner_radius=6,padx=6,anchor="center",justify="center").grid(row=0,column=2,sticky="ew",padx=8,pady=6)
            div(); ri[0]+=1

        mat_changes=build_material_changes()
        if mat_changes:
            sec("MATERIAL CHANGES")
            mh=ctk.CTkFrame(body,fg_color=C["accent"],corner_radius=0,height=26)
            mh.grid(row=ri[0],column=0,sticky="ew",padx=10); ri[0]+=1
            for i in range(3): mh.grid_columnconfigure(i,weight=1)
            for col,txt in enumerate(["Material Type","Finish Applied","Parts Changed"]):
                ctk.CTkLabel(mh,text=txt,font=ctk.CTkFont(size=10,weight="bold"),text_color="#ffffff",anchor="center",justify="center").grid(row=0,column=col,sticky="ew",padx=8,pady=4)
            for i,m in enumerate(mat_changes):
                bg=C["card"] if i%2==0 else C["card2"]
                mr=ctk.CTkFrame(body,fg_color=bg,corner_radius=0)
                mr.grid(row=ri[0],column=0,sticky="ew",padx=10); ri[0]+=1
                for j in range(3): mr.grid_columnconfigure(j,weight=1)
                ctk.CTkLabel(mr,text=m["name"],font=ctk.CTkFont(size=11,weight="bold"),text_color=C["text"],anchor="center",justify="center").grid(row=0,column=0,sticky="ew",padx=8,pady=6)
                ctk.CTkLabel(mr,text=m["finish"],font=ctk.CTkFont(size=11,weight="bold"),fg_color=C["accent_light"],text_color=C["accent"],corner_radius=6,padx=6,anchor="center",justify="center").grid(row=0,column=1,sticky="ew",padx=8,pady=6)
                ctk.CTkLabel(mr,text=str(m["count"]),font=ctk.CTkFont(size=11),text_color=C["text_dim"],anchor="center",justify="center").grid(row=0,column=2,sticky="ew",padx=8,pady=6)
            div(); ri[0]+=1

        def tbl_sec(title,entries):
            if not entries: return
            new_num_entries=[e for e in entries if e.get("status_display","")=="New Number"]
            bom_added=sum(1 for e in new_num_entries if e.get("bom_appended"))
            bom_failed=sum(1 for e in new_num_entries if e.get("bom_appended")==False)
            if new_num_entries:
                bom_txt=f" | {bom_added} added to Master BOM"
                if bom_failed: bom_txt+=f" | {bom_failed} failed"
                full_title=f"{title} ({len(entries)} total{bom_txt})"
            else:
                full_title=f"{title} ({len(entries)} total)"
            sec(full_title)
            make_col_hdr(body,ri,["Old Number","Description","Updates","New/Existing Number","Status"])
            new_nums=[]
            for i,entry in enumerate(entries):
                bg=C["card"] if i%2==0 else C["card2"]
                fr=ctk.CTkFrame(body,fg_color=bg,corner_radius=0)
                fr.grid(row=ri[0],column=0,sticky="ew",padx=10)
                weights=[0,1,0,0,0]; minsizes=[145,0,110,130,110]
                for j,(w,ms) in enumerate(zip(weights,minsizes)): fr.grid_columnconfigure(j,weight=w,minsize=ms)
                ri[0]+=1
                ctk.CTkFrame(body,fg_color=C["border"],height=1,corner_radius=0).grid(row=ri[0],column=0,sticky="ew",padx=10); ri[0]+=1
                st=entry.get("status_display",entry.get("status",""))
                if   st=="New Number": sc2,sb2=C["success"],C["success_bg"]
                elif st=="Replaced":   sc2,sb2=C["accent"],C["accent_light"]
                elif st in ("File Missing","Drawing Missing"): sc2,sb2=C["warning"],C["warning_bg"]
                else: sc2,sb2=C["danger"],C["danger_bg"]
                of=ctk.CTkFrame(fr,fg_color="transparent"); of.grid(row=0,column=0,pady=7,padx=6)
                ctk.CTkLabel(of,text=entry["old_number"],font=ctk.CTkFont(size=11,weight="bold"),text_color=C["text"]).pack(side="left")
                ctk.CTkButton(of,text="\u29c9",font=ctk.CTkFont(size=10),fg_color="transparent",hover_color=C["border"],
                              text_color=C["accent"],width=22,height=22,corner_radius=6,
                              command=lambda n=entry["old_number"]:(copy_cb(n),messagebox.showinfo("Copied",f"Copied: {n}"))
                              ).pack(side="left",padx=(3,0))
                ctk.CTkLabel(fr,text=entry["desc"],font=ctk.CTkFont(size=11),text_color=C["text_dim"],anchor="center",justify="center").grid(row=0,column=1,sticky="ew",padx=8,pady=7)
                # Update Type column — derived from entry flags + material name
                upd_text = _entry_action_label(entry)
                upd_color = C["text_muted"] if upd_text == "—" else C["text"]
                ctk.CTkLabel(fr, text=upd_text,
                             font=ctk.CTkFont(size=10, weight="bold"),
                             text_color=upd_color, anchor="center", justify="center",
                             ).grid(row=0, column=2, sticky="ew", padx=6, pady=7)
                nf=ctk.CTkFrame(fr,fg_color="transparent"); nf.grid(row=0,column=3,pady=7,padx=6)
                num_show=entry.get("new_number","— ")
                ctk.CTkLabel(nf,text=num_show,font=ctk.CTkFont(size=11,weight="bold"),text_color=sc2 if st in ("New Number","Replaced") else C["text_muted"]).pack(side="left")
                if st in ("New Number","Replaced") and num_show not in ("— ","Pending..."):
                    ctk.CTkButton(nf,text="\u29c9",font=ctk.CTkFont(size=10),fg_color="transparent",hover_color=C["border"],
                                  text_color=sc2,width=22,height=22,corner_radius=6,
                                  command=lambda n=num_show:(copy_cb(n),messagebox.showinfo("Copied",f"Copied: {n}"))
                                  ).pack(side="left",padx=(3,0))
                new_nums.append(num_show)
                ctk.CTkLabel(fr,text=st,font=ctk.CTkFont(size=10,weight="bold"),fg_color=sb2,text_color=sc2,corner_radius=6,padx=6,anchor="center",justify="center").grid(row=0,column=4,sticky="ew",padx=8,pady=7)
            if new_nums:
                cf=ctk.CTkFrame(body,fg_color="transparent"); cf.grid(row=ri[0],column=0,sticky="ew",padx=10,pady=6); ri[0]+=1
                ctk.CTkButton(cf,text="\u29c9  Copy All New Numbers",font=ctk.CTkFont(size=11),
                              fg_color=C["success_bg"],hover_color=C["success"],text_color=C["success"],corner_radius=10,height=30,
                              command=lambda nums=new_nums:copy_cb("\n".join(nums))).pack(side="left")

        parts_entries=[e for e in all_entries if e["type"]=="part"]
        assy_entries =[e for e in all_entries if e["type"]=="assembly"]
        tbl_sec("PARTS",parts_entries); tbl_sec("SUBASSEMBLIES",assy_entries)

    def build_counts():
        parts=[e for e in all_entries if e["type"]=="part"]
        assy =[e for e in all_entries if e["type"]=="assembly"]
        def st(e): return e.get("status_display",e.get("status",""))
        return {
            "parts_total":    len(parts),
            "parts_new":      sum(1 for e in parts if st(e)=="New Number"),
            "parts_existing": sum(1 for e in parts if st(e)=="Replaced"),
            "parts_missing":  sum(1 for e in parts if st(e) in ("Drawing Missing","File Missing")),
            "assy_total":     len(assy),
            "assy_new":       sum(1 for e in assy if st(e)=="New Number"),
            "assy_existing":  sum(1 for e in assy if st(e)=="Replaced"),
            "assy_failed":    sum(1 for e in assy if st(e)=="Failed"),
        }

    def build_material_changes():
        """Group the parts that actually got re-materialed this run by their
        material selector, so the summary can show which finishes were applied —
        mirroring the dimension-changes section. Only selectors with at least one
        changed part appear. A group selector (e.g. Interior) spans gauges, so we
        report the FINISH word, not a code."""
        agg={}; order=[]
        for e in all_entries:
            if e.get("type")!="part": continue
            if not (e.get("apply_material") and e.get("target_material")): continue
            ab=normalize_abbrev(e.get("abbrev"))
            sel=None
            for k in MATERIAL_CHOICES:
                if normalize_abbrev(k)==ab: sel=k; break
            if sel is None:
                for gname,members in MATERIAL_GROUP_CODES.items():
                    if ab in members: sel=gname; break
            if sel is None:
                sel=(e.get("abbrev") or "Material").strip()
            kind=_finish_kind(e["target_material"])
            disp="Painted" if kind=="painted" else ("Brushed SS" if kind=="brushed" else e["target_material"])
            if sel not in agg: agg[sel]={"finish":disp,"count":0}; order.append(sel)
            agg[sel]["count"]+=1
        return [{"name":s,"finish":agg[s]["finish"],"count":agg[s]["count"]} for s in order]

    def build_summary_data():
        counts=build_counts()
        dims=[{"name":pl,"original":f"{format_value(orig_vals_snap.get(pn,''))} in","new":f"{format_value(new_vals_snap.get(pn,''))} in"}
              for pn,pl in PARAMS if orig_vals_snap.get(pn) and new_vals_snap.get(pn) and orig_vals_snap.get(pn)!=new_vals_snap.get(pn)]
        parts=[e for e in all_entries if e["type"]=="part"]
        assy =[e for e in all_entries if e["type"]=="assembly"]
        for e in parts+assy:
            if "status_display" not in e:
                raw=e.get("status","")
                if   raw=="existing":     e["status_display"]="Part Found"
                elif raw=="file_missing": e["status_display"]="File Missing"
                elif raw=="needs_update": e["status_display"]="Pending..."
                else:                     e["status_display"]=raw
            e["updates_label"]=_entry_action_label(e)   # same "Updates" text the panel shows
        return {"assembly":assy_title,"reference":job_reference[0],
                "timestamp":update_timestamp[0],"time_taken":total_time_str[0],
                "user":get_user_display_name(),
                "dimensions":dims,"materials":build_material_changes(),
                "counts":counts,"parts":parts,"assemblies":assy}

    def do_export_excel():
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel files","*.xlsx")],
                                          initialfile=f"SW_Update_Summary_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        if not path: return
        try: export_to_excel(build_summary_data(),path); messagebox.showinfo("Exported",f"Excel saved:\n{path}")
        except Exception as e: messagebox.showerror("Export Failed",str(e))

    def do_export_pdf():
        path=filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF files","*.pdf")],
                                          initialfile=f"SW_Update_Summary_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
        if not path: return
        try: export_to_pdf(build_summary_data(),path); messagebox.showinfo("Exported",f"PDF saved:\n{path}")
        except Exception as e: messagebox.showerror("Export Failed",str(e))

    # ─────────────────────────────────────────
    # Run update
    # ─────────────────────────────────────────
    def run_update():
        nonlocal main_path, assy_title, sw_model
        if _busy[0]: return
        if not sw_model: messagebox.showerror("Error","Not connected to SolidWorks."); return
        if not excel_ok: messagebox.showerror("Error","Guideline Excel not found."); return
        names,values=collect()
        if names is None: return
        automation_started[0] = True
        summary_generated[0] = False
        qc_answers.clear()
        qc_measure_started.clear()
        qc_macro_complete[0] = False
        qc_report_path[0] = ""
        qc_record_path[0] = ""
        qc_completed[0] = False
        qc_question_index[0] = 0
        _busy[0]=True
        all_entries.clear(); gen_done[0]=False
        # Fresh run — clear the step indicators so a previous run's nav_done
        # (esp. 'replace') can't disable Confirm & Replace this time around.
        nav_done.update({"replace":False,"newnum":False,"summary":False})
        resolved_comps_all.clear()
        orig_vals_snap.clear(); new_vals_snap.clear()
        stored_sel[0]=names; stored_sel[1]=values
        # Lock all inputs the moment Run Update is clicked; they stay locked
        # until the user presses Connect to Model (which calls set_inputs_locked(False)).
        set_inputs_locked(True)
        # ── Save As job number (optional, FIRST) ─────────────────────────────
        # If a job number was entered, save the open model as that number and
        # switch the working file to it BEFORE anything reads or changes the model,
        # so every update below runs on the job copy and the template stays clean.
        _job = ""
        try: _job = job_var.get().strip()
        except Exception: _job = ""
        job_reference[0] = ""
        if _job:
            log("="*52+f"\n JOB: save model as '{_job}'\n"+"="*52+"\n")
            # Remember the number the model is CURRENTLY open as — this becomes the
            # "Saved From / Reference" on the report (the source the job was cut from).
            _orig_ref = get_part_number_from_path(main_path) or os.path.splitext(os.path.basename(main_path))[0]
            set_status("Saving as job number...",C["warning"]); root.update()
            _newp = _save_as_job(sw_app, sw_model, main_path, _job, log)
            if _newp:
                main_path = _newp
                job_reference[0] = _orig_ref
                # The Save-As may have swapped the active document — re-acquire the
                # model pointer from the active doc so everything downstream (scan,
                # phase 2, sub-assembly config) works on the live handle, not a stale one.
                try:
                    _adoc = sw_app.ActiveDoc
                    if _adoc is not None:
                        try:    _ap = _adoc.GetPathName()
                        except TypeError: _ap = _adoc.GetPathName
                        if _ap and os.path.normcase(os.path.abspath(_ap)) == os.path.normcase(os.path.abspath(_newp)):
                            sw_model = _adoc
                except Exception: pass
                try:    assy_title = sw_model.GetTitle()
                except Exception: assy_title = os.path.basename(_newp)
                try:    title_lbl.configure(text=assy_title)
                except Exception: pass
            log("="*52+"\n")
        # ─────────────────────────────────────────────────────────────────────
        # Capture the user's material radio selections into SELECTED_MATERIALS.
        # This is read later by the walker via get_selected_material(abbrev).
        SELECTED_MATERIALS.clear()
        DOCS_OPENED_IN_RUN.clear()
        CONFIG_TARGETS.clear()
        clear_design_ref_cache()   # fresh read of the Design Ref this run
        for _abbrev, _var in material_radio_vars.items():
            try:    SELECTED_MATERIALS[_abbrev] = _var.get()
            except Exception: pass
        # Capture the base-frame configuration selection (None on an axis = leave
        # the base frame's active config as-is for that axis).
        BASE_FRAME_CFG["leveling_pad"] = {"lp": True, "nolp": False}.get(bf_lp_var.get())
        BASE_FRAME_CFG["refrigerant"]  = {"DX": "DX", "GLYCOL": "GLYCOL"}.get(bf_ref_var.get())
        orig_vals_snap.update(get_current_equation_values(sw_model))
        update_timestamp[0]=datetime.now().strftime("%d %b %Y %H:%M:%S")
        run_btn.configure(state="disabled",fg_color=C["accent_mid"],text_color="#ffffff")
        _reset_conns(); _set_btn_busy(run_btn,"Scanning tree\u2026"); _begin_conn(0,5,20)
        timer_reset(); timer_start_fn("Scanning...")
        set_progress(5,"Loading...")
        set_status("Loading lookup tables...",C["warning"]); root.update()
        log("="*52+"\n STEP 1: Loading Guideline Excel...\n")
        lookup,err=load_excel_lookup(excel_path)
        if err:
            log(f"  ERROR: {err}\n")
            _busy[0]=False
            _clear_btn_busy(); _reset_conns()
            run_btn.configure(state="normal",fg_color=C["accent"],text_color="#ffffff",text="▶ Run Update"); return
        log(f"  Loaded {len(lookup)} descriptions\n")
        set_progress(10,"Guideline loaded")
        log("  STEP 2: Loading Master BOM...\n")
        parts_rows,assy_rows,bom_err=load_master_bom()
        if bom_err:
            log(f"  WARNING: Master BOM error: {bom_err}\n  Proceeding without BOM check\n"); parts_rows=[]; assy_rows=[]
        else:
            log(f"  Loaded {len(parts_rows)} parts, {len(assy_rows)} assemblies from Master BOM\n")
        # Cache for PHASE 2 so it doesn't re-parse both workbooks.
        cached_lookup[0]     = lookup
        cached_parts_rows[0] = parts_rows
        cached_assy_rows[0]  = assy_rows
        log("="*52+"\n\n")
        log("  STEP 3: Walking assembly tree (read only)...\n")
        set_status("Scanning assembly tree...",C["warning"]); root.update()
        folder=os.path.dirname(main_path) if main_path else ""
        visited=set()
        if main_path: visited.add(main_path.lower())
        main_desc=get_description(sw_model)
        main_app_names,main_app_values=get_applicable_params(lookup,main_desc,names,values)
        if main_app_names:
            main_existing=find_in_master_bom(main_desc,main_app_names,main_app_values,True,parts_rows,assy_rows)
            log(f"  Main assembly [{main_desc}]: {'BOM FOUND →'+main_existing if main_existing else 'needs update'}\n")
        # Detect the model's variant from the TOP-LEVEL SUBASSEMBLY descriptions
        # (not the top assembly's own description). Fall back to the top
        # description only if the subassemblies yield no front-panel tag.
        log("  Detecting variant from direct-child front-panel & raceway assemblies...\n")
        sub_fp, sub_rcw = detect_variant_from_subassemblies(sw_model, sw_app, log)
        main_fp  = sub_fp if sub_fp else extract_front_panel(main_desc)
        main_rcw = sub_rcw            # {'RR'} / {'FR'} / {'BB'} from the raceway assembly
        if not sub_fp:
            log("  (no front-panel assembly found among direct children — fell back to top description)\n")
        assy_config = {"front_panel": main_fp, "raceways": main_rcw}
        fp_lbl  = main_fp if isinstance(main_fp, str) else (",".join(sorted(main_fp)) if main_fp else "—")
        rcw_lbl = ",".join(sorted(main_rcw)) if main_rcw else "—"
        log(f"  Assembly variant: FP={fp_lbl} RCW={rcw_lbl}\n")
        set_progress(15,"Scanning tree...")
        # Pass current SW equation values so BOM matching uses correct values
        # for params that are applicable to a part but NOT selected by the user
        cur_eq_vals = get_current_equation_values(sw_model)
        node_info = {}
        entries=walk_assembly_tree(sw_app,sw_model,names,values,lookup,parts_rows,assy_rows,
                                   visited,log,lambda t:set_status(t,C["warning"]),root,folder,
                                   assy_config=assy_config, cur_eq_vals=cur_eq_vals,
                                   resolved_comps=resolved_comps_all, node_info=node_info)
        all_entries.extend(entries)

        # ── Cascade new numbers up the tree ──────────────────────────────────
        # When a child gets a NEW number its parent assembly's BOM changes, so
        # the parent must also get a new number — and so on up the tree — EXCEPT
        # the top-level main assembly, which keeps its number. Seed from every
        # needs_update node and walk UP node_info, flagging each ancestor
        # assembly that isn't already flagged. Generation later runs bottom-up so
        # each parent picks up its renamed child.
        try:
            main_lc = main_path.lower()
            have = {e["old_path"].lower() for e in all_entries}
            cascade_added = {}
            for seed in [e for e in all_entries
                         if e.get("status") in ("needs_update", "existing")]:
                _seed_lbl = f"{seed.get('old_number','?')} [{seed.get('desc','')}]"
                p = node_info.get(seed["old_path"].lower(), {}).get("parent", "")
                while p and p.lower() != main_lc:
                    pk = p.lower()
                    ninfo = node_info.get(pk)
                    if not ninfo:
                        break
                    if pk not in have and pk not in cascade_added:
                        cascade_added[pk] = {
                            "old_path":        p,
                            "old_number":      ninfo.get("old_number") or get_part_number_from_path(p),
                            "desc":            ninfo.get("desc", ""),
                            "parent_path":     ninfo.get("parent", main_path) or main_path,
                            "depth":           ninfo.get("depth", 0),
                            "type":            "assembly",
                            "_cascade_seed":   _seed_lbl,
                            "app_names":       [],
                            "app_values":      [],
                            "param_values":    {},
                            "existing_number": None,
                            "new_number":      "Pending...",
                            "new_path":        "",
                            "status":          "needs_update",
                            "_status_lbl":     None,
                            "_cascade":        True,
                        }
                    p = ninfo.get("parent", "")
            if cascade_added:
                for v in cascade_added.values():
                    all_entries.append(v)
                log(f"  Cascade: {len(cascade_added)} parent assembly(ies) also flagged "
                    f"for new numbers (a child changed):\n")
                for v in sorted(cascade_added.values(), key=lambda e: e["depth"], reverse=True):
                    log(f"    ↑ {v['old_number']}  [{v['desc']}]  ← child changed: {v.get('_cascade_seed','?')}\n")
        except Exception as e:
            log(f"  Cascade warning: {e}\n")

        # Give each Front Panel ASSEMBLY the same material as its Front Panel
        # part. Walk up from every Front Panel part (a part with material
        # options, i.e. target_material set) and stamp fp_material on each
        # ancestor assembly whose own description carries an FP tag — those are
        # the "front panel assemblies". append_to_master_bom uses it for the row.
        try:
            asm_by_path = {e["old_path"].lower(): e for e in all_entries if e["type"] == "assembly"}
            for pe in [e for e in all_entries if e["type"] == "part" and e.get("target_material")
                       and normalize_abbrev(e.get("abbrev")) == "front panel"]:
                changed = bool(pe.get("apply_material"))   # True only if the material was actually changed
                fp_mat = (pe["target_material"] if changed
                          else (pe.get("current_material") or pe.get("target_material")))
                if not fp_mat:
                    continue
                p = node_info.get(pe["old_path"].lower(), {}).get("parent", "")
                while p and p.lower() != main_path.lower():
                    pk = p.lower()
                    ae = asm_by_path.get(pk)
                    ninfo = node_info.get(pk, {})
                    adesc = (ae.get("desc") if ae else ninfo.get("desc", "")) or ""
                    if ae is not None and extract_front_panel(adesc):
                        if not ae.get("fp_material"):
                            ae["fp_material"] = fp_mat
                            log(f"    FP-material: {ae.get('old_number','?')} → {fp_mat}\n")
                        # A descendant Front Panel part's material actually CHANGED → record
                        # it so the Updates column shows "Material (...)" for this assembly
                        # (it's getting a new number for that reason, not an equation).
                        if changed and not ae.get("_mat_changed_name"):
                            ae["_mat_changed_name"] = pe.get("target_material") or fp_mat
                    p = ninfo.get("parent", "")
        except Exception as e:
            log(f"  FP-material propagation warning: {e}\n")

        # Same idea for the OTHER material part types (toekick / end cover /
        # Case Top / Interior): when such a part's material actually CHANGED,
        # stamp EVERY ancestor assembly up to (not including) the main assembly,
        # so each one's Updates column reads "Material" instead of "—". Those
        # ancestors are all getting a new number (via the cascade) because this
        # part changed, so the column should say so — mirroring the front panel.
        # An assembly's own description need NOT match the part's (a Frame Kit
        # can contain a Deck Pan Support); the label omits the gauge code for
        # assemblies anyway, since one assembly may span several gauges.
        try:
            asm_by_path = {e["old_path"].lower(): e for e in all_entries if e["type"] == "assembly"}
            for pe in [e for e in all_entries if e["type"] == "part"
                       and e.get("apply_material") and e.get("target_material")
                       and normalize_abbrev(e.get("abbrev")) != "front panel"]:
                tgt = pe["target_material"]
                p = node_info.get(pe["old_path"].lower(), {}).get("parent", "")
                while p and p.lower() != main_path.lower():
                    pk = p.lower()
                    ae = asm_by_path.get(pk)
                    ninfo = node_info.get(pk, {})
                    if ae is not None and not ae.get("_mat_changed_name"):
                        ae["_mat_changed_name"] = tgt
                    p = ninfo.get("parent", "")
        except Exception as e:
            log(f"  material-display propagation warning: {e}\n")

        # For each cascade-flagged assembly, mirror the PART replacement check:
        # look it up in the Master BOM by description + applicable dimensions +
        # (FP) material + variant tiebreak. If a DIFFERENT existing assembly
        # qualifies, replace with it instead of generating a new number. If only
        # the same number matches (old children) or nothing matches, keep it as
        # needs_update. Dimension values mirror the scan: selected value if the
        # user selected it, else the main assembly's current equation value.
        try:
            _folder = os.path.dirname(main_path)
            for e in [x for x in all_entries if x.get("_cascade")]:
                _desc = e.get("desc", "")
                _grow = _guideline_row(lookup, _desc) or {}
                _bn, _bv = [], []
                for p_name, _ in PARAMS:
                    if not _grow.get(p_name, False):
                        continue
                    if p_name in names:
                        _bn.append(p_name); _bv.append(values[names.index(p_name)])
                    elif p_name in cur_eq_vals:
                        _bn.append(p_name); _bv.append(cur_eq_vals[p_name])
                _req = e.get("fp_material") or material_for_description(lookup, _desc)
                _hit = find_in_master_bom(_desc, _bn, _bv, True, parts_rows, assy_rows,
                                          assy_config, required_material=_req)
                if _hit and _hit.upper() != (e.get("old_number") or "").upper():
                    _ext = os.path.join(_folder, _hit + ".SLDASM")
                    _on_disk = os.path.exists(_ext)
                    e["existing_number"] = _hit
                    e["new_number"]      = _hit
                    e["status"]          = "existing" if _on_disk else "file_missing"
                    log(f"    Cascade assy {e.get('old_number')}: BOM "
                        f"{'FOUND' if _on_disk else 'MISSING'} {_hit}"
                        f"{' (mat '+_req+')' if _req else ''} → replace\n")
                else:
                    log(f"    Cascade assy {e.get('old_number')}: no BOM match → new number\n")
        except Exception as e:
            log(f"  Cascade BOM-check warning: {e}\n")

        log(f"\n  Found {len(all_entries)} components to process\n")
        log(f"  → {sum(1 for e in all_entries if e['status']=='existing')} existing (ready to replace)\n")
        log(f"  → {sum(1 for e in all_entries if e['status']=='file_missing')} file missing\n")
        log(f"  → {sum(1 for e in all_entries if e['status']=='needs_update')} need new numbers\n")
        log("="*52+"\n")
        log("  PHASE 1 complete — showing replacement page\n")
        _clear_btn_busy(); _finish_conn(0)
        run_btn.configure(state="normal",fg_color=C["accent"],text_color="#ffffff",text="▶ Run Update")
        set_status("Review replacements →",C["success"])
        update_done[0]=True
        view_results_btn.place(relx=1.0,x=-130,rely=0.5,anchor="e")
        timer_pause("Waiting for user...")
        _busy[0]=False
        last_results[0]="replace"; build_replace_panel(); show_panel(results_panel)

    def reset():
        _clear_btn_busy(); _reset_conns()
        summary_generated[0] = False
        qc_answers.clear()
        qc_measure_started.clear()
        qc_macro_complete[0] = False
        qc_report_path[0] = ""
        qc_record_path[0] = ""
        qc_completed[0] = False
        qc_question_index[0] = 0
        for name,_ in PARAMS:
            check_vars[name].set(False); entry_vars[name].set("")
            entry_widgets[name].configure(state="disabled",fg_color=C["card2"],text_color=C["text_muted"],border_color=C["border"])
            row_frames[name].configure(fg_color=C["card"],border_width=1,border_color=C["border_light"])
        output.configure(state="normal"); output.delete("1.0","end"); output.configure(state="disabled")
        # Hide new_bar, restore layout to default (no new_bar)
        new_bar.grid_remove()
        orig_bar.grid(row=5,column=0,sticky="new",padx=(0,4))
        all_entries.clear(); orig_vals_snap.clear(); new_vals_snap.clear()
        stored_sel[0]=[]; stored_sel[1]=[]
        # Reset every material radio back to its default (Painted) and repaint the
        # cards — otherwise a finish chosen for the previous model stays selected.
        for _ab,_rv in material_radio_vars.items():
            try:
                _opts=MATERIAL_CHOICES.get(_ab) or []
                if _opts: _rv.set(_opts[0][1])
            except Exception: pass
        for _rf in material_refreshers.values():
            try: _rf()
            except Exception: pass
        # Reset the base-frame configuration selectors to their defaults.
        try:
            bf_lp_var.set("nolp"); bf_ref_var.set("DX")
            _refresh_cfg_selectors()
        except Exception: pass
        # Reset the step indicators / navigation so the 1-2-3 ticks clear and the
        # Replace step can run again (a stale nav_done['replace'] disables Confirm
        # & Replace, which is what blocked progress with '0 Parts Replaced').
        nav_enabled.update({"input":True,"log":True,"replace":False,"newnum":False,"summary":False})
        nav_done.update({"input":False,"log":False,"replace":False,"newnum":False,"summary":False})
        try: refresh_nav()
        except Exception: pass
        timer_reset(); last_results[0]="replace"
        update_done[0]=False; gen_done[0]=False; view_results_btn.place_forget()
        show_panel(input_panel); set_status("Ready to update",C["success"])

    def reconnect():
        """Re-connect to whatever assembly is currently open in SolidWorks,
        re-detect the product line, refresh original values, clear previous
        results and unlock the inputs — lets the user switch models without
        restarting the app."""
        if _busy[0]:
            messagebox.showinfo("Please wait",
                "An update is currently running. Let it finish before connecting to a different model.")
            return
        nonlocal sw_app, sw_model, main_path, assy_title, current_vals
        try:
            new_app=win32com.client.Dispatch("SldWorks.Application")
            doc=new_app.ActiveDoc
        except Exception as e:
            messagebox.showerror("Connect to Model", f"Could not reach SolidWorks:\n{e}")
            return
        if doc is None:
            messagebox.showwarning("Connect to Model", "No active document in SolidWorks.\nOpen an assembly and try again.")
            return
        try:   doc_type=int(doc.GetType())
        except TypeError: doc_type=int(doc.GetType)
        except Exception: doc_type=-1
        if doc_type != 2:
            messagebox.showwarning("Connect to Model", "The active SolidWorks document is not an assembly.")
            return
        # Read description and re-detect product line
        try:   new_desc=get_description(doc) or ""
        except Exception: new_desc=""
        ndu=new_desc.strip().upper()
        if ndu.startswith("PFI-"):  picked="PFI"
        elif ndu.startswith("PF-"): picked="PF"
        elif ndu.startswith("PT-"): picked="PT"
        else:                       picked=None
        if picked is None:
            messagebox.showerror("Product Line Not Detected",
                "Could not determine the product line from the open assembly.\n\n"
                "The Description custom property must start with 'PF-', 'PFI-' or 'PT-'.\n"
                f"Got: '{new_desc}'")
            return
        # Commit new connection state
        sw_app=new_app
        sw_model=doc
        try:   assy_title=doc.GetTitle()
        except TypeError: assy_title=doc.GetTitle
        try:   main_path=doc.GetPathName()
        except TypeError: main_path=doc.GetPathName
        cfg=PRODUCT_LINES[picked]
        ACTIVE_LINE["key"]=picked; ACTIVE_LINE["cfg"]=cfg
        ACTIVE_LINE["bom_path"]=os.path.join(_tool_dir(), cfg["bom_filename"])
        # Refresh original values from the new model
        current_vals=get_current_equation_values(sw_model) if sw_model else {}
        # Update header/info UI
        try:
            title_lbl.configure(text=assy_title)
            _dshow=new_desc if len(new_desc)<=48 else new_desc[:46]+"..."
            desc_lbl.configure(text=_dshow)
            conn_pill.configure(text="● Connected",fg_color=C["success_bg"],text_color=C["success"])
        except Exception: pass
        # Refresh the ORIGINAL VALUES bar value labels in place
        for _nm,_lbl in orig_bar_labels.items():
            try:
                _v=format_value(current_vals.get(_nm,"— "))
                _disp=f"{_v} in" if _v!="— " else "— "
                _lbl.configure(text=f"{PARAM_SHORT[_nm]}\n{_disp}")
            except Exception: pass
        # Unlock inputs and clear previous results
        set_inputs_locked(False)
        reset()
        set_status(f"Connected to {assy_title}", C["success"])

    connect_btn.configure(command=reconnect)

    run_btn=ctk.CTkButton(bf,text="\u25b6  Run Update",
                          font=ctk.CTkFont(size=13,weight="bold"),
                          fg_color=C["accent"],hover_color=C["accent_hover"],
                          text_color="#ffffff",corner_radius=18,height=44,width=176,
                          command=run_update)
    run_btn.grid(row=0,column=1,sticky="e",padx=(8,4),pady=10)
    ctk.CTkButton(bf,text="\u21ba  Reset",
                  font=ctk.CTkFont(size=13,weight="bold"),
                  fg_color=C["card"],hover_color=C["card2"],
                  text_color=C["text_muted"],border_width=1,
                  border_color=C["border"],corner_radius=18,height=44,width=124,
                  command=reset).grid(row=0,column=0,sticky="w",padx=(4,8),pady=10)

    show_panel(input_panel); root.mainloop()

if __name__ == "__main__":
    build_app()
