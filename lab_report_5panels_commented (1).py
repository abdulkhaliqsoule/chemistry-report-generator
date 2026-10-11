"""
==============================================================================
LAB REPORT GENERATOR — 5 PANEL VERSION (FULLY COMMENTED)
==============================================================================
This script creates a professional PDF lab report with 5 colour-coded sections:
  1. Liver Function Tests (LFT)     — Green
  2. Kidney Function Tests (KFT)    — Blue
  3. Lipid Profile                  — Purple
  4. Glycated Haemoglobin (HbA1c)  — Orange
  5. Uric Acid                      — Teal

HOW TO RUN IT (type this in your terminal / PowerShell):
  With demo data:
    python lab_report_5panels_commented.py --demo

  With a real PDF from your analyser:
    python lab_report_5panels_commented.py "patient.pdf" "My Lab Name"

  With all options:
    python lab_report_5panels_commented.py "patient.pdf" "My Lab Name"
        --address "123 Street, Accra"
        --phone "+233 XX XXX XXXX"
        --director "Bio. Mensah"
        --output "report.pdf"
==============================================================================
"""


# ==============================================================================
# SECTION 1 — IMPORTS
# "Importing" means borrowing tools that other people already wrote.
# Think of it like getting equipment out of a toolbox before starting a job.
# ==============================================================================

import argparse

# ── GUI imports (tkinter + drag-and-drop) ─────────────────────────────────────
# These are wrapped in a try/except so the script can still be imported or run
# in CLI mode on systems where tkinter is not installed (e.g. minimal Linux servers).
# On Windows and macOS, tkinter is always available.
import threading    # Runs PDF generation in the background so the GUI stays responsive
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter import font as tkfont
    TK_AVAILABLE = True
except ImportError:
    TK_AVAILABLE = False
    tk = None

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    TkinterDnD = None

# argparse = the tool that reads what you type after "python script.py ..."
# For example it reads --demo, --output, "patient.pdf" etc.
# Without this, the script would have no way to accept instructions from you.

from datetime import datetime
# datetime = gives us access to the current date and time.
# We use it to stamp "Report generated: 2024-06-10 21:07" at the bottom.

from reportlab.lib.pagesizes import letter
# reportlab = the big PDF-creation library we rely on for everything visual.
# "letter" = the standard US paper size (8.5 inches wide × 11 inches tall).
# If you wanted A4 instead, you'd write: from reportlab.lib.pagesizes import A4

from reportlab.lib import colors
# colors = gives us named colours and hex colour codes.
# colors.black = pure black
# colors.white = pure white
# colors.HexColor("#cc0000") = a custom red using a 6-digit hex code
# (Hex codes are how computers describe colours — #cc0000 means "dark red")

from reportlab.lib.units import inch
# inch = a measurement unit so we can write "0.6*inch" instead of raw numbers.
# ReportLab works in "points" internally (72 points = 1 inch).
# Using inch makes our margin values human-readable.

from reportlab.lib.styles import ParagraphStyle
# ParagraphStyle = lets us define reusable text styles (like Styles in Word).
# We create styles like "big bold black" or "small grey" once and reuse them.

from reportlab.platypus import (
    Paragraph,    # A block of text with a style applied to it
    Spacer,       # An invisible gap — like pressing Enter to add blank space
    Table,        # A grid of cells — like a mini spreadsheet inside the PDF
    TableStyle,   # Rules that control how a Table looks (colours, borders, padding)
    HRFlowable,   # A horizontal line drawn across the page (like <hr> in HTML)
    KeepTogether  # Tries to keep a group of elements on the same page (prevents orphans)
)
# "platypus" is ReportLab's layout engine.
# It takes a list of Paragraphs, Tables, Spacers etc. and arranges them
# top-to-bottom across pages automatically — like a word processor.

from reportlab.lib.enums import TA_LEFT, TA_RIGHT
# These are text alignment options:
# TA_LEFT  = text starts from the left edge (normal reading direction)
# TA_RIGHT = text is pushed to the right edge (useful for page numbers etc.)

from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
# BaseDocTemplate = a powerful document type that lets us customise every page.
#   (We extend this to add our own running header and footer.)
# Frame = defines the rectangular "content box" on the page where text goes.
#   (Everything inside the margins sits inside a Frame.)
# PageTemplate = tells ReportLab what a page looks like and what to do each time
#   a new page starts (e.g. draw our header/footer automatically).


# ==============================================================================
# SECTION 2 — LAB CONFIGURATION
# These are the details about YOUR lab that appear at the top of every report.
# ✏️  EDIT THESE FOUR LINES to match your actual lab before using for real patients.
# ==============================================================================

LAB_NAME     = "Your Laboratory Name"           # e.g. "City Hospital Lab"
LAB_ADDRESS  = "123 Lab Street, City, State 00000"  # e.g. "123 High Street, Accra"
LAB_PHONE    = "+1 (000) 000-0000"              # e.g. "+233 20 000 0000"
LAB_DIRECTOR = "Laboratory Scientist"            # e.g. "Bio. Ama Owusu"


# ==============================================================================
# SECTION 3 — PANEL DEFINITIONS
# A "panel" is one coloured section of the report (e.g. Liver Function Tests).
# Each panel has:
#   "color"    = the dark colour used for the heading banner
#   "light"    = a pale version of that colour for zebra-striping the rows
#   "keywords" = a list of words the script looks for in test names
#                to decide which panel a test belongs to
#
# HOW THE AUTO-SORTING WORKS:
#   When a test result comes in (e.g. "ALT (SGPT)"), the script converts
#   the name to lowercase ("alt (sgpt)") and checks if any keyword is
#   contained inside it. "alt" IS inside "alt (sgpt)", so it goes to LFT.
#   If no keyword matches, the test goes to the "Other" section at the bottom.
#
# ✏️  You can add keywords here if your analyser uses different test names.
# ==============================================================================

PANEL_DEFINITIONS = {

    # ── Panel 1: Liver Function Tests ────────────────────────────────────────
    "Liver Function Tests (LFT)": {
        "color":    colors.HexColor("#1a6b3c"),
        "light":    colors.HexColor("#e8f5ee"),
        "keywords": [
            # Short BS240 codes (Layout B)
            "t-bil", "d-bil", "ibil", "ind bill",
            # Standard names and variants
            "alt", "alanine", "ast", "aspartate",
            "alp", "alkaline",
            "ggt", "gamma", "γ-gt",
            "bilirubin", "albumin", "globulin",
            "total protein", "tp-o",
            "a/g ratio", "prothrombin",
            "sgot", "sgpt",
        ],
    },

    # ── Panel 2: Kidney Function Tests ───────────────────────────────────────
    "Kidney Function Tests (KFT)": {
        "color":    colors.HexColor("#1a3c5e"),
        "light":    colors.HexColor("#dce8f5"),
        "keywords": [
            # Short BS240 codes (Layout B)
            "crea", "k+", "na+", "cl-",
            # Standard names
            "urea", "bun", "creatinine", "egfr",
            "sodium", "potassium", "chloride",
            "carbon dioxide", "bicarbonate",
            "calcium", "phosphate", "magnesium",
        ],
    },

    # ── Panel 3: Lipid Profile ────────────────────────────────────────────────
    "Lipid Profile": {
        "color":    colors.HexColor("#7b2d8b"),
        "light":    colors.HexColor("#f3e8f7"),
        "keywords": [
            # Short BS240 codes (Layout B)
            "tc", "tg", "hdl-c", "ldl-c", "ldl", "vldl",
            "coronary r", "coronary risk",
            # Standard names
            "cholesterol", "triglyceride",
            "hdl", "lipid", "ldl/hdl", "chol/hdl", "non-hdl",
        ],
    },

    # ── Panel 4: Glycated Haemoglobin ─────────────────────────────────────────
    "Glycated Haemoglobin (HbA1c)": {
        "color":    colors.HexColor("#b85c00"),
        "light":    colors.HexColor("#fdf0e0"),
        "keywords": [
            # The ONLY code we want from HbA1c tests is A1c(NGSP)
            # Other HbA1c sub-fractions (A1c IFCC, eAG etc.) are excluded
            "a1c(ngsp)", "a1c (ngsp)",
            "hemoglobin a1c", "haemoglobin a1c",
            "hba1c", "hb a1c",
            "glycated", "glycosylated",
        ],
    },

    # ── Panel 5: Uric Acid ────────────────────────────────────────────────────
    "Uric Acid": {
        "color":    colors.HexColor("#006b6b"),
        "light":    colors.HexColor("#e0f5f5"),
        "keywords": [
            "uric acid", "urate", "gout",
        ],
    },
}

# Panel display order (top to bottom in the report)
PANEL_ORDER = [
    "Liver Function Tests (LFT)",
    "Kidney Function Tests (KFT)",
    "Lipid Profile",
    "Glycated Haemoglobin (HbA1c)",
    "Uric Acid",
    "Other",
]


# ==============================================================================
# SECTION 3B  TEST_CODE_MAP
# Maps every Mindray BS240 short code to (Full Display Name, Panel Name).
# Format:  "short_code_lowercase": ("Full Name", "Panel or empty string")
# "__SKIP__" removes a row entirely (used for HbA1c IFCC / eAG sub-fractions).
# ==============================================================================

TEST_CODE_MAP = {
    # Liver Function Tests
    "t-bil":               ("Total Bilirubin",                   "Liver Function Tests (LFT)"),
    "tbil":                ("Total Bilirubin",                   "Liver Function Tests (LFT)"),
    "tb":                  ("Total Bilirubin",                   "Liver Function Tests (LFT)"),
    "total bilirubin":     ("Total Bilirubin",                   "Liver Function Tests (LFT)"),
    "bilirubin, total":    ("Total Bilirubin",                   "Liver Function Tests (LFT)"),
    "d-bil":               ("Direct Bilirubin",                  "Liver Function Tests (LFT)"),
    "dbil":                ("Direct Bilirubin",                  "Liver Function Tests (LFT)"),
    "direct bil":          ("Direct Bilirubin",                  "Liver Function Tests (LFT)"),
    "direct bilirubin":    ("Direct Bilirubin",                  "Liver Function Tests (LFT)"),
    "ibil":                ("Indirect Bilirubin",                "Liver Function Tests (LFT)"),
    "ind bill":            ("Indirect Bilirubin",                "Liver Function Tests (LFT)"),
    "ind bil":             ("Indirect Bilirubin",                "Liver Function Tests (LFT)"),
    "indirect bilirubin":  ("Indirect Bilirubin",                "Liver Function Tests (LFT)"),
    "γ-gt":           ("Gamma GT (γ-Glutamyl Transferase)", "Liver Function Tests (LFT)"),
    "y-gt":                ("Gamma GT (γ-Glutamyl Transferase)", "Liver Function Tests (LFT)"),
    "ggt":                 ("Gamma GT (γ-Glutamyl Transferase)", "Liver Function Tests (LFT)"),
    "gamma gt":            ("Gamma GT (γ-Glutamyl Transferase)", "Liver Function Tests (LFT)"),
    "ggtp":                ("Gamma GT (γ-Glutamyl Transferase)", "Liver Function Tests (LFT)"),
    "alb":                 ("Albumin",                           "Liver Function Tests (LFT)"),
    "albumin":             ("Albumin",                           "Liver Function Tests (LFT)"),
    "tp-o":                ("Total Protein",                     "Liver Function Tests (LFT)"),
    "tp":                  ("Total Protein",                     "Liver Function Tests (LFT)"),
    "total protein":       ("Total Protein",                     "Liver Function Tests (LFT)"),
    "protein, total":      ("Total Protein",                     "Liver Function Tests (LFT)"),
    "glob":                ("Globulin",                          "Liver Function Tests (LFT)"),
    "globulin":            ("Globulin",                          "Liver Function Tests (LFT)"),
    "alt":                 ("ALT (Alanine Aminotransferase)",    "Liver Function Tests (LFT)"),
    "alt (sgpt)":          ("ALT (Alanine Aminotransferase)",    "Liver Function Tests (LFT)"),
    "sgpt":                ("ALT (Alanine Aminotransferase)",    "Liver Function Tests (LFT)"),
    "alanine aminotrans":  ("ALT (Alanine Aminotransferase)",    "Liver Function Tests (LFT)"),
    "ast":                 ("AST (Aspartate Aminotransferase)",  "Liver Function Tests (LFT)"),
    "ast (sgot)":          ("AST (Aspartate Aminotransferase)",  "Liver Function Tests (LFT)"),
    "sgot":                ("AST (Aspartate Aminotransferase)",  "Liver Function Tests (LFT)"),
    "alp":                 ("ALP (Alkaline Phosphatase)",        "Liver Function Tests (LFT)"),
    "alkaline phosphatase":("ALP (Alkaline Phosphatase)",        "Liver Function Tests (LFT)"),
    "alkaline phos.":      ("ALP (Alkaline Phosphatase)",        "Liver Function Tests (LFT)"),
    "ast/alt":             ("AST/ALT Ratio",                     "Liver Function Tests (LFT)"),
    "a/g ratio":           ("Albumin/Globulin Ratio",            "Liver Function Tests (LFT)"),
    # Lipid Profile
    "tc":                  ("Total Cholesterol",                 "Lipid Profile"),
    "total chol.":         ("Total Cholesterol",                 "Lipid Profile"),
    "total chol":          ("Total Cholesterol",                 "Lipid Profile"),
    "cholesterol, total":  ("Total Cholesterol",                 "Lipid Profile"),
    "tg":                  ("Triglycerides",                     "Lipid Profile"),
    "triglycerides":       ("Triglycerides",                     "Lipid Profile"),
    "coronary r":          ("Coronary Risk",                     "Lipid Profile"),
    "coronary risk":       ("Coronary Risk",                     "Lipid Profile"),
    "hdl-c":               ("HDL Cholesterol",                   "Lipid Profile"),
    "hdl":                 ("HDL Cholesterol",                   "Lipid Profile"),
    "hdl-chol.":           ("HDL Cholesterol",                   "Lipid Profile"),
    "ldl":                 ("LDL Cholesterol",                   "Lipid Profile"),
    "ldl-c":               ("LDL Cholesterol",                   "Lipid Profile"),
    "ldl chol calc (nih)": ("LDL Cholesterol",                   "Lipid Profile"),
    "vldl":                ("VLDL Cholesterol",                  "Lipid Profile"),
    "ldl/hdl":             ("LDL/HDL Ratio",                     "Lipid Profile"),
    "chol/hdl":            ("Cholesterol/HDL Ratio",             "Lipid Profile"),
    "non-hdl":             ("Non-HDL Cholesterol",               "Lipid Profile"),
    # Kidney Function Tests
    "crea":                ("Creatinine",                        "Kidney Function Tests (KFT)"),
    "creatinine":          ("Creatinine",                        "Kidney Function Tests (KFT)"),
    "urea":                ("Urea",                              "Kidney Function Tests (KFT)"),
    "bun":                 ("Blood Urea Nitrogen (BUN)",         "Kidney Function Tests (KFT)"),
    "egfr":                ("eGFR (Estimated GFR)",              "Kidney Function Tests (KFT)"),
    "k+":                  ("Potassium (K<super>+</super>)",               "Kidney Function Tests (KFT)"),
    "potassium":           ("Potassium (K<super>+</super>)",               "Kidney Function Tests (KFT)"),
    "cl-":                 ("Chloride (Cl<super>-</super>)",               "Kidney Function Tests (KFT)"),
    "chloride":            ("Chloride (Cl<super>-</super>)",               "Kidney Function Tests (KFT)"),
    "na+":                 ("Sodium (Na<super>+</super>)",                 "Kidney Function Tests (KFT)"),
    "sodium":              ("Sodium (Na<super>+</super>)",                 "Kidney Function Tests (KFT)"),
    "bicarbonate":         ("Bicarbonate (HCO3<super>-</super>)",     "Kidney Function Tests (KFT)"),
    "calcium":             ("Calcium",                           "Kidney Function Tests (KFT)"),
    "phosphate":           ("Phosphate",                         "Kidney Function Tests (KFT)"),
    "magnesium":           ("Magnesium",                         "Kidney Function Tests (KFT)"),
    # Glycated Haemoglobin — ONLY NGSP is shown; all others are skipped
    "a1c(ngsp)":           ("HbA1c (NGSP %)",                   "Glycated Haemoglobin (HbA1c)"),
    "a1c (ngsp)":          ("HbA1c (NGSP %)",                   "Glycated Haemoglobin (HbA1c)"),
    "ngsp":                ("HbA1c (NGSP %)",                   "Glycated Haemoglobin (HbA1c)"),
    "hba1c":               ("HbA1c (NGSP %)",                   "Glycated Haemoglobin (HbA1c)"),
    "hb a1c":              ("HbA1c (NGSP %)",                   "Glycated Haemoglobin (HbA1c)"),
    "hemoglobin a1c":      ("HbA1c (NGSP %)",                   "Glycated Haemoglobin (HbA1c)"),
    "haemoglobin a1c":     ("HbA1c (NGSP %)",                   "Glycated Haemoglobin (HbA1c)"),
    "glycated":            ("HbA1c (NGSP %)",                   "Glycated Haemoglobin (HbA1c)"),
    "a1c(ifcc)":           ("__SKIP__",                         ""),
    "a1c (ifcc)":          ("__SKIP__",                         ""),
    "ifcc":                ("__SKIP__",                         ""),
    "eag":                 ("__SKIP__",                         ""),
    "estimated average glucose": ("__SKIP__",                   ""),
    # Uric Acid
    "uric acid":           ("Uric Acid",                         "Uric Acid"),
    "urate":               ("Uric Acid",                         "Uric Acid"),
    "ua":                  ("Uric Acid",                         "Uric Acid"),
}


def resolve_test_name(raw_name):
    key = raw_name.strip().lower()
    # Handle "CODE — Full Name" format from table parser
    if " — " in key:
        code_key = key.split(" — ")[0].strip()
        if code_key in TEST_CODE_MAP:
            return TEST_CODE_MAP[code_key]
    if key in TEST_CODE_MAP:
        return TEST_CODE_MAP[key]
    return (raw_name.strip(), "")


def apply_test_name_map(results):
    resolved = []
    for r in results:
        full_name, panel_override = resolve_test_name(r["test"])
        if full_name == "__SKIP__":
            continue
        new_r = dict(r)
        new_r["test"] = full_name
        new_r["panel_override"] = panel_override
        resolved.append(new_r)
    return resolved



# ==============================================================================
# SECTION 3B — TEST CODE MAP  &  NAME RESOLUTION
#
# The Mindray BS240 prints short codes (e.g. "T-bil", "CREA", "γ-GT") that are
# not immediately clear to patients or clinicians reading the report.
# This dictionary maps every known BS240 short code to its full clinical name
# and the panel it belongs to.
#
# FORMAT:
#   "short_code_lowercase": ("Full Display Name", "Panel Name or ''")
#
# The second value is a PANEL OVERRIDE — it forces the result into the named
# panel, bypassing the keyword matcher.  Set to "" to let keywords decide.
#
# SPECIAL VALUE __SKIP__:
#   If the full name is "__SKIP__", the row is silently removed from the report.
#   Used to hide the HbA1c IFCC and eAG sub-rows (we only show NGSP %).
#
# ✏️  Add new codes here whenever your analyser introduces a new abbreviation.
# ==============================================================================

TEST_CODE_MAP = {

    # ── Liver Function Tests ───────────────────────────────────────────────────
    "t-bil":              ("Total Bilirubin",                    "Liver Function Tests (LFT)"),
    "tbil":               ("Total Bilirubin",                    "Liver Function Tests (LFT)"),
    "total bilirubin":    ("Total Bilirubin",                    "Liver Function Tests (LFT)"),
    "bilirubin, total":   ("Total Bilirubin",                    "Liver Function Tests (LFT)"),
    "d-bil":              ("Direct Bilirubin",                   "Liver Function Tests (LFT)"),
    "dbil":               ("Direct Bilirubin",                   "Liver Function Tests (LFT)"),
    "direct bil":         ("Direct Bilirubin",                   "Liver Function Tests (LFT)"),
    "direct bilirubin":   ("Direct Bilirubin",                   "Liver Function Tests (LFT)"),
    "ibil":               ("Indirect Bilirubin",                 "Liver Function Tests (LFT)"),
    "ind bill":           ("Indirect Bilirubin",                 "Liver Function Tests (LFT)"),
    "ind bil":            ("Indirect Bilirubin",                 "Liver Function Tests (LFT)"),
    "indirect bilirubin": ("Indirect Bilirubin",                 "Liver Function Tests (LFT)"),
    "γ-gt":               ("Gamma GT (γ-GT)",                    "Liver Function Tests (LFT)"),
    "y-gt":               ("Gamma GT (γ-GT)",                    "Liver Function Tests (LFT)"),
    "ggt":                ("Gamma GT (γ-GT)",                    "Liver Function Tests (LFT)"),
    "gamma gt":           ("Gamma GT (γ-GT)",                    "Liver Function Tests (LFT)"),
    "ggtp":               ("Gamma GT (γ-GT)",                    "Liver Function Tests (LFT)"),
    "alb":                ("Albumin",                            "Liver Function Tests (LFT)"),
    "albumin":            ("Albumin",                            "Liver Function Tests (LFT)"),
    "tp-o":               ("Total Protein",                      "Liver Function Tests (LFT)"),
    "tp":                 ("Total Protein",                      "Liver Function Tests (LFT)"),
    "total protein":      ("Total Protein",                      "Liver Function Tests (LFT)"),
    "protein, total":     ("Total Protein",                      "Liver Function Tests (LFT)"),
    "glob":               ("Globulin",                           "Liver Function Tests (LFT)"),
    "globulin":           ("Globulin",                           "Liver Function Tests (LFT)"),
    "globulin, total":    ("Globulin",                           "Liver Function Tests (LFT)"),
    "alt":                ("ALT (Alanine Aminotransferase)",     "Liver Function Tests (LFT)"),
    "alt (sgpt)":         ("ALT (Alanine Aminotransferase)",     "Liver Function Tests (LFT)"),
    "sgpt":               ("ALT (Alanine Aminotransferase)",     "Liver Function Tests (LFT)"),
    "alanine aminotrans": ("ALT (Alanine Aminotransferase)",     "Liver Function Tests (LFT)"),
    "ast":                ("AST (Aspartate Aminotransferase)",   "Liver Function Tests (LFT)"),
    "ast (sgot)":         ("AST (Aspartate Aminotransferase)",   "Liver Function Tests (LFT)"),
    "sgot":               ("AST (Aspartate Aminotransferase)",   "Liver Function Tests (LFT)"),
    "alp":                ("ALP (Alkaline Phosphatase)",         "Liver Function Tests (LFT)"),
    "alkaline phosphatase":("ALP (Alkaline Phosphatase)",        "Liver Function Tests (LFT)"),
    "alkaline phos.":     ("ALP (Alkaline Phosphatase)",         "Liver Function Tests (LFT)"),
    "ast/alt":            ("AST/ALT Ratio",                      "Liver Function Tests (LFT)"),
    "a/g ratio":          ("Albumin/Globulin (A/G) Ratio",       "Liver Function Tests (LFT)"),

    # ── Lipid Profile ──────────────────────────────────────────────────────────
    "tc":                 ("Total Cholesterol",                  "Lipid Profile"),
    "total chol.":        ("Total Cholesterol",                  "Lipid Profile"),
    "total chol":         ("Total Cholesterol",                  "Lipid Profile"),
    "cholesterol, total": ("Total Cholesterol",                  "Lipid Profile"),
    "tg":                 ("Triglycerides",                      "Lipid Profile"),
    "triglycerides":      ("Triglycerides",                      "Lipid Profile"),
    "coronary r":         ("Coronary Risk",                      "Lipid Profile"),
    "coronary risk":      ("Coronary Risk",                      "Lipid Profile"),
    "hdl-c":              ("HDL Cholesterol",                    "Lipid Profile"),
    "hdl":                ("HDL Cholesterol",                    "Lipid Profile"),
    "hdl-chol.":          ("HDL Cholesterol",                    "Lipid Profile"),
    "hdl cholesterol":    ("HDL Cholesterol",                    "Lipid Profile"),
    "ldl":                ("LDL Cholesterol",                    "Lipid Profile"),
    "ldl-c":              ("LDL Cholesterol",                    "Lipid Profile"),
    "ldl chol calc (nih)":("LDL Cholesterol",                    "Lipid Profile"),
    "vldl":               ("VLDL Cholesterol",                   "Lipid Profile"),
    "vldl cholesterol cal":("VLDL Cholesterol",                  "Lipid Profile"),
    "ldl/hdl":            ("LDL/HDL Ratio",                      "Lipid Profile"),
    "chol/hdl":           ("Cholesterol/HDL Ratio",              "Lipid Profile"),
    "non-hdl":            ("Non-HDL Cholesterol",                "Lipid Profile"),

    # ── Kidney Function Tests ──────────────────────────────────────────────────
    "crea":               ("Creatinine",                         "Kidney Function Tests (KFT)"),
    "creatinine":         ("Creatinine",                         "Kidney Function Tests (KFT)"),
    "urea":               ("Urea",                               "Kidney Function Tests (KFT)"),
    "bun":                ("Blood Urea Nitrogen (BUN)",          "Kidney Function Tests (KFT)"),
    "bun/creatinine ratio":("BUN/Creatinine Ratio",              "Kidney Function Tests (KFT)"),
    "egfr":               ("eGFR",                               "Kidney Function Tests (KFT)"),
    "estimated gfr":      ("eGFR",                               "Kidney Function Tests (KFT)"),
    "k+":                 ("Potassium (K+)",                     "Kidney Function Tests (KFT)"),
    "potassium":          ("Potassium (K+)",                     "Kidney Function Tests (KFT)"),
    "cl-":                ("Chloride (Cl-)",                     "Kidney Function Tests (KFT)"),
    "chloride":           ("Chloride (Cl-)",                     "Kidney Function Tests (KFT)"),
    "na+":                ("Sodium (Na+)",                       "Kidney Function Tests (KFT)"),
    "sodium":             ("Sodium (Na+)",                       "Kidney Function Tests (KFT)"),
    "carbon dioxide, total":("Carbon Dioxide (CO2)",             "Kidney Function Tests (KFT)"),
    "bicarbonate":        ("Bicarbonate (HCO3-)",                "Kidney Function Tests (KFT)"),
    "calcium":            ("Calcium",                            "Kidney Function Tests (KFT)"),
    "phosphate":          ("Phosphate",                          "Kidney Function Tests (KFT)"),
    "magnesium":          ("Magnesium",                          "Kidney Function Tests (KFT)"),

    # ── Glycated Haemoglobin — NGSP only, all others hidden ───────────────────
    "a1c(ngsp)":          ("HbA1c (NGSP %)",                    "Glycated Haemoglobin (HbA1c)"),
    "a1c (ngsp)":         ("HbA1c (NGSP %)",                    "Glycated Haemoglobin (HbA1c)"),
    "ngsp":               ("HbA1c (NGSP %)",                    "Glycated Haemoglobin (HbA1c)"),
    "hba1c":              ("HbA1c (NGSP %)",                    "Glycated Haemoglobin (HbA1c)"),
    "hb a1c":             ("HbA1c (NGSP %)",                    "Glycated Haemoglobin (HbA1c)"),
    "hemoglobin a1c":     ("HbA1c (NGSP %)",                    "Glycated Haemoglobin (HbA1c)"),
    "haemoglobin a1c":    ("HbA1c (NGSP %)",                    "Glycated Haemoglobin (HbA1c)"),
    "glycated":           ("HbA1c (NGSP %)",                    "Glycated Haemoglobin (HbA1c)"),
    "a1c(ifcc)":          ("__SKIP__", ""),   # IFCC row — not shown in report
    "a1c (ifcc)":         ("__SKIP__", ""),
    "ifcc":               ("__SKIP__", ""),
    "eag":                ("__SKIP__", ""),
    "estimated average glucose": ("__SKIP__", ""),

    # ── Uric Acid ──────────────────────────────────────────────────────────────
    "uric acid":          ("Uric Acid",                          "Uric Acid"),
    "urate":              ("Uric Acid",                          "Uric Acid"),
    "ua":                 ("Uric Acid",                          "Uric Acid"),
}


def resolve_test_name(raw_name):
    """
    Looks up a raw BS240 test code in TEST_CODE_MAP.
    Returns: (full_display_name, panel_override_or_empty_string)

    If the code is not in the map, the original name is returned unchanged.
    If the full_display_name is "__SKIP__", the caller must discard this result.
    """
    key = raw_name.strip().lower()
    if key in TEST_CODE_MAP:
        return TEST_CODE_MAP[key]
    return raw_name, ""


def apply_test_name_map(results):
    """
    Processes all results from parse_pdf_data():
      - Replaces short codes with full clinical names using TEST_CODE_MAP
      - Removes rows marked __SKIP__ (e.g. HbA1c IFCC, eAG sub-rows)
      - Adds "panel_override" key to each result for classify_test()

    Must be called BEFORE group_results_by_panel().
    Returns a new filtered list — does not modify the input.
    """
    resolved = []
    for r in results:
        full_name, panel_override = resolve_test_name(r["test"])
        if full_name == "__SKIP__":
            continue
        new_r = dict(r)
        new_r["test"]           = full_name
        new_r["panel_override"] = panel_override
        resolved.append(new_r)
    return resolved




# ==============================================================================
# SECTION 4 — classify_test() FUNCTION
# This function takes a test name and returns which panel it belongs to.
# It's called automatically for every test result before building the PDF.
# ==============================================================================

def classify_test(test_name, panel_override=""):
    """
    Given a test name, returns which panel it belongs to.

    Parameters:
      test_name      — the FULL display name (already resolved via TEST_CODE_MAP)
      panel_override — if non-empty, skip keyword matching and use this directly
                       (set by apply_test_name_map() from TEST_CODE_MAP entries)

    How it works:
      1. If panel_override is set (e.g. "Liver Function Tests (LFT)"), use it
      2. Otherwise convert test_name to lowercase and check panel keywords
      3. If no keyword matches, return "Other"
    """

    # ── Fast path: panel was explicitly set in TEST_CODE_MAP ──────────────────
    # This guarantees e.g. "Gamma GT" always goes to LFT even if the keyword
    # matcher wouldn't find it (because "gamma gt" isn't a keyword, "gamma" is).
    if panel_override and panel_override in PANEL_DEFINITIONS:
        return panel_override

    # ── Keyword matching path ─────────────────────────────────────────────────
    # Convert to lowercase so the comparison isn't case-sensitive
    # Convert to lowercase. Also strip " — Full Name" suffix from table-parser codes
    # e.g. "ALT — ALANINE AMINOTRANS" → use just "alt" for keyword matching
    raw = test_name.split(" — ")[0].strip() if " — " in test_name else test_name
    name_lower = raw.lower()

    for panel_name, panel_info in PANEL_DEFINITIONS.items():
        for keyword in panel_info["keywords"]:
            if keyword in name_lower:
                return panel_name

    # If we get here, no keyword matched any panel
    return "Other"


# ==============================================================================
# SECTION 5 — styles() FUNCTION
# Defines ALL the text styles used across the entire report.
# Think of this like the "Styles" panel in Microsoft Word — we define
# "heading1", "body text", "caption" etc. once, then apply them by name.
#
# This function returns a Python dictionary (a collection of labelled values).
# We store it in a variable called "s" and access styles like s["flag"] or s["pt_name"].
# ==============================================================================

def styles():

    # ── Step 1: Define our colour palette with friendly names ─────────────────
    # Instead of remembering "#cc0000" every time, we just write FLAG_RED.

    BLACK      = colors.black                  # Pure black — main body text
    DARK_GREY  = colors.HexColor("#333333")    # Very dark grey — labels and sub-text
    MED_GREY   = colors.HexColor("#666666")    # Medium grey — secondary info
    LIGHT_GREY = colors.HexColor("#f2f2f2")    # Very light grey — table backgrounds
    FLAG_RED   = colors.HexColor("#cc0000")    # Bold red — for High/Low flags
    HEADER_BG  = colors.HexColor("#e8e8e8")    # Slightly darker grey — column headers
    WHITE      = colors.white                  # Pure white — main row background

    # ── Step 2: Return everything in one dictionary ───────────────────────────
    # "dict()" creates a dictionary. We include both raw colours AND text styles
    # so we only need to call styles() once and everything is in one place.
    return dict(

        # Raw colours — used directly in TableStyle backgrounds, borders etc.
        BLACK=BLACK,
        DARK_GREY=DARK_GREY,
        MED_GREY=MED_GREY,
        LIGHT_GREY=LIGHT_GREY,
        FLAG_RED=FLAG_RED,
        HEADER_BG=HEADER_BG,
        WHITE=WHITE,

        # ── Text styles (ParagraphStyle objects) ──────────────────────────────
        # ParagraphStyle("InternalName", fontSize=..., fontName=..., textColor=...)
        # "InternalName" is just a unique ID ReportLab needs — it doesn't show on screen.
        # fontSize is in "points" (12pt = normal document text, 8pt = small print)
        # fontName: "Helvetica" = plain, "Helvetica-Bold" = bold

        # Big bold lab name shown at the very top of page 1
        lab_title = ParagraphStyle(
            "LabTitle",
            fontSize=13,                 # Relatively large — most prominent text
            fontName="Helvetica-Bold",   # Bold weight
            textColor=BLACK,
            alignment=TA_LEFT            # Left-aligned
        ),

        # Right-aligned request/control number shown at top right of page 1
        req_num = ParagraphStyle(
            "ReqNum",
            fontSize=9,
            fontName="Helvetica-Bold",
            textColor=DARK_GREY,
            alignment=TA_RIGHT           # Pushed to the right side
        ),

        # Large bold text for the patient's name
        pt_name = ParagraphStyle(
            "PtName",
            fontSize=13,
            fontName="Helvetica-Bold",
            textColor=BLACK
        ),

        # Medium bold for short labels like "Final" and "Ordering Physician"
        pt_status = ParagraphStyle(
            "PtStatus",
            fontSize=9,
            fontName="Helvetica-Bold",
            textColor=BLACK
        ),

        # Small grey text for patient detail lines (DOB, Age, Gender, ID, dates)
        pt_detail = ParagraphStyle(
            "PtDetail",
            fontSize=8,
            fontName="Helvetica",        # Regular (not bold) weight
            textColor=DARK_GREY
        ),

        # WHITE bold text for panel section headings — displayed on the coloured banner
        # e.g. "LIVER FUNCTION TESTS (LFT)" in white text on a green background
        panel_hdr = ParagraphStyle(
            "PanelHdr",
            fontSize=10,
            fontName="Helvetica-Bold",
            textColor=WHITE              # White so it's readable on dark backgrounds
        ),

        # Bold grey text for column header row labels (TEST | RESULT | FLAG etc.)
        col_hdr = ParagraphStyle(
            "ColHdr",
            fontSize=8,
            fontName="Helvetica-Bold",
            textColor=DARK_GREY
        ),

        # Normal black text for the test name column (e.g. "ALT (SGPT)")
        test_name = ParagraphStyle(
            "TestName",
            fontSize=8,
            fontName="Helvetica",
            textColor=BLACK
        ),

        # Normal black text for the result value column (e.g. "13", "1.34")
        result = ParagraphStyle(
            "Result",
            fontSize=8,
            fontName="Helvetica",
            textColor=BLACK
        ),

        # Bold RED text used exclusively for flags like "High" or "Low"
        # This makes abnormal results immediately obvious to the reader
        flag = ParagraphStyle(
            "Flag",
            fontSize=8,
            fontName="Helvetica-Bold",   # Bold makes it stand out even more
            textColor=FLAG_RED           # Red = warning colour
        ),

        # Small grey text for the reference range column (e.g. "0-44", ">59")
        ref_range = ParagraphStyle(
            "RefRange",
            fontSize=8,
            fontName="Helvetica",
            textColor=DARK_GREY          # Grey because it's supporting info, not primary
        ),

        # Very small grey text for interpretation notes shown below certain results
        # leftIndent=10 pushes the text slightly inward to visually show it's a sub-item
        note = ParagraphStyle(
            "Note",
            fontSize=7,                  # Smallest text in the report
            fontName="Helvetica",
            textColor=MED_GREY,
            leftIndent=10               # Indent from left edge = looks like a sub-item
        ),

        # Tiny grey text used in the footer area at the bottom of each page
        footer = ParagraphStyle(
            "Footer",
            fontSize=7,
            fontName="Helvetica",
            textColor=MED_GREY,
            alignment=TA_LEFT
        ),

        # Small grey text for the mini patient info header that appears on pages 2+
        page_hdr = ParagraphStyle(
            "PageHdr",
            fontSize=8,
            fontName="Helvetica",
            textColor=DARK_GREY
        ),
    )


# ==============================================================================
# SECTION 6 — DEMO DATA
# This is fake but realistic data used to test the script or show a sample report.
# When you run with --demo, the script uses these values instead of a real PDF.
# ✏️  You can change these to see how different data looks in the report.
# ==============================================================================

# ── Demo patient information ───────────────────────────────────────────────────
# This is a Python "dictionary" — a collection of named values.
# Format: "key": "value"
# You access a value with patient["name"], patient["dob"] etc.
DEMO_PATIENT = {
    "name":      "Zedni, Salim",          # Patient full name (Last, First)
    "dob":       "11/20/2001",            # Date of birth
    "age":       "22",                    # Age at time of testing
    "gender":    "Male",                  # Gender
    "fasting":   "Y",                     # Was patient fasting? Y = Yes, N = No
    "id":        "ZESA000010",            # Patient ID in your system
    "req":       "9719734",              # Requisition/control number for this order
    "physician": "MIRANDA, A",           # Doctor who ordered the tests
    "collected": "06/05/2024 13:41",     # Date and time the blood was taken
    "received":  "06/05/2024",           # Date the lab received the sample
    "reported":  "06/10/2024 21:07",     # Date and time results were finalised
    "status":    "Final",                # "Final" = results are complete and verified
}

# ── Demo test results ──────────────────────────────────────────────────────────
# This is a Python "list" of dictionaries — a list where each item is a test result.
# The script automatically sorts each result into the right panel using classify_test().
#
# Each result dictionary has these fields:
#   "test"   — the name of the test (used to decide which panel it goes into)
#   "result" — the measured value as a string (strings because some values have < or >)
#   "flag"   — "High", "Low", or "" (empty string means normal/no flag)
#   "ref"    — the normal reference range (e.g. "0-44" or ">59")
#   "unit"   — the unit of measurement (e.g. "mg/dL", "mmol/L", "%")
#   "notes"  — (optional) a list of interpretation lines shown below the result
DEMO_RESULTS = [

    # ── Liver Function Tests (these match "alt", "ast", "bilirubin" etc.) ──────
    {"test": "ALT (SGPT)",           "result": "13",   "flag": "",     "ref": "0-44",      "unit": "IU/L"},
    {"test": "AST (SGOT)",           "result": "15",   "flag": "",     "ref": "0-40",      "unit": "IU/L"},
    {"test": "Alkaline Phosphatase", "result": "59",   "flag": "",     "ref": "44-121",    "unit": "IU/L"},
    {"test": "Bilirubin, Total",     "result": "1.5",  "flag": "High", "ref": "0.0-1.2",   "unit": "mg/dL"},
    # "High" above means the result (1.5) is above the normal range (0.0-1.2)
    # The word "High" will appear in RED in the flag column of the report
    {"test": "Protein, Total",       "result": "6.7",  "flag": "",     "ref": "6.0-8.5",   "unit": "g/dL"},
    {"test": "Albumin",              "result": "4.6",  "flag": "",     "ref": "4.3-5.2",   "unit": "g/dL"},
    {"test": "Globulin, Total",      "result": "2.1",  "flag": "",     "ref": "1.5-4.5",   "unit": "g/dL"},
    {"test": "A/G Ratio",            "result": "2.2",  "flag": "",     "ref": "1.2-2.2",   "unit": ""},

    # ── Kidney Function Tests (these match "bun", "creatinine", "sodium" etc.) ─
    {"test": "BUN",                  "result": "18",   "flag": "",     "ref": "6-20",      "unit": "mg/dL"},
    {"test": "Creatinine",           "result": "1.34", "flag": "High", "ref": "0.76-1.27", "unit": "mg/dL"},
    {"test": "eGFR",                 "result": "77",   "flag": "",     "ref": ">59",       "unit": "mL/min/1.73"},
    {"test": "BUN/Creatinine Ratio", "result": "13",   "flag": "",     "ref": "9-20",      "unit": ""},
    {"test": "Sodium",               "result": "139",  "flag": "",     "ref": "134-144",   "unit": "mmol/L"},
    {"test": "Potassium",            "result": "4.5",  "flag": "",     "ref": "3.5-5.2",   "unit": "mmol/L"},
    {"test": "Chloride",             "result": "100",  "flag": "",     "ref": "96-106",    "unit": "mmol/L"},
    {"test": "Carbon Dioxide, Total","result": "20",   "flag": "",     "ref": "20-29",     "unit": "mmol/L"},
    {"test": "Calcium",              "result": "9.2",  "flag": "",     "ref": "8.7-10.2",  "unit": "mg/dL"},

    # ── Lipid Profile (these match "cholesterol", "triglyceride", "hdl" etc.) ──
    {"test": "Cholesterol, Total",   "result": "202",  "flag": "High", "ref": "100-199",   "unit": "mg/dL"},
    {"test": "Triglycerides",        "result": "69",   "flag": "",     "ref": "0-149",     "unit": "mg/dL"},
    {"test": "HDL Cholesterol",      "result": "45",   "flag": "",     "ref": ">39",       "unit": "mg/dL"},
    {"test": "VLDL Cholesterol Cal", "result": "13",   "flag": "",     "ref": "5-40",      "unit": "mg/dL"},
    {"test": "LDL Chol Calc (NIH)",  "result": "144",  "flag": "High", "ref": "0-99",      "unit": "mg/dL"},
    {
        "test": "LDL/HDL Ratio",
        "result": "3.2",
        "flag": "",
        "ref": "0.0-3.6",
        "unit": "ratio",
        # "notes" is a list of strings — each one becomes a separate pale yellow note row
        "notes": [
            "Men:   1/2 Avg. Risk 1.0  |  Avg. Risk 3.6  |  2X Avg. Risk 6.2  |  3X Avg. Risk 8.0",
            "Women: 1/2 Avg. Risk 1.5  |  Avg. Risk 3.2  |  2X Avg. Risk 5.0  |  3X Avg. Risk 6.1",
        ]
    },

    # ── HbA1c (matches "hemoglobin a1c", "a1c" etc.) ─────────────────────────
    {
        "test": "Hemoglobin A1c",
        "result": "6.0",
        "flag": "High",          # 6.0% is above the normal 4.8-5.6% range
        "ref": "4.8-5.6",
        "unit": "%",
        "notes": [
            "Normal (Non-diabetic): 4.8 - 5.6%",
            "Prediabetes:           5.7 - 6.4%",
            "Diabetes:              > 6.4%",
            "Target for diabetics:  < 7.0%",
        ]
    },

    # ── Uric Acid (matches "uric acid", "urate") ──────────────────────────────
    {
        "test": "Uric Acid",
        "result": "6.2",
        "flag": "",              # 6.2 is within normal range (3.4-7.0) so no flag
        "ref": "3.4-7.0",
        "unit": "mg/dL",
        "notes": [
            "Elevated uric acid (hyperuricaemia) may be associated with gout, kidney stones, or renal impairment.",
            "Normal range: Men 3.4-7.0 mg/dL  |  Women 2.4-6.0 mg/dL",
        ]
    },
]

# ── Demo lab information ───────────────────────────────────────────────────────
# This list contains the performing laboratory details shown in the footer table.
# "code" must be "01", "02" etc. — short identifiers for each lab location.
DEMO_LABS = [
    {
        "code":     "01",          # Short code shown in the "Lab" column
        "name":     LAB_NAME,      # Uses the LAB_NAME variable defined at top of file
        "address":  LAB_ADDRESS,   # Uses the LAB_ADDRESS variable
        "phone":    LAB_PHONE,     # Uses the LAB_PHONE variable
        "director": LAB_DIRECTOR,  # Uses the LAB_DIRECTOR variable
    },
]


# ==============================================================================
# SECTION 7 — LabReportDoc CLASS
# A "class" is a blueprint for creating a custom type of object.
# Here we create a special PDF document that automatically draws a
# patient info header on pages 2+ and a footer on every page.
#
# We "extend" (inherit from) ReportLab's BaseDocTemplate, meaning we take
# everything BaseDocTemplate already does and ADD our own header/footer on top.
# ==============================================================================

class LabReportDoc(BaseDocTemplate):
    # BaseDocTemplate is the "parent" — it handles all the page layout machinery.
    # LabReportDoc is the "child" — it adds our custom header and footer behaviour.

    def __init__(self, filename, patient, **kwargs):
        """
        __init__ is called when we CREATE a new LabReportDoc object.
        It's like the "set up" step that runs once when the document is born.

        Parameters:
          filename — the file path where the PDF will be saved (e.g. "report.pdf")
          patient  — the patient dictionary so we can print their name in the header
          **kwargs — any extra settings (pagesize, margins etc.) passed through
        """

        # Call the parent class's __init__ first so it sets up all its own machinery
        # "super()" means "the class we inherited from" (BaseDocTemplate)
        super().__init__(filename, **kwargs)

        # Save the patient data as part of this document object
        # "self" means "this specific document" — like labelling a box
        self.patient = patient

        # Load all text styles and save them to this document object
        self.s = styles()

        # Define the "Frame" — the rectangular area inside the margins
        # where the actual content (paragraphs, tables) will flow
        frame = Frame(
            self.leftMargin,    # Left edge of the content area (left margin)
            self.bottomMargin,  # Bottom edge of the content area (bottom margin)
            self.width,         # Width of usable space (page width minus both margins)
            self.height,        # Height of usable space (page height minus top/bottom)
            id='normal'         # A name/ID for this frame (required by ReportLab)
        )

        # Register the PageTemplate — this tells ReportLab:
        # "Every page should use this frame, AND run _on_page() after drawing each page"
        # onPage=self._on_page means "call our _on_page method when each page is rendered"
        self.addPageTemplates([
            PageTemplate(id='All', frames=frame, onPage=self._on_page)
        ])

    def _on_page(self, canvas, doc):
        """
        This method is called AUTOMATICALLY every time ReportLab starts a new page.
        "canvas" is like a blank drawing surface — we use it to draw text and lines.
        "doc" is the document itself — gives us info like the current page number.

        We use this to draw:
          - A running patient info header on pages 2 and beyond
          - A confidential notice footer on every single page
        """

        # saveState() saves all current drawing settings (font, colour, line width etc.)
        # This is important because if we change settings here, we don't want those
        # changes to bleed into the main content below.
        canvas.saveState()

        # Shorthand variable names to make the code below less verbose
        p = self.patient   # p = patient dictionary, so p["name"] = patient name
        s = self.s         # s = styles dictionary, so s["DARK_GREY"] = dark grey colour

        # ── Running header: only shows on page 2 and beyond ──────────────────
        # Page 1 has its own full patient info block, so we don't repeat it there.
        # doc.page gives us the current page number (1, 2, 3 etc.)
        if doc.page > 1:

            # Set font to Helvetica size 8 for the header text
            canvas.setFont("Helvetica", 8)

            # Set text colour to dark grey
            canvas.setFillColor(s["DARK_GREY"])

            # Build the three lines of patient info text using f-strings
            # f"..." = an f-string — curly braces {} insert variable values into the text
            line1 = f"Name: {p['name']}    Ordering Physician: {p['physician']}"
            line2 = f"Age: {p['age']} Years    DOB: {p['dob']}    Req/Control #: {p['req']}"
            line3 = f"Gender: {p['gender']}    ID: {p['id']}    Page: {doc.page}"

            # drawString(x, y, text) draws text starting at position (x, y)
            # In ReportLab, y=0 is the BOTTOM of the page, y increases going UP
            # doc.height + doc.bottomMargin puts us at the TOP of the content area
            canvas.drawString(
                doc.leftMargin,                          # Start at left margin
                doc.height + doc.bottomMargin + 10,      # 10 points above content area
                line1
            )
            canvas.drawString(
                doc.leftMargin,
                doc.height + doc.bottomMargin - 2,       # Slightly lower — second line
                line2
            )

            # drawRightString is like drawString but text is aligned to the RIGHT
            # at the given x coordinate (doc.width + doc.leftMargin = right edge)
            canvas.drawRightString(
                doc.width + doc.leftMargin,              # Right edge of content area
                doc.height + doc.bottomMargin + 10,
                line3
            )

            # Draw a thin horizontal dividing line under the header text
            canvas.setStrokeColor(colors.HexColor("#aaaaaa"))  # Light grey line colour
            canvas.setLineWidth(0.5)                           # Very thin line

            # canvas.line(x1, y1, x2, y2) draws a straight line from point 1 to point 2
            canvas.line(
                doc.leftMargin,                          # Start x = left margin
                doc.height + doc.bottomMargin + 20,      # Start y = above header text
                doc.width + doc.leftMargin,              # End x = right margin
                doc.height + doc.bottomMargin + 20       # End y = same height (horizontal)
            )

        # ── Footer: shows on EVERY page ───────────────────────────────────────
        canvas.setFont("Helvetica", 7)           # Tiny font for footer
        canvas.setFillColor(s["MED_GREY"])       # Medium grey colour

        # doc.bottomMargin - 15 places the text just BELOW the content area
        # (in the bottom margin space)
        canvas.drawString(
            doc.leftMargin,
            doc.bottomMargin - 15,
            f"CONFIDENTIAL PATIENT REPORT  |  {LAB_NAME}  |  Page {doc.page}"
        )

        # restoreState() puts ALL drawing settings back to what they were before
        # saveState() — this prevents our header/footer colours from affecting
        # the main content that ReportLab draws after this function returns
        canvas.restoreState()


# ==============================================================================
# SECTION 8 — group_results_by_panel() FUNCTION
# Takes the flat list of all results and sorts them into separate groups
# (one group per panel) using classify_test().
# ==============================================================================

def group_results_by_panel(results):
    """
    Translates short BS240 test codes to full clinical names, filters __SKIP__
    rows, then sorts all results into their correct panel groups.

    Input:  flat list of raw result dicts from parse_pdf_data()
    Output: dict keyed by panel name, each value a list of enriched result dicts

    Step 1 — apply_test_name_map():
      Replaces short codes with full names using TEST_CODE_MAP
      (e.g. "T-bil" → "Total Bilirubin", "γ-GT" → "Gamma GT")
      Removes rows flagged as __SKIP__ (e.g. HbA1c IFCC / eAG rows)
      Adds a "panel_override" key to force correct panel assignment

    Step 2 — classify_test():
      Uses panel_override if set, otherwise falls back to keyword matching

    Example output:
      {
        "Liver Function Tests (LFT)": [
            {"test": "Total Bilirubin", "result": "12.76", ...},
            {"test": "ALT (Alanine Aminotransferase)", "result": "43.7", ...},
        ],
        "Kidney Function Tests (KFT)": [
            {"test": "Creatinine", "result": "123.1", ...},
        ],
        ...
      }
    """

    # ── Step 1: Resolve short codes → full names, remove skip rows ───────────
    resolved = apply_test_name_map(results)

    # ── Step 2: Initialise one empty bucket per panel ─────────────────────────
    grouped = {panel: [] for panel in PANEL_ORDER}

    # ── Step 3: Sort each resolved result into its panel ──────────────────────
    for r in resolved:
        # Use panel_override if apply_test_name_map() set it, otherwise ""
        override = r.get("panel_override", "")

        # classify_test() respects the override first, then falls back to keywords
        panel = classify_test(r["test"], override)

        if panel not in grouped:
            panel = "Other"

        grouped[panel].append(r)

    # ── Step 4: HbA1c panel — keep ONLY the NGSP row ─────────────────────────
    # The BS240 and other analysers sometimes output several HbA1c sub-fractions
    # (NGSP %, IFCC mmol/mol, eAG mg/dL etc.).  The report should only show the
    # NGSP percentage — the internationally recognised patient-facing value.
    #
    # Strategy: scan the HbA1c bucket for any result whose name contains "ngsp"
    # (case-insensitive).  If one is found, keep only that row.  If none is
    # labelled "ngsp", fall back to keeping just the first row so the panel is
    # never empty.
    hba1c_key = "Glycated Haemoglobin (HbA1c)"
    hba1c_rows = grouped.get(hba1c_key, [])
    if len(hba1c_rows) > 1:
        ngsp_rows = [r for r in hba1c_rows if "ngsp" in r["test"].lower()]
        if ngsp_rows:
            grouped[hba1c_key] = ngsp_rows[:1]   # exactly one NGSP row
        else:
            grouped[hba1c_key] = hba1c_rows[:1]  # fallback: first row only

    return grouped


# ==============================================================================
# SECTION 9 — build_panel() FUNCTION
# Builds all the visual elements for ONE panel and returns them as a list.
# Called once per panel inside generate_report().
# ==============================================================================

def build_panel(panel_name, panel_results, W, s):
    """
    Builds the visual elements for a single panel (e.g. Liver Function Tests).
    Returns a list of ReportLab elements: [banner, col_headers, row1, row2, ...]

    Parameters:
      panel_name    — string like "Liver Function Tests (LFT)"
      panel_results — list of result dicts that belong to this panel
      W             — usable page width in points (used to size columns)
      s             — styles dictionary from styles()
    """

    # "block" will collect all the elements for this panel
    # We build it as a list, then return it at the end
    block = []

    # ── Look up this panel's colours ─────────────────────────────────────────
    # Check if this panel has a definition in PANEL_DEFINITIONS
    if panel_name in PANEL_DEFINITIONS:
        # Get the dark colour for the heading banner
        panel_color = PANEL_DEFINITIONS[panel_name]["color"]
        # Get the light colour for alternating row backgrounds
        panel_light = PANEL_DEFINITIONS[panel_name]["light"]
    else:
        # "Other" panel has no definition, so fall back to neutral grey colours
        panel_color = colors.HexColor("#555555")   # Dark grey
        panel_light = colors.HexColor("#f0f0f0")   # Very light grey


    # ── Element 1: Coloured heading banner ────────────────────────────────────
    # This is a full-width table with one cell containing the panel name in white text.
    # The cell background is the panel's unique colour (green, blue, purple etc.)

    banner = Table(
        # The table data: one row, one cell, containing a Paragraph
        # .upper() converts "Liver Function Tests (LFT)" to "LIVER FUNCTION TESTS (LFT)"
        [[Paragraph(panel_name.upper(), s["panel_hdr"])]],
        colWidths=[W]   # One column that spans the full page width
    )
    banner.setStyle(TableStyle([
        # BACKGROUND sets the background colour of the cells
        # (0,0) = top-left cell, (-1,-1) = bottom-right cell (in a 1x1 table, both are the same)
        ("BACKGROUND",    (0,0), (-1,-1), panel_color),  # Panel's unique dark colour
        ("TOPPADDING",    (0,0), (-1,-1), 6),   # 6pt gap above the text inside the cell
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),   # 6pt gap below the text
        ("LEFTPADDING",   (0,0), (-1,-1), 8),   # 8pt gap on the left of the text
    ]))
    block.append(banner)   # Add the banner to our list of elements


    # ── Element 2: Column header row ─────────────────────────────────────────
    # A light grey row with labels: TEST | RESULT | FLAG | REFERENCE RANGE | UNIT

    col_hdr_tbl = Table([[
        Paragraph("TEST",            s["col_hdr"]),   # Column 1 label
        Paragraph("RESULT",          s["col_hdr"]),   # Column 2 label
        Paragraph("FLAG",            s["col_hdr"]),   # Column 3 label
        Paragraph("REFERENCE RANGE", s["col_hdr"]),   # Column 4 label
        Paragraph("UNIT",            s["col_hdr"]),   # Column 5 label
    ]], colWidths=[
        W*0.38,   # Column 1 (Test name) = 38% of page width — widest column
        W*0.12,   # Column 2 (Result)    = 12% of page width
        W*0.10,   # Column 3 (Flag)      = 10% of page width
        W*0.25,   # Column 4 (Ref range) = 25% of page width
        W*0.15,   # Column 5 (Unit)      = 15% of page width
        # Total = 38+12+10+25+15 = 100% — fills the full width perfectly
    ])

    col_hdr_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), s["HEADER_BG"]),     # Light grey background
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 4),
        # LINEBELOW draws a coloured line under the column headers
        # The colour matches this panel — a nice visual touch tying header to panel
        ("LINEBELOW",     (0,0), (-1,-1), 0.5, panel_color),
    ]))
    block.append(col_hdr_tbl)


    # ── Elements 3+: One row per test result ──────────────────────────────────
    # enumerate() loops through the list AND gives us a counter (i) at the same time
    # i = 0 for the first result, i = 1 for the second, etc.
    # r = the current result dictionary (e.g. {"test": "ALT", "result": "13", ...})
    for i, r in enumerate(panel_results):

        # ── HbA1c panel: always inject interpretation notes ───────────────────
        # The HbA1c panel only shows the NGSP % value.
        # We always add the standard interpretation table as notes so the
        # clinician and patient can see what the value means at a glance.
        if "Glycated" in panel_name or "HbA1c" in panel_name:
            if not r.get("notes"):
                try:
                    val = float(r.get("result", "0"))
                except (ValueError, TypeError):
                    val = 0.0
                if val >= 6.5:
                    interp = "⚠  Consistent with Diabetes (≥6.5%)"
                elif val >= 5.7:
                    interp = "⚠  Prediabetes range (5.7–6.4%)"
                else:
                    interp = "✓  Normal / Non-diabetic (<5.7%)"
                r = dict(r)   # copy so we don't mutate the shared list
                r["notes"] = [
                    f"Interpretation: {interp}",
                    "Normal (Non-diabetic): <5.7%   |   Prediabetes: 5.7–6.4%   |   Diabetes: ≥6.5%   |   Target for known diabetics: <7.0%",
                ]

        # Get the flag text — e.g. "High", "Low", or "" (empty = normal)
        # .get("flag", "") safely returns "" if no "flag" key exists
        flag_txt = r.get("flag", "")

        # Check if this flag word indicates an abnormal result
        # "in" checks if flag_txt matches any item in the list
        is_flag = flag_txt in ("High", "Low", "high", "low", "H", "L", "*", "Immune")

        # Choose the right text style:
        # If it's a flag → use the red bold "flag" style
        # If it's normal → use the plain black "result" style
        # This is a Python "ternary expression": value_if_true if condition else value_if_false
        f_style = s["flag"] if is_flag else s["result"]

        # Build the row as a one-row Table with 5 columns
        row_data = [[
            Paragraph(r["test"],             s["test_name"]),   # Full clinical name (already resolved by apply_test_name_map)
            Paragraph(r["result"],         s["result"]),      # Result value (col 2)
            Paragraph(flag_txt,            f_style),          # Flag — red if abnormal (col 3)
            Paragraph(r.get("ref", ""),    s["ref_range"]),   # Reference range (col 4)
            Paragraph(r.get("unit", ""),   s["ref_range"]),   # Unit (col 5)
        ]]

        # Create the table with the same column widths as the header row
        row_tbl = Table(row_data,
                        colWidths=[W*0.38, W*0.12, W*0.10, W*0.25, W*0.15])

        # Zebra striping: alternate between white and the panel's light colour
        # i % 2 = the remainder when i is divided by 2
        #   i=0 → 0 % 2 = 0 → white (even)
        #   i=1 → 1 % 2 = 1 → panel light colour (odd)
        #   i=2 → 2 % 2 = 0 → white (even) ... and so on
        bg = colors.white if i % 2 == 0 else panel_light

        row_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), bg),   # Alternating row background
            ("TOPPADDING",    (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ("LEFTPADDING",   (0,0), (-1,-1), 4),
            # LINEBELOW draws a very faint grey dividing line under each row
            ("LINEBELOW",     (0,0), (-1,-1), 0.2, colors.HexColor("#dddddd")),
        ]))
        block.append(row_tbl)

        # ── Optional: interpretation note rows ────────────────────────────────
        # Some results have extra "notes" (e.g. HbA1c ranges, risk tables).
        # If the result has a "notes" key, loop through each note string.
        # r.get("notes", []) returns an empty list [] if no "notes" key exists,
        # which means this loop simply doesn't run for results without notes.
        for note in r.get("notes", []):

            # Create a full-width single-cell table for this note
            nt = Table(
                [[Paragraph(note, s["note"])]],
                colWidths=[W]    # Full page width
            )
            nt.setStyle(TableStyle([
                # Pale yellow background — visually signals "extra information"
                ("BACKGROUND",    (0,0), (-1,-1), colors.HexColor("#fffef0")),
                ("TOPPADDING",    (0,0), (-1,-1), 1),
                ("BOTTOMPADDING", (0,0), (-1,-1), 1),
                ("LEFTPADDING",   (0,0), (-1,-1), 20),   # Indented to look like a sub-item
            ]))
            block.append(nt)

    # Add a small vertical gap after all the rows in this panel
    # Spacer(width, height) — width doesn't matter for vertical spacers
    block.append(Spacer(1, 8))   # 8 points of blank space

    return block   # Return the completed list of elements for this panel


# ==============================================================================
# SECTION 10 — generate_report() FUNCTION
# The master function that assembles everything and builds the final PDF.
# Called once with all the data — patient info, results, labs, output path.
# ==============================================================================

def generate_report(patient, results, labs, output_path):
    """
    Builds and saves the complete PDF report.

    Parameters:
      patient     — dict with patient info (name, DOB, age etc.)
      results     — flat list of all test result dicts
      labs        — list of performing lab dicts (name, address etc.)
      output_path — string file path where the PDF is saved (e.g. "report.pdf")
    """

    # Create the document using our custom LabReportDoc class
    # This sets up the page size and margins
    doc = LabReportDoc(
        output_path,        # Where to save the file
        patient,            # Patient data (needed by LabReportDoc for running headers)
        pagesize=letter,    # US letter size (8.5" × 11") — change to A4 if needed
        leftMargin=0.6*inch,    # 0.6 inches from left edge of page
        rightMargin=0.6*inch,   # 0.6 inches from right edge
        topMargin=0.7*inch,     # 0.7 inches from top
        bottomMargin=0.5*inch,  # 0.5 inches from bottom
    )

    s     = styles()    # Load all text styles into variable "s"
    story = []          # "story" = the list of ALL elements to put in the PDF
                        # ReportLab reads this list top-to-bottom and flows it across pages
    W     = doc.width   # Usable page width = page width minus left and right margins


    # ── PAGE 1 TOP: Lab name (left) + Request number (right) ─────────────────
    # A two-column table: lab name on the left, request number on the right

    hdr = Table([[
        # <b>...</b> is HTML-style bold markup that ReportLab's Paragraph understands
        Paragraph(f"<b>{LAB_NAME}</b>", s["lab_title"]),
        Paragraph(f"Req/Control #: <b>{patient['req']}</b>", s["req_num"]),
    ]], colWidths=[
        W*0.6,   # Lab name column = 60% of width
        W*0.4,   # Req number column = 40% of width
    ])
    hdr.setStyle(TableStyle([
        ("VALIGN",        (0,0), (-1,-1), "TOP"),   # Align both cells to the top
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    story.append(hdr)

    # Draw a thick dark horizontal line under the lab name/req number
    # HRFlowable = a Horizontal Rule (a straight line)
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#333333")))

    # Add a small 4-point gap below the line before the patient block starts
    story.append(Spacer(1, 4))


    # ── PATIENT INFORMATION BLOCK ────────────────────────────────────────────
    # Two columns side by side:
    #   Left  column = patient demographics (name, DOB, age, gender, fasting)
    #   Right column = collection info (physician, collected/received/reported dates)

    # Build the LEFT column as a list of Paragraphs and Spacers
    left_col = [
        Paragraph(f"<b>{patient['name']}</b>",    s["pt_name"]),    # Bold large name
        Paragraph(f"<b>{patient['status']}</b>",  s["pt_status"]),  # e.g. "Final"
        Spacer(1, 4),   # Small vertical gap between status and the detail lines
        Paragraph(f"DOB: {patient['dob']}",        s["pt_detail"]),
        Paragraph(f"AGE: {patient['age']} Years",  s["pt_detail"]),
        Paragraph(f"GENDER: {patient['gender']}",  s["pt_detail"]),
        Paragraph(f"FASTING: {patient['fasting']}", s["pt_detail"]),
    ]

    # Build the RIGHT column
    right_col = [
        Paragraph("<b>Ordering Physician</b>",          s["pt_status"]),
        Paragraph(patient['physician'],                 s["pt_detail"]),
        Spacer(1, 4),
        Paragraph(f"COLLECTED: {patient['collected']}",  s["pt_detail"]),
        Paragraph(f"RECEIVED:  {patient['received']}",   s["pt_detail"]),
        Paragraph(f"REPORTED:  {patient['reported']}",   s["pt_detail"]),
        Paragraph(f"ID: {patient['id']}",                s["pt_detail"]),
    ]

    # Helper function to wrap a list of elements into a single-column table
    # This is needed because Table cells must contain Table or Paragraph objects,
    # not raw Python lists. This little function handles that conversion.
    def col_table(items):
        # Turn each item into a single-item list: [item] → becomes a table row
        # [[item1], [item2], [item3]] = a table with 3 rows and 1 column
        t = Table([[item] for item in items], colWidths=[W*0.45])
        t.setStyle(TableStyle([
            ("TOPPADDING",    (0,0), (-1,-1), 1),
            ("BOTTOMPADDING", (0,0), (-1,-1), 1),
        ]))
        return t

    # Put the two columns side by side inside an outer Table
    # [[left, right]] = one row with two cells
    pt_block = Table(
        [[col_table(left_col), col_table(right_col)]],
        colWidths=[W*0.5, W*0.5]   # Each column is exactly half the page width
    )
    pt_block.setStyle(TableStyle([
        ("VALIGN",        (0,0), (-1,-1), "TOP"),              # Both columns start at the top
        ("BACKGROUND",    (0,0), (-1,-1), s["LIGHT_GREY"]),    # Light grey background for the block
        ("BOX",           (0,0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),  # Thin grey border around block
        ("TOPPADDING",    (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LEFTPADDING",   (0,0), (-1,-1), 8),
    ]))
    story.append(pt_block)
    story.append(Spacer(1, 10))   # 10-point gap before the colour legend


    # ── PANEL COLOUR LEGEND ───────────────────────────────────────────────────
    # A small strip showing a coloured square next to each panel name.
    # Helps the reader navigate the report quickly.
    # Example:  🟩 LIVER FUNCTION TESTS   🟦 KIDNEY FUNCTION TESTS  ...

    legend_items = []   # Will hold alternating [coloured_square, panel_name_text] pairs

    for panel_name in PANEL_ORDER:
        # Skip "Other" — we don't want to show it in the legend
        if panel_name == "Other":
            continue

        pdef = PANEL_DEFINITIONS[panel_name]   # Get this panel's colour info

        # Create a tiny coloured square — a 1-cell Table with a coloured background
        swatch = Table(
            [[""]],                       # Empty cell — colour comes from background
            colWidths=[0.18*inch],        # Very narrow square width
            rowHeights=[0.14*inch]        # Very short square height
        )
        swatch.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), pdef["color"]),   # Fill with panel colour
        ]))

        # Add the coloured square and the panel name text to our legend list
        legend_items.append(swatch)
        legend_items.append(Paragraph(f" {panel_name}   ", s["pt_detail"]))

    # Place all legend items in a single horizontal row
    # colWidths=[0.2*inch, W/5 - 0.2*inch] * 5 creates pairs of [swatch_width, text_width]
    # for all 5 panels — total width = full page width
    legend_row = Table(
        [legend_items],
        colWidths=[0.2*inch, W/5 - 0.2*inch] * 5
    )
    legend_row.setStyle(TableStyle([
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),              # Vertically centre everything
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 2),
        ("BACKGROUND",    (0,0), (-1,-1), colors.HexColor("#f9f9f9")),  # Off-white background
        ("BOX",           (0,0), (-1,-1), 0.3, colors.HexColor("#cccccc")),  # Thin border
    ]))
    story.append(legend_row)
    story.append(Spacer(1, 10))


    # ── SORT RESULTS INTO PANELS ──────────────────────────────────────────────
    # Call our grouping function which uses classify_test() to sort all results
    # Returns a dictionary like: {"Liver...": [...results...], "Kidney...": [...], ...}
    grouped = group_results_by_panel(results)


    # ── BUILD AND ADD EACH PANEL ──────────────────────────────────────────────
    # Loop through panels in the defined order (LFT → KFT → Lipid → HbA1c → Uric Acid)
    for panel_name in PANEL_ORDER:

        # Get the list of results for this panel (empty list if none)
        panel_results = grouped.get(panel_name, [])

        # If this panel has no results, skip it entirely — don't show an empty heading
        if not panel_results:
            continue   # "continue" skips the rest of this loop iteration and goes to next panel

        # Build the visual elements for this panel (returns a list of ReportLab objects)
        panel_block = build_panel(panel_name, panel_results, W, s)

        # KeepTogether wraps the first 4 elements (banner + col headers + first 2 rows)
        # and tells ReportLab to try to keep them on the same page.
        # This prevents a panel heading from getting stranded alone at the bottom of a page.
        story.append(KeepTogether(panel_block[:4]))
        # panel_block[:4] = a slice — elements at index 0, 1, 2, 3 (the first 4)

        # Add the remaining rows normally — they can flow across pages if needed
        for item in panel_block[4:]:
            # panel_block[4:] = all elements from index 4 onwards
            story.append(item)


    # ── PERFORMING LABORATORY FOOTER TABLE ───────────────────────────────────
    # A table at the very end listing which lab(s) performed the tests

    story.append(Spacer(1, 12))   # Gap before the footer section

    # Thin grey horizontal line as a visual separator
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#aaaaaa")))
    story.append(Spacer(1, 4))

    # Bold label above the table
    story.append(Paragraph("<b>Performing Laboratory Information:</b>", s["footer"]))
    story.append(Spacer(1, 3))

    # Build the lab table — starting with the column header row
    lab_rows = [[
        Paragraph("<b>Code</b>",     s["col_hdr"]),
        Paragraph("<b>Name</b>",     s["col_hdr"]),
        Paragraph("<b>Address</b>",  s["col_hdr"]),
        Paragraph("<b>Phone</b>",    s["col_hdr"]),
        Paragraph("<b>Lab Scientist</b>", s["col_hdr"]),
    ]]

    # Add one data row for each lab in the labs list
    for lab in labs:
        lab_rows.append([
            Paragraph(lab["code"],     s["footer"]),
            Paragraph(lab["name"],     s["footer"]),
            Paragraph(lab["address"],  s["footer"]),
            Paragraph(lab["phone"],    s["footer"]),
            Paragraph(lab["director"], s["footer"]),
        ])

    # Create the table with proportional column widths adding up to 100%
    lab_tbl = Table(lab_rows,
                    colWidths=[W*0.06, W*0.22, W*0.34, W*0.18, W*0.20])
    lab_tbl.setStyle(TableStyle([
        ("BACKGROUND",     (0,0),  (-1,0),  s["HEADER_BG"]),              # Grey header row
        ("ROWBACKGROUNDS", (0,1),  (-1,-1), [colors.white, s["LIGHT_GREY"]]),  # Alternating data rows
        ("TOPPADDING",     (0,0),  (-1,-1), 2),
        ("BOTTOMPADDING",  (0,0),  (-1,-1), 2),
        ("LEFTPADDING",    (0,0),  (-1,-1), 4),
        ("GRID",           (0,0),  (-1,-1), 0.2, colors.HexColor("#cccccc")),  # Very light grid lines
    ]))
    story.append(lab_tbl)
    story.append(Spacer(1, 4))

    # Timestamp + confidentiality notice at the very bottom
    # datetime.today() gets the current date and time
    # .strftime("%Y-%m-%d %H:%M") formats it as "2024-06-10 21:07"
    story.append(Paragraph(
        f"Report generated: {datetime.today().strftime('%Y-%m-%d %H:%M')}  |  "
        "This report is confidential and intended solely for the named ordering physician.",
        s["footer"]
    ))

    # ── FINAL STEP: Build and save the PDF ────────────────────────────────────
    # multiBuild() processes the story list and writes the PDF to disk.
    # We use multiBuild (not build) because we're using a custom BaseDocTemplate
    # with page templates — multiBuild handles the two-pass rendering needed
    # for features like "Page X of Y" to work correctly.
    doc.multiBuild(story)

    # Print a success message to the terminal
    print(f"✅ Report saved: {output_path}")


# ==============================================================================
# SECTION 11 — extract_text_from_pdf() FUNCTION
#
# Opens a PDF exported by the Mindray BS240 analyser and extracts all text
# AND positional word data so we can reconstruct the column layout.
#
# WHY POSITIONAL EXTRACTION?
#   The BS240 produces PDFs in (at least) two different layouts:
#
#   Layout A — Old-style (Weija Gbawe format):
#     Uses real table borders → pdfplumber's extract_tables() works perfectly.
#     Columns: Test | Full Name | Concentration | Unit | Results | Remark | Reference
#
#   Layout B — New-style (Yaa Specimen / CHE900 format):
#     NO table borders — just text positioned at fixed x-coordinates.
#     extract_tables() returns 0 tables so the old approach gets nothing.
#     The columns are identified purely by their horizontal (x) position:
#       x ≈   7  → Test name (e.g. "T-bil", "ALT", "CORONARY R")
#       x ≈ 116  → Result value (e.g. "43.7", "0.8")
#       x ≈ 228  → Unit (e.g. "U/L", "mmol/L", "μmol/L")
#       x ≈ 303  → Flag (just "H" or "L" — only present if abnormal)
#       x ≈ 398  → Reference range (e.g. "0.0-41.0", "44.0-106.0")
#
# This function now returns a dict with THREE keys:
#   "header_text"   → full plain text string (used for patient field regex)
#   "table_rows"    → rows from extract_tables() — populated in Layout A only
#   "word_rows"     → rows built from positional words — populated for Layout B
#                     Each row: [test_name, result, unit, flag, ref_range]
# ==============================================================================

def extract_text_from_pdf(pdf_path):
    """
    Extracts header text, table rows (if any), and positional word rows
    from a Mindray BS240 PDF export.

    Returns:
      {
        "header_text": "Patient: YAA SPECIMEN ...",
        "table_rows":  [...],   # non-empty only when real table borders exist
        "word_rows":   [        # always populated from positional grouping
            ["ALT",  "43.7",  "U/L",    "H", "0.0-41.0"],
            ["AST",  "34.1",  "U/L",    "",  "0.0-40.0"],
            ["CREA", "123.1", "μmol/L", "H", "44.0-106.0"],
            ...
        ]
      }

    Requires pdfplumber — install with:  pip install pdfplumber
    """

    try:
        import pdfplumber
    except ImportError:
        print("❌ pdfplumber is not installed.")
        print("   Run:  pip install pdfplumber")
        raise

    header_text = ""   # Plain text for patient field regex
    table_rows  = []   # Rows from extract_tables() — only for bordered-table PDFs
    word_rows   = []   # Rows reconstructed from x-position grouping

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            # ── Plain text: used for patient header fields ─────────────────────
            plain = page.extract_text()
            if plain:
                header_text += plain + "\n"

            # ── Table extraction: works when table borders exist ───────────────
            # This succeeds for Layout A (old Weija Gbawe style).
            # For Layout B it returns an empty list — we handle that below.
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or all(c is None or str(c).strip() == "" for c in row):
                        continue
                    cleaned = [str(c).strip() if c is not None else "" for c in row]
                    table_rows.append(cleaned)

            # ── Positional word grouping: works for both layouts ───────────────
            # extract_words() returns every word with its x,y coordinates.
            # We group words that share the same vertical position (same row)
            # then assign each word to a column based on its x coordinate.
            #
            # COLUMN BOUNDARIES (from inspecting the actual PDF word positions):
            #   x < 100     → Column 0: Test name
            #   100 ≤ x < 200 → Column 1: Result value
            #   200 ≤ x < 280 → Column 2: Unit
            #   280 ≤ x < 360 → Column 3: Flag (H or L)
            #   360 ≤ x       → Column 4: Reference range
            #
            # These boundaries work for both the old and new BS240 PDF formats
            # because the result columns are always in the same relative positions.
            #
            # We use a tolerance of ±4 points when grouping by vertical position
            # because slight differences in baseline can shift top values by 1-2pt.

            words = page.extract_words()

            # Group words into rows by their vertical (top) coordinate
            # "rows_by_y" is a dictionary: {y_value: [list_of_word_dicts]}
            rows_by_y = {}
            for word in words:
                # Round to nearest 2 points to group words on the same line
                # (small floating-point differences in top values otherwise
                #  split one visual row into multiple groups)
                y_key = round(word["top"] / 2) * 2
                if y_key not in rows_by_y:
                    rows_by_y[y_key] = []
                rows_by_y[y_key].append(word)

            # Process each row in top-to-bottom order
            for y_key in sorted(rows_by_y.keys()):
                row_words = sorted(rows_by_y[y_key], key=lambda w: w["x0"])

                # Assign each word to a column slot based on its x position
                # slots[0] = test name words, slots[1] = result, etc.
                slots = ["", "", "", "", ""]   # 5 slots for 5 columns

                for word in row_words:
                    x = word["x0"]
                    text = word["text"]

                    if x < 100:
                        # Column 0: Test name — words can be multi-word so we append
                        slots[0] = (slots[0] + " " + text).strip()
                    elif x < 200:
                        # Column 1: Result value
                        slots[1] = (slots[1] + " " + text).strip()
                    elif x < 280:
                        # Column 2: Unit (e.g. "U/L", "mmol/L", "μmol/L")
                        slots[2] = (slots[2] + " " + text).strip()
                    elif x < 360:
                        # Column 3: Flag — only "H" or "L" will appear here
                        slots[3] = (slots[3] + " " + text).strip()
                    else:
                        # Column 4: Reference range (e.g. "0.0-41.0")
                        slots[4] = (slots[4] + " " + text).strip()

                # Only keep rows that have something in both col 0 (test name)
                # AND col 1 (result value), and where col 1 contains a digit.
                # This filters out the page header, column label row, footer etc.
                import re
                if slots[0] and slots[1] and re.search(r"\d", slots[1]):
                    word_rows.append(slots)

    return {
        "header_text": header_text,
        "table_rows":  table_rows,
        "word_rows":   word_rows,
    }


# ==============================================================================
# SECTION 12 — parse_pdf_data() FUNCTION
#
# Takes the extracted data from extract_text_from_pdf() and turns it into
# two clean Python structures:
#   1. patient  — a dictionary with patient details (name, age, gender etc.)
#   2. results  — a list of test result dictionaries
#
# PARSER PRIORITY (chosen automatically):
#   1. parse_bs240_table()    — when bordered table rows were found (Layout A)
#   2. parse_bs240_words()    — when positional word rows were found (Layout B)
#   3. parse_generic_fallback() — last resort for unknown formats
# ==============================================================================

def parse_pdf_data(extracted):
    """
    Master parser. Accepts the dict returned by extract_text_from_pdf().
    Returns: (patient_dict, list_of_result_dicts)

    patient dict keys:
      name, age, gender, dob, id, req, physician,
      collected, received, reported, status, fasting,
      sample_type, blood_type, bed_no, priority, diagnosis

    result dict keys:
      test, full_name, result, unit, flag, remark, ref
    """

    import re

    # ── Unpack ────────────────────────────────────────────────────────────────
    if isinstance(extracted, dict):
        header_text = extracted.get("header_text", "")
        table_rows  = extracted.get("table_rows",  [])
        word_rows   = extracted.get("word_rows",   [])
    else:
        # Backwards compatibility: plain string passed in
        header_text = extracted
        table_rows  = []
        word_rows   = []

    # ── Helper: find a regex match in header_text ─────────────────────────────
    def find(pattern, default="N/A"):
        m = re.search(pattern, header_text, re.IGNORECASE)
        return m.group(1).strip() if m else default

    # ==========================================================================
    # PART A — PATIENT INFORMATION
    #
    # This parser handles BOTH BS240 header formats:
    #
    # FORMAT 1 (Old / Weija Gbawe style):
    #   Patient: Enoch Sipitey che 320   Gender: Male   Age: 65Year   Sample ID: 26
    #   Sample: Serum    MRN.:    Zone:    Bed No.:
    #   Sent from:   Priority: No
    #   Test Date:  12/18/2023 3:36:41 PM
    #
    # FORMAT 2 (New / Yaa Specimen style):
    #   Patient: YAA SPECIMEN          Sample ID: CHE900
    #   Patient ID: CHE900               Sample Type: Serum
    #   Date of Birth:                   Bar Code: CHE900
    #   Age: 68 Year                     Collection Date: 06/23/2025
    #   Gender: Female                   Collection Time:
    #   Physician:                       Department:
    #   Diagnosis:                       Comment:
    #   Ordering Date/Time: 06/23/2025   Print Date/Time: 06/23/2025 04:18:58 PM
    # ==========================================================================

    patient = {

        # Patient name — stops at 2+ spaces OR known next-column label
        # Handles both "Patient:" and "Patient ID:" (we want the first one)
        "name": find(
            r"Patient\s*:\s*([A-Za-z][A-Za-z0-9 '\-\.]+?)(?:\s{2,}|Sample\s*ID|Patient\s*ID|\n)",
            "Unknown"
        ),

        # Age — digits followed by optional "Year"/"Yr"
        "age": find(r"Age\s*:\s*(\d+\s*(?:Year|Yr|Y)?)", "N/A"),

        # Gender
        "gender": find(r"Gender\s*:\s*(Male|Female|Other|\w+)", "N/A"),

        # Date of Birth — Format 2 has "Date of Birth:" field
        # Matches dates like 01/15/1980 or 1980-01-15
        "dob": find(
            r"Date\s*of\s*Birth\s*:\s*([\d]{1,2}[\/\-][\d]{1,2}[\/\-][\d]{2,4}|[\d]{4}[\/\-][\d]{2}[\/\-][\d]{2})",
            "N/A"
        ),

        # Sample ID — both formats use "Sample ID:"
        "id": find(r"Sample\s*ID\s*:\s*(\S+)", "N/A"),

        # Patient ID — Format 2 has a separate "Patient ID:" field
        # We use this as req/requisition number if MRN is absent
        "req": find(r"Patient\s*ID\s*:\s*(\S+)",
               find(r"MRN\.?\s*:\s*((?!Zone|Bed|Sample|Sender|\s)\S+)", "N/A")),

        # Sample type — Format 1: "Sample: Serum", Format 2: "Sample Type: Serum"
        "sample_type": find(r"Sample\s*(?:Type)?\s*:\s*(\w+)", "N/A"),

        # Physician — Format 2 has "Physician:" field
        "physician": find(r"Physician\s*:\s*([A-Za-z][A-Za-z ,\.]+?)(?:\s{2,}|\n|Department)", "N/A"),

        # Diagnosis — Format 2 has "Diagnosis:" field
        "diagnosis": find(r"Diagnosis\s*:\s*([^\n]+?)(?:\s{2,}|Comment\s*:|\s*$)", "N/A"),

        # Blood type — only if explicitly filled in
        "blood_type": find(r"Blood\s*(?:Type|Group)\s*:\s*([ABO][B]?[+\-]?(?:pos|neg|itive|ative)?)", "N/A"),

        # Bed number (old format only)
        "bed_no": find(r"Bed\s*No\.?\s*:\s*((?!Sender|Sent|\s)\S+)", "N/A"),

        # Priority (old format only)
        "priority": find(r"Priority\s*:\s*(\S+)", "N/A"),

        # Collection date —
        #   Format 1: "Test Date: 12/18/2023 3:36:41 PM"
        #   Format 2: "Collection Date: 06/23/2025" + "Ordering Date/Time: 06/23/2025"
        "collected": find(
            r"(?:Collection\s*Date|Ordering\s*Date(?:/Time)?|Test\s*Date)\s*[:/]\s*([\d/]+(?:\s+[\d:]+\s*(?:AM|PM)?)?)",
            "N/A"
        ),

        # Received date —
        #   Format 1: "Send Date: 12/18/2023 3:36:41 PM"
        #   Format 2: "Test Date/Time: 06/23/2025"
        "received": find(
            r"(?:Test\s*Date(?:/Time)?|Send\s*Date)\s*[:/]\s*([\d/]+(?:\s+[\d:]+\s*(?:AM|PM)?)?)",
            "N/A"
        ),

        # Reported / Print date
        "reported": find(
            r"Print\s*Date(?:/Time)?\s*[:/]\s*([\d/]+\s+[\d:]+\s*(?:AM|PM)?)",
            "N/A"
        ),

        # Fasting status — not present in either BS240 format
        "fasting": "N/A",

        # Always Final for a completed BS240 printout
        "status": "Final",
    }

    # ── Clean up age (remove "Year"/"Yr" suffix) ──────────────────────────────
    age_raw   = patient["age"]
    age_clean = re.sub(r"[^\d]", "", age_raw)
    patient["age"] = age_clean if age_clean else age_raw

    # ── Clean up name (strip trailing spaces) ─────────────────────────────────
    patient["name"] = patient["name"].strip()

    # ── Clean up physician — "N/A" if only whitespace was captured ───────────
    if patient["physician"].strip() in ("", "N/A"):
        patient["physician"] = "N/A"

    # ==========================================================================
    # PART B — CHOOSE RESULTS PARSER
    #
    # Priority:
    #   1. Bordered-table rows  (Layout A — old Weija Gbawe style)
    #   2. Positional word rows (Layout B — new Yaa Specimen / CHE900 style)
    #   3. Generic text fallback (unknown formats)
    # ==========================================================================

    # Count usable table rows (must have a number in the Concentration column)
    data_row_count = sum(
        1 for row in table_rows
        if len(row) >= 4 and re.search(r"\d", row[2] if len(row) > 2 else "")
    )

    # Count usable word rows (must have a number in the Result slot)
    word_row_count = sum(
        1 for row in word_rows
        if len(row) >= 2 and re.search(r"\d", row[1])
    )

    if data_row_count >= 1:
        print(f"   Using BS240 table parser  ({data_row_count} bordered-table rows)")
        results = parse_bs240_table(table_rows)

    elif word_row_count >= 1:
        print(f"   Using BS240 word parser   ({word_row_count} positional word rows)")
        results = parse_bs240_words(word_rows)

    else:
        print("   No structured data found — using generic text fallback parser")
        results = parse_generic_fallback(header_text)

    # ==========================================================================
    # PART C — ENRICH RESULTS
    # After parsing, run every result through two enrichment steps:
    #
    #   1. FULL NAME MAPPING
    #      The BS240 prints short codes like "T-bil", "γ-GT", "CREA".
    #      We replace them with their proper clinical full names so the report
    #      reads clearly to patients and clinicians alike.
    #      e.g. "T-bil" → "Total Bilirubin"
    #           "γ-GT"  → "Gamma-Glutamyl Transferase (γ-GT)"
    #           "CREA"  → "Creatinine"
    #
    #   2. FALLBACK REFERENCE RANGES
    #      Some tests have no reference range printed on the BS240 report
    #      (e.g. VLDL, LDL, K+, Na+, Cl-).  We supply standard clinical
    #      reference ranges for these so the report is complete.
    #      A range from the PDF always takes priority — we only fill in when
    #      the field is blank.
    # ==========================================================================

    results = enrich_results(results)

    # ==========================================================================
    # PART D — CALCULATE eGFR (CKD-EPI 2021 formula)
    #
    # eGFR (estimated Glomerular Filtration Rate) measures how well the kidneys
    # are filtering waste from the blood. It is NOT printed on the BS240 report
    # because the analyser only measures Creatinine — eGFR must be calculated.
    #
    # FORMULA: CKD-EPI 2021 (race-free version, the current clinical standard)
    #
    #   eGFR = 142 × min(Scr/κ, 1)^α × max(Scr/κ, 1)^(-1.200)
    #          × 0.9938^Age × (1.012 if Female)
    #
    #   Where:
    #     Scr = serum creatinine in mg/dL
    #     κ   = 0.7 (Female) or 0.9 (Male)
    #     α   = -0.241 (Female) or -0.302 (Male)
    #     Age = patient age in years
    #
    # UNIT CONVERSION:
    #   The BS240 reports creatinine in µmol/L.
    #   To convert to mg/dL: divide by 88.42
    #
    # REFERENCE RANGES (KDIGO staging):
    #   >=90         Normal or high (G1)
    #   60-89        Mildly decreased (G2)
    #   45-59        Mildly to moderately decreased (G3a)
    #   30-44        Moderately to severely decreased (G3b)
    #   15-29        Severely decreased (G4)
    #   <15          Kidney failure (G5)
    #   We report simply: >=60 as the lower acceptable bound.
    # ==========================================================================

    results = calculate_egfr(results, patient)

    return patient, results


# ==============================================================================
# HELPER — calculate_egfr()
#
# Calculates eGFR using the CKD-EPI 2021 (race-free) formula and inserts it
# as a new result into the KFT panel, directly after Creatinine.
#
# This runs automatically whenever Creatinine, Age, and Gender are available.
# If any of those are missing, it silently skips the calculation.
# ==============================================================================

def calculate_egfr(results, patient):
    """
    Calculates eGFR (CKD-EPI 2021) from Creatinine + patient Age + Gender.
    Inserts the eGFR result directly after the Creatinine row.

    Args:
        results  — list of result dicts from the parser
        patient  — patient info dict (needs 'age' and 'gender')

    Returns:
        Updated results list with eGFR inserted after Creatinine (if calculable)
    """
    import math

    # ── Step 1: find the creatinine result ────────────────────────────────────
    crea_idx   = None    # Index of the Creatinine row in results
    crea_val   = None    # Creatinine value (float)
    crea_unit  = ""      # Unit as printed (µmol/L or mg/dL)

    for i, r in enumerate(results):
        # Match by test name — the full name has been applied by enrich_results()
        # so we look for "Creatinine" (case-insensitive)
        if "creatinine" in r["test"].lower() and "ratio" not in r["test"].lower():
            try:
                crea_val  = float(r["result"])
                crea_unit = r.get("unit", "").strip()
                crea_idx  = i
                break
            except (ValueError, TypeError):
                pass   # Non-numeric result — skip

    if crea_val is None or crea_idx is None:
        # Creatinine not found or not numeric — can't calculate eGFR
        return results

    # ── Step 2: get patient age ───────────────────────────────────────────────
    try:
        age = int(patient.get("age", "0") or "0")
    except (ValueError, TypeError):
        age = 0

    if age <= 0:
        # Age unknown — can't calculate eGFR
        return results

    # ── Step 3: get patient gender ────────────────────────────────────────────
    gender_raw = (patient.get("gender") or "").strip().lower()
    is_female  = gender_raw in ("female", "f", "woman", "w")
    # Default to Male constants if gender is unknown

    # ── Step 4: convert creatinine to mg/dL if needed ────────────────────────
    # The BS240 typically reports in µmol/L.
    # CKD-EPI formula requires mg/dL.
    # Conversion: 1 mg/dL = 88.42 µmol/L
    unit_lower = crea_unit.lower()
    if "µmol" in unit_lower or "umol" in unit_lower or "μmol" in unit_lower:
        # Convert µmol/L → mg/dL
        scr = crea_val / 88.42
    elif "mg" in unit_lower:
        # Already in mg/dL
        scr = crea_val
    else:
        # Unit unknown — assume µmol/L (most common for BS240)
        scr = crea_val / 88.42

    # ── Step 5: apply CKD-EPI 2021 formula ───────────────────────────────────
    # Constants differ by sex:
    #   κ (kappa) = sex-specific Scr normalisation factor
    #   α (alpha) = exponent for the first term
    if is_female:
        kappa = 0.7
        alpha = -0.241
        sex_factor = 1.012   # Multiplier for female patients
    else:
        kappa = 0.9
        alpha = -0.302
        sex_factor = 1.0     # No multiplier for male patients

    # Calculate the two ratio terms:
    #   ratio = Scr / κ
    #   term1 = min(ratio, 1) raised to the power alpha
    #   term2 = max(ratio, 1) raised to the power -1.200
    ratio = scr / kappa
    term1 = min(ratio, 1.0) ** alpha
    term2 = max(ratio, 1.0) ** (-1.200)

    # Age term: 0.9938 raised to the power of Age
    age_term = 0.9938 ** age

    # Full formula
    egfr = 142 * term1 * term2 * age_term * sex_factor

    # Round to 1 decimal place
    egfr_rounded = round(egfr, 1)

    # ── Step 6: determine eGFR flag ──────────────────────────────────────────
    # Normal lower bound is >=60 mL/min/1.73m2
    # Values below 60 indicate chronic kidney disease
    if egfr_rounded < 60:
        egfr_flag = "Low"    # Below normal — flag as Low
    else:
        egfr_flag = ""       # Normal range — no flag

    # ── Step 7: determine CKD stage note ─────────────────────────────────────
    # Included as a brief interpretation note shown under the eGFR row
    if egfr_rounded >= 90:
        stage_note = "G1: Normal or High"
    elif egfr_rounded >= 60:
        stage_note = "G2: Mildly Decreased"
    elif egfr_rounded >= 45:
        stage_note = "G3a: Mildly-Moderately Decreased"
    elif egfr_rounded >= 30:
        stage_note = "G3b: Moderately-Severely Decreased"
    elif egfr_rounded >= 15:
        stage_note = "G4: Severely Decreased"
    else:
        stage_note = "G5: Kidney Failure — seek urgent care"

    # ── Step 8: build the eGFR result dict ───────────────────────────────────
    egfr_result = {
        "test":      "eGFR (CKD-EPI 2021)",
        "full_name": "Estimated Glomerular Filtration Rate",
        "result":    str(egfr_rounded),
        "unit":      "mL/min/1.73m<super>2</super>",
        "flag":      egfr_flag,
        "remark":    stage_note,   # CKD stage shown in the Remark column
        "ref":       ">=60",
        # "notes" must be a LIST of strings — build_panel iterates over it.
        # A plain string would be iterated character-by-character → PDF corruption.
        "notes": [
            f"CKD-EPI 2021  |  Stage {stage_note}  |  "
            f"From: Creatinine {crea_val} {crea_unit}, Age {age} yrs, Sex: {'Female' if is_female else 'Male'}"
        ],
    }

    # ── Step 9: insert after Creatinine ──────────────────────────────────────
    # list.insert(index, item) inserts BEFORE the given index
    # So inserting at crea_idx + 1 places it directly after Creatinine
    results = list(results)   # make a mutable copy
    results.insert(crea_idx + 1, egfr_result)

    return results


# ==============================================================================
# HELPER — enrich_results()
#
# Applies full-name mapping and fallback reference ranges to every result.
# Called at the end of parse_pdf_data() so it covers all three parsers.
# ==============================================================================

def enrich_results(results):
    """
    Enriches a list of result dicts in-place by:
      - Replacing short test codes with proper full clinical names
      - Filling in reference ranges for tests where the PDF left them blank

    The mapping keys are lower-case so matching is case-insensitive.
    Returns the enriched list.
    """

    # ── Full name mapping ──────────────────────────────────────────────────────
    # Keys:   lower-case version of the raw test code from the PDF
    # Values: full clinical name to display in the report
    #
    # Add new entries here whenever you encounter a new test code.
    FULL_NAMES = {
        # ── Liver Function Tests (LFT) ─────────────────────────────────────────
        "t-bil":        "Total Bilirubin",
        "tbil":         "Total Bilirubin",
        "tb":           "Total Bilirubin",
        "d-bil":        "Direct Bilirubin",
        "dbil":         "Direct Bilirubin",
        "d-bil":        "Direct Bilirubin",
        "ibil":         "Indirect Bilirubin",
        "ind bill":     "Indirect Bilirubin",
        "ind-bil":      "Indirect Bilirubin",
        "alt":          "Alanine Aminotransferase (ALT)",
        "ast":          "Aspartate Aminotransferase (AST)",
        "ast/alt":      "AST/ALT Ratio",
        "alp":          "Alkaline Phosphatase (ALP)",
        "γ-gt":         "Gamma-Glutamyl Transferase (γ-GT)",
        "ggt":          "Gamma-Glutamyl Transferase (GGT)",
        "gamma gt":     "Gamma-Glutamyl Transferase (GGT)",
        "alb":          "Albumin",
        "tp-o":         "Total Protein",
        "tp":           "Total Protein",
        "glob":         "Globulin",

        # ── Kidney Function Tests (KFT) ────────────────────────────────────────
        "crea":         "Creatinine",
        "creatinine":   "Creatinine",
        "urea":         "Urea (BUN)",
        "bun":          "Blood Urea Nitrogen (BUN)",
        "egfr":         "eGFR (Estimated Glomerular Filtration Rate)",
        "na+":          "Sodium (Na<super>+</super>)",
        "na":           "Sodium (Na<super>+</super>)",
        "k+":           "Potassium (K<super>+</super>)",
        "k":            "Potassium (K<super>+</super>)",
        "cl-":          "Chloride (Cl<super>-</super>)",
        "cl":           "Chloride (Cl<super>-</super>)",
        "ca":           "Calcium",
        "ca2+":         "Calcium (Ca<super>2+</super>)",
        "phos":         "Phosphate",
        "mg":           "Magnesium",
        "uric acid":    "Uric Acid",
        "ua":           "Uric Acid",

        # ── Lipid Profile ──────────────────────────────────────────────────────
        "tc":           "Total Cholesterol",
        "tg":           "Triglycerides",
        "hdl-c":        "HDL Cholesterol (Good)",
        "hdl":          "HDL Cholesterol (Good)",
        "ldl-c":        "LDL Cholesterol (Bad)",
        "ldl":          "LDL Cholesterol (Bad)",
        "vldl":         "VLDL Cholesterol",
        "coronary r":   "Coronary Risk Ratio",
        "coronary risk":"Coronary Risk Ratio",

        # ── Glycaemic Tests ────────────────────────────────────────────────────
        "hba1c":        "Glycated Haemoglobin (HbA1c)",
        "hemoglobin a1c":"Glycated Haemoglobin (HbA1c)",
        "haemoglobin a1c":"Glycated Haemoglobin (HbA1c)",
        "a1c":          "Glycated Haemoglobin (HbA1c)",
        "glucose":      "Fasting Blood Glucose",
        "fbg":          "Fasting Blood Glucose",
        "rbs":          "Random Blood Sugar",

        # ── Old Weija Gbawe format full names (already verbose — keep as-is) ──
        # These arrive as "ALT — ALANINE AMINOTRANS" from parse_bs240_table()
        # The mapping handles the short code before the " — " is added.
    }

    # ── Fallback reference ranges ──────────────────────────────────────────────
    # Keys:   lower-case test code OR lower-case full name (after mapping above)
    # Values: reference range string to show when the PDF left it blank
    #
    # Sources: standard adult reference ranges (SI units where applicable).
    # The PDF's own range always takes priority — this only fills blank fields.
    FALLBACK_REFS = {
        # Liver
        "ast/alt ratio":                    "0.8 - 1.4",
        "indirect bilirubin":               "0.00 - 17.00",
        "ibil":                             "0.00 - 17.00",

        # Kidney / Electrolytes
        "sodium (na<super>+</super>)":  "135.0 - 150.0",
        "potassium (k<super>+</super>)": "3.5 - 5.0",
        "chloride (cl<super>-</super>)": "98.0 - 107.0",
        "calcium":                          "2.15 - 2.55",
        "calcium (ca<super>2+</super>)": "2.15 - 2.55",
        "phosphate":                        "0.81 - 1.45",
        "magnesium":                        "0.75 - 1.00",

        # Lipid
        "vldl cholesterol":                 "0.10 - 1.00",
        "ldl cholesterol (bad)":            "<=3.40",
        "hdl cholesterol (good)":           ">=1.04",
        "coronary risk ratio":              "<=5.0",

        # Glycaemic
        "fasting blood glucose":            "3.90 - 6.10",
        "random blood sugar":               "3.90 - 7.80",
        "glycated haemoglobin (hba1c)":     "20.0 - 42.0",
    }

    enriched = []
    for r in results:
        r = dict(r)   # make a copy so we don't mutate the original

        # ── Step 1: apply full name mapping ──────────────────────────────────
        # The raw test name from the PDF (e.g. "T-bil", "γ-GT", "Na+")
        raw_name  = r["test"].strip()
        lookup    = raw_name.lower()

        # Check if a full name exists for this code
        if lookup in FULL_NAMES:
            r["test"] = FULL_NAMES[lookup]
        # else: keep the original name — no mapping defined for this code yet

        # ── Step 2: apply fallback reference range if blank ───────────────────
        if not r.get("ref", "").strip():
            # Try matching the (possibly already-mapped) full name
            full_lower = r["test"].lower()
            if full_lower in FALLBACK_REFS:
                r["ref"] = FALLBACK_REFS[full_lower]
            else:
                # Also try the raw original code in case the name wasn't mapped
                if lookup in FALLBACK_REFS:
                    r["ref"] = FALLBACK_REFS[lookup]

        enriched.append(r)

    return enriched


# ==============================================================================
# HELPER — parse_bs240_table()
#
# Used for Layout A PDFs that have real table borders.
# Columns arrive already separated by pdfplumber's extract_tables():
#   [0] Test code  [1] Full Name  [2] Concentration  [3] Unit
#   [4] Results (Normal/high/low)  [5] Remark (RFH/RFL)  [6] Reference
# ==============================================================================

def parse_bs240_table(table_rows):
    """
    Parses bordered-table rows (Layout A — old Weija Gbawe format).
    Returns a list of result dicts.
    """
    import re
    results = []

    HEADER_WORDS = {
        "test", "full name", "concentration", "unit", "results",
        "remark", "reference", "result", "fullname", "chemistry",
    }

    for row in table_rows:
        if len(row) < 4:
            continue

        test_code  = row[0] if len(row) > 0 else ""
        full_name  = row[1] if len(row) > 1 else ""
        conc       = row[2] if len(row) > 2 else ""
        unit       = row[3] if len(row) > 3 else ""
        result_txt = row[4] if len(row) > 4 else ""
        remark     = row[5] if len(row) > 5 else ""
        ref        = row[6] if len(row) > 6 else ""

        if test_code.lower().strip() in HEADER_WORDS:
            continue

        conc_clean = conc.strip()
        if not re.search(r"\d", conc_clean):
            continue

        result_lower = result_txt.strip().lower()
        if result_lower == "high" or remark.upper() in ("RFH",):
            flag = "High"
        elif result_lower == "low" or remark.upper() in ("RFL",):
            flag = "Low"
        else:
            flag = ""

        if full_name and full_name.strip() and full_name.strip() != test_code.strip():
            display_name = f"{test_code.strip()} — {full_name.strip()}"
        else:
            display_name = test_code.strip()

        results.append({
            "test":      display_name,
            "full_name": full_name.strip(),
            "result":    conc_clean,
            "unit":      unit.strip(),
            "flag":      flag,
            "remark":    remark.strip(),
            "ref":       ref.strip().replace("–", "-").replace("—", "-"),
        })

    return results


# ==============================================================================
# HELPER — parse_bs240_words()
#
# Used for Layout B PDFs — no table borders, columns identified by x-position.
# Each row in word_rows is already a 5-slot list built in extract_text_from_pdf:
#   [0] test_name  [1] result  [2] unit  [3] flag (H/L/"")  [4] ref_range
#
# This format comes from the "Yaa Specimen / CHE900" style PDF where:
#   - Flag column contains just "H" or "L" (not "Normal"/"high"/"low")
#   - Some rows have no unit and no ref range (e.g. AST/ALT ratio, CORONARY R)
#   - Test names can contain special chars: γ-GT, K+, Cl-, Na+, T-bil
# ==============================================================================

def parse_bs240_words(word_rows):
    """
    Parses positional word rows (Layout B — new CHE900 / Yaa Specimen format).
    Returns a list of result dicts.
    """
    import re
    results = []

    # Words that identify header rows or footer rows — skip these
    SKIP_NAMES = {
        "chemistry", "result", "unit", "flag", "ref", "range",
        "ordering", "operator", "reviewer", "page", "the",
        "print", "test", "collection", "patient", "gender",
        "age", "physician", "diagnosis", "sample", "date",
        "ordering date/time:", "print date/time:", "test date/time:",
        "send date/time:", "the result", "the results",
    }

    # Also skip any row where the test name contains "date" or "time" or "page"
    # These are footer/header lines, not test results
    DATE_TIME_SKIP = re.compile(r"(?:date|time|page|operator|reviewer|result\s+are)", re.IGNORECASE)

    for row in word_rows:
        if len(row) < 2:
            continue

        test_name  = row[0].strip()
        result_val = row[1].strip()
        unit       = row[2].strip() if len(row) > 2 else ""
        flag_raw   = row[3].strip() if len(row) > 3 else ""
        ref        = row[4].strip() if len(row) > 4 else ""

        # Skip header/footer rows — by exact name match
        if test_name.lower() in SKIP_NAMES:
            continue
        # Skip any row whose test name looks like a date/time label or footer line
        if DATE_TIME_SKIP.search(test_name):
            continue

        # Skip if result has no digit
        if not re.search(r"\d", result_val):
            continue

        # Normalise flag:
        #   Layout B uses "H" or "L" in the Flag column
        #   Layout A uses "Normal"/"high"/"low" in the Results column
        flag_upper = flag_raw.upper()
        if flag_upper in ("H", "HIGH", "RFH"):
            flag = "High"
        elif flag_upper in ("L", "LOW", "RFL"):
            flag = "Low"
        else:
            flag = ""

        results.append({
            "test":      test_name,
            "full_name": "",       # Layout B doesn't have a separate Full Name column
            "result":    result_val,
            "unit":      unit,
            "flag":      flag,
            "remark":    flag_raw,  # Store original H/L for reference
            "ref":       ref.replace("–", "-"),
        })

    return results


# ==============================================================================
# HELPER — parse_generic_fallback()
#
# Last-resort parser used when neither table extraction nor positional word
# grouping produces results. Scans every line of plain text with regex.
# Less accurate but works for unknown analyser formats and simple text PDFs.
# ==============================================================================

def parse_generic_fallback(text):
    """
    Regex line-scanner fallback. Returns result dicts in the same format
    as parse_bs240_table() and parse_bs240_words().
    """
    import re
    results = []

    SKIP_WORDS = {
        "test", "full name", "concentration", "unit", "results", "chemistry",
        "remark", "reference", "patient", "sample", "gender", "flag", "ref",
        "sender", "sent", "diagnosis", "test date", "send date", "print date",
        "ordering", "operator", "reviewer", "collection", "physician",
    }

    # Pattern: TEST_NAME   NUMBER   UNIT   (H|L|flag)   (ref_range)
    pattern = re.compile(
        r"^([A-Za-zγµ][A-Za-z0-9 /\-\(\)\.\+γµ]+?)\s+"
        r"([<>]?\d+\.?\d*)\s*"
        r"([µμa-zA-Z/%\.]+(?:/[a-zA-Z\.]+)?)?\s*"
        r"(Normal|high|low|H\b|L\b|RFH|RFL|\*)?\s*"
        r"([<>=]{0,2}[\d\.\-–]+(?:\s*[-–]\s*[\d\.]+)?)?",
        re.MULTILINE | re.IGNORECASE
    )

    for m in pattern.finditer(text):
        test_name  = m.group(1).strip()
        result_val = m.group(2).strip()
        unit       = m.group(3).strip() if m.group(3) else ""
        flag_word  = m.group(4).strip() if m.group(4) else ""
        ref        = m.group(5).strip() if m.group(5) else ""

        if test_name.lower() in SKIP_WORDS or len(test_name) < 2:
            continue

        flag_upper = flag_word.upper()
        if flag_upper in ("HIGH", "H", "RFH"):
            flag = "High"
        elif flag_upper in ("LOW", "L", "RFL"):
            flag = "Low"
        else:
            flag = ""

        results.append({
            "test":      test_name,
            "full_name": "",
            "result":    result_val,
            "unit":      unit,
            "flag":      flag,
            "remark":    flag_word,
            "ref":       ref.replace("–", "-"),
        })

    return results




# ==============================================================================
# SECTION 13 — main() FUNCTION
# This is the entry point — the first function called when you run the script.
# It reads what you typed in the terminal and decides what to do.
# ==============================================================================

def main():
    """
    Reads command-line arguments and runs the appropriate mode:
      --demo      → generates a sample report using the built-in demo data
      input_pdf   → reads a real PDF and generates a report from it
    """

    # Set up the argument parser — this reads what you type after "python script.py"
    # formatter_class=RawDescriptionHelpFormatter preserves the formatting of the epilog
    parser = argparse.ArgumentParser(
        description="Lab Report Generator — LFT, KFT, Lipid, HbA1c, Uric Acid",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        # epilog = extra text shown at the bottom of the --help page
        epilog="""
Examples:
  Run with demo data:
    python lab_report_5panels_commented.py --demo

  Run with a real PDF from your analyser:
    python lab_report_5panels_commented.py "patient_results.pdf" "My Lab Name"

  Full example with all options:
    python lab_report_5panels_commented.py "results.pdf" "City Hospital Lab"
        --address "123 Street, Accra"
        --phone "+233 XX XXX XXXX"
        --director "Bio. Mensah"
        --output "report.pdf"
        """
    )

    # ── Positional arguments — typed WITHOUT a -- prefix, in order ────────────
    # nargs="?" means "optional" — if you don't type it, the value becomes None
    parser.add_argument(
        "input_pdf",
        nargs="?",   # Optional — don't have to provide it
        help="Path to the lab PDF exported from your analyser (e.g. BS240)"
        # Example: 'CHES10501 HANNAN SHERIFF.pdf'
    )
    parser.add_argument(
        "lab_name",
        nargs="?",   # Optional
        help="Your lab's name to show in the report header"
        # Example: 'City Hospital Lab'
    )

    # ── Optional named arguments — typed WITH a -- prefix ─────────────────────
    # default=None means if not provided, the value is None (we check for this later)
    parser.add_argument("--address",  default=None, help="Your lab's address")
    parser.add_argument("--phone",    default=None, help="Your lab's phone number")
    parser.add_argument("--director", default=None, help="Laboratory scientist's name")
    parser.add_argument("--output",   default=None, help="Output PDF filename (auto-named if not set)")

    # action="store_true" means: if --demo is present, args.demo = True, otherwise False
    parser.add_argument("--demo", action="store_true", help="Generate report with built-in demo data")

    # Actually parse the arguments the user typed
    # args now holds all the values, e.g. args.demo, args.input_pdf, args.output
    args = parser.parse_args()


    # ── MODE 1: Demo mode ─────────────────────────────────────────────────────
    # If the user typed --demo, use the built-in DEMO_PATIENT and DEMO_RESULTS
    if args.demo:
        print("🔬 Generating demo 5-panel lab report...")

        # Use the provided output name, or fall back to a default filename
        # "or" returns the right side if the left side is None (falsy)
        out = args.output or "lab_report_5panels_demo.pdf"

        generate_report(DEMO_PATIENT, DEMO_RESULTS, DEMO_LABS, out)
        return   # Stop here — don't fall through to the PDF mode below


    # ── MODE 2: Real PDF mode ─────────────────────────────────────────────────
    # If the user provided a PDF file path, read and process it
    if args.input_pdf:

        # "global" lets us modify the module-level variables LAB_NAME etc.
        # Without "global", changes inside a function would be local only and
        # the original values at the top of the file wouldn't be updated.
        global LAB_NAME, LAB_ADDRESS, LAB_PHONE, LAB_DIRECTOR

        # Override the lab details if the user provided them on the command line
        # "if args.lab_name:" checks whether lab_name was provided (not None)
        if args.lab_name:
            LAB_NAME = args.lab_name       # e.g. "labhanan"
        if args.address:
            LAB_ADDRESS = args.address     # e.g. "123 Street, Accra"
        if args.phone:
            LAB_PHONE = args.phone         # e.g. "+233 20 000 0000"
        if args.director:
            LAB_DIRECTOR = args.director   # e.g. "Dr. Mensah"

        # Tell the user what's happening
        print(f"📄 Reading PDF: {args.input_pdf}")
        print(f"🏥 Lab name:   {LAB_NAME}")

        # Try to extract text from the PDF
        # "try/except" = attempt something, and if it fails, handle the error gracefully
        try:
            extracted = extract_text_from_pdf(args.input_pdf)

        except FileNotFoundError:
            # FileNotFoundError happens when Python can't find the file at that path
            print(f"❌ File not found: {args.input_pdf}")
            print("   Check that the path is correct and the file exists.")
            print("   Tip: On Windows, use quotes around paths with spaces.")
            return   # Stop the script

        except Exception as e:
            # "Exception as e" catches any other kind of error
            # "e" contains the error message
            print(f"❌ Could not read PDF: {e}")
            return

        # Parse the extracted text into patient info and test results
        patient, results = parse_pdf_data(extracted)

        # If no results were found, the PDF layout probably didn't match our patterns
        if not results:
            print("⚠️  No test results could be parsed from this PDF.")
            print("   The PDF layout may not match the expected format.")
            print("   Try --demo first to confirm the script works correctly.")
            print("\n── First 600 characters of extracted text ──")
            # Show raw text and table rows to help diagnose parsing failures
            header_preview = extracted.get("header_text", "")[:600] if isinstance(extracted, dict) else str(extracted)[:600]
            print(header_preview)
            if isinstance(extracted, dict):
                print("\n── Table rows found ──")
                for row in extracted.get("table_rows", [])[:8]:
                    print(row)
            return

        # Tell the user how many results were found
        print(f"✅ Found {len(results)} test results")
        print(f"   Patient: {patient['name']}")

        # Build the labs list using the (possibly updated) global config
        labs = [{
            "code":     "01",
            "name":     LAB_NAME,
            "address":  LAB_ADDRESS,
            "phone":    LAB_PHONE,
            "director": LAB_DIRECTOR,
        }]

        # Decide on the output filename
        if args.output:
            out = args.output   # Use whatever the user specified with --output
        else:
            # Auto-generate a filename from the patient's name
            # .replace(",", "") removes commas: "Zedni, Salim" → "Zedni  Salim"
            # .replace(" ", "_") replaces spaces with underscores: "Zedni  Salim" → "Zedni__Salim"
            # .strip("_") removes any leading/trailing underscores
            safe_name = patient["name"].replace(",", "").replace(" ", "_").strip("_")
            out = f"{safe_name}_report.pdf"
            # Result: "Zedni_Salim_report.pdf"

        generate_report(patient, results, labs, out)
        return


    # ── No arguments given ────────────────────────────────────────────────────
    # If the user ran the script with no arguments at all, show the help message
    parser.print_help()



# ==============================================================================
# SECTION 15 — GRAPHICAL USER INTERFACE (GUI)
#
# A drag-and-drop desktop application built with tkinter.
# This section is only active when the script is run directly with no arguments.
#
# HOW TO RUN THE GUI:
#   python lab_report_5panels_commented.py        ← opens the GUI window
#
# HOW TO RUN IN CLI MODE:
#   python lab_report_5panels_commented.py --demo
#   python lab_report_5panels_commented.py patient.pdf "Lab Name"
#
# REQUIREMENTS:
#   pip install reportlab pdfplumber tkinterdnd2   ← tkinterdnd2 is optional
# ==============================================================================

import os
import threading
import sys

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter import font as tkfont
    TK_AVAILABLE = True
except ImportError:
    TK_AVAILABLE = False
    tk = None

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    TkinterDnD = None


# ==============================================================================
# COLOUR THEME
# All GUI colours in one place — edit here to restyle the whole app at once.
# ==============================================================================

THEME = {
    "bg":           "#1e1e2e",
    "card":         "#2a2a3e",
    "border":       "#3a3a5c",
    "accent":       "#7c6af7",
    "accent_hover": "#9d8fff",
    "accent_dark":  "#5a4fd4",
    "text":         "#e0e0f0",
    "text_dim":     "#8888aa",
    "text_bright":  "#ffffff",
    "success":      "#4caf84",
    "error":        "#e05555",
    "warning":      "#e0a040",
    "drop_idle":    "#252538",
    "drop_hover":   "#2a3a2a",
    "drop_border":  "#7c6af7",
    "input_bg":     "#16162a",
    "input_border": "#44446a",
}


# ==============================================================================
# APP CLASS
# The main window. All GUI logic lives here.
# ==============================================================================

def _build_gui_app():
    """
    Creates and returns the App class.
    Defined inside a function so the class body is only evaluated when
    tkinter is confirmed available — avoids errors on headless servers.
    """

    BaseClass = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk

    class App(BaseClass):

        def __init__(self):
            super().__init__()
            self.title("Lab Report Generator")
            self.geometry("780x860")
            self.minsize(680, 720)
            self.configure(bg=THEME["bg"])
            self.resizable(True, True)

            # ── State variables ───────────────────────────────────────────────
            self.input_pdfs    = []   # List of PDF paths to process (multiple supported)
            self.lab_name_var   = tk.StringVar()
            self.lab_addr_var   = tk.StringVar()
            self.lab_phone_var  = tk.StringVar()
            self.lab_dir_var    = tk.StringVar()
            self.output_dir_var = tk.StringVar()
            self.filename_var   = tk.StringVar(value="auto")

            self._build_ui()
            self._load_settings()

            # Centre window on screen
            self.update_idletasks()
            x = (self.winfo_screenwidth()  - self.winfo_width())  // 2
            y = (self.winfo_screenheight() - self.winfo_height()) // 2
            self.geometry(f"+{x}+{y}")

        # ──────────────────────────────────────────────────────────────────────
        # UI CONSTRUCTION
        # ──────────────────────────────────────────────────────────────────────

        def _build_ui(self):
            # Outer scrollable canvas so the window can be made smaller
            self.canvas = tk.Canvas(self, bg=THEME["bg"], highlightthickness=0)
            scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
            self.canvas.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side="right", fill="y")
            self.canvas.pack(side="left", fill="both", expand=True)

            self.inner = tk.Frame(self.canvas, bg=THEME["bg"])
            self.canvas_window = self.canvas.create_window(
                (0, 0), window=self.inner, anchor="nw"
            )
            self.inner.bind("<Configure>",   self._on_frame_configure)
            self.canvas.bind("<Configure>",  self._on_canvas_configure)
            self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

            self._build_header()
            self._build_drop_zone()
            self._build_lab_details()
            self._build_output_section()
            self._build_generate_btn()
            self._build_status_bar()

        def _build_header(self):
            hdr = tk.Frame(self.inner, bg=THEME["accent_dark"], pady=14)
            hdr.pack(fill="x")

            tk.Label(
                hdr, text="🔬  Lab Report Generator",
                font=("Segoe UI", 18, "bold"),
                bg=THEME["accent_dark"], fg=THEME["text_bright"]
            ).pack()

            tk.Label(
                hdr,
                text="LFT  ·  KFT (+ eGFR)  ·  Lipid Profile  ·  HbA1c  ·  Uric Acid",
                font=("Segoe UI", 9),
                bg=THEME["accent_dark"], fg=THEME["text_dim"]
            ).pack()

        def _build_drop_zone(self):
            self._section_label("1.  Input PDFs  (add one or more patient files)")

            # ── Drop target area ──────────────────────────────────────────────
            self.drop_frame = tk.Frame(
                self.inner, bg=THEME["drop_idle"],
                highlightbackground=THEME["drop_border"],
                highlightthickness=2, pady=16, padx=20
            )
            self.drop_frame.pack(fill="x", padx=16, pady=(0, 6))

            self.drop_icon = tk.Label(
                self.drop_frame, text="📄",
                font=("Segoe UI", 26),
                bg=THEME["drop_idle"], fg=THEME["text_dim"]
            )
            self.drop_icon.pack()

            self.drop_label = tk.Label(
                self.drop_frame,
                text="Drag & drop PDFs here  (multiple files supported)" if DND_AVAILABLE
                     else "Use the buttons below to add patient PDF files",
                font=("Segoe UI", 10),
                bg=THEME["drop_idle"], fg=THEME["text_dim"]
            )
            self.drop_label.pack(pady=(4, 6))

            btn_row = tk.Frame(self.drop_frame, bg=THEME["drop_idle"])
            btn_row.pack()
            self._make_button(btn_row, "➕  Add PDF(s)…",
                              self._browse_input, style="secondary").pack(side="left", padx=4)
            self._make_button(btn_row, "🗑  Remove Selected",
                              self._remove_selected, style="secondary").pack(side="left", padx=4)
            self._make_button(btn_row, "✖  Clear All",
                              self._clear_all_pdfs, style="secondary").pack(side="left", padx=4)

            # Register drag-and-drop if available
            if DND_AVAILABLE:
                for widget in [self.drop_frame, self.drop_icon, self.drop_label]:
                    widget.drop_target_register(DND_FILES)
                    widget.dnd_bind("<<Drop>>",       self._on_drop)
                    widget.dnd_bind("<<DragEnter>>",  self._on_drag_enter)
                    widget.dnd_bind("<<DragLeave>>",  self._on_drag_leave)

            # ── File list box ─────────────────────────────────────────────────
            list_frame = tk.Frame(self.inner, bg=THEME["bg"])
            list_frame.pack(fill="x", padx=16, pady=(0, 8))

            self.file_listbox = tk.Listbox(
                list_frame,
                font=("Segoe UI", 9),
                bg=THEME["input_bg"], fg=THEME["text"],
                selectbackground=THEME["accent"], selectforeground=THEME["text_bright"],
                relief="flat", highlightthickness=1,
                highlightbackground=THEME["input_border"],
                height=5, activestyle="none",
                selectmode=tk.EXTENDED   # Allow Ctrl+click / Shift+click multi-select
            )
            self.file_listbox.pack(side="left", fill="x", expand=True)

            sb = ttk.Scrollbar(list_frame, orient="vertical",
                               command=self.file_listbox.yview)
            sb.pack(side="right", fill="y")
            self.file_listbox.configure(yscrollcommand=sb.set)

            # Counter label below the list
            self.file_count_label = tk.Label(
                self.inner, text="No files added yet",
                font=("Segoe UI", 8), bg=THEME["bg"], fg=THEME["text_dim"], anchor="w", padx=16
            )
            self.file_count_label.pack(fill="x")

        def _build_lab_details(self):
            self._section_label("2.  Laboratory Details")
            card = self._card()

            self.lab_name_var  = tk.StringVar()
            self.lab_addr_var  = tk.StringVar()
            self.lab_phone_var = tk.StringVar()
            self.lab_dir_var   = tk.StringVar()

            self._labeled_entry(card, "Lab Name *",    self.lab_name_var,  "e.g. Weija Gbawe Municipal Hospital")
            self._labeled_entry(card, "Address",       self.lab_addr_var,  "e.g. Weija, Accra, Ghana")
            self._labeled_entry(card, "Phone Number",  self.lab_phone_var, "e.g. +233 20 000 0000")
            self._labeled_entry(card, "Laboratory Scientist",  self.lab_dir_var, "e.g. Bio. Asante")

        def _build_output_section(self):
            self._section_label("3.  Output Settings")
            card = self._card()

            self.output_dir_var = tk.StringVar()
            self.filename_var   = tk.StringVar(value="auto")

            # Output folder
            row = tk.Frame(card, bg=THEME["card"])
            row.pack(fill="x", pady=(0, 8))
            tk.Label(
                row, text="Save to folder",
                font=("Segoe UI", 9), bg=THEME["card"], fg=THEME["text_dim"], width=14, anchor="w"
            ).pack(side="left")
            tk.Entry(
                row, textvariable=self.output_dir_var,
                font=("Segoe UI", 10), bg=THEME["input_bg"],
                fg=THEME["text"], insertbackground=THEME["text"],
                relief="flat", highlightthickness=1,
                highlightbackground=THEME["input_border"],
                highlightcolor=THEME["accent"]
            ).pack(side="left", fill="x", expand=True, padx=(6, 6))
            self._make_button(row, "Browse…", self._browse_output_dir,
                              style="secondary", width=8).pack(side="left")

            self._labeled_entry(card, "Filename", self.filename_var,
                                "Leave as 'auto' to name from patient name")

        def _build_generate_btn(self):
            btn_frame = tk.Frame(self.inner, bg=THEME["bg"], pady=10)
            btn_frame.pack(fill="x", padx=16)

            self.generate_btn = self._make_button(
                btn_frame, "⚡  Generate Report", self._generate,
                style="primary", height=2
            )
            self.generate_btn.pack(fill="x")

            self.progress = ttk.Progressbar(
                btn_frame, mode="indeterminate", length=200
            )
            self.progress.pack(fill="x", pady=(6, 0))

        def _build_status_bar(self):
            self.status_var = tk.StringVar(value="Ready — drag a PDF or click Browse to start")
            self.status_bar = tk.Label(
                self.inner, textvariable=self.status_var,
                font=("Segoe UI", 9), bg=THEME["border"],
                fg=THEME["text_dim"], anchor="w", padx=10, pady=5
            )
            self.status_bar.pack(fill="x", side="bottom")

        # ──────────────────────────────────────────────────────────────────────
        # WIDGET HELPERS
        # ──────────────────────────────────────────────────────────────────────

        def _section_label(self, text):
            tk.Label(
                self.inner, text=text,
                font=("Segoe UI", 10, "bold"),
                bg=THEME["bg"], fg=THEME["accent"],
                anchor="w", padx=16
            ).pack(fill="x", pady=(10, 2))

        def _card(self):
            frame = tk.Frame(
                self.inner, bg=THEME["card"],
                highlightbackground=THEME["border"],
                highlightthickness=1, padx=14, pady=10
            )
            frame.pack(fill="x", padx=16, pady=(0, 6))
            return frame

        def _labeled_entry(self, parent, label_text, variable, placeholder=""):
            row = tk.Frame(parent, bg=THEME["card"])
            row.pack(fill="x", pady=4)

            tk.Label(
                row, text=label_text,
                font=("Segoe UI", 9), bg=THEME["card"],
                fg=THEME["text_dim"], width=14, anchor="w"
            ).pack(side="left")

            entry = tk.Entry(
                row, textvariable=variable,
                font=("Segoe UI", 10), bg=THEME["input_bg"],
                fg=THEME["text"], insertbackground=THEME["text"],
                relief="flat", highlightthickness=1,
                highlightbackground=THEME["input_border"],
                highlightcolor=THEME["accent"]
            )
            entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

            if placeholder:
                self._add_placeholder(entry, variable, placeholder)

            # Auto-save settings whenever the user leaves this field
            entry.bind("<FocusOut>", lambda e: self._save_settings())

            return entry

        def _add_placeholder(self, entry, variable, placeholder):
            if not variable.get():
                entry.insert(0, placeholder)
                entry.config(fg=THEME["text_dim"])

            def on_focus_in(event):
                if entry.get() == placeholder:
                    entry.delete(0, "end")
                    entry.config(fg=THEME["text"])

            def on_focus_out(event):
                if not entry.get():
                    entry.insert(0, placeholder)
                    entry.config(fg=THEME["text_dim"])

            def on_key(event):
                if entry.cget("fg") == THEME["text_dim"]:
                    entry.delete(0, "end")
                    entry.config(fg=THEME["text"])

            entry.bind("<FocusIn>",  on_focus_in)
            entry.bind("<FocusOut>", on_focus_out)
            entry.bind("<Key>",      on_key)

        def _make_button(self, parent, text, command,
                         style="primary", width=None, height=1):
            colours = {
                "primary":   (THEME["accent"],      THEME["text_bright"], THEME["accent_hover"]),
                "secondary": (THEME["card"],         THEME["text"],        THEME["border"]),
                "success":   (THEME["success"],      THEME["text_bright"], "#3da870"),
            }
            bg, fg, hover = colours.get(style, colours["primary"])

            btn = tk.Button(
                parent, text=text, command=command,
                font=("Segoe UI", 10, "bold" if style == "primary" else "normal"),
                bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
                relief="flat", cursor="hand2",
                padx=14, pady=6 * height,
                width=width or 0
            )

            btn.bind("<Enter>", lambda e: btn.config(bg=hover))
            btn.bind("<Leave>", lambda e: btn.config(bg=bg))
            return btn

        # ──────────────────────────────────────────────────────────────────────
        # EVENT HANDLERS
        # ──────────────────────────────────────────────────────────────────────

        def _on_drop(self, event):
            # tkinterdnd2 delivers multiple files as space-separated paths
            # wrapped in {} braces. Parse them all.
            raw = event.data.strip()
            # Split on "} {" to handle multiple files from drag
            paths = [p.strip().strip("{}") for p in raw.replace("} {", "\n").split("\n") if p.strip()]
            added = 0
            for path in paths:
                if self._add_pdf(path):
                    added += 1
            if added:
                self._set_status(f"✅  Added {added} file(s)", "success")
            self._on_drag_leave(event)

        def _on_drag_enter(self, event):
            self.drop_frame.config(bg=THEME["drop_hover"],
                                   highlightbackground=THEME["success"])
            for w in [self.drop_icon, self.drop_label]:
                w.config(bg=THEME["drop_hover"])

        def _on_drag_leave(self, event):
            self.drop_frame.config(bg=THEME["drop_idle"],
                                   highlightbackground=THEME["drop_border"])
            for w in [self.drop_icon, self.drop_label]:
                w.config(bg=THEME["drop_idle"])

        def _browse_input(self):
            paths = filedialog.askopenfilenames(
                title="Select Patient PDF(s)",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
            )
            added = 0
            for path in paths:
                if self._add_pdf(path):
                    added += 1
            if added:
                self._set_status(f"✅  Added {added} file(s)", "success")

        def _add_pdf(self, path):
            """Add a single PDF path to the list. Returns True if added."""
            if not path.lower().endswith(".pdf"):
                self._set_status(f"⚠️  Skipped (not a PDF): {os.path.basename(path)}", "warning")
                return False
            if not os.path.isfile(path):
                self._set_status(f"⚠️  File not found: {path}", "warning")
                return False
            if path in self.input_pdfs:
                return False   # Already in list — silently skip duplicate
            self.input_pdfs.append(path)
            self.file_listbox.insert(tk.END, os.path.basename(path))
            # Auto-fill output folder from the first file added
            if len(self.input_pdfs) == 1 and not self.output_dir_var.get():
                self.output_dir_var.set(os.path.dirname(path))
            self._update_file_count()
            return True

        def _remove_selected(self):
            """Remove whichever rows are selected in the listbox."""
            selected = list(self.file_listbox.curselection())
            for idx in reversed(selected):   # reverse so indices stay valid
                self.file_listbox.delete(idx)
                del self.input_pdfs[idx]
            self._update_file_count()

        def _clear_all_pdfs(self):
            self.input_pdfs.clear()
            self.file_listbox.delete(0, tk.END)
            self._update_file_count()
            self._set_status("File list cleared.", "info")

        def _update_file_count(self):
            n = len(self.input_pdfs)
            if n == 0:
                self.file_count_label.config(text="No files added yet", fg=THEME["text_dim"])
            elif n == 1:
                self.file_count_label.config(text="1 file ready to process", fg=THEME["success"])
            else:
                self.file_count_label.config(text=f"{n} files ready — one report per patient", fg=THEME["success"])

        def _browse_output_dir(self):
            folder = filedialog.askdirectory(title="Select Output Folder")
            if folder:
                self.output_dir_var.set(folder)

        def _get_clean_value(self, var, placeholder=""):
            val = var.get().strip()
            if val == placeholder:
                return ""
            return val

        # ──────────────────────────────────────────────────────────────────────
        # REPORT GENERATION
        # ──────────────────────────────────────────────────────────────────────

        def _generate(self):
            # ── Validate inputs ───────────────────────────────────────────────
            if not self.input_pdfs:
                self._set_status("⚠️  Please add at least one PDF first.", "warning")
                return

            lab_name = self._get_clean_value(self.lab_name_var,
                                             "e.g. Weija Gbawe Municipal Hospital")
            if not lab_name:
                self._set_status("⚠️  Lab Name is required.", "warning")
                messagebox.showwarning("Missing Field",
                                       "Please enter your Lab Name before generating.")
                return

            lab_addr  = self._get_clean_value(self.lab_addr_var,  "e.g. Weija, Accra, Ghana")
            lab_phone = self._get_clean_value(self.lab_phone_var, "e.g. +233 20 000 0000")
            lab_dir   = self._get_clean_value(self.lab_dir_var,   "e.g. Bio. Asante")

            out_dir      = self.output_dir_var.get().strip() or                            os.path.dirname(self.input_pdfs[0])
            filename_raw = self.filename_var.get().strip()
            # Custom filename only makes sense for a single file
            custom_name  = "" if filename_raw in ("", "auto",
                "Leave as 'auto' to name from patient name") else filename_raw
            if custom_name and not custom_name.lower().endswith(".pdf"):
                custom_name += ".pdf"

            # ── Start generation in background thread ─────────────────────────
            self.generate_btn.config(state="disabled", text="⏳  Generating…")
            self.progress.start(12)
            n = len(self.input_pdfs)
            self._set_status(f"⏳  Processing {n} file(s)…", "info")

            thread = threading.Thread(
                target=self._run_generation,
                args=(list(self.input_pdfs), lab_name, lab_addr,
                      lab_phone, lab_dir, out_dir, custom_name),
                daemon=True
            )
            thread.start()

        def _run_generation(self, pdf_paths, lab_name, lab_addr,
                            lab_phone, lab_dir, out_dir, custom_name):
            """Runs in background thread. Processes every PDF in pdf_paths."""
            errors  = []
            outputs = []

            # Set global lab config once for all files
            global LAB_NAME, LAB_ADDRESS, LAB_PHONE, LAB_DIRECTOR
            LAB_NAME     = lab_name
            LAB_ADDRESS  = lab_addr
            LAB_PHONE    = lab_phone
            LAB_DIRECTOR = lab_dir

            labs = [{
                "code":     "01",
                "name":     LAB_NAME,
                "address":  LAB_ADDRESS,
                "phone":    LAB_PHONE,
                "director": LAB_DIRECTOR,
            }]

            for i, input_pdf in enumerate(pdf_paths, 1):
                try:
                    self.after(0, lambda i=i, n=len(pdf_paths):
                        self._set_status(f"⏳  Processing file {i} of {n}…", "info"))

                    extracted = extract_text_from_pdf(input_pdf)
                    patient, results = parse_pdf_data(extracted)

                    if not results:
                        raise ValueError(f"No results parsed from: {os.path.basename(input_pdf)}")

                    # Determine output path for this file
                    if custom_name and len(pdf_paths) == 1:
                        out_path = os.path.join(out_dir, custom_name)
                    else:
                        safe = (patient["name"]
                                .replace(",", "").replace(" ", "_").strip("_") or "report")
                        out_path = os.path.join(out_dir, f"{safe}_report.pdf")

                    generate_report(patient, results, labs, out_path)
                    outputs.append(out_path)

                except Exception as exc:
                    errors.append(f"{os.path.basename(input_pdf)}: {exc}")

            self.after(0, lambda: self._generation_done(outputs, errors))

        def _generation_done(self, outputs, errors):
            """Called on the GUI thread when all files are done."""
            self.progress.stop()
            self.generate_btn.config(state="normal", text="⚡  Generate Report")
            self._save_settings()

            if errors and not outputs:
                # All failed
                msg = "\n".join(errors)
                self._set_status(f"❌  All {len(errors)} file(s) failed.", "error")
                messagebox.showerror("Generation Failed", msg)
                return

            if errors:
                # Partial success
                messagebox.showwarning(
                    "Partial Success",
                    f"{len(outputs)} report(s) saved.\n\nFailed:\n" + "\n".join(errors)
                )

            # Success
            n = len(outputs)
            self._set_status(f"✅  {n} report(s) saved to: {os.path.dirname(outputs[0])}", "success")

            if n == 1:
                msg = f"Report saved:\n{outputs[0]}\n\nOpen it now?"
                if messagebox.askyesno("Report Ready", msg):
                    self._open_file(outputs[0])
            else:
                msg = f"{n} reports saved to:\n{os.path.dirname(outputs[0])}\n\nOpen folder?"
                if messagebox.askyesno("Reports Ready", msg):
                    self._open_file(os.path.dirname(outputs[0]))

        def _open_file(self, path):
            """Opens a file or folder using the OS default viewer/explorer."""
            import subprocess
            try:
                if sys.platform == "win32":
                    os.startfile(path)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
            except Exception:
                pass

        def _set_status(self, message, level="info"):
            colours = {
                "info":    THEME["text_dim"],
                "success": THEME["success"],
                "warning": THEME["warning"],
                "error":   THEME["error"],
            }
            self.status_var.set(message)
            self.status_bar.config(fg=colours.get(level, THEME["text_dim"]))

        def _on_frame_configure(self, event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def _on_canvas_configure(self, event):
            self.canvas.itemconfig(self.canvas_window, width=event.width)

        def _on_mousewheel(self, event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # ──────────────────────────────────────────────────────────────────────
        # SETTINGS PERSISTENCE
        # ──────────────────────────────────────────────────────────────────────

        def _settings_path(self):
            return os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "lab_report_settings.txt"
            )

        def _save_settings(self):
            try:
                with open(self._settings_path(), "w") as f:
                    f.write(f"lab_name={self._get_clean_value(self.lab_name_var)}\n")
                    f.write(f"lab_addr={self._get_clean_value(self.lab_addr_var)}\n")
                    f.write(f"lab_phone={self._get_clean_value(self.lab_phone_var)}\n")
                    f.write(f"lab_dir={self._get_clean_value(self.lab_dir_var)}\n")
                    f.write(f"output_dir={self.output_dir_var.get()}\n")
            except Exception:
                pass

        def _load_settings(self):
            try:
                path = self._settings_path()
                if not os.path.exists(path):
                    return
                settings = {}
                with open(path) as f:
                    for line in f:
                        line = line.strip()
                        if "=" in line:
                            key, _, value = line.partition("=")
                            settings[key] = value
                if settings.get("lab_name"):
                    self.lab_name_var.set(settings["lab_name"])
                if settings.get("lab_addr"):
                    self.lab_addr_var.set(settings["lab_addr"])
                if settings.get("lab_phone"):
                    self.lab_phone_var.set(settings["lab_phone"])
                if settings.get("lab_dir"):
                    self.lab_dir_var.set(settings["lab_dir"])
                if settings.get("output_dir"):
                    self.output_dir_var.set(settings["output_dir"])
            except Exception:
                pass

    return App


# ==============================================================================
# SECTION 16 — SCRIPT ENTRY TRIGGER
#
# Double-click the file  →  GUI opens
# python script.py --demo  →  CLI demo
# python script.py file.pdf "Lab Name"  →  CLI with real PDF
# ==============================================================================

if __name__ == "__main__":

    if len(sys.argv) > 1:
        # Arguments provided — run in CLI mode
        main()

    else:
        # No arguments — launch the GUI
        if not TK_AVAILABLE:
            print("⚠️  tkinter is not available on this system.")
            print("   CLI mode still works. Example:")
            print("     python lab_report_5panels_commented.py --demo")
            sys.exit(0)

        # Check required packages
        missing = []
        try:
            import reportlab
        except ImportError:
            missing.append("reportlab")
        try:
            import pdfplumber
        except ImportError:
            missing.append("pdfplumber")

        if missing:
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror(
                    "Missing Packages",
                    f"Please install required packages first:\n\n"
                    f"  pip install {' '.join(missing)}\n\n"
                    f"Then run the script again."
                )
                root.destroy()
            except Exception:
                print(f"❌ Missing packages: pip install {' '.join(missing)}")
            sys.exit(1)

        if not DND_AVAILABLE:
            print("ℹ️  Drag-and-drop unavailable. Run: pip install tkinterdnd2")

        # Build and run the app
        AppClass = _build_gui_app()
        app = AppClass()
        app.mainloop()
