# -*- coding: utf-8 -*-
import io
import math
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# PDF export
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

DISCLAIMER = (
    "These results are estimates based on the specific assumptions and data points provided and are intended for illustrative purposes only. "
    "They do not constitute a guarantee of actual returns or a promise of future financial performance."
)

APP_DIR = Path(__file__).parent
ASSETS = APP_DIR / "assets"
LOGO_PATH = ASSETS / "commscoach_logo_reverse.png"
LOGO_PDF_PATH = ASSETS / "commscoach_logo_navy.png"
FAVICON_PATH = ASSETS / "favicon.png"

MODULE_SUMMARIES = {
    "QA": {
        "title": "CommsCoach QA",
        "body": (
            "Annual subscription for the single agency named on the sales order. Includes: "
            "Call and Radio Evaluations, Post Event Audio Transcription, Keyword Search, Review Queues, "
            "Shift Goals, Dashboards, Reports, and Evaluator Feedback."
        ),
    },
    "TRAIN": {
        "title": "CommsCoach Train",
        "body": (
            "Includes AI-driven training simulations created from actual events in agency CAD and audio, a simulations library, "
            "and the ability to create your own. Phased Training Templates, Task Lists, Automated Evaluations (when connected to QA), "
            "Observation Summaries, Dashboards, Reports, and Trainer Feedback. Observations can include and summarize evaluations performed over events."
        ),
    },
    "HIRE": {
        "title": "CommsCoach HIRE",
        "body": (
            "Annual subscription for the single agency identified on the order form. Provides pre-hire candidate assessments "
            "using simulations, interactive questions, evaluations, and reporting."
        ),
    },
    "ASSIST": {
        "title": "CommsCoach Assist",
        "body": (
            "Annual subscription for the single agency named on the order form. Provides real-time call transcription and translation "
            "with in-call guidance and evaluations, with notifications based on agency settings."
        ),
    },
}


def money(x: float) -> str:
    try:
        return f"${x:,.0f}"
    except Exception:
        return "$0"


def pct_from_ratio(x: float) -> str:
    try:
        return f"{x * 100:,.1f}%"
    except Exception:
        return "0.0%"


def safe_div(a: float, b: float) -> float:
    return a / b if b else 0.0


def build_excel_export(inputs: dict, results: dict, breakdown_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(list(inputs.items()), columns=["Input", "Value"]).to_excel(writer, sheet_name="Inputs", index=False)
        pd.DataFrame(list(results.items()), columns=["Metric", "Value"]).to_excel(writer, sheet_name="Results", index=False)
        breakdown_df.to_excel(writer, sheet_name="Savings Breakdown", index=False)
    buf.seek(0)
    return buf.read()


def _pdf_new_page(c: canvas.Canvas, title: str | None = None):
    c.showPage()
    if title:
        c.setFont("Helvetica-Bold", 14)
        c.drawString(0.75 * inch, letter[1] - 0.75 * inch, title)


# ---------------------------------------------------------------------------
# PDF color palette matching the dark Streamlit UI
# ---------------------------------------------------------------------------
PDF_BG        = (0.082, 0.082, 0.082)   # #141414  page background
PDF_HEADER_BG = (0.082, 0.141, 0.239)   # #152639  dark-navy header band
PDF_BANNER_BG = (0.122, 0.188, 0.298)   # #1F304C  payback banner
PDF_ROW_ALT   = (0.118, 0.118, 0.118)   # #1E1E1E  table alternate row
PDF_ROW_EVEN  = (0.090, 0.090, 0.090)   # #171717
PDF_TEXT_WH   = (1.0,   1.0,   1.0)     # white
PDF_TEXT_DIM  = (0.65,  0.65,  0.65)    # dimmed / caption
PDF_ACCENT    = (0.259, 0.569, 0.871)   # #4291DE  blue accent
PDF_GREEN     = (0.118, 0.741, 0.424)   # metric green


def _set_fill(c, rgb):
    c.setFillColorRGB(*rgb)


def _set_stroke(c, rgb):
    c.setStrokeColorRGB(*rgb)


def _draw_rect(c, x, y, w, h, fill_rgb, stroke_rgb=None):
    _set_fill(c, fill_rgb)
    if stroke_rgb:
        _set_stroke(c, stroke_rgb)
        c.rect(x, y, w, h, stroke=1, fill=1)
    else:
        c.rect(x, y, w, h, stroke=0, fill=1)


def _wrap_text(text: str, max_chars: int) -> list:
    """Simple word-wrap returning list of lines."""
    words = text.split()
    lines, line = [], ""
    for w in words:
        test = (line + " " + w).strip()
        if len(test) > max_chars:
            lines.append(line)
            line = w
        else:
            line = test
    if line:
        lines.append(line)
    return lines


def build_pdf_report(
    agency_name: str,
    scenario_name: str,
    selected_modules: list,
    inputs: dict,
    results: dict,
    breakdown_df: pd.DataFrame,
    baseline_estimates: dict,
    logo_path: Path | None = None,
    # extra kwargs forwarded from the live calculation so the PDF can
    # reproduce the exact numbers shown in the UI
    qa_calls_reviewed_year: float = 0,
    manual_qa_coverage_pct: float = 0,
    annual_calls: int = 0,
    avg_qa_minutes_per_call: float = 0,
    baseline_qa_hours_year_est: float = 0,
    baseline_qa_hours_week_est: float = 0,
    supervisor_baseline_hours_week: float = 0,
    supervisor_hours_saved_week: float = 0,
    sup_hours_saved_week_capped: float = 0,
    qa_baseline_hours_week_capacity: float = 0,
    qa_specialists_fte: float = 0,
    qa_hours_saved_week: float = 0,
    qa_hours_saved_week_capped: float = 0,
    baseline_training_hours_year: float = 0,
    training_hours_saved_year_capped: float = 0,
    supervisor_hourly_fully_loaded: float = 0,
    qa_hourly_loaded: float = 0,
    trainer_hourly_fully_loaded: float = 0,
) -> bytes:
    """Customer-ready PDF that mirrors the dark-theme Streamlit dashboard."""
    from reportlab.lib.colors import HexColor

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    W, H = letter          # 612 x 792 pt
    PAD = 0.45 * inch      # outer horizontal margin

    # -----------------------------------------------------------------------
    # Helper: fill entire page with dark background
    # -----------------------------------------------------------------------
    def _bg():
        _draw_rect(c, 0, 0, W, H, PDF_BG)

    # -----------------------------------------------------------------------
    # Helper: draw the top header band
    # -----------------------------------------------------------------------
    def _header_band(title: str, subtitle: str = ""):
        band_h = 1.05 * inch
        _draw_rect(c, 0, H - band_h, W, band_h, PDF_HEADER_BG)
        _set_fill(c, PDF_TEXT_WH)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(PAD, H - 0.45 * inch, title)
        if subtitle:
            _set_fill(c, PDF_TEXT_DIM)
            c.setFont("Helvetica", 8)
            c.drawString(PAD, H - 0.65 * inch, subtitle)

    # -----------------------------------------------------------------------
    # Helper: footer
    # -----------------------------------------------------------------------
    def _footer(page_num: int):
        _set_fill(c, PDF_TEXT_DIM)
        c.setFont("Helvetica", 7)
        label = f"GovWorx | CommsCoach ROI Summary  |  {datetime.now().strftime('%Y-%m-%d')}  |  Page {page_num}"
        c.drawCentredString(W / 2, 0.30 * inch, label)

    # -----------------------------------------------------------------------
    # Helper: section heading on dark bg
    # -----------------------------------------------------------------------
    def _section_heading(text: str, y: float, size: int = 11) -> float:
        _set_fill(c, PDF_TEXT_WH)
        c.setFont("Helvetica-Bold", size)
        c.drawString(PAD, y, text)
        return y - 0.20 * inch

    # -----------------------------------------------------------------------
    # Helper: bullet item (bold label + normal value)
    # -----------------------------------------------------------------------
    def _bullet(label: str, value: str, y: float, x: float = None) -> float:
        if x is None:
            x = PAD
        _set_fill(c, PDF_TEXT_DIM)
        c.setFont("Helvetica", 8)
        c.drawString(x, y, "•")
        c.setFont("Helvetica", 8)
        _set_fill(c, PDF_TEXT_WH)
        c.drawString(x + 0.12 * inch, y, f"{label}: ")
        tw = c.stringWidth(f"{label}: ", "Helvetica", 8)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 0.12 * inch + tw, y, value)
        return y - 0.165 * inch

    # -----------------------------------------------------------------------
    # PAGE 1  --  main dashboard mirror
    # -----------------------------------------------------------------------
    _bg()
    _header_band(
        "GovWorx | CommsCoach ROI Calculator",
        "Hour-based ROI model with split inputs, manual QA effort estimator, and customer-ready PDF export.",
    )

    gross  = float(results.get("Annual gross savings ($)", 0) or 0)
    invest = float(results.get("Annual investment ($)", 0) or 0)
    net    = float(results.get("Net annual benefit ($)", 0) or 0)
    roi    = float(results.get("ROI ratio", 0) or 0)
    pb     = results.get("Payback months", None)

    # --- 4 metric tiles ---
    tile_y_top  = H - 1.05 * inch
    tile_h      = 0.90 * inch
    tile_gap    = 0.08 * inch
    tile_w      = (W - 2 * PAD - 3 * tile_gap) / 4

    metrics = [
        ("Annual Gross Savings", money(gross)),
        ("Annual Investment",    money(invest)),
        ("Net Annual Benefit",   money(net)),
        ("ROI",                  pct_from_ratio(roi)),
    ]
    for i, (lbl, val) in enumerate(metrics):
        tx = PAD + i * (tile_w + tile_gap)
        ty = tile_y_top - tile_h
        _draw_rect(c, tx, ty, tile_w, tile_h, PDF_ROW_ALT)
        _set_fill(c, PDF_TEXT_DIM)
        c.setFont("Helvetica", 7)
        c.drawString(tx + 6, ty + tile_h - 14, lbl)
        _set_fill(c, PDF_TEXT_WH)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(tx + 6, ty + 10, val)

1Password menu is available. Press down arrow to select.
