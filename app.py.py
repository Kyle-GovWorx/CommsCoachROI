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

    # --- payback banner ---
    banner_y = tile_y_top - tile_h - 0.08 * inch
    banner_h = 0.26 * inch
    _draw_rect(c, PAD, banner_y - banner_h, W - 2 * PAD, banner_h, PDF_BANNER_BG)
    _set_fill(c, PDF_ACCENT)
    c.setFont("Helvetica", 8)
    pb_text = f"Estimated payback period: {float(pb):.1f} months" if pb is not None else "Payback: N/A"
    c.drawString(PAD + 6, banner_y - banner_h + 6, pb_text)

    body_top = banner_y - banner_h - 0.18 * inch

    # --- two columns ---
    col_gap   = 0.25 * inch
    col_l_w   = (W - 2 * PAD - col_gap) * 0.555
    col_r_w   = (W - 2 * PAD - col_gap) * 0.445
    col_l_x   = PAD
    col_r_x   = PAD + col_l_w + col_gap

    # ---- LEFT COLUMN ----
    ly = body_top
    # section heading pinned to left column x
    _set_fill(c, PDF_TEXT_WH)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(col_l_x, ly, "Savings breakdown")
    ly -= 0.20 * inch

    # table header row
    th = 0.18 * inch
    _draw_rect(c, col_l_x, ly - th + 4, col_l_w, th, PDF_ROW_ALT)
    _set_fill(c, PDF_TEXT_DIM)
    c.setFont("Helvetica", 7)
    c.drawString(col_l_x + 4, ly - th + 8, "Bucket")
    c.drawRightString(col_l_x + col_l_w - 4, ly - th + 8, "Annual Value ($)")
    ly -= th

    row_h = 0.175 * inch
    for idx, (_, row) in enumerate(breakdown_df.iterrows()):
        bg = PDF_ROW_ALT if idx % 2 == 0 else PDF_ROW_EVEN
        _draw_rect(c, col_l_x, ly - row_h + 4, col_l_w, row_h, bg)
        bucket = str(row.get("Bucket", ""))
        val    = float(row.get("Annual Value ($)", 0) or 0)
        _set_fill(c, PDF_TEXT_WH)
        c.setFont("Helvetica", 7.5)
        c.drawString(col_l_x + 4, ly - row_h + 7, bucket)
        c.drawRightString(col_l_x + col_l_w - 4, ly - row_h + 7, f"{val:,.0f}")
        ly -= row_h

    ly -= 0.20 * inch
    _set_fill(c, PDF_TEXT_WH)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(col_l_x, ly, "Product summary (selected modules)")
    ly -= 0.20 * inch

    for key in selected_modules:
        item = MODULE_SUMMARIES.get(key)
        if not item:
            continue
        _set_fill(c, PDF_TEXT_WH)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(col_l_x, ly, item["title"])
        ly -= 0.155 * inch
        c.setFont("Helvetica", 7.5)
        _set_fill(c, PDF_TEXT_DIM)
        for ln in _wrap_text(item["body"], max_chars=72):
            c.drawString(col_l_x, ly, ln)
            ly -= 0.135 * inch
            if ly < 0.55 * inch:
                break
        ly -= 0.08 * inch

    # ---- RIGHT COLUMN ----
    ry = body_top
    _set_fill(c, PDF_TEXT_WH)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(col_r_x, ry, "Manual QA effort estimator (baseline)")
    ry -= 0.20 * inch
    ry = _bullet("Calls reviewed/year",
                 f"{qa_calls_reviewed_year:,.0f} (coverage {manual_qa_coverage_pct:.1f}% of {annual_calls:,})",
                 ry, col_r_x)
    ry = _bullet("Avg QA minutes per reviewed call", f"{avg_qa_minutes_per_call:.0f}", ry, col_r_x)
    ry = _bullet("Estimated baseline QA hours/year", f"{baseline_qa_hours_year_est:,.0f}", ry, col_r_x)
    ry = _bullet("Estimated baseline QA hours/week", f"{baseline_qa_hours_week_est:,.1f}", ry, col_r_x)

    ry -= 0.12 * inch
    _set_fill(c, PDF_TEXT_WH)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(col_r_x, ry, "Caps and time-saved inputs")
    ry -= 0.20 * inch
    ry = _bullet("Supervisor baseline QA time (estimated or provided)",
                 f"{supervisor_baseline_hours_week:,.1f} hrs/week", ry, col_r_x)
    ry = _bullet("Supervisor hours saved (entered)",
                 f"{supervisor_hours_saved_week:,.1f} hrs/week", ry, col_r_x)
    ry = _bullet("Supervisor hours saved (capped)",
                 f"{sup_hours_saved_week_capped:,.1f} hrs/week", ry, col_r_x)
    ry -= 0.05 * inch
    ry = _bullet("QA baseline capacity",
                 f"{qa_baseline_hours_week_capacity:,.1f} hrs/week ({qa_specialists_fte:.2f} FTE)",
                 ry, col_r_x)
    ry = _bullet("QA hours saved (entered)",
                 f"{qa_hours_saved_week:,.1f} hrs/week", ry, col_r_x)
    ry = _bullet("QA hours saved (capped)",
                 f"{qa_hours_saved_week_capped:,.1f} hrs/week", ry, col_r_x)
    ry -= 0.05 * inch
    ry = _bullet("Training baseline",
                 f"{baseline_training_hours_year:,.0f} hrs/year", ry, col_r_x)
    ry = _bullet("Training hours saved (capped)",
                 f"{training_hours_saved_year_capped:,.0f} hrs/year", ry, col_r_x)

    _footer(1)
    c.showPage()

    # -----------------------------------------------------------------------
    # PAGE 2  --  Math explanation
    # -----------------------------------------------------------------------
    _bg()
    _header_band("How the numbers are calculated", "Step-by-step math behind every figure on the dashboard.")

    # Build math explanation lines
    sup_hrs_yr   = sup_hours_saved_week_capped * 52
    qa_hrs_yr    = qa_hours_saved_week_capped  * 52
    sup_savings  = sup_hrs_yr  * supervisor_hourly_fully_loaded
    qa_savings   = qa_hrs_yr   * qa_hourly_loaded
    tr_savings   = training_hours_saved_year_capped * trainer_hourly_fully_loaded
    turnover_val = float(results.get("Annual gross savings ($)", 0) or 0) - sup_savings - qa_savings - tr_savings

    sections = [
        (
            "1. Manual QA Effort Baseline",
            [
                f"  Calls reviewed / year  =  Annual calls ({annual_calls:,}) x Coverage ({manual_qa_coverage_pct:.1f}% / 100)",
                f"                         =  {qa_calls_reviewed_year:,.0f} calls",
                f"  Baseline QA hrs / year =  Calls reviewed ({qa_calls_reviewed_year:,.0f}) x Avg minutes ({avg_qa_minutes_per_call:.0f}) / 60",
                f"                         =  {baseline_qa_hours_year_est:,.0f} hrs",
                f"  Baseline QA hrs / week =  {baseline_qa_hours_year_est:,.0f} / 52  =  {baseline_qa_hours_week_est:,.1f} hrs",
            ],
        ),
        (
            "2. Supervisor Time Savings",
            [
                f"  Entered hrs saved / week  =  Supervisors in scope x Hrs saved per supervisor",
                f"                            =  {supervisor_hours_saved_week:,.1f} hrs/week (before cap)",
                f"  Cap                       =  Supervisor baseline QA time: {supervisor_baseline_hours_week:,.1f} hrs/week",
                f"  Capped hrs saved / week   =  min(entered, cap)  =  {sup_hours_saved_week_capped:,.1f} hrs/week",
                f"  Annual hrs saved           =  {sup_hours_saved_week_capped:,.1f} x 52  =  {sup_hrs_yr:,.0f} hrs",
                f"  Dollar value               =  {sup_hrs_yr:,.0f} hrs x ${supervisor_hourly_fully_loaded:,.2f}/hr  =  {money(sup_savings)}",
            ],
        ),
        (
            "3. QA Labor Savings",
            [
                f"  QA baseline capacity       =  {qa_specialists_fte:.2f} FTE x 40 hrs/week  =  {qa_baseline_hours_week_capacity:,.1f} hrs/week",
                f"  Entered hrs saved / week   =  QA staff in scope x Hrs saved per QA staff  =  {qa_hours_saved_week:,.1f} hrs/week",
                f"  Capped hrs saved / week    =  min(entered, capacity)  =  {qa_hours_saved_week_capped:,.1f} hrs/week",
                f"  Annual hrs saved            =  {qa_hours_saved_week_capped:,.1f} x 52  =  {qa_hrs_yr:,.0f} hrs",
                f"  QA hourly rate              =  Annual cost / 2,080  =  ${qa_hourly_loaded:,.2f}/hr",
                f"  Dollar value               =  {qa_hrs_yr:,.0f} hrs x ${qa_hourly_loaded:,.2f}/hr  =  {money(qa_savings)}",
            ],
        ),
        (
            "4. Training Savings",
            [
                f"  Training baseline          =  New hires/yr x Training hrs/hire  =  {baseline_training_hours_year:,.0f} hrs/year",
                f"  Raw hrs saved / year       =  max(trainer weekly savings, per-hire savings)",
                f"  Capped hrs saved / year    =  min(raw, baseline)  =  {training_hours_saved_year_capped:,.0f} hrs",
                f"  Dollar value               =  {training_hours_saved_year_capped:,.0f} hrs x ${trainer_hourly_fully_loaded:,.2f}/hr  =  {money(tr_savings)}",
            ],
        ),
        (
            "5. Summary Roll-Up",
            [
                f"  QA labor savings           =  {money(qa_savings)}",
                f"  Supervisor time savings    =  {money(sup_savings)}",
                f"  Training savings           =  {money(tr_savings)}",
                f"  Turnover / other savings   =  {money(max(0.0, turnover_val))}",
                f"  ─────────────────────────────────────────────",
                f"  Annual gross savings       =  {money(gross)}",
                f"  Annual investment          =  {money(invest)}",
                f"  Net annual benefit         =  Gross - Investment  =  {money(net)}",
                f"  ROI                        =  Net / Investment  =  {pct_from_ratio(roi)}",
                f"  Payback period             =  (Investment / Gross) x 12  =  {float(pb):.1f} months" if pb is not None else "  Payback period  =  N/A",
            ],
        ),
    ]

    my = H - 1.15 * inch
    for sec_title, lines in sections:
        if my < 1.0 * inch:
            _footer(2)
            c.showPage()
            _bg()
            my = H - 0.55 * inch
        _set_fill(c, PDF_ACCENT)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(PAD, my, sec_title)
        my -= 0.165 * inch
        for ln in lines:
            if my < 0.70 * inch:
                _footer(2)
                c.showPage()
                _bg()
                my = H - 0.55 * inch
            _set_fill(c, PDF_TEXT_WH if not ln.strip().startswith("─") else PDF_TEXT_DIM)
            c.setFont("Courier", 7.5)
            c.drawString(PAD, my, ln)
            my -= 0.145 * inch
        my -= 0.12 * inch

    # disclaimer at bottom
    if my > 0.8 * inch:
        my = max(my, 0.8 * inch)
    _set_fill(c, PDF_TEXT_DIM)
    c.setFont("Helvetica-Oblique", 6.5)
    for ln in _wrap_text(DISCLAIMER, 120):
        c.drawString(PAD, my, ln)
        my -= 0.115 * inch

    _footer(2)
    c.showPage()

    # -----------------------------------------------------------------------
    # PAGE 3  --  Assumptions detail
    # -----------------------------------------------------------------------
    _bg()
    _header_band("Assumptions detail", "All inputs used to produce this report.")
    ay = H - 1.15 * inch
    row_h2 = 0.165 * inch
    page_n = 3
    for idx, (k, v) in enumerate(inputs.items()):
        bg = PDF_ROW_ALT if idx % 2 == 0 else PDF_ROW_EVEN
        _draw_rect(c, PAD, ay - row_h2 + 3, W - 2 * PAD, row_h2, bg)
        _set_fill(c, PDF_TEXT_WH)
        c.setFont("Helvetica", 7.5)
        c.drawString(PAD + 4, ay - row_h2 + 6, str(k)[:60])
        c.drawRightString(W - PAD - 4, ay - row_h2 + 6, str(v)[:35])
        ay -= row_h2
        if ay < 0.65 * inch:
            _footer(page_n)
            c.showPage()
            _bg()
            _header_band("Assumptions detail (continued)")
            ay = H - 1.15 * inch
            page_n += 1

    _footer(page_n)
    c.save()
    buf.seek(0)
    return buf.read()


icon = str(FAVICON_PATH) if FAVICON_PATH.exists() else "ROI"
st.set_page_config(page_title="GovWorx | CommsCoach ROI Calculator", page_icon=icon, layout="wide")

st.title("GovWorx | CommsCoach ROI Calculator")
st.caption("Hour-based ROI model with split inputs, manual QA effort estimator, and customer-ready PDF export.")

# ---------------------------
# Sidebar inputs
# ---------------------------
with st.sidebar:
    st.subheader("Inputs")
    agency_name = st.text_input("Agency name (for exports)", value="")
    scenario_name = st.text_input("Scenario name", value="Standard")

    st.markdown("#### Included CommsCoach modules")
    c1, c2 = st.columns(2)
    with c1:
        include_qa = st.toggle("QA", value=True)
        include_hire = st.toggle("HIRE", value=False)
    with c2:
        include_assist = st.toggle("ASSIST", value=False)
        include_train = st.toggle("TRAIN", value=True)

    selected_modules = []
    if include_qa:
        selected_modules.append("QA")
    if include_train:
        selected_modules.append("TRAIN")
    if include_hire:
        selected_modules.append("HIRE")
    if include_assist:
        selected_modules.append("ASSIST")

    st.markdown("#### Investment")
    annual_investment = st.number_input("Annual CommsCoach cost ($)", min_value=0.0, value=225000.0, step=5000.0)

    st.markdown("#### Agency size (context)")
    annual_calls = st.number_input("Annual calls eligible for QA", min_value=0, value=600000, step=10000)
    dispatchers = st.number_input("Dispatchers / telecommunicators", min_value=0, value=75, step=5)

    st.markdown("#### Labor rates")
    supervisor_hourly_fully_loaded = st.number_input("Supervisor fully loaded hourly rate ($/hr)", min_value=0.0, value=110.0, step=1.0)

    qa_specialists_fte = st.number_input("QA specialists (FTE, baseline)", min_value=0.0, value=1.0, step=0.25)
    qa_fte_fully_loaded_cost = st.number_input("QA specialist fully loaded annual cost ($)", min_value=0.0, value=140000.0, step=5000.0)

    trainers_count = st.number_input("Trainers (headcount, context)", min_value=0, value=2, step=1)
    trainer_hourly_fully_loaded = st.number_input("Trainer fully loaded hourly rate ($/hr)", min_value=0.0, value=95.0, step=1.0)

    st.markdown("#### Manual QA baseline (context)")
    estimate_missing_baselines = st.toggle("Agency does not have baseline metrics (estimate for me)", value=True)
    supervisor_hours_per_week_on_qa_manual = st.number_input(
        "Supervisor hours per week currently spent on QA (total)",
        min_value=0.0,
        value=25.0,
        step=1.0,
        disabled=estimate_missing_baselines,
    )
    manual_qa_coverage_pct = st.number_input("Manual QA coverage (% of calls)", min_value=0.0, max_value=100.0, value=2.0, step=0.5)

    st.markdown("#### CommsCoach time savings (split inputs)")
    st.caption("Enter hours saved per person per week. The calculator totals it and caps at baseline where applicable.")
    supervisors_in_scope = st.number_input("Supervisors impacted (count)", min_value=0, value=4, step=1)
    hours_saved_per_supervisor_week = st.number_input("Hours saved per supervisor per week", min_value=0.0, value=3.0, step=0.5)

    qa_specialists_in_scope = st.number_input("QA staff impacted (count)", min_value=0, value=1, step=1)
    hours_saved_per_qa_week = st.number_input("Hours saved per QA staff per week", min_value=0.0, value=4.0, step=0.5)

    st.markdown("#### Training time savings (split inputs)")
    new_hires_per_year = st.number_input("New hires per year", min_value=0, value=15, step=1)
    training_hours_per_new_hire = st.number_input("Training hours per new hire (baseline)", min_value=0.0, value=80.0, step=1.0)
    # Hours saved per trainer per week (auto-estimated if baseline metrics are unknown)
    trainer_hours_saved_default = 2.0
    if estimate_missing_baselines:
        # Conservative default for automating daily observation reports + roleplay/sim admin
        trainer_hours_saved_default = 2.0

    hours_saved_per_trainer_week = st.number_input(
        "Hours saved per trainer per week (DOR + roleplay admin)",
        min_value=0.0,
        value=float(trainer_hours_saved_default),
        step=0.5,
        disabled=estimate_missing_baselines,
    )
    # optional: per new hire savings
    training_hours_saved_per_new_hire = st.number_input("Optional: hours saved per new hire (if known)", min_value=0.0, value=0.0, step=1.0)

    st.markdown("#### Turnover (optional, count-based)")
    st.caption("If you include retention impact, enter the number of dispatchers you believe this helps retain per year.")
    dispatchers_retained_per_year = st.number_input("Dispatchers retained per year (count)", min_value=0.0, value=0.0, step=0.5)
    # Replacement cost per dispatcher (auto-estimated if baseline metrics are unknown)
    # Leadership-defensible defaults:
    # - $50k conservative baseline
    # - $75k typical mid case (often cited by centers as closer to reality)
    # - $100k higher-end case sometimes cited for fully trained replacement
    #
    # References used to set these options:
    # - EMSWorld (HMP Global) reports ~ $100k to replace a fully trained 9-1-1 operator in one cited example.
    # - CritiCall highlights that costs stack when multiple trainees wash out before one stays.
    if estimate_missing_baselines:
        replacement_cost_choice = st.selectbox(
            "Replacement cost per dispatcher (estimate)",
            options=[50000, 75000, 100000],
            index=1,
            format_func=lambda x: f"${x:,.0f}",
            help=(
                "Estimate of the all-in cost to replace one dispatcher (recruiting, hiring, onboarding, training time, and staffing impact). "
                "Use $50k for a conservative case, $75k for a typical mid case, or $100k for a higher-end case. "
                "Optional washout factor below can account for the reality that it may take multiple trainees to get one who stays."
            ),
        )
        washout_factor = st.slider(
            "Washout factor (optional)",
            min_value=1.0,
            max_value=2.0,
            value=1.0,
            step=0.1,
            help=(
                "If it often takes more than one trainee to fill one seat (for example, 1.3x), this increases the effective replacement cost accordingly. "
                "Set to 1.0x to disable."
            ),
        )
        replacement_cost_per_dispatcher = float(replacement_cost_choice) * float(washout_factor)
    else:
        washout_factor = 1.0
        replacement_cost_per_dispatcher = st.number_input(
            "Replacement cost per dispatcher ($)",
            min_value=0.0,
            value=75000.0,
            step=1000.0,
        )

    st.markdown("#### Optional value buckets")
    annual_productivity_value = st.number_input("Annual productivity / effectiveness value ($)", min_value=0.0, value=0.0, step=10000.0)
    annual_risk_reduction_value = st.number_input("Annual risk reduction value ($)", min_value=0.0, value=0.0, step=10000.0)

    # Manual QA effort estimator
    st.markdown("#### Manual QA effort estimator")
    st.caption("Use this if the agency cannot estimate current QA effort. Outputs baseline QA hours based on coverage and review time.")
    # Average QA minutes per reviewed call (auto-estimated if baseline metrics are unknown)
    qa_minutes_default = 25.0
    if estimate_missing_baselines:
        # Common manual QA range for a 5-minute call is ~20-30 minutes including playback + notes.
        qa_minutes_default = 25.0

    avg_qa_minutes_per_call = st.number_input(
        "Average QA minutes per reviewed call",
        min_value=0.0,
        value=float(qa_minutes_default),
        step=1.0,
        disabled=estimate_missing_baselines,
    )
    include_playback_admin = st.toggle("Include admin time in QA minutes", value=True)

# ---------------------------
# Derived baseline estimates and caps
# ---------------------------
weeks_per_year = 52.0
qa_hourly_loaded = qa_fte_fully_loaded_cost / 2080.0 if qa_fte_fully_loaded_cost else 0.0

# Manual QA effort estimator (baseline QA hours/year)
qa_coverage_ratio = max(0.0, min(1.0, manual_qa_coverage_pct / 100.0))
qa_calls_reviewed_year = annual_calls * qa_coverage_ratio
qa_minutes = avg_qa_minutes_per_call
baseline_qa_hours_year_est = (qa_calls_reviewed_year * qa_minutes) / 60.0
baseline_qa_hours_week_est = baseline_qa_hours_year_est / weeks_per_year if weeks_per_year else 0.0

# Baseline capacity caps
# Supervisor baseline QA hours/week:
# If the agency does not know this number, estimate it from the manual QA effort estimator by assuming
# supervisors cover QA work that exceeds dedicated QA capacity (QA FTE * 40 hrs/week).
if estimate_missing_baselines:
    supervisor_baseline_hours_week = max(0.0, baseline_qa_hours_week_est - (qa_specialists_fte * 40.0))
else:
    supervisor_baseline_hours_week = supervisor_hours_per_week_on_qa_manual
qa_baseline_hours_week_capacity = qa_specialists_fte * 40.0

# Saved hours computed from split inputs
supervisor_hours_saved_week = supervisors_in_scope * hours_saved_per_supervisor_week
qa_hours_saved_week = qa_specialists_in_scope * hours_saved_per_qa_week

sup_hours_saved_week_capped = min(supervisor_hours_saved_week, supervisor_baseline_hours_week)

# QA cap: if estimated baseline QA effort is larger than QA capacity, cap at the higher of the two? safest is cap at capacity.
qa_hours_saved_week_capped = min(qa_hours_saved_week, qa_baseline_hours_week_capacity)

# Training savings: pick the larger of (trainer weekly saved) vs (per new hire saved * hires / year / 52), but cap at baseline training hours.
baseline_training_hours_year = new_hires_per_year * training_hours_per_new_hire
training_hours_saved_year_from_trainers = trainers_count * hours_saved_per_trainer_week * weeks_per_year
training_hours_saved_year_from_hires = training_hours_saved_per_new_hire * new_hires_per_year
training_hours_saved_year_raw = max(training_hours_saved_year_from_trainers, training_hours_saved_year_from_hires)
training_hours_saved_year_capped = min(training_hours_saved_year_raw, baseline_training_hours_year)

# ---------------------------
# Savings calculations (hour-based)
# ---------------------------
supervisor_time_savings = sup_hours_saved_week_capped * weeks_per_year * supervisor_hourly_fully_loaded
qa_labor_savings = qa_hours_saved_week_capped * weeks_per_year * qa_hourly_loaded
training_savings = training_hours_saved_year_capped * trainer_hourly_fully_loaded
turnover_savings = dispatchers_retained_per_year * replacement_cost_per_dispatcher

total_gross_savings = (
    qa_labor_savings
    + supervisor_time_savings
    + turnover_savings
    + training_savings
    + annual_productivity_value
    + annual_risk_reduction_value
)
net_benefit = total_gross_savings - annual_investment
roi_ratio = safe_div(net_benefit, annual_investment)
payback_months = (annual_investment / total_gross_savings) * 12.0 if total_gross_savings else math.inf

breakdown_df = pd.DataFrame(
    [
        {"Bucket": "QA labor savings (hour-based)", "Annual Value ($)": qa_labor_savings},
        {"Bucket": "Supervisor time savings (hour-based)", "Annual Value ($)": supervisor_time_savings},
        {"Bucket": "Turnover savings (dispatchers retained)", "Annual Value ($)": turnover_savings},
        {"Bucket": "Training savings (hour-based)", "Annual Value ($)": training_savings},
        {"Bucket": "Productivity value", "Annual Value ($)": annual_productivity_value},
        {"Bucket": "Risk reduction value", "Annual Value ($)": annual_risk_reduction_value},
    ]
)

baseline_estimates = {
    "supervisor_hours_saved_year": sup_hours_saved_week_capped * weeks_per_year,
    "qa_hours_saved_year": qa_hours_saved_week_capped * weeks_per_year,
    "training_hours_saved_year": training_hours_saved_year_capped,
}

# ---------------------------
# Page outputs
# ---------------------------
m1, m2, m3, m4 = st.columns(4)
m1.metric("Annual Gross Savings", money(total_gross_savings))
m2.metric("Annual Investment", money(annual_investment))
m3.metric("Net Annual Benefit", money(net_benefit))
m4.metric("ROI", pct_from_ratio(roi_ratio))

if math.isfinite(payback_months):
    st.info(f"Estimated payback period: {payback_months:,.1f} months")
else:
    st.warning("Payback period cannot be computed because gross savings is $0.")

left, right = st.columns([1.1, 0.9])

with left:
    st.subheader("Savings breakdown")
    display_df = breakdown_df.copy()
    display_df["Annual Value ($)"] = display_df["Annual Value ($)"].map(lambda x: round(float(x), 0))
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    st.subheader("Product summary (selected modules)")
    if not selected_modules:
        st.write("Select one or more modules in the sidebar to see what is included.")
    else:
        for key in selected_modules:
            item = MODULE_SUMMARIES.get(key)
            if item:
                st.markdown(f"**{item['title']}**")
                st.write(item["body"])

with right:
    st.subheader("Manual QA effort estimator (baseline)")
    st.markdown(
        f"""
- Calls reviewed/year: **{qa_calls_reviewed_year:,.0f}** (coverage {manual_qa_coverage_pct:.1f}% of {annual_calls:,})
- Avg QA minutes per reviewed call: **{avg_qa_minutes_per_call:.0f}**
- Estimated baseline QA hours/year: **{baseline_qa_hours_year_est:,.0f}**
- Estimated baseline QA hours/week: **{baseline_qa_hours_week_est:,.1f}**
"""
    )

    st.subheader("Caps and time-saved inputs")
    st.markdown(
        f"""
- Supervisor baseline QA time (estimated or provided): **{supervisor_baseline_hours_week:,.1f} hrs/week**
- Supervisor hours saved (entered): **{supervisor_hours_saved_week:,.1f} hrs/week**
- Supervisor hours saved (capped): **{sup_hours_saved_week_capped:,.1f} hrs/week**

- QA baseline capacity: **{qa_baseline_hours_week_capacity:,.1f} hrs/week** ({qa_specialists_fte:.2f} FTE)
- QA hours saved (entered): **{qa_hours_saved_week:,.1f} hrs/week**
- QA hours saved (capped): **{qa_hours_saved_week_capped:,.1f} hrs/week**

- Training baseline: **{baseline_training_hours_year:,.0f} hrs/year**
- Training hours saved (capped): **{training_hours_saved_year_capped:,.0f} hrs/year**
"""
    )

st.subheader("Export")

inputs = {
    "Agency name": agency_name,
    "Scenario name": scenario_name,
    "Baseline metrics estimated?": estimate_missing_baselines,
    "Modules selected": ", ".join(selected_modules),
    "Annual CommsCoach cost ($)": annual_investment,
    "Annual calls eligible for QA": annual_calls,
    "Dispatchers / telecommunicators": dispatchers,
    "Supervisor fully loaded hourly rate ($/hr)": supervisor_hourly_fully_loaded,
    "QA specialists (FTE, baseline)": qa_specialists_fte,
    "QA specialist fully loaded annual cost ($)": qa_fte_fully_loaded_cost,
    "Trainer hourly fully loaded rate ($/hr)": trainer_hourly_fully_loaded,
    "Supervisors impacted (count)": supervisors_in_scope,
    "Hours saved per supervisor per week": hours_saved_per_supervisor_week,
    "Supervisor hours saved per week (total)": supervisor_hours_saved_week,
    "Supervisor hours saved per week (capped)": sup_hours_saved_week_capped,
    "QA staff impacted (count)": qa_specialists_in_scope,
    "Hours saved per QA staff per week": hours_saved_per_qa_week,
    "QA hours saved per week (total)": qa_hours_saved_week,
    "QA hours saved per week (capped)": qa_hours_saved_week_capped,
    "Trainers (headcount, context)": trainers_count,
    "Hours saved per trainer per week": hours_saved_per_trainer_week,
    "New hires per year": new_hires_per_year,
    "Training hours per new hire (baseline)": training_hours_per_new_hire,
    "Optional: hours saved per new hire": training_hours_saved_per_new_hire,
    "Training hours saved per year (capped)": training_hours_saved_year_capped,
    "Manual QA coverage (%)": manual_qa_coverage_pct,
    "Avg QA minutes per reviewed call": avg_qa_minutes_per_call,
    "Estimated baseline QA hours/year": baseline_qa_hours_year_est,
    "Dispatchers retained per year (count)": dispatchers_retained_per_year,
    "Replacement cost per dispatcher ($)": replacement_cost_per_dispatcher,
    "Washout factor (if estimated)": washout_factor,
    "Annual productivity / effectiveness value ($)": annual_productivity_value,
    "Annual risk reduction value ($)": annual_risk_reduction_value,
}

results = {
    "Annual gross savings ($)": total_gross_savings,
    "Annual investment ($)": annual_investment,
    "Net annual benefit ($)": net_benefit,
    "ROI ratio": roi_ratio,
    "Payback months": None if not math.isfinite(payback_months) else payback_months,
}

csv_bytes = pd.DataFrame(list(results.items()), columns=["Metric", "Value"]).to_csv(index=False).encode("utf-8")
st.download_button("Download results as CSV", data=csv_bytes, file_name="commscoach_roi_results.csv", mime="text/csv")

excel_bytes = build_excel_export(inputs, results, breakdown_df)
st.download_button(
    "Download full export as Excel",
    data=excel_bytes,
    file_name="commscoach_roi_export.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

pdf_bytes = build_pdf_report(
    agency_name=agency_name,
    scenario_name=scenario_name,
    selected_modules=selected_modules,
    inputs=inputs,
    results=results,
    breakdown_df=breakdown_df,
    baseline_estimates=baseline_estimates,
    logo_path=LOGO_PATH if LOGO_PATH.exists() else None,
    # live calculation values so the PDF mirrors the UI exactly
    qa_calls_reviewed_year=qa_calls_reviewed_year,
    manual_qa_coverage_pct=manual_qa_coverage_pct,
    annual_calls=annual_calls,
    avg_qa_minutes_per_call=avg_qa_minutes_per_call,
    baseline_qa_hours_year_est=baseline_qa_hours_year_est,
    baseline_qa_hours_week_est=baseline_qa_hours_week_est,
    supervisor_baseline_hours_week=supervisor_baseline_hours_week,
    supervisor_hours_saved_week=supervisor_hours_saved_week,
    sup_hours_saved_week_capped=sup_hours_saved_week_capped,
    qa_baseline_hours_week_capacity=qa_baseline_hours_week_capacity,
    qa_specialists_fte=qa_specialists_fte,
    qa_hours_saved_week=qa_hours_saved_week,
    qa_hours_saved_week_capped=qa_hours_saved_week_capped,
    baseline_training_hours_year=baseline_training_hours_year,
    training_hours_saved_year_capped=training_hours_saved_year_capped,
    supervisor_hourly_fully_loaded=supervisor_hourly_fully_loaded,
    qa_hourly_loaded=qa_hourly_loaded,
    trainer_hourly_fully_loaded=trainer_hourly_fully_loaded,
)
pdf_name = f"{agency_name.strip().replace(' ', '_')}_CommsCoach_ROI_Summary.pdf" if agency_name.strip() else "CommsCoach_ROI_Summary.pdf"
st.download_button("Download results as PDF", data=pdf_bytes, file_name=pdf_name, mime="application/pdf")
st.caption(DISCLAIMER)
