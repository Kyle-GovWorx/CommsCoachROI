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


def build_pdf_report(
    agency_name: str,
    scenario_name: str,
    selected_modules: list,
    inputs: dict,
    results: dict,
    breakdown_df: pd.DataFrame,
    baseline_estimates: dict,
    logo_path: Path | None = None,
) -> bytes:
    """Customer-ready PDF summary with cover page + assumptions."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter
    left = 0.75 * inch
    right = width - 0.75 * inch
    y = height - 0.75 * inch

    # COVER PAGE
    pdf_logo = logo_path
    if "LOGO_PDF_PATH" in globals() and LOGO_PDF_PATH.exists():
        pdf_logo = LOGO_PDF_PATH

    if pdf_logo and pdf_logo.exists():
        try:
            img = ImageReader(str(pdf_logo))
            c.drawImage(img, left, y - 0.55 * inch, width=3.0 * inch, height=0.55 * inch, mask="auto")
        except Exception:
            pass

    c.setFont("Helvetica-Bold", 18)
    c.drawRightString(right, y - 0.10 * inch, "CommsCoach ROI Summary")
    c.setFont("Helvetica", 10)
    c.drawRightString(right, y - 0.30 * inch, datetime.now().strftime("%Y-%m-%d"))
    y -= 0.90 * inch

    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Agency:")
    c.setFont("Helvetica", 12)
    c.drawString(left + 1.1 * inch, y, agency_name or "Not specified")
    y -= 0.22 * inch

    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Scenario:")
    c.setFont("Helvetica", 12)
    c.drawString(left + 1.1 * inch, y, scenario_name or "Standard")
    y -= 0.22 * inch

    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Modules:")
    c.setFont("Helvetica", 12)
    c.drawString(left + 1.1 * inch, y, ", ".join(selected_modules) if selected_modules else "Not specified")
    y -= 0.35 * inch

    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Key results (annual)")
    y -= 0.20 * inch

    def kv(label: str, value: str):
        nonlocal y
        c.setFont("Helvetica-Bold", 10)
        c.drawString(left, y, label)
        c.setFont("Helvetica", 10)
        c.drawRightString(right, y, value)
        y -= 0.18 * inch

    kv("Annual gross savings", money(float(results.get("Annual gross savings ($)", 0) or 0)))
    kv("Annual investment", money(float(results.get("Annual investment ($)", 0) or 0)))
    kv("Net annual benefit", money(float(results.get("Net annual benefit ($)", 0) or 0)))
    kv("ROI", pct_from_ratio(float(results.get("ROI ratio", 0) or 0)))
    pb = results.get("Payback months", None)
    kv("Payback period (months)", f"{float(pb):.1f}" if pb is not None else "N/A")

    y -= 0.10 * inch
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Estimated time saved (annual)")
    y -= 0.22 * inch
    c.setFont("Helvetica", 10)
    kv("Supervisor time saved", f"{float(baseline_estimates.get('supervisor_hours_saved_year', 0) or 0):,.0f} hrs")
    kv("QA time saved", f"{float(baseline_estimates.get('qa_hours_saved_year', 0) or 0):,.0f} hrs")
    kv("Trainer time saved", f"{float(baseline_estimates.get('training_hours_saved_year', 0) or 0):,.0f} hrs")

    c.setFont("Helvetica", 8)
    
    y -= 0.10 * inch
    if y < 1.7 * inch:
        y = 1.7 * inch
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left, y, "Disclaimer")
    y -= 0.18 * inch
    c.setFont("Helvetica", 9)
    words = DISCLAIMER.split()
    line = ""
    for w in words:
        test = (line + " " + w).strip()
        if len(test) > 95:
            c.drawString(left, y, line)
            y -= 0.14 * inch
            line = w
        else:
            line = test
    if line:
        c.drawString(left, y, line)
        y -= 0.14 * inch

    _pdf_new_page(c, "Assumptions and savings detail")

    # PAGE 2: ASSUMPTIONS TABLE
    y = height - 1.05 * inch
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Assumptions used")
    y -= 0.22 * inch
    c.setFont("Helvetica", 9)

    # Assumptions list (top 18 or so)
    rows = []
    for k, v in inputs.items():
        rows.append((str(k), str(v)))

    max_rows = 18
    for i, (k, v) in enumerate(rows[:max_rows]):
        c.setFont("Helvetica-Bold", 9)
        c.drawString(left, y, k[:48])
        c.setFont("Helvetica", 9)
        c.drawRightString(right, y, v[:40])
        y -= 0.16 * inch
        if y < 2.2 * inch:
            break

    if y < 2.4 * inch:
        _pdf_new_page(c, "Savings breakdown")
        y = height - 1.05 * inch
    else:
        y -= 0.18 * inch

    # Savings breakdown
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Savings breakdown (annual)")
    y -= 0.22 * inch
    c.setFont("Helvetica-Bold", 9)
    c.drawString(left, y, "Bucket")
    c.drawRightString(right, y, "Value")
    y -= 0.14 * inch
    c.line(left, y, right, y)
    y -= 0.12 * inch
    c.setFont("Helvetica", 9)

    for _, row in breakdown_df.iterrows():
        bucket = str(row.get("Bucket", ""))[:70]
        val = float(row.get("Annual Value ($)", 0) or 0)
        c.drawString(left, y, bucket)
        c.drawRightString(right, y, money(val))
        y -= 0.16 * inch
        if y < 1.3 * inch:
            _pdf_new_page(c, "Savings breakdown (continued)")
            y = height - 1.05 * inch
            c.setFont("Helvetica", 9)

    # Module summary (short)
    _pdf_new_page(c, "What is included (selected modules)")
    y = height - 1.05 * inch
    c.setFont("Helvetica", 10)
    for key in selected_modules:
        item = MODULE_SUMMARIES.get(key)
        if not item:
            continue
        c.setFont("Helvetica-Bold", 11)
        c.drawString(left, y, item["title"])
        y -= 0.18 * inch
        c.setFont("Helvetica", 9)
        # wrap text manually
        words = item["body"].split()
        line = ""
        for w in words:
            test = (line + " " + w).strip()
            if len(test) > 95:
                c.drawString(left, y, line)
                y -= 0.14 * inch
                line = w
                if y < 1.2 * inch:
                    _pdf_new_page(c, "What is included (continued)")
                    y = height - 1.05 * inch
                    c.setFont("Helvetica", 9)
            else:
                line = test
        if line:
            c.drawString(left, y, line)
            y -= 0.20 * inch
        y -= 0.10 * inch

    c.setFont("Helvetica", 8)
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
)
pdf_name = f"{agency_name.strip().replace(' ', '_')}_CommsCoach_ROI_Summary.pdf" if agency_name.strip() else "CommsCoach_ROI_Summary.pdf"
st.download_button("Download results as PDF", data=pdf_bytes, file_name=pdf_name, mime="application/pdf")
st.caption(DISCLAIMER)
