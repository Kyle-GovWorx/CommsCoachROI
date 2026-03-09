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

APP_DIR = Path(__file__).parent
ASSETS = APP_DIR / "assets"
LOGO_PATH = ASSETS / "commscoach_logo_reverse.png"
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


def build_pdf_report(
    agency_name: str,
    selected_modules: list,
    inputs: dict,
    results: dict,
    breakdown_df: pd.DataFrame,
    logo_path: Path | None = None,
) -> bytes:
    """Printable PDF summary for agency leadership."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter
    left = 0.75 * inch
    right = width - 0.75 * inch
    y = height - 0.75 * inch

    def new_page():
        nonlocal y
        c.showPage()
        y = height - 0.75 * inch

    def ensure_space(min_y: float):
        nonlocal y
        if y < min_y:
            new_page()

    # Header
    if logo_path and logo_path.exists():
        try:
            img = ImageReader(str(logo_path))
            c.drawImage(img, left, y - 0.55 * inch, width=2.6 * inch, height=0.55 * inch, mask="auto")
        except Exception:
            pass

    c.setFont("Helvetica-Bold", 16)
    c.drawRightString(right, y - 0.15 * inch, "CommsCoach ROI Summary")
    c.setFont("Helvetica", 10)
    c.drawRightString(right, y - 0.33 * inch, datetime.now().strftime("%Y-%m-%d"))
    y -= 0.85 * inch

    # Agency + modules
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Agency:")
    c.setFont("Helvetica", 12)
    c.drawString(left + 1.1 * inch, y, agency_name or "Not specified")
    y -= 0.25 * inch

    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Modules:")
    c.setFont("Helvetica", 12)
    c.drawString(left + 1.1 * inch, y, ", ".join(selected_modules) if selected_modules else "Not specified")
    y -= 0.35 * inch

    # Key results (omit 0 values)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Key results (annual)")
    y -= 0.25 * inch

    def draw_kv(label: str, value_str: str, numeric_value: float | None = None):
        nonlocal y
        if numeric_value is not None and abs(numeric_value) < 1e-9:
            return
        c.setFont("Helvetica-Bold", 10)
        c.drawString(left, y, label)
        c.setFont("Helvetica", 10)
        c.drawRightString(right, y, value_str)
        y -= 0.18 * inch

    gross = float(results.get("Annual gross savings ($)", 0) or 0)
    invest = float(results.get("Annual investment ($)", 0) or 0)
    net = float(results.get("Net annual benefit ($)", 0) or 0)
    roi = float(results.get("ROI ratio", 0) or 0)
    pb = results.get("Payback months", None)

    draw_kv("Annual gross savings", money(gross), gross)
    draw_kv("Annual investment", money(invest), invest)
    draw_kv("Net annual benefit", money(net), net)
    draw_kv("ROI", pct_from_ratio(roi), roi)

    if pb is not None:
        try:
            pb_f = float(pb)
            if pb_f != 0:
                draw_kv("Payback period (months)", f"{pb_f:.1f}", pb_f)
        except Exception:
            # If it's not numeric and not empty, include it
            if str(pb).strip():
                draw_kv("Payback period (months)", str(pb).strip(), None)

    y -= 0.10 * inch

    # Savings breakdown (omit rows with 0)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Savings breakdown (annual)")
    y -= 0.25 * inch
    c.setFont("Helvetica-Bold", 9)
    c.drawString(left, y, "Bucket")
    c.drawRightString(right, y, "Value")
    y -= 0.15 * inch
    c.line(left, y, right, y)
    y -= 0.12 * inch
    c.setFont("Helvetica", 9)

    any_rows = False
    for _, row in breakdown_df.iterrows():
        val = float(row.get("Annual Value ($)", 0) or 0)
        if abs(val) < 1e-9:
            continue
        any_rows = True
        bucket = str(row.get("Bucket", ""))[:70]
        c.drawString(left, y, bucket)
        c.drawRightString(right, y, money(val))
        y -= 0.16 * inch
        ensure_space(1.5 * inch)

    if not any_rows:
        c.setFont("Helvetica", 9)
        c.drawString(left, y, "No savings buckets were included (all were $0).")
        y -= 0.16 * inch

    y -= 0.20 * inch
    ensure_space(2.8 * inch)

    # Product summary (selected modules)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "Product summary (selected modules)")
    y -= 0.22 * inch
    c.setFont("Helvetica", 10)

    if not selected_modules:
        c.drawString(left, y, "No modules selected.")
        y -= 0.18 * inch
    else:
        for key in selected_modules:
            item = MODULE_SUMMARIES.get(key)
            if not item:
                continue
            title = item.get("title", key)
            body = item.get("body", "")
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, title)
            y -= 0.16 * inch
            c.setFont("Helvetica", 9)

            # Simple wrap
            wrap_width = 105  # characters, approximate
            words = body.split()
            line = ""
            lines = []
            for w in words:
                if len(line) + len(w) + 1 <= wrap_width:
                    line = (line + " " + w).strip()
                else:
                    lines.append(line)
                    line = w
            if line:
                lines.append(line)

            for ln in lines:
                c.drawString(left, y, ln)
                y -= 0.14 * inch
                ensure_space(1.2 * inch)

            y -= 0.10 * inch
            ensure_space(1.2 * inch)

    y -= 0.10 * inch
    ensure_space(2.2 * inch)

    # Narrative (always include, but keep short)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, y, "What drives the savings")
    y -= 0.22 * inch
    c.setFont("Helvetica", 10)
    bullets = [
        "Reduced supervisor time spent on manual QA review and follow-up.",
        "Reduced dedicated QA labor required to score, document, and trend performance.",
        "Reduced turnover and onboarding burden through faster feedback loops and targeted coaching.",
        "Reduced training time via simulations and structured training workflows (when TRAIN is included).",
    ]
    for b in bullets:
        c.drawString(left + 0.12 * inch, y, "- " + b)
        y -= 0.18 * inch
        ensure_space(1.0 * inch)

    y -= 0.10 * inch
    ensure_space(1.6 * inch)

    # Time saved (omit 0 values)
    sup_week = float(inputs.get("Supervisor hours per week spent on QA", 0) or 0)
    sup_red = float(inputs.get("Reduction in supervisor QA time (%)", 0) or 0) / 100.0
    qa_fte = float(inputs.get("Dedicated QA specialists (FTE)", 0) or 0)
    qa_red = float(inputs.get("Reduction in manual QA labor (%)", 0) or 0) / 100.0

    sup_hours_saved = sup_week * 52 * sup_red
    qa_hours_saved = qa_fte * 2080 * qa_red

    if sup_hours_saved > 0 or qa_hours_saved > 0:
        c.setFont("Helvetica-Bold", 12)
        c.drawString(left, y, "Estimated time saved")
        y -= 0.22 * inch
        c.setFont("Helvetica", 10)

        if sup_hours_saved > 0:
            c.drawString(left + 0.12 * inch, y, f"- Supervisor time saved: ~{sup_hours_saved:,.0f} hours per year")
            y -= 0.18 * inch
        if qa_hours_saved > 0:
            c.drawString(left + 0.12 * inch, y, f"- QA labor time saved: ~{qa_hours_saved:,.0f} hours per year")
            y -= 0.18 * inch

    c.setFont("Helvetica", 8)
    c.drawString(left, 0.6 * inch, "GovWorx | CommsCoach ROI Calculator. Estimates only; results depend on adoption and baseline practices.")
    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()


icon = str(FAVICON_PATH) if FAVICON_PATH.exists() else "ROI"
st.set_page_config(page_title="GovWorx | CommsCoach ROI Calculator", page_icon=icon, layout="wide")

st.title("GovWorx | CommsCoach ROI Calculator")
st.caption("Estimate annual value, net benefit, ROI, and payback period for CommsCoach. Export to CSV, Excel, and PDF.")

with st.sidebar:
    st.subheader("Inputs")
    agency_name = st.text_input("Agency name (for exports)", value="")

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

    st.markdown("#### Agency size")
    annual_calls = st.number_input("Annual calls eligible for QA", min_value=0, value=600000, step=10000)
    dispatchers = st.number_input("Number of dispatchers / telecommunicators", min_value=0, value=75, step=5)

    st.markdown("#### Manual QA baseline")
    qa_specialists_fte = st.number_input("Dedicated QA specialists (FTE)", min_value=0.0, value=1.0, step=0.25)
    qa_fte_fully_loaded_cost = st.number_input("QA specialist fully loaded annual cost ($)", min_value=0.0, value=95000.0, step=5000.0)
    supervisor_hours_per_week_on_qa = st.number_input("Supervisor hours per week spent on QA", min_value=0.0, value=25.0, step=1.0)
    supervisor_hourly_fully_loaded = st.number_input("Supervisor fully loaded hourly rate ($/hr)", min_value=0.0, value=110.0, step=1.0)
    manual_qa_coverage_pct = st.number_input("Manual QA coverage (% of calls)", min_value=0.0, max_value=100.0, value=2.0, step=1.0)

    st.markdown("#### CommsCoach impact")
    qa_labor_reduction_pct = st.number_input("Reduction in manual QA labor (%)", min_value=0.0, max_value=100.0, value=70.0, step=5.0)
    supervisor_time_reduction_pct = st.number_input("Reduction in supervisor QA time (%)", min_value=0.0, max_value=100.0, value=60.0, step=5.0)

    st.markdown("#### Turnover")
    annual_turnover_pct = st.number_input("Annual dispatcher turnover (%)", min_value=0.0, max_value=100.0, value=12.0, step=1.0)
    replacement_cost_per_dispatcher = st.number_input("Replacement cost per dispatcher ($)", min_value=0.0, value=35000.0, step=1000.0)
    turnover_reduction_pct = st.number_input("Turnover reduction due to CommsCoach (%)", min_value=0.0, max_value=100.0, value=5.0, step=1.0)

    st.markdown("#### Training")
    new_hires_per_year = st.number_input("New hires per year", min_value=0, value=15, step=1)
    training_hours_per_new_hire = st.number_input("Training hours per new hire", min_value=0.0, value=80.0, step=1.0)
    trainer_hourly_fully_loaded = st.number_input("Trainer fully loaded hourly rate ($/hr)", min_value=0.0, value=95.0, step=1.0)
    training_time_reduction_pct = st.number_input("Training time reduction (%)", min_value=0.0, max_value=100.0, value=15.0, step=1.0)

    st.markdown("#### Optional value buckets")
    annual_productivity_value = st.number_input("Annual productivity / effectiveness value ($)", min_value=0.0, value=0.0, step=10000.0)
    annual_risk_reduction_value = st.number_input("Annual risk reduction value ($)", min_value=0.0, value=0.0, step=10000.0)

weeks_per_year = 52.0
manual_qa_labor_cost = qa_specialists_fte * qa_fte_fully_loaded_cost
manual_supervisor_qa_cost = supervisor_hours_per_week_on_qa * weeks_per_year * supervisor_hourly_fully_loaded
manual_total_cost = manual_qa_labor_cost + manual_supervisor_qa_cost

qa_labor_savings = manual_qa_labor_cost * (qa_labor_reduction_pct / 100.0)
supervisor_time_savings = manual_supervisor_qa_cost * (supervisor_time_reduction_pct / 100.0)

annual_turnovers = dispatchers * (annual_turnover_pct / 100.0)
turnover_cost_baseline = annual_turnovers * replacement_cost_per_dispatcher
turnover_savings = turnover_cost_baseline * (turnover_reduction_pct / 100.0)

baseline_training_cost = new_hires_per_year * training_hours_per_new_hire * trainer_hourly_fully_loaded
training_savings = baseline_training_cost * (training_time_reduction_pct / 100.0)

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
        {"Bucket": "QA specialist labor savings", "Annual Value ($)": qa_labor_savings},
        {"Bucket": "Supervisor time savings", "Annual Value ($)": supervisor_time_savings},
        {"Bucket": "Turnover savings", "Annual Value ($)": turnover_savings},
        {"Bucket": "Training savings", "Annual Value ($)": training_savings},
        {"Bucket": "Productivity value", "Annual Value ($)": annual_productivity_value},
        {"Bucket": "Risk reduction value", "Annual Value ($)": annual_risk_reduction_value},
    ]
)

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
    st.subheader("Context")
    st.markdown(
        f"""
- Manual QA labor baseline: **{money(manual_qa_labor_cost)}**
- Manual supervisor QA baseline: **{money(manual_supervisor_qa_cost)}**
- Manual total QA baseline: **{money(manual_total_cost)}**
- Annual turnovers at baseline: **{annual_turnovers:,.1f}**
- Turnover cost baseline: **{money(turnover_cost_baseline)}**
- Baseline training cost: **{money(baseline_training_cost)}**
"""
    )

st.subheader("Export")

inputs = {
    "Annual CommsCoach cost ($)": annual_investment,
    "Annual calls eligible for QA": annual_calls,
    "Dispatchers / telecommunicators": dispatchers,
    "Dedicated QA specialists (FTE)": qa_specialists_fte,
    "QA specialist fully loaded annual cost ($)": qa_fte_fully_loaded_cost,
    "Supervisor hours per week spent on QA": supervisor_hours_per_week_on_qa,
    "Supervisor hourly fully loaded rate ($/hr)": supervisor_hourly_fully_loaded,
    "Manual QA coverage (%)": manual_qa_coverage_pct,
    "Reduction in manual QA labor (%)": qa_labor_reduction_pct,
    "Reduction in supervisor QA time (%)": supervisor_time_reduction_pct,
    "Annual dispatcher turnover (%)": annual_turnover_pct,
    "Replacement cost per dispatcher ($)": replacement_cost_per_dispatcher,
    "Turnover reduction due to CommsCoach (%)": turnover_reduction_pct,
    "New hires per year": new_hires_per_year,
    "Training hours per new hire": training_hours_per_new_hire,
    "Trainer fully loaded hourly rate ($/hr)": trainer_hourly_fully_loaded,
    "Training time reduction (%)": training_time_reduction_pct,
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
    selected_modules=selected_modules,
    inputs=inputs,
    results=results,
    breakdown_df=breakdown_df,
    logo_path=LOGO_PATH if LOGO_PATH.exists() else None,
)
pdf_name = f"{agency_name.strip().replace(' ', '_')}_CommsCoach_ROI_Summary.pdf" if agency_name.strip() else "CommsCoach_ROI_Summary.pdf"
st.download_button("Download results as PDF", data=pdf_bytes, file_name=pdf_name, mime="application/pdf")
