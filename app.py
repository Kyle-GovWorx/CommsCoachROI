import io
import math
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).parent
ASSETS = APP_DIR / "assets"
LOGO_PATH = ASSETS / "commscoach_logo_reverse.png"     # Replace with official GovWorx/CommsCoach logo per brand folder
FAVICON_PATH = ASSETS / "favicon.png"


def inject_brand_css():
    # Import Inter + Manrope per GovWorx brand guidelines, then apply a simple hierarchy.
    # Streamlit theme controls primary colors; CSS handles typography + some layout polish.
    st.markdown(
        """
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Manrope:wght@400;500;700&display=swap');

          html, body, [class*="css"]  {
            font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
          }

          h1, h2, h3, h4 {
            font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
            letter-spacing: -0.02em;
          }

          .gw-eyebrow {
            font-family: Manrope, Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
            letter-spacing: 0.10em;
            text-transform: uppercase;
            font-weight: 700;
            font-size: 0.78rem;
            color: rgba(14, 51, 81, 0.85);
          }

          .block-container { padding-top: 1.4rem; }
          header { visibility: hidden; }
          footer { visibility: hidden; }

          .gw-hero {
            background: linear-gradient(135deg, #0E1F2B 0%, #0E3351 100%);
            border-radius: 18px;
            padding: 18px 18px 10px 18px;
            margin-bottom: 1rem;
          }

          .gw-hero h2, .gw-hero p, .gw-hero span { color: #FFFFFF; }

          .gw-pill {
            display: inline-block;
            padding: 0.28rem 0.65rem;
            border-radius: 999px;
            border: 1px solid rgba(255, 255, 255, 0.25);
            margin-right: 0.5rem;
            margin-top: 0.25rem;
            font-size: 0.85rem;
            color: #FFFFFF;
            background: rgba(255, 255, 255, 0.08);
          }

          .gw-footer {
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid rgba(15, 23, 42, 0.12);
            color: rgba(15, 23, 42, 0.75);
            font-size: 0.9rem;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


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
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(list(inputs.items()), columns=["Input", "Value"]).to_excel(writer, sheet_name="Inputs", index=False)
        pd.DataFrame(list(results.items()), columns=["Metric", "Value"]).to_excel(writer, sheet_name="Results", index=False)
        breakdown_df.to_excel(writer, sheet_name="Savings Breakdown", index=False)
    buffer.seek(0)
    return buffer.read()


icon = str(FAVICON_PATH) if FAVICON_PATH.exists() else "📈"
st.set_page_config(page_title="GovWorx | CommsCoach ROI Calculator", page_icon=icon, layout="wide")
inject_brand_css()

# Hero header (dark navy to sentinel blue gradient per brand guidelines)
st.markdown('<div class="gw-hero">', unsafe_allow_html=True)
cols = st.columns([0.55, 0.45])
with cols[0]:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=280)
    st.markdown('<div class="gw-eyebrow">Mission-first, people-first</div>', unsafe_allow_html=True)
    st.markdown("## CommsCoach ROI Calculator")
    st.write("Estimate annual value, net benefit, ROI, and payback period for CommsCoach.")
with cols[1]:
    st.markdown(
        """
        <div style="text-align:right;">
          <span class="gw-pill">QA</span>
          <span class="gw-pill">TRAIN</span>
          <span class="gw-pill">HIRE</span>
          <span class="gw-pill">ASSIST</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
st.markdown("</div>", unsafe_allow_html=True)

# Inputs (sidebar)
with st.sidebar:
    st.subheader("Inputs")
    st.caption("Use your center’s best available numbers. Keep optional buckets at $0 for a conservative case.")

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

# Calculations
weeks_per_year = 52.0
manual_qa_labor_cost = qa_specialists_fte * qa_fte_fully_loaded_cost
manual_supervisor_qa_cost = supervisor_hours_per_week_on_qa * weeks_per_year * supervisor_hourly_fully_loaded
qa_labor_savings = manual_qa_labor_cost * (qa_labor_reduction_pct / 100.0)
supervisor_time_savings = manual_supervisor_qa_cost * (supervisor_time_reduction_pct / 100.0)

annual_turnovers = dispatchers * (annual_turnover_pct / 100.0)
turnover_cost_baseline = annual_turnovers * replacement_cost_per_dispatcher
turnover_savings = turnover_cost_baseline * (turnover_reduction_pct / 100.0)

baseline_training_cost = new_hires_per_year * training_hours_per_new_hire * trainer_hourly_fully_loaded
training_savings = baseline_training_cost * (training_time_reduction_pct / 100.0)

manual_cov = max(0.0, min(1.0, manual_qa_coverage_pct / 100.0))
coverage_uplift_pct = max(0.0, 1.0 - manual_cov) * 100.0

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

# Outputs
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
    display_df = breakdown_df.copy()
    display_df["Annual Value ($)"] = display_df["Annual Value ($)"].map(lambda x: round(float(x), 0))
    st.dataframe(display_df, use_container_width=True, hide_index=True)

with right:
    st.subheader("Context")
    st.markdown(
        f"""
- Manual QA labor baseline: **{money(manual_qa_labor_cost)}**
- Manual supervisor QA baseline: **{money(manual_supervisor_qa_cost)}**
- Coverage uplift vs manual review: **{coverage_uplift_pct:,.1f}%**
- Annual turnovers at baseline: **{annual_turnovers:,.1f}**
- Turnover cost baseline: **{money(turnover_cost_baseline)}**
- Baseline training cost: **{money(baseline_training_cost)}**
"""
    )

st.subheader("Export")
results = {
    "Annual gross savings ($)": total_gross_savings,
    "Annual investment ($)": annual_investment,
    "Net annual benefit ($)": net_benefit,
    "ROI ratio": roi_ratio,
    "Payback months": None if not math.isfinite(payback_months) else payback_months,
}
csv_bytes = pd.DataFrame(list(results.items()), columns=["Metric", "Value"]).to_csv(index=False).encode("utf-8")
st.download_button("Download results as CSV", data=csv_bytes, file_name="commscoach_roi_results.csv", mime="text/csv")

excel_bytes = build_excel_export({}, results, breakdown_df)
st.download_button(
    "Download full export as Excel",
    data=excel_bytes,
    file_name="commscoach_roi_export.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.markdown(
    f"""
    <div class="gw-footer">
      <div><strong>GovWorx</strong> | CommsCoach ROI Calculator</div>
      <div>Last calculated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
    </div>
    """,
    unsafe_allow_html=True,
)
