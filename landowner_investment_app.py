from __future__ import annotations

from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

MODEL_PATH = Path("V4 ANA Feasibility Financial Model 11.06.25 CS.xlsx")
YEARS = [1, 2, 3, 4]


def safe_float(value, default=0.0):
    if value is None:
        return float(default)
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(default)


def solve_irr(cashflows):
    # Bisection IRR; returns None if sign does not change.
    if not (any(cf > 0 for cf in cashflows) and any(cf < 0 for cf in cashflows)):
        return None

    def npv(rate):
        return sum(cf / ((1 + rate) ** t) for t, cf in enumerate(cashflows))

    lo, hi = -0.95, 5.0
    f_lo, f_hi = npv(lo), npv(hi)
    if f_lo * f_hi > 0:
        return None

    for _ in range(250):
        mid = (lo + hi) / 2
        f_mid = npv(mid)
        if abs(f_mid) < 1e-10:
            return mid
        if f_lo * f_mid <= 0:
            hi, f_hi = mid, f_mid
        else:
            lo, f_lo = mid, f_mid
    return mid


@st.cache_data
def load_defaults(model_path: str) -> dict:
    wb = openpyxl.load_workbook(model_path, data_only=True)

    sales_mix = [
        safe_float(wb["Sales Phasing"]["D2"].value, 0.25),
        safe_float(wb["Sales Phasing"]["D6"].value, 0.25),
        safe_float(wb["Sales Phasing"]["D10"].value, 0.25),
        safe_float(wb["Sales Phasing"]["D14"].value, 0.25),
    ]
    total_mix = sum(sales_mix) or 1.0
    sales_mix = [x / total_mix for x in sales_mix]

    return {
        "resi_area": safe_float(wb["FSI & Area Cal"]["E48"].value),
        "comm_area": safe_float(wb["FSI & Area Cal"]["E49"].value),
        "start_price_resi": safe_float(wb["Input"]["C7"].value),
        "start_price_comm": safe_float(wb["Input"]["C8"].value),
        "sales_growth": safe_float(wb["Input"]["C11"].value, 0.045),
        "landowner_share": safe_float(wb["ANA Summary"]["E46"].value, 0.234),
        "deposit": safe_float(wb["ANA Summary"]["F48"].value, 75.0),
        "sales_mix": sales_mix,
    }


def compute_landowner_model(
    resi_area,
    comm_area,
    start_price_resi,
    start_price_comm,
    sales_growth,
    sales_mix,
    landowner_share,
    deposit,
    annual_landowner_cost,
    tax_rate,
    entry_land_value,
    discount_rate,
):
    rows = []

    for i, year in enumerate(YEARS):
        resi_price = start_price_resi * ((1 + sales_growth) ** i)
        comm_price = start_price_comm * ((1 + sales_growth) ** i)
        gross_revenue = (
            resi_area * sales_mix[i] * resi_price + comm_area * sales_mix[i] * comm_price
        ) / 1e7
        lo_gross_share = gross_revenue * landowner_share

        rows.append(
            {
                "Year": f"Y{year}",
                "Sales Mix": sales_mix[i],
                "Resi Price (Rs/sft)": resi_price,
                "Comm Price (Rs/sft)": comm_price,
                "Gross Project Revenue (Cr)": gross_revenue,
                "Landowner Gross Share (Cr)": lo_gross_share,
            }
        )

    lo_total_share = sum(r["Landowner Gross Share (Cr)"] for r in rows)
    deposit_recovery_ratio = min(deposit / lo_total_share, 1.0) if lo_total_share > 0 else 0.0

    for i, row in enumerate(rows):
        deposit_adjustment = -row["Landowner Gross Share (Cr)"] * deposit_recovery_ratio
        net_share_after_adjustment = row["Landowner Gross Share (Cr)"] + deposit_adjustment
        deposit_inflow = deposit if i == 0 else 0.0
        landowner_cost = annual_landowner_cost
        taxable_profit = max(net_share_after_adjustment - landowner_cost, 0.0)
        tax = taxable_profit * tax_rate
        net_cashflow = net_share_after_adjustment + deposit_inflow - landowner_cost - tax

        row["Deposit Adjustment (Cr)"] = deposit_adjustment
        row["Net Share After Adjustment (Cr)"] = net_share_after_adjustment
        row["Deposit Inflow (Cr)"] = deposit_inflow
        row["Landowner Cost (Cr)"] = landowner_cost
        row["Tax (Cr)"] = tax
        row["Net Landowner Cashflow (Cr)"] = net_cashflow

    df = pd.DataFrame(rows)

    yearly_cashflows = df["Net Landowner Cashflow (Cr)"].tolist()
    full_cashflows = [-entry_land_value] + yearly_cashflows

    npv = sum(cf / ((1 + discount_rate) ** t) for t, cf in enumerate(full_cashflows))
    irr = solve_irr(full_cashflows)

    cumulative = -entry_land_value
    payback_year = None
    for i, cf in enumerate(yearly_cashflows, start=1):
        cumulative += cf
        if cumulative >= 0 and payback_year is None:
            payback_year = i

    summary = {
        "total_project_revenue": df["Gross Project Revenue (Cr)"].sum(),
        "landowner_gross_entitlement": lo_total_share,
        "deposit_recovery_ratio": deposit_recovery_ratio,
        "net_landowner_inflow": df["Net Landowner Cashflow (Cr)"].sum(),
        "npv": npv,
        "irr": irr,
        "payback_year": payback_year,
        "profit_over_land_value": df["Net Landowner Cashflow (Cr)"].sum() - entry_land_value,
        "moic": (df["Net Landowner Cashflow (Cr)"].sum() / entry_land_value) if entry_land_value > 0 else None,
        "cashflows": full_cashflows,
    }
    return df, summary


def pct_to_str(x):
    return "NA" if x is None else f"{x * 100:.2f}%"


st.set_page_config(page_title="Landowner Investment Model", layout="wide")
st.title("Landowner Investment Financial Model")
st.caption("Adjust assumptions manually and evaluate Landowner returns, profit, NPV and IRR.")

if not MODEL_PATH.exists():
    st.error(f"Source model not found at: {MODEL_PATH}")
    st.stop()

defaults = load_defaults(str(MODEL_PATH))

with st.sidebar:
    st.header("Assumptions")
    resi_area = st.number_input("Residential Sale Area (sft)", min_value=0.0, value=defaults["resi_area"], step=1000.0)
    comm_area = st.number_input("Commercial Sale Area (sft)", min_value=0.0, value=defaults["comm_area"], step=500.0)
    start_price_resi = st.number_input(
        "Starting Resi Price (Rs/sft)", min_value=0.0, value=defaults["start_price_resi"], step=500.0
    )
    start_price_comm = st.number_input(
        "Starting Comm Price (Rs/sft)", min_value=0.0, value=defaults["start_price_comm"], step=500.0
    )
    sales_growth_pct = st.slider("Sales Growth YOY (%)", min_value=-5.0, max_value=20.0, value=defaults["sales_growth"] * 100, step=0.1)
    landowner_share_pct = st.slider(
        "Landowner Revenue Share (%)", min_value=0.0, max_value=100.0, value=defaults["landowner_share"] * 100, step=0.1
    )
    deposit = st.number_input("Upfront Refundable Deposit (Cr)", min_value=0.0, value=defaults["deposit"], step=1.0)

    st.subheader("Sales Mix by Year")
    mix_values = []
    for i, d in enumerate(defaults["sales_mix"], start=1):
        mix_values.append(
            st.number_input(f"Y{i} Sales Mix", min_value=0.0, max_value=1.0, value=float(d), step=0.01, format="%.2f")
        )

    mix_total = sum(mix_values)
    if mix_total <= 0:
        sales_mix = [0.25, 0.25, 0.25, 0.25]
    else:
        sales_mix = [x / mix_total for x in mix_values]

    st.caption(f"Sales mix auto-normalized to 100%. Current total entered: {mix_total:.2f}")

    st.subheader("Landowner Cost & Valuation")
    annual_landowner_cost = st.number_input("Annual Landowner Cost (Cr)", min_value=0.0, value=0.0, step=1.0)
    tax_rate_pct = st.slider("Tax Rate on Annual Landowner Profit (%)", min_value=0.0, max_value=50.0, value=0.0, step=0.5)
    entry_land_value = st.number_input(
        "Landowner Entry Land Value / Cost Basis (Cr)", min_value=0.0, value=0.0, step=5.0
    )
    discount_rate_pct = st.slider("Discount Rate for NPV (%)", min_value=0.0, max_value=30.0, value=15.0, step=0.5)

sales_growth = sales_growth_pct / 100
landowner_share = landowner_share_pct / 100
tax_rate = tax_rate_pct / 100
discount_rate = discount_rate_pct / 100

results_df, summary = compute_landowner_model(
    resi_area=resi_area,
    comm_area=comm_area,
    start_price_resi=start_price_resi,
    start_price_comm=start_price_comm,
    sales_growth=sales_growth,
    sales_mix=sales_mix,
    landowner_share=landowner_share,
    deposit=deposit,
    annual_landowner_cost=annual_landowner_cost,
    tax_rate=tax_rate,
    entry_land_value=entry_land_value,
    discount_rate=discount_rate,
)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Project Revenue", f"{summary['total_project_revenue']:.2f} Cr")
k2.metric("Landowner Gross Entitlement", f"{summary['landowner_gross_entitlement']:.2f} Cr")
k3.metric("Net Landowner Inflow", f"{summary['net_landowner_inflow']:.2f} Cr")
k4.metric("Profit Over Land Value", f"{summary['profit_over_land_value']:.2f} Cr")

k5, k6, k7, k8 = st.columns(4)
k5.metric("NPV", f"{summary['npv']:.2f} Cr")
k6.metric("IRR", pct_to_str(summary["irr"]))
k7.metric("Payback Year", "NA" if summary["payback_year"] is None else f"Y{summary['payback_year']}")
k8.metric("MOIC", "NA" if summary["moic"] is None else f"{summary['moic']:.2f}x")

st.subheader("Landowner Cashflow Schedule")
st.dataframe(
    results_df.style.format(
        {
            "Sales Mix": "{:.2%}",
            "Resi Price (Rs/sft)": "{:,.0f}",
            "Comm Price (Rs/sft)": "{:,.0f}",
            "Gross Project Revenue (Cr)": "{:,.2f}",
            "Landowner Gross Share (Cr)": "{:,.2f}",
            "Deposit Adjustment (Cr)": "{:,.2f}",
            "Net Share After Adjustment (Cr)": "{:,.2f}",
            "Deposit Inflow (Cr)": "{:,.2f}",
            "Landowner Cost (Cr)": "{:,.2f}",
            "Tax (Cr)": "{:,.2f}",
            "Net Landowner Cashflow (Cr)": "{:,.2f}",
        }
    ),
    use_container_width=True,
)

st.subheader("Cashflow Trend")
trend = pd.DataFrame(
    {
        "Period": ["Entry"] + [f"Y{i}" for i in YEARS],
        "Cashflow (Cr)": summary["cashflows"],
    }
).set_index("Period")
st.bar_chart(trend)

st.subheader("Quick Sensitivity (Landowner Share vs Sales Growth)")
growth_cases = [max(sales_growth - 0.02, -0.2), sales_growth, sales_growth + 0.02]
share_cases = [max(landowner_share - 0.02, 0.0), landowner_share, min(landowner_share + 0.02, 1.0)]

sens_rows = []
for g in growth_cases:
    for s in share_cases:
        _, sm = compute_landowner_model(
            resi_area=resi_area,
            comm_area=comm_area,
            start_price_resi=start_price_resi,
            start_price_comm=start_price_comm,
            sales_growth=g,
            sales_mix=sales_mix,
            landowner_share=s,
            deposit=deposit,
            annual_landowner_cost=annual_landowner_cost,
            tax_rate=tax_rate,
            entry_land_value=entry_land_value,
            discount_rate=discount_rate,
        )
        sens_rows.append(
            {
                "Sales Growth": f"{g*100:.1f}%",
                "LO Share": f"{s*100:.1f}%",
                "Net Inflow (Cr)": round(sm["net_landowner_inflow"], 2),
                "NPV (Cr)": round(sm["npv"], 2),
                "IRR": pct_to_str(sm["irr"]),
            }
        )

st.dataframe(pd.DataFrame(sens_rows), use_container_width=True)

st.download_button(
    label="Download Landowner Cashflow CSV",
    data=results_df.to_csv(index=False),
    file_name="landowner_cashflow_output.csv",
    mime="text/csv",
)
