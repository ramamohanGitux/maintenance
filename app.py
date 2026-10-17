from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

WORKBOOK_PATH = Path(__file__).with_name("Monthly_Maintenance.xlsx")
REQUIRED_COLUMNS = [
    "Apartment No.",
    "Monthly Contribution (INR)",
    "Payment Date",
    "UPI UTR Number (Sample)",
    "Amount Paid (INR)",
]

st.set_page_config(
    page_title="Maintenance Ledger",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
    :root { --ink: #173a32; --muted: #68766f; --green: #26745f; --mint: #e5f1eb; --gold: #d99a2b; --coral: #c85f48; --paper: #f7f8f4; }
    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
    .stApp { background: var(--paper); color: var(--ink); }
    h1, h2, h3 { font-family: 'Manrope', sans-serif; letter-spacing: 0; color: var(--ink); }
    h1 { font-size: 2rem !important; font-weight: 800 !important; }
    [data-testid="stSidebar"] { background: #eef3ee; border-right: 1px solid #dce5dd; }
    [data-testid="stMetric"] { background: #fff; border: 1px solid #dce5dd; border-radius: 6px; padding: 16px 18px; }
    [data-testid="stMetricLabel"] { color: var(--muted); font-weight: 600; }
    [data-testid="stMetricValue"] { color: var(--ink); font-family: 'Manrope', sans-serif; }
    [data-testid="stDataFrame"] { border: 1px solid #dce5dd; border-radius: 6px; }
    div[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 6px; }
    .eyebrow { color: var(--green); font-size: 0.75rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; }
    .subtle { color: var(--muted); font-size: 0.9rem; }
    @media (max-width: 700px) { h1 { font-size: 1.55rem !important; } [data-testid="stMetric"] { padding: 12px; } }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data
def load_ledger(path: str, modified_at: float) -> tuple[pd.DataFrame, str]:
    from openpyxl import load_workbook

    header_probe = pd.read_excel(path, sheet_name="Monthly Maintenance", header=None, nrows=5)
    header_row = next(
        (
            row_number
            for row_number, row in header_probe.iterrows()
            if set(REQUIRED_COLUMNS).issubset({str(value).strip() for value in row.dropna()})
        ),
        None,
    )
    if header_row is None:
        raise ValueError("Could not find all required maintenance columns in the first five rows.")
    frame = pd.read_excel(
        path,
        sheet_name="Monthly Maintenance",
        header=header_row,
        converters={
            "UPI UTR Number (Sample)": lambda value: "" if pd.isna(value) else str(value).removesuffix(".0")
        },
    )

    frame = frame[REQUIRED_COLUMNS].copy()
    frame["Apartment No."] = pd.to_numeric(frame["Apartment No."], errors="coerce")
    frame = frame.dropna(subset=["Apartment No."]).copy()
    frame["Apartment No."] = frame["Apartment No."].astype("Int64").astype(str)
    for column in ("Monthly Contribution (INR)", "Amount Paid (INR)"):
        frame[column] = pd.to_numeric(frame[column], errors="coerce").fillna(0)
    frame["Payment Date"] = pd.to_datetime(frame["Payment Date"], errors="coerce")
    frame["UPI UTR Number (Sample)"] = frame["UPI UTR Number (Sample)"].fillna("").astype(str)

    paid = frame["Amount Paid (INR)"]
    due = frame["Monthly Contribution (INR)"]
    frame["Status"] = "Unpaid"
    frame.loc[(paid > 0) & (paid < due), "Status"] = "Partial"
    frame.loc[paid >= due, "Status"] = "Paid"
    frame["Outstanding (INR)"] = (due - paid).clip(lower=0)

    workbook = load_workbook(path, read_only=True, data_only=True)
    month = workbook["Monthly Maintenance"]["B2"].value if header_row >= 2 else None
    workbook.close()
    return frame, str(month).strip() if month else "Month not set"


if not WORKBOOK_PATH.exists():
    st.error("Monthly_Maintenance.xlsx was not found beside app.py.")
    st.stop()

try:
    ledger, month_label = load_ledger(str(WORKBOOK_PATH), WORKBOOK_PATH.stat().st_mtime)
except Exception as error:
    st.error(f"Could not read the maintenance workbook: {error}")
    st.stop()

st.sidebar.markdown("### Ledger filters")
statuses = st.sidebar.multiselect(
    "Payment status",
    options=["Paid", "Partial", "Unpaid"],
    default=["Paid", "Partial", "Unpaid"],
)
apartment_search = st.sidebar.text_input("Find apartment", placeholder="For example, 301")
st.sidebar.caption("UTR values in this workbook are randomly generated samples.")

expected = float(ledger["Monthly Contribution (INR)"].sum())
collected = float(ledger["Amount Paid (INR)"].sum())
outstanding = max(expected - collected, 0)
collection_rate = collected / expected if expected else 0

st.markdown('<p class="eyebrow">RESIDENTIAL OPERATIONS</p>', unsafe_allow_html=True)
st.title("Maintenance ledger")
st.caption(f"Monthly collections · {month_label} · {len(ledger)} apartments")

metric_columns = st.columns(4)
metric_columns[0].metric("Expected", f"₹{expected:,.0f}")
metric_columns[1].metric("Collected", f"₹{collected:,.0f}", f"{collection_rate:.0%} of expected")
metric_columns[2].metric("Outstanding", f"₹{outstanding:,.0f}")
metric_columns[3].metric("Paid in full", f"{int((ledger['Status'] == 'Paid').sum())} / {len(ledger)}")

chart_left, chart_right = st.columns([0.8, 1.6], gap="large")
with chart_left:
    st.subheader("Payment status")
    status_order = ["Paid", "Partial", "Unpaid"]
    status_counts = (
        ledger["Status"].value_counts().reindex(status_order, fill_value=0)
        .rename_axis("Status").reset_index(name="Apartments")
    )
    status_colors = {"Paid": "#26745f", "Partial": "#d99a2b", "Unpaid": "#c85f48"}
    status_chart = px.bar(
        status_counts,
        x="Status",
        y="Apartments",
        color="Status",
        color_discrete_map=status_colors,
        text="Apartments",
        category_orders={"Status": status_order},
    )
    status_chart.update_layout(
        showlegend=False,
        margin=dict(l=8, r=8, t=18, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        yaxis=dict(title=None, rangemode="tozero", gridcolor="#e4e9e3"),
        xaxis=dict(title=None),
        font=dict(family="DM Sans, sans-serif", color="#173a32"),
    )
    st.plotly_chart(status_chart, use_container_width=True, config={"displayModeBar": False})

with chart_right:
    st.subheader("Paid and outstanding by apartment")
    apartment_chart = px.bar(
        ledger.sort_values("Apartment No."),
        x=["Amount Paid (INR)", "Outstanding (INR)"],
        y="Apartment No.",
        orientation="h",
        barmode="stack",
        color_discrete_map={"Amount Paid (INR)": "#26745f", "Outstanding (INR)": "#edc46e"},
        labels={"value": "Amount (INR)", "variable": "", "Apartment No.": "Apartment"},
    )
    apartment_chart.update_layout(
        margin=dict(l=8, r=8, t=18, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title=None, tickprefix="₹", gridcolor="#e4e9e3"),
        yaxis=dict(title=None, autorange="reversed"),
        legend=dict(orientation="h", y=1.12, x=0),
        font=dict(family="DM Sans, sans-serif", color="#173a32"),
    )
    st.plotly_chart(apartment_chart, use_container_width=True, config={"displayModeBar": False})

st.subheader("Apartment details")
filtered = ledger[ledger["Status"].isin(statuses)]
if apartment_search.strip():
    filtered = filtered[filtered["Apartment No."].str.contains(apartment_search.strip(), case=False)]

display_columns = [
    "Apartment No.",
    "Monthly Contribution (INR)",
    "Payment Date",
    "UPI UTR Number (Sample)",
    "Amount Paid (INR)",
    "Status",
]
st.dataframe(
    filtered[display_columns],
    hide_index=True,
    use_container_width=True,
    height=min(620, 38 + max(len(filtered), 1) * 35),
    column_config={
        "Apartment No.": st.column_config.TextColumn("Apartment", width="small"),
        "Monthly Contribution (INR)": st.column_config.NumberColumn("Monthly due", format="₹%.0f"),
        "Payment Date": st.column_config.DateColumn("Payment date", format="DD MMM YYYY"),
        "UPI UTR Number (Sample)": st.column_config.TextColumn("Sample UTR", width="medium"),
        "Amount Paid (INR)": st.column_config.NumberColumn("Paid", format="₹%.0f"),
        "Status": st.column_config.TextColumn("Status", width="small"),
    },
)
st.caption("Dashboard data is read from Monthly_Maintenance.xlsx in the same folder as this app.")
