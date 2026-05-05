"""
AWS Feed Dashboard
------------------
Streamlit dashboard to visualize the AWS Feed Processor output.
Reads from aws_feed_master.xlsx and displays:
- Summary metrics
- Category and Service charts
- Timeline of announcements
- Cost summary
- Full searchable data table

Run with: streamlit run dashboard.py
"""

import pandas as pd
import streamlit as st
import plotly.express as px
from email.utils import parsedate_to_datetime

# ─────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────
MASTER_FILE = "aws_feed_master.xlsx"

REQUIRED_COLUMNS = {"Title", "Link", "Published", "LLM_Output", "Category", "AWS_Service"}

# ─────────────────────────────────────────────
# Page Setup
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="AWS What's New Dashboard",
    page_icon="☁️",
    layout="wide"
)

st.title("☁️ AWS What's New — Feed Dashboard")
st.caption("Powered by Amazon Bedrock Nova Lite | Data from aws.amazon.com/new/feed/")

# ─────────────────────────────────────────────
# Data Loading
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)  # Re-read file every 5 minutes to pick up fresh runs
def load_data(filepath: str):
    """Load AWS Feed and Cost Summary sheets from master Excel file."""
    feed_df = pd.read_excel(filepath, sheet_name="AWS Feed")
    cost_df = pd.read_excel(filepath, sheet_name="Cost Summary")  # header=0 by default

    # Validate required columns exist
    missing = REQUIRED_COLUMNS - set(feed_df.columns)
    if missing:
        raise ValueError(f"AWS Feed sheet is missing columns: {missing}")

    # Parse RFC 2822 published date (e.g. 'Fri, 01 May 2026 22:00:00 GMT')
    def safe_parse(val):
        try:
            return parsedate_to_datetime(val).strftime("%d %b %Y")
        except Exception:
            return val

    feed_df["Published"] = feed_df["Published"].apply(safe_parse)
    feed_df["Date"] = pd.to_datetime(feed_df["Published"], format="%d %b %Y", errors="coerce").dt.date

    return feed_df, cost_df


# ─────────────────────────────────────────────
# Load Data
# ─────────────────────────────────────────────
try:
    feed_df, cost_df = load_data(MASTER_FILE)
except FileNotFoundError:
    st.error(f"Master file `{MASTER_FILE}` not found. Please run `aws_feed_processor.py` first.")
    st.stop()
except ValueError as e:
    st.error(f"Data error: {e}")
    st.stop()

# ─────────────────────────────────────────────
# Sidebar Filters
# ─────────────────────────────────────────────
st.sidebar.header("Filters")

if st.sidebar.button("🔄 Refresh Data"):
    st.cache_data.clear()
    st.rerun()

all_categories = sorted(feed_df["Category"].dropna().unique().tolist())
all_services   = sorted(feed_df["AWS_Service"].dropna().unique().tolist())

selected_categories = st.sidebar.multiselect("Category", all_categories, default=all_categories)
selected_services   = st.sidebar.multiselect("AWS Service", all_services, default=all_services)
search_text         = st.sidebar.text_input("Search Title", "")

# Apply filters
filtered_df = feed_df[
    (feed_df["Category"].isin(selected_categories)) &
    (feed_df["AWS_Service"].isin(selected_services))
]
if search_text:
    filtered_df = filtered_df[filtered_df["Title"].str.contains(search_text, case=False, na=False)]

# ─────────────────────────────────────────────
# Section 1 — Summary Metrics
# ─────────────────────────────────────────────
st.subheader("Summary")
col1, col2, col3, col4 = st.columns(4)

# col1/col3/col4 intentionally use feed_df (total counts, unaffected by filters)
col1.metric("Total Announcements", len(feed_df))
col2.metric("Filtered Entries",    len(filtered_df))
col3.metric("Categories",          feed_df["Category"].nunique())
col4.metric("AWS Services",        feed_df["AWS_Service"].nunique())

st.divider()

# ─────────────────────────────────────────────
# Section 2 — Charts
# ─────────────────────────────────────────────
st.subheader("Analytics")

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    category_counts = filtered_df["Category"].value_counts().reset_index()
    category_counts.columns = ["Category", "Count"]
    fig1 = px.bar(
        category_counts, x="Category", y="Count",
        title="Announcements by Category",
        color="Category", text="Count"
    )
    fig1.update_layout(showlegend=False, xaxis_tickangle=-30)
    st.plotly_chart(fig1, width='stretch')

with chart_col2:
    service_counts = filtered_df["AWS_Service"].value_counts().head(15).reset_index()
    service_counts.columns = ["AWS_Service", "Count"]
    fig2 = px.bar(
        service_counts, x="Count", y="AWS_Service",
        title="Top 15 AWS Services",
        orientation="h", color="Count",
        color_continuous_scale="Blues", text="Count"
    )
    fig2.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
    st.plotly_chart(fig2, width='stretch')

# Timeline chart
timeline_df = filtered_df.groupby("Date").size().reset_index(name="Count")
fig3 = px.line(
    timeline_df, x="Date", y="Count",
    title="Announcements Over Time",
    markers=True
)
fig3.update_traces(line_color="#4472C4")
st.plotly_chart(fig3, width='stretch')

st.divider()

# ─────────────────────────────────────────────
# Section 3 — Cost Summary
# ─────────────────────────────────────────────
st.subheader("Cost Summary (Latest Run)")

cost_col1, cost_col2 = st.columns([1, 2])

with cost_col1:
    for _, row in cost_df.iterrows():
        st.markdown(f"**{row['Metric']}:** {row['Value']}")

with cost_col2:
    cost_values = cost_df[cost_df["Metric"].isin(["Input Cost (USD)", "Output Cost (USD)"])]
    if not cost_values.empty:
        fig4 = px.pie(
            cost_values, names="Metric", values="Value",
            title="Cost Breakdown",
            color_discrete_sequence=["#4472C4", "#70AD47"]
        )
        st.plotly_chart(fig4, width='stretch')

st.divider()

# ─────────────────────────────────────────────
# Section 4 — Data Table
# ─────────────────────────────────────────────
st.subheader(f"Announcements ({len(filtered_df)} entries)")

st.dataframe(
    filtered_df[["Title", "Published", "AWS_Service", "Category", "LLM_Output", "Link"]],
    width='stretch',
    hide_index=True,
    column_config={
        "Link":       st.column_config.LinkColumn("Link"),
        "Published":  st.column_config.TextColumn("Published"),
        "LLM_Output": st.column_config.TextColumn("Summary", width="large"),
    }
)

# ─────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────
st.caption("Refresh the page or click 'Refresh Data' in the sidebar after running aws_feed_processor.py")
