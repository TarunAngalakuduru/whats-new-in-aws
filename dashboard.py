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
from openpyxl import load_workbook

MASTER_FILE = "aws_feed_master.xlsx"

st.set_page_config(
    page_title="AWS What's New Dashboard",
    page_icon="☁️",
    layout="wide"
)

st.title("☁️ AWS What's New — Feed Dashboard")
st.caption("Powered by Amazon Bedrock Nova Lite | Data from aws.amazon.com/new/feed/")

@st.cache_data
def load_data(filepath: str):
    wb       = load_workbook(filepath)
    feed_df  = pd.read_excel(filepath, sheet_name="AWS Feed")
    cost_df  = pd.read_excel(filepath, sheet_name="Cost Summary", header=None, names=["Metric", "Value"])

    from email.utils import parsedate_to_datetime
    def safe_parse(val):
        try:
            return parsedate_to_datetime(val).strftime("%d %b %Y")
        except Exception:
            return val

    feed_df["Published"] = feed_df["Published"].apply(safe_parse)
    feed_df["Date"]      = pd.to_datetime(feed_df["Published"], format="%d %b %Y", errors="coerce").dt.date

    return feed_df, cost_df


try:
    feed_df, cost_df = load_data(MASTER_FILE)
except FileNotFoundError:
    st.error(f"Master file `{MASTER_FILE}` not found.")
    st.stop()

# Sidebar
st.sidebar.header("Filters")

all_categories = sorted(feed_df["Category"].dropna().unique().tolist())
all_services   = sorted(feed_df["AWS_Service"].dropna().unique().tolist())

selected_categories = st.sidebar.multiselect("Category", all_categories, default=all_categories)
selected_services   = st.sidebar.multiselect("AWS Service", all_services, default=all_services)
search_text         = st.sidebar.text_input("Search Title", "")

filtered_df = feed_df[
    (feed_df["Category"].isin(selected_categories)) &
    (feed_df["AWS_Service"].isin(selected_services))
]

if search_text:
    filtered_df = filtered_df[filtered_df["Title"].str.contains(search_text, case=False, na=False)]

# Summary
st.subheader("Summary")
col1, col2, col3, col4 = st.columns(4)

col1.metric("Total Announcements", len(feed_df))
col2.metric("Filtered Entries",    len(filtered_df))
col3.metric("Categories",          feed_df["Category"].nunique())
col4.metric("AWS Services",        feed_df["AWS_Service"].nunique())

st.divider()

# Charts
st.subheader("Analytics")

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    category_counts = filtered_df["Category"].value_counts().reset_index()
    category_counts.columns = ["Category", "Count"]

    fig1 = px.bar(category_counts, x="Category", y="Count", color="Category", text="Count")
    fig1.update_layout(showlegend=False, xaxis_tickangle=-30)

    st.plotly_chart(fig1, use_container_width=True)

with chart_col2:
    service_counts = filtered_df["AWS_Service"].value_counts().head(15).reset_index()
    service_counts.columns = ["AWS_Service", "Count"]

    fig2 = px.bar(service_counts, x="Count", y="AWS_Service", orientation="h", color="Count", text="Count")
    fig2.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)

    st.plotly_chart(fig2, use_container_width=True)

timeline_df = filtered_df.groupby("Date").size().reset_index(name="Count")

fig3 = px.line(timeline_df, x="Date", y="Count", markers=True)
fig3.update_traces(line_color="#4472C4")

st.plotly_chart(fig3, use_container_width=True)

st.divider()

# Cost
st.subheader("Cost Summary")

cost_col1, cost_col2 = st.columns([1, 2])

with cost_col1:
    for _, row in cost_df.iterrows():
        st.markdown(f"**{row['Metric']}:** {row['Value']}")

with cost_col2:
    cost_values = cost_df[cost_df["Metric"].isin(["Input Cost (USD)", "Output Cost (USD)"])]

    if not cost_values.empty:
        fig4 = px.pie(cost_values, names="Metric", values="Value")
        st.plotly_chart(fig4, use_container_width=True)

st.divider()

# Table
st.subheader(f"Announcements ({len(filtered_df)} entries)")

filtered_df = filtered_df.sort_values("Date", ascending=False)

st.dataframe(
    filtered_df[["Title", "Published", "AWS_Service", "Category", "LLM_Output", "Link"]],
    use_container_width=True,
    hide_index=True,
    column_config={
        "Link":      st.column_config.LinkColumn("Link"),
        "Published": st.column_config.TextColumn("Published"),
        "LLM_Output": st.column_config.TextColumn("Summary", width="large"),
    }
)

st.caption("Refresh page to load latest data")
