# AWS What's New — Feed Dashboard

**Public Dashboard:** https://whats-new-in-aws.streamlit.app/

Automated pipeline that fetches AWS announcements, summarizes them using Amazon Bedrock Nova Lite, and displays them in a Streamlit dashboard.

## What it does

* Fetches latest AWS announcements from the RSS feed daily via GitHub Actions
* Processes only new entries (deduplication against master Excel file)
* Uses Amazon Bedrock Nova Lite to:

  * Generate a short summary of each announcement
  * Identify the outcome (impact)
  * Suggest where the feature can be used (use case)
  * Extract AWS Service and Category
* Tracks token usage and cost per run
* Displays everything in an interactive Streamlit dashboard


## Files

| File                                | Description                                             |
| ----------------------------------- | ------------------------------------------------------- |
| `aws_feed_processor.py`             | Main processor — fetches, processes, and saves to Excel |
| `dashboard.py`                      | Streamlit dashboard to visualize the data               |
| `requirements.txt`                  | Python dependencies                                     |
| `.github/workflows/update_feed.yml` | GitHub Actions workflow for daily automation            |

## Setup

### 1. Clone the repo

```bash
git clone https://github.com/TarunAngalakuduru/whats-new-in-aws.git
cd whats-new-in-aws
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure AWS credentials

Add these to GitHub Secrets (Settings → Secrets → Actions):

* `AWS_ACCESS_KEY_ID`
* `AWS_SECRET_ACCESS_KEY`
* `AWS_REGION`

### 4. Run locally

```bash
python aws_feed_processor.py
streamlit run dashboard.py
```

## Automation

GitHub Actions runs the processor daily between **12:00 PM and 2:00 PM IST (UTC+5:30)**, updates `aws_feed_master.xlsx`, and commits it back to the repo. Streamlit Cloud picks up the change automatically.

## Tech Stack

* Amazon Bedrock Nova Lite (ap-south-1)
* Python, feedparser, openpyxl
* Streamlit, Plotly
* GitHub Actions
