# 🤖 AI-Powered Data Analytics Pipeline

> **Automated end-to-end analytics:** Drop a CSV → Get a cleaned dataset, web dashboard, Word report, and PNG screenshot delivered to your inbox — fully automated, weekly.

![Pipeline Flow](https://img.shields.io/badge/Status-Live-brightgreen) ![Python](https://img.shields.io/badge/Python-3.14-blue) ![Claude](https://img.shields.io/badge/AI-Claude%20Sonnet-orange) ![Make.com](https://img.shields.io/badge/Automation-Make.com-purple)

---

## 📋 Table of Contents

- [Overview](#overview)
- [Live Demo](#live-demo)
- [Pipeline Architecture](#pipeline-architecture)
- [Tech Stack](#tech-stack)
- [How It Works](#how-it-works)
- [Setup Guide](#setup-guide)
- [Project Structure](#project-structure)
- [Sample Output](#sample-output)
- [Key Design Decisions](#key-design-decisions)
- [Limitations & Known Issues](#limitations--known-issues)
- [Future Improvements](#future-improvements)

---

## Overview

This project implements a **fully automated AI data analytics pipeline** that runs on a weekly schedule with zero human intervention after setup.

**The problem it solves:** Small teams and SMEs without dedicated BI infrastructure or data science resources can now get professional-grade analytics reports automatically — just from a CSV file in Google Drive.

**What it does in ~3 minutes:**
1. Picks up your latest CSV from Google Drive
2. Runs an AI agent loop to detect and fix data quality issues
3. Analyses the clean data and extracts KPIs, trends, and anomalies
4. Generates an interactive HTML dashboard with Chart.js bar charts
5. Screenshots the dashboard to a PNG using Screenshotone.com
6. Writes a full business report as a Word document
7. Emails all three outputs (Word report, HTML dashboard, PNG screenshot) to your inbox

**Data source:** [OWID Monkeypox Dataset](https://github.com/owid/monkeypox) — regularly updated global surveillance data (165,106 rows, 15 columns, 151 countries). The pipeline fetches the latest CSV directly from this repository on every run.

---

## Live Demo

### 📄 Sample Report Findings
- **180,545** total confirmed global cases
- **506** deaths — 0.28% case fatality rate
- **Peak:** August 1, 2022 — 5,997 new cases in a single day
- **151 countries** affected over 48 months
- Top 5: USA (37,789) · DRC (34,054) · Brazil (15,027) · Spain (9,454) · Uganda (8,512)

### 🔄 Pipeline Run Time
~3 minutes end-to-end on Render.com free tier

---

## Pipeline Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        MAKE.COM (Orchestrator)                   │
│                                                                  │
│  ⏰ Schedule          📂 Google Drive      ⬇ Google Drive        │
│  (Weekly/Daily)  →   Upload CSV       →   Download CSV           │
│                                                │                 │
│                              ┌─────────────────┘                 │
│                              ▼                                   │
│                    🔗 HTTP POST /run                             │
│                    (to Render.com)                               │
│                              │                                   │
│                    ◀─────────┘ response                          │
│               (docx_b64 + dashboard_html + dashboard_png)        │
│                              │                                   │
│                    📧 Gmail → Send Email                         │
│          (Report.docx + Dashboard.html + Dashboard.png attached) │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼ HTTP POST
┌─────────────────────────────────────────────────────────────────┐
│                      RENDER.COM (Python Server)                 │
│                                                                   │
│  Step 1: Load full CSV (165K rows)                               │
│       ↓                                                           │
│  Step 2: Pre-compute global stats from ALL rows (Python)         │
│          Uses most-recent row for cumulative fields              │
│          (total_cases, total_deaths) — NOT .max()                │
│       ↓                                                           │
│  Step 3: Stratified sample → 1,000 representative rows           │
│       ↓                                                           │
│  Step 4: AI Agent Cleaning Loop (Claude API)                     │
│          Inspect → Decide → Fix → Verify → Repeat                │
│       ↓                                                           │
│  Step 5: AI Analysis (Claude API)                                │
│          Clean CSV + Global Stats → JSON insights                │
│       ↓                                                           │
│  Step 6: Dashboard Generation (Claude API → HTML + Chart.js)     │
│          animation:false, responsive:false, overflow:hidden      │
│       ↓                                                           │
│  Step 7: Screenshot Dashboard → PNG (Screenshotone.com API)      │
│          wait_until:["networkidle0"], delay:4s                   │
│       ↓                                                           │
│  Step 8: Report Writing (Claude API → Word .docx)                │
│       ↓                                                           │
│  Return: base64 files → Make.com                                 │
└─────────────────────────────────────────────────────────────────┘
                         
```

---

## Tech Stack

| Layer | Tool | Purpose |
|-------|------|---------|
| **Orchestration** | Make.com | Schedule, file retrieval, email delivery |
| **Compute** | Render.com (free tier) | Hosts the Python Flask API |
| **AI Engine** | Claude Sonnet (Anthropic API) | Cleaning, analysis, writing, dashboard HTML |
| **Data** | Python + Pandas | CSV loading, sampling, cleaning |
| **Dashboard** | Chart.js (CDN) | Interactive bar chart in HTML dashboard |
| **Screenshot** | Screenshotone.com API | Converts HTML dashboard to PNG |
| **Report** | python-docx | Word document generation |
| **Storage** | Google Drive | CSV input file storage |
| **Data Source** | [OWID Monkeypox GitHub](https://github.com/owid/monkeypox) | Regularly updated surveillance CSV, fetched fresh each run |
| **Delivery** | Gmail (via Make.com) | Email delivery of all 3 outputs |

---

## How It Works

### Step 1 — Data Ingestion
Make.com runs on a schedule (weekly by default). It searches a designated Google Drive folder for the latest CSV file and downloads it as raw text, then sends it to the Render.com API via HTTP POST.

### Step 2 — Pre-compute Global Statistics
Before any sampling, Python computes accurate statistics from the **entire dataset** — total cases, deaths, peak date, top countries, case fatality rate.

**Important:** For cumulative datasets (like OWID), `total_cases` and `total_deaths` are read from the **most recent row** (sorted by date, last non-null value) — not `.max()`, which would incorrectly sum across all dates and produce inflated numbers. Continental and OWID aggregate rows (World, Asia, Europe, etc.) are excluded from country rankings.

### Step 3 — Stratified Sampling
For large datasets, the pipeline samples down to 1,000 rows using **stratified sampling by location** — ensuring all countries/segments are proportionally represented.

### Step 4 — AI Agent Cleaning Loop
An agentic loop runs up to N iterations. Each iteration:
- Claude inspects a statistical summary of the dataset
- Identifies one data quality issue (nulls, duplicates, outliers, format errors, sentinel values, cross-column inconsistencies)
- Generates a Python fix — executed safely
- Claude verifies the fix worked
- Loop continues until no more issues found or max iterations reached
- Low-confidence fixes are flagged for human review instead of auto-applied

### Step 5 — AI Analysis
Claude receives both the pre-computed global statistics AND the sampled clean CSV. It produces structured JSON: KPIs, trend observations, anomalies, top items, and recommendations tailored to the data domain.

### Step 6 — Dashboard Generation
Claude writes a complete self-contained HTML dashboard from the JSON insights. Chart.js renders a bar chart of top countries/items with these critical settings to ensure correct headless rendering:
- `animation: false` — chart paints final state on first frame (no mid-tween blank canvas)
- `responsive: false` + explicit canvas `width`/`height` — no dependency on a layout/resize pass
- `overflow: hidden` wrapper — prevents gridlines bleeding outside the chart area
- `layout.padding` and `drawBorder: false` — clean chart boundaries

### Step 7 — Dashboard Screenshot
Screenshotone.com API converts the HTML dashboard to a PNG:
- `wait_until: ["networkidle0"]` — waits for Chart.js CDN to fully load (must be an **array**, not a string)
- `delay: 4` — additional buffer for rendering
- Full-page capture at 1400×900 viewport

### Step 8 — Report Writing
Claude writes a professional business report with proper section structure. python-docx formats it into a Word document with a KPI table, headings, bullet points, and a cleaning log appendix.

---

## Setup Guide

### Prerequisites
- Python 3.10+
- Anthropic API key (platform.anthropic.com)
- Screenshotone.com API key (screenshotone.com)
- Google account
- Make.com account (free tier)
- Render.com account (free tier)
- GitHub account

### Part 1 — Deploy the Python API

**1. Clone this repository**
```bash
git clone https://github.com/YOUR_USERNAME/ai-analytics-pipeline.git
cd ai-analytics-pipeline
```

**2. Install dependencies locally (optional, for testing)**
```bash
pip install -r requirements.txt
```

**3. Deploy to Render.com**
- Go to render.com → New → Web Service
- Connect your GitHub repo
- Set the following:
  - **Language:** Python 3
  - **Build Command:** `pip install -r requirements.txt`
  - **Start Command:** `gunicorn pipeline:app -c gunicorn.conf.py`
- Add environment variables:
  - `ANTHROPIC_API_KEY` = your Anthropic API key
  - `SCREENSHOT_API_KEY` = your Screenshotone.com API key
- Click **Deploy Web Service**
- Note your service URL: `https://your-service.onrender.com`

### Part 2 — Set Up Make.com Automation

**1. Create a new scenario in Make.com**

**2. Add these modules in order:**

```
Module 1: HTTP — GET CSV
  - URL: https://raw.githubusercontent.com/.../owid-monkeypox-data.csv
  - Method: GET
  - Parse response: No

Module 2: Google Drive — Upload a File
  - Connection: your Google Drive
  - File name: owid-monkeypox-data.csv
  - Data: map from Module 1 response

Module 3: HTTP — Make a Request (POST /run)
  - URL: https://your-service.onrender.com/run
  - Method: POST
  - Body content type: application/json
  - Body: { "csv": "<Module 1 data>" }
  - Parse response: Yes
  - Timeout: 300

Module 4: Gmail — Send an Email
  - To: your email address
  - Subject: Monkeypox Analytics Report — Weekly Update
  - Body type: Collection of contents (text only — do NOT add an Image element)
  - Body: Your automated Monkeypox analytics report is ready.
          Please find the Word report and Dashboard page attached.
  - Attachment 1:
      File name: Report.docx
      Data: toBinary(6.data.docx_b64; base64)
  - Attachment 2:
      File name: Dashboard.html
      Data: 6.data.dashboard_html   ← plain text, no toBinary needed
  - Attachment 3:
      File name: Dashboard.png
      Data: toBinary(6.data.dashboard_png; base64)
```

**3. Set schedule**
- Click the clock icon on Module 1
- Set to Weekly, Monday, 09:00, your timezone

**4. Test**
- Click **Run once**
- Check your inbox in ~3 minutes — you should receive an email with 3 attachments

### Part 3 — Customise for Your Domain

Open `pipeline.py` and edit these two sections:

```python
# Line 1: Your company/project name
COMPANY_NAME = "Your Company Name"

# Line 2: Tell Claude about your data
ANALYSIS_SYSTEM_PROMPT = """
You are a senior data analyst for [YOUR INDUSTRY].
The dataset contains: [list your columns].
Key KPIs to track: [list your metrics].
...
"""
```

---

## Project Structure

```
ai-analytics-pipeline/
│
├── pipeline.py          # Main Flask API — all pipeline steps
├── gunicorn.conf.py     # Gunicorn server config (optimised for free tier)
├── requirements.txt     # Python dependencies
├── README.md            # This file
│
└── sample_output/
    ├── Dashboard.html        # Sample HTML dashboard
    ├── Dashboard.png         # Sample dashboard screenshot (PNG)
    └── Report.docx           # Sample Word report
```

---

## Sample Output

### Word Report Structure
```
Global Monkeypox Tracker — Analytics Report
Period: May 2022 - April 2026 | Rows analysed: 165,106

KPI Snapshot (table)
├── Total Global Cases:     180,545
├── Total Global Deaths:    506
├── Case Fatality Rate:     0.28%
├── Peak Daily Cases:       5,997
└── Countries Affected:     151

# Executive Summary
# Key Performance Indicators
# Trend Analysis
# Anomalies & Risk Signals
# Public Health Recommendations
# Appendix: Data Cleaning Log
```

### HTML Dashboard
- 7 KPI cards (cases, deaths, CFR, peak daily cases, locations, avg daily cases/deaths)
- Bar chart: Top 8 countries by total cases (countries only — World/continent aggregates excluded)
- Key Trends section
- Public Health Recommendations section

### Dashboard PNG
- Full-page screenshot of the HTML dashboard
- Chart fully rendered (Chart.js animations disabled for headless capture)
- Delivered as email attachment alongside the HTML and Word report

---

## Key Design Decisions

### Why use the last row (not `.max()`) for cumulative KPIs?
OWID data is cumulative — each row for "World" has a running total. Using `.max()` picks the highest value across all dates, which for COVID-era data produces wildly inflated numbers (471M instead of the correct ~180K). The fix: sort by date, take the last non-null row.

### Why is `wait_until` an array in the Screenshotone API?
Screenshotone requires `wait_until` to be a JSON array (e.g. `["networkidle0"]`), not a plain string. Passing a string causes a `400 Bad Request: "wait_until" must be an array` — the request fails silently and `dashboard_png` returns empty.

### Why stratified sampling instead of random?
Random sampling of 1,000 rows from 165,000 risks overrepresenting one time period or country. Stratified sampling by location ensures all countries appear proportionally — critical for geographic analysis.

### Why pre-compute global stats before sampling?
Claude's context window can't hold 165,000 rows. But KPIs like "total global cases" require the full dataset. Solution: compute all aggregate statistics in Python first, pass them to Claude as context alongside the sample.

### Why an agent loop instead of one-pass cleaning?
A single cleaning pass applies fixed rules. An agent loop can reason about meaning — it knows that `999999` in a Revenue column is likely a sentinel value, that Start Date > End Date is a logical error. Each iteration it identifies one issue, fixes it, and verifies before moving on.

### Why Render.com instead of a serverless function?
The pipeline takes 2–5 minutes to run — too long for Lambda/Cloud Functions (typically 15–30s timeout). Render's persistent server handles long-running requests reliably on the free tier.

---

## Limitations & Known Issues

| Issue | Impact | Status |
|-------|--------|--------|
| Free tier spins down after inactivity | ~50s cold start delay | Known — acceptable for weekly runs |
| 512MB RAM limit | Large datasets (>100K rows) must be sampled | Mitigated by stratified sampling |
| Claude API cost | ~$0.10–0.30 per run | Low — ~$5 budget covers months of weekly runs |
| Make.com free tier | 1,000 operations/month | Sufficient for weekly schedule |
| Screenshotone.com free tier | 100 screenshots/month | Sufficient for weekly schedule |
| base64 response size | Large files may hit Make.com limits | Mitigated by file size limits in pipeline |
| Gmail body Image element | Leaving blank triggers BundleValidationError | Use Text-only body content |

---

## Debugging Tips

### `dashboard_png` is empty in Make.com output
Check your Render logs for a `[screenshot]` line:
- `HTTPError 400: "wait_until" must be an array` → you're passing a string, not `["networkidle0"]`
- `HTTPError 401/403` → check `SCREENSHOT_API_KEY` environment variable on Render
- `HTTPError 429` → Screenshotone free tier limit (100/month) reached
- No `[screenshot]` line at all → new code hasn't deployed yet; redeploy on Render

### KPI numbers are inflated (e.g. 471M cases instead of 180K)
Your dataset is cumulative. The pipeline uses `.iloc[-1]` on the date-sorted World rows — if this is broken, check that the `date` column parses correctly as a datetime and that `world_sorted` is not empty.

### Bar chart shows "World" dwarfing all country bars
The aggregate exclusion list may not match your dataset's location names. Check `df_full['location'].unique()` for the exact strings and add them to `agg_labels` in `compute_global_stats()`.

### Gmail `BundleValidationError: Missing value of required parameter 'data'`
One of your Gmail module's file fields has an empty Data mapping. Most common cause: an Image element added to the Body contents with no data mapped. Remove it — use Text-only body content.

---

## Future Improvements

- [ ] **Add PowerPoint output** — re-add slide generation step removed in current version
- [ ] **Upload outputs directly to Google Drive** — avoid Make.com response size limits
- [ ] **Multi-sheet Excel support** — currently CSV only
- [ ] **Predictive analytics** — add forecasting step using Claude + statsmodels
- [ ] **Custom report templates** — allow users to upload a .docx template
- [ ] **Slack/Teams notification** — send summary message alongside email
- [ ] **Upgrade to Render paid tier** — eliminate cold start, increase RAM to 2GB
- [ ] **Dashboard auto-deploy** — push HTML to Vercel/Netlify automatically

---

## Cost Breakdown

| Service | Cost |
|---------|------|
| Render.com | Free tier ($0/month) |
| Make.com | Free tier ($0/month) |
| Anthropic API | ~$0.10–0.30 per pipeline run |
| Screenshotone.com | Free tier ($0/month, 100 screenshots) |
| Google Drive | Free |
| Gmail | Free |
| **Total monthly (weekly runs)** | **~$0.40–1.20/month** |

---

## Author

Built as part of exploring **Agentic AI for Decision Making** 

---

## License

MIT License — free to use, modify, and distribute.
