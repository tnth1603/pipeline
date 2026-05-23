# 🤖 AI-Powered Data Analytics Pipeline

> **Automated end-to-end analytics:** Drop a CSV → Get a cleaned dataset, web dashboard, Word report, and PowerPoint slides delivered to your inbox — fully automated, weekly.

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
4. Generates an interactive HTML dashboard
5. Writes a full business report as a Word document
6. Creates a PowerPoint slide deck
7. Emails all outputs to your inbox

**Demo dataset:** Global Monkeypox surveillance data (165,106 rows, 15 columns, 151 countries)

---

## Live Demo

### 📄 Sample Report Findings
- **179,029** total confirmed global cases
- **503** deaths — 0.281% case fatality rate
- **Peak:** August 1, 2022 — 5,997 new cases in a single day
- **151 countries** affected over 48 months
- Top 5: USA (37,530) · DRC (34,011) · Brazil (15,027) · Spain (9,454) · Uganda (8,512)

### 🔄 Pipeline Run Time
~3 minutes end-to-end on Render.com free tier

---

## Pipeline Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        MAKE.COM (Orchestrator)                   │
│                                                                   │
│  ⏰ Schedule          📂 Google Drive      ⬇ Google Drive        │
│  (Weekly/Daily)  →   Search Folder    →   Download CSV           │
│                                                │                  │
│                              ┌─────────────────┘                  │
│                              ▼                                    │
│                    🔗 HTTP POST /run                              │
│                    (to Render.com)                                │
│                              │                                    │
│                    ◀─────────┘ response                          │
│                    (docx_b64 + pptx_b64)                         │
│                              │                                    │
│                    📧 Gmail → Send Email                          │
│                    (Word + PowerPoint attached)                   │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼ HTTP POST
┌─────────────────────────────────────────────────────────────────┐
│                      RENDER.COM (Python Server)                   │
│                                                                   │
│  Step 1: Load full CSV (165K rows)                               │
│       ↓                                                           │
│  Step 2: Pre-compute global stats from ALL rows (Python)         │
│       ↓                                                           │
│  Step 3: Stratified sample → 1,000 representative rows           │
│       ↓                                                           │
│  Step 4: AI Agent Cleaning Loop (Claude API)                     │
│          Inspect → Decide → Fix → Verify → Repeat                │
│       ↓                                                           │
│  Step 5: AI Analysis (Claude API)                                │
│          Clean CSV + Global Stats → JSON insights                │
│       ↓                                                           │
│  Step 6: Dashboard Generation (Claude API → HTML)                │
│       ↓                                                           │
│  Step 7: Report Writing (Claude API → Word .docx)                │
│       ↓                                                           │
│  Step 8: Slide Generation (Claude API → PowerPoint .pptx)        │
│       ↓                                                           │
│  Return: base64 files → Make.com                                 │
└─────────────────────────────────────────────────────────────────┘
                              │
                    ┌─────────┴─────────┐
                    ▼                   ▼
              CLAUDE API          ANTHROPIC
              (claude-sonnet-     API Credits
               4-20250514)        ~$0.10-0.30/run
```

---

## Tech Stack

| Layer | Tool | Purpose |
|-------|------|---------|
| **Orchestration** | Make.com | Schedule, file retrieval, email delivery |
| **Compute** | Render.com (free tier) | Hosts the Python Flask API |
| **AI Engine** | Claude Sonnet (Anthropic API) | Cleaning, analysis, writing, slides |
| **Data** | Python + Pandas | CSV loading, sampling, cleaning |
| **Report** | python-docx | Word document generation |
| **Slides** | python-pptx | PowerPoint generation |
| **Storage** | Google Drive | CSV input, output file storage |
| **Delivery** | Gmail (via Make.com) | Email delivery |

---

## How It Works

### Step 1 — Data Ingestion
Make.com runs on a schedule (weekly by default). It searches a designated Google Drive folder for the latest CSV file and downloads it as raw text, then sends it to the Render.com API via HTTP POST.

### Step 2 — Pre-compute Global Statistics
Before any sampling, Python computes accurate statistics from the **entire dataset** — total cases, deaths, peak date, top countries, case fatality rate. This ensures KPI accuracy regardless of sampling.

### Step 3 — Stratified Sampling
For large datasets, the pipeline samples down to 1,000 rows using **stratified sampling by location** — ensuring all countries/segments are proportionally represented, not randomly discarded.

### Step 4 — AI Agent Cleaning Loop
An agentic loop runs up to N iterations. Each iteration:
- Claude inspects a statistical summary of the dataset
- Identifies one data quality issue (nulls, duplicates, outliers, format errors, sentinel values, cross-column inconsistencies)
- Generates a Python fix — executed safely in a sandboxed environment
- Claude verifies the fix worked
- Loop continues until no more issues found or max iterations reached
- Low-confidence fixes are flagged for human review instead of auto-applied

### Step 5 — AI Analysis
Claude receives both the pre-computed global statistics AND the sampled clean CSV. It produces structured JSON: KPIs, trend observations, anomalies, top items, and recommendations tailored to the data domain.

### Step 6 — Dashboard Generation
Claude writes a complete self-contained HTML dashboard (Chart.js) from the JSON insights — KPI cards, interactive bar charts, trends and recommendations sections.

### Step 7 — Report Writing
Claude writes a professional business report with proper section structure. python-docx formats it into a Word document with a KPI table, headings, bullet points, and a cleaning log appendix.

### Step 8 — Slide Generation
Claude plans a 7-slide deck as structured JSON. python-pptx builds the .pptx file with dark theme, headline stats, bullet points, and speaker notes per slide.

---

## Setup Guide

### Prerequisites
- Python 3.10+
- Anthropic API key (platform.anthropic.com)
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
- Add environment variable:
  - `ANTHROPIC_API_KEY` = your Anthropic API key
- Click **Deploy Web Service**
- Note your service URL: `https://your-service.onrender.com`

### Part 2 — Set Up Make.com Automation

**1. Create a new scenario in Make.com**

**2. Add these modules in order:**

```
Module 1: Google Drive — Search for Files/Folders
  - Connection: your Google Drive
  - Drive: My Drive
  - Search Scope: Limit to chosen folder
  - Folder: your CSV folder (e.g. /MakeData)
  - Retrieve: Files
  - Limit: 1

Module 2: Google Drive — Download a File
  - Connection: same Google Drive
  - File ID: map from Module 1 → File ID

Module 3: HTTP — Make a Request
  - URL: https://your-service.onrender.com/run
  - Method: POST
  - Body content type: application/json
  - Body input method: Data structure
  - Body structure: create with field "csv" (Text type)
  - Map csv field → Module 2 Data
  - Parse response: Yes
  - Timeout: 300

Module 4: Gmail — Send an Email
  - To: your email address
  - Subject: Analytics Report — Weekly Update
  - Body: Your automated report is ready.
  - Attachment 1:
      File name: Report.docx
      Data: {{toBinary(6.data.docx_b64; "base64")}}
  - Attachment 2:
      File name: Slides.pptx
      Data: {{toBinary(6.data.pptx_b64; "base64")}}
```

**3. Set schedule**
- Click the clock icon on Module 1
- Set to Weekly, Monday, 09:00, your timezone

**4. Test**
- Drop your CSV into the Google Drive folder
- Click **Run once**
- Check your inbox in ~3 minutes

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
├── pipeline.py          # Main Flask API — all 8 pipeline steps
├── gunicorn.conf.py     # Gunicorn server config (optimised for free tier)
├── requirements.txt     # Python dependencies
├── README.md            # This file
│
└── sample_output/
    ├── Report_sample.docx    # Sample Word report
    └── Slides_sample.pptx    # Sample PowerPoint deck
```

---

## Sample Output

### Word Report Structure
```
Global Monkeypox Tracker — Analytics Report
Period: May 2022 - April 2026 | Rows analysed: 165,106

KPI Snapshot (table)
├── Total Global Cases:     179,029
├── Total Global Deaths:    503
├── Case Fatality Rate:     0.281%
├── Peak Daily Cases:       5,997
└── Countries Affected:     151

# Executive Summary
# Key Performance Indicators
# Trend Analysis
# Anomalies & Risk Signals
# Public Health Recommendations
# Appendix: Data Cleaning Log
```

### PowerPoint Deck (7 slides)
1. Title — Global Health Crisis Overview
2. Global KPI Overview
3. Epidemic Trend Analysis
4. Peak Impact & Timeline
5. Top Affected Countries
6. Anomalies & Risk Signals
7. Public Health Recommendations

---

## Key Design Decisions

### Why stratified sampling instead of random?
Random sampling of 1,000 rows from 165,000 risks overrepresenting one time period or country. Stratified sampling by location ensures all countries appear proportionally — critical for geographic analysis.

### Why pre-compute global stats before sampling?
Claude's context window can't hold 165,000 rows. But KPIs like "total global cases" require the full dataset. Solution: compute all aggregate statistics in Python first, pass them to Claude as context alongside the sample.

### Why an agent loop instead of one-pass cleaning?
A single cleaning pass applies fixed rules (drop nulls, remove duplicates). An agent loop can reason about meaning — it knows that `999999` in a Revenue column is likely a sentinel value, that Start Date > End Date is a logical error, that `A/B/C` in a Status column may need human clarification. Each iteration it identifies one issue, fixes it, and verifies before moving on.

### Why Render.com instead of a serverless function?
The pipeline takes 2-5 minutes to run — too long for Lambda/Cloud Functions (typically 15-30s timeout). Render's persistent server handles long-running requests reliably on the free tier.

---

## Limitations & Known Issues

| Issue | Impact | Status |
|-------|--------|--------|
| Free tier spins down after inactivity | ~50s cold start delay | Known — acceptable for weekly runs |
| 512MB RAM limit | Large datasets (>100K rows) must be sampled | Mitigated by stratified sampling |
| Claude API cost | ~$0.10-0.30 per run | Low — ~$5 budget covers months of weekly runs |
| Make.com free tier | 1,000 operations/month | Sufficient for weekly schedule |
| base64 response size | Large files may hit Make.com limits | Mitigated by file size limits in pipeline |

---

## Future Improvements

- [ ] **Upload outputs directly to Google Drive** from pipeline.py — avoid Make.com response size limits
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
| Anthropic API | ~$0.10-0.30 per pipeline run |
| Google Drive | Free |
| Gmail | Free |
| **Total monthly (weekly runs)** | **~$0.40-1.20/month** |

---

## Author

Built as part of exploring **Agentic AI for Decision Making** — demonstrating that non-technical users can deploy production-grade AI automation with no code beyond a single Python file.

---

## License

MIT License — free to use, modify, and distribute.
