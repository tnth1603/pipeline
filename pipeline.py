# ── IMPORTS ───────────────────────────────────────────────────────────────────
from flask import Flask, request, jsonify
import anthropic, json, os, base64, time
import pandas as pd
import numpy as np
from io import StringIO, BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor as PptxRGB

# ── APP SETUP ─────────────────────────────────────────────────────────────────
app = Flask(__name__)
client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

# ── CONFIG ────────────────────────────────────────────────────────────────────
COMPANY_NAME    = "Global Monkeypox Tracker"
MAX_ROWS        = 1000   # max rows passed to Claude (memory limit)
MAX_ITERATIONS  = 1      # cleaning iterations (save credits)
MAX_RETRIES     = 3      # Claude API retry attempts

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Claude call with retry + exponential backoff
# ─────────────────────────────────────────────────────────────────────────────
def call_claude(messages, max_tokens=1000, system=None):
    """Call Claude API with retry logic. Raises on final failure."""
    for attempt in range(MAX_RETRIES):
        try:
            kwargs = {
                "model": "claude-sonnet-4-20250514",
                "max_tokens": max_tokens,
                "messages": messages
            }
            if system:
                kwargs["system"] = system
            msg = client.messages.create(**kwargs)
            return msg.content[0].text
        except Exception as e:
            if attempt == MAX_RETRIES - 1:
                raise RuntimeError(f"Claude API failed after {MAX_RETRIES} attempts: {str(e)}")
            wait = 2 ** attempt  # 1s, 2s, 4s
            time.sleep(wait)

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Pre-compute real global statistics from the FULL dataset
# This runs in Python (no memory risk) before any sampling
# ─────────────────────────────────────────────────────────────────────────────
def compute_global_stats(df_full):
    """Compute accurate stats from all rows before sampling."""
    try:
        stats = {
            "total_rows": int(len(df_full)),
            "unique_locations": int(df_full['location'].nunique()),
            "date_range": f"{df_full['date'].min()} to {df_full['date'].max()}",
        }

        # Global case/death totals (sum of new_cases across all dates, world-level)
        world = df_full[df_full['location'].str.contains('World|OWID_WRL', na=False)]
        if not world.empty:
            stats["total_global_cases"] = int(world['total_cases'].max())
            stats["total_global_deaths"] = int(world['total_deaths'].max())
            peak_row = world.loc[world['new_cases'].idxmax()]
            stats["peak_date"] = str(peak_row['date'])
            stats["peak_daily_cases"] = int(peak_row['new_cases'])
        else:
            # fallback: sum new_cases per date across all non-aggregate locations
            daily = df_full.groupby('date')['new_cases'].sum()
            stats["total_global_cases"] = int(df_full['new_cases'].sum())
            stats["total_global_deaths"] = int(df_full['new_deaths'].sum())
            stats["peak_date"] = str(daily.idxmax())
            stats["peak_daily_cases"] = int(daily.max())

        # Top 10 countries by total cases (exclude aggregate regions)
        countries = df_full[~df_full['iso_code'].str.startswith('OWID', na=False)]
        top10 = (
            countries.groupby('location')['total_cases']
            .max()
            .nlargest(10)
            .reset_index()
            .rename(columns={'total_cases': 'total_cases'})
            .to_dict(orient='records')
        )
        stats["top_10_countries"] = top10

        # Case fatality rate
        tc = stats.get("total_global_cases", 0)
        td = stats.get("total_global_deaths", 0)
        stats["case_fatality_rate_pct"] = round((td / tc * 100), 4) if tc > 0 else 0.0

        return stats
    except Exception as e:
        return {"error": f"Could not compute global stats: {str(e)}"}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Stratified sample — keeps all locations represented
# ─────────────────────────────────────────────────────────────────────────────
def stratified_sample(df_full, max_rows=MAX_ROWS):
    """Sample proportionally by location so all countries are represented."""
    if len(df_full) <= max_rows:
        return df_full
    try:
        sampled = (
            df_full.groupby('location', group_keys=False)
            .apply(lambda x: x.sample(
                max(1, int(max_rows * len(x) / len(df_full))),
                random_state=42
            ))
        )
        return sampled.head(max_rows)
    except Exception:
        # Fallback to simple random sample
        return df_full.sample(n=max_rows, random_state=42)

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: AGENT CLEANING
# Runs on full dataset summary stats (not raw rows) to avoid memory issues
# ─────────────────────────────────────────────────────────────────────────────
CLEANING_AGENT_PROMPT = """
You are an expert data cleaning agent. You will receive a statistical summary of a dataset.
Identify ONE data quality issue that has not yet been fixed.
Return ONLY a JSON object:

{
  "issue_found": true,
  "issue_type": "null_values | duplicates | outlier | format_inconsistency | placeholder",
  "column": "column name affected, or 'multiple'",
  "description": "Plain English description of the issue",
  "reasoning": "Why this is an error, not valid data",
  "confidence": "high | medium | low",
  "action": {
    "type": "one of: fillna_median | fillna_zero | fillna_mode | drop_duplicates | drop_nulls | replace_value | normalize_case | to_datetime | clip_outliers",
    "column": "column name to apply action to, or 'all'",
    "value": "only needed for replace_value — the value to replace",
    "replacement": "only needed for replace_value — what to replace it with"
  },
  "needs_human_review": false,
  "human_review_reason": ""
}

If no issues remain: { "issue_found": false }

Rules:
- ONE issue per response only
- action.type must be exactly one of the allowed values listed above
- Never suggest dropping more than 5% of rows
- Consider epidemiological context: zeroes in new_cases are valid
"""

VERIFY_PROMPT = """
Verify this data fix was applied correctly.
Return ONLY JSON: { "verified": true, "note": "brief note" }
"""

def summarise_df(df, fixed_so_far):
    """Generate a concise profile for the cleaning agent to inspect."""
    summary = {
        "shape": {"rows": int(df.shape[0]), "cols": int(df.shape[1])},
        "columns": {},
        "sample_rows": df.head(3).to_dict(orient='records'),
        "issues_already_fixed": [x.get('description', '') for x in fixed_so_far]
    }
    for col in df.columns:
        col_info = {
            "dtype": str(df[col].dtype),
            "nulls": int(df[col].isnull().sum()),
            "unique": int(df[col].nunique()),
            "sample": [str(x) for x in df[col].dropna().unique()[:5]]
        }
        if pd.api.types.is_numeric_dtype(df[col]) and not df[col].isnull().all():
            col_info["min"]  = float(df[col].min())
            col_info["max"]  = float(df[col].max())
            col_info["mean"] = round(float(df[col].mean()), 4)
        summary["columns"][col] = col_info
    return json.dumps(summary, indent=2, default=str)[:4000]

# Pre-approved safe action map — Claude picks from this list only
# No exec(), no arbitrary code, no sandbox escaping possible

SAFE_ACTIONS = {
    "fillna_median": lambda df, col, **_: df.assign(
        **{col: df[col].fillna(df[col].median())}
    ) if col in df.columns and pd.api.types.is_numeric_dtype(df[col]) else df,

    "fillna_zero": lambda df, col, **_: df.assign(
        **{col: df[col].fillna(0)}
    ) if col in df.columns else df,

    "fillna_mode": lambda df, col, **_: df.assign(
        **{col: df[col].fillna(df[col].mode()[0])}
    ) if col in df.columns and not df[col].mode().empty else df,

    "drop_duplicates": lambda df, **_: df.drop_duplicates(),

    "drop_nulls": lambda df, col, **_: df.dropna(
        subset=[col]
    ) if col in df.columns else df,

    "replace_value": lambda df, col, value, replacement, **_: df.assign(
        **{col: df[col].replace(value, replacement)}
    ) if col in df.columns else df,

    "normalize_case": lambda df, col, **_: df.assign(
        **{col: df[col].str.strip().str.title()}
    ) if col in df.columns and df[col].dtype == object else df,

    "to_datetime": lambda df, col, **_: df.assign(
        **{col: pd.to_datetime(df[col], errors='coerce')}
    ) if col in df.columns else df,

    "clip_outliers": lambda df, col, **_: df.assign(
        **{col: df[col].clip(
            lower=df[col].quantile(0.01),
            upper=df[col].quantile(0.99)
        )}
    ) if col in df.columns and pd.api.types.is_numeric_dtype(df[col]) else df,
}

def safe_exec_fix(df, action):
    """
    Execute a pre-approved cleaning action from the SAFE_ACTIONS map.
    Claude returns a structured action dict — we look it up and run it.
    No exec(), no eval(), no arbitrary code execution.
    """
    action_type = action.get("type", "")
    col         = action.get("column", "")
    value       = action.get("value", None)
    replacement = action.get("replacement", None)

    if action_type not in SAFE_ACTIONS:
        raise ValueError(f"Unknown action type: '{action_type}'. Must be one of: {list(SAFE_ACTIONS.keys())}")

    return SAFE_ACTIONS[action_type](
        df,
        col=col,
        value=value,
        replacement=replacement
    )

def agent_clean_csv(df_sample, max_iterations=MAX_ITERATIONS):
    """
    Run AI agent cleaning loop on the sampled DataFrame.
    Returns cleaned DataFrame + logs.
    """
    df = df_sample.copy()
    rows_before  = len(df)
    cleaning_log = []
    flags        = []

    for iteration in range(1, max_iterations + 1):
        try:
            summary  = summarise_df(df, cleaning_log)
            raw      = call_claude(
                system=CLEANING_AGENT_PROMPT,
                messages=[{"role": "user", "content": f"Inspect this dataset summary:\n{summary}"}],
                max_tokens=500
            )
            decision = json.loads(raw.replace("```json","").replace("```","").strip())
        except Exception as e:
            cleaning_log.append({"iteration": iteration, "action": f"error: {str(e)}"})
            break

        if not decision.get("issue_found", False):
            break

        action       = decision.get("action", {})
        description  = decision.get("description", "")
        confidence   = decision.get("confidence", "low")
        needs_review = decision.get("needs_human_review", False)

        if needs_review or confidence == "low" or not action:
            entry = {
                "iteration":   iteration,
                "description": description,
                "action":      "flagged for human review"
            }
            flags.append(entry)
            cleaning_log.append(entry)
            continue

        try:
            rows_before_fix = len(df)
            df = safe_exec_fix(df, action)
            cleaning_log.append({
                "iteration":    iteration,
                "issue_type":   decision.get("issue_type"),
                "description":  description,
                "action_taken": action,
                "rows_before":  rows_before_fix,
                "rows_after":   len(df),
                "action":       "applied"
            })
        except Exception as e:
            cleaning_log.append({
                "iteration":   iteration,
                "description": description,
                "action":      f"fix_failed: {str(e)}"
            })

    return {
        "clean_df":        df,
        "cleaning_log":    cleaning_log,
        "flags_for_human": flags,
        "rows_before":     rows_before,
        "rows_after":      len(df)
    }

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: AI ANALYSIS
# Passes both sampled CSV AND pre-computed global stats to Claude
# ─────────────────────────────────────────────────────────────────────────────
ANALYSIS_SYSTEM_PROMPT = """
You are a senior epidemiological data analyst working with global Monkeypox data.
You will receive:
1. Pre-computed accurate global statistics (from the FULL dataset)
2. A sample of raw rows for pattern analysis

Use the pre-computed stats for all KPI values — they are accurate.
Use the raw rows only to identify trends and patterns.

Return ONLY a valid JSON object:
{
  "period": "date range e.g. May 2022 - Dec 2023",
  "row_count": 0,
  "kpis": {
    "total_global_cases": 0,
    "total_global_deaths": 0,
    "case_fatality_rate_pct": 0.0,
    "peak_daily_new_cases": 0,
    "peak_date": "YYYY-MM-DD",
    "countries_affected": 0
  },
  "trends": ["trend 1", "trend 2", "trend 3"],
  "anomalies": ["anomaly 1", "anomaly 2"],
  "top_items": [{"name": "country", "value": 0, "label": "total_cases"}],
  "recommendations": ["recommendation 1", "recommendation 2", "recommendation 3"]
}
No markdown. No explanation. JSON only.
"""

def analyse_with_claude(clean_csv, global_stats):
    """Analyse data using both global stats and sampled rows."""
    prompt = f"""
ACCURATE GLOBAL STATISTICS (computed from full dataset of {global_stats.get('total_rows','?')} rows):
{json.dumps(global_stats, indent=2)}

SAMPLE ROWS FOR PATTERN ANALYSIS:
{clean_csv[:5000]}

Use the global statistics for all KPI values.
Analyse patterns and trends from the sample rows.
"""
    raw = call_claude(
        system=ANALYSIS_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=1000
    )
    return json.loads(raw.replace("```json","").replace("```","").strip())

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4: DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────
def generate_dashboard(insights, company=COMPANY_NAME):
    prompt = f"""
Create a complete self-contained HTML dashboard for {company}.
Use Chart.js from https://cdn.jsdelivr.net/npm/chart.js
Include:
- Header: {company} | Period: {insights.get('period','')}
- KPI cards row: {json.dumps(insights.get('kpis',{}))}
- Horizontal bar chart of top affected countries: {json.dumps(insights.get('top_items',[])[:8])}
- Trends section (bullet list)
- Recommendations section (numbered list)
Dark professional theme (#0f172a background, white text, teal accents #4ecdc4).
Return ONLY complete HTML. No markdown. No backticks.
"""
    html = call_claude(
        messages=[{"role": "user", "content": prompt}],
        max_tokens=3000
    )
    return html.replace("```html","").replace("```","").strip()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5: WORD REPORT
# ─────────────────────────────────────────────────────────────────────────────
REPORT_PROMPT = """
Write a professional epidemiological analytics report based on these insights.
Use these exact section headings (prefix each with #):
# Executive Summary
# Key Performance Indicators
# Trend Analysis
# Anomalies & Risk Signals
# Public Health Recommendations

Concise, evidence-based, written for senior public health managers.
Insights: {insights}
"""

def write_report(insights):
    return call_claude(
        messages=[{"role": "user", "content": REPORT_PROMPT.format(insights=json.dumps(insights, indent=2))}],
        max_tokens=1500
    )

def make_docx(report_text, insights):
    doc = Document()
    title = doc.add_heading(f"{COMPANY_NAME} — Analytics Report", 0)
    if title.runs:
        title.runs[0].font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)
    doc.add_paragraph(f"Period: {insights.get('period','N/A')} | Total rows analysed: {insights.get('row_count','N/A')}")
    doc.add_paragraph("")

    # KPI table
    doc.add_heading("KPI Snapshot", 1)
    kpis = insights.get("kpis", {})
    if kpis:
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text = "Metric"
        hdr[1].text = "Value"
        for k, v in kpis.items():
            row = table.add_row().cells
            row[0].text = str(k).replace("_"," ").title()
            row[1].text = str(v)
    doc.add_paragraph("")

    # Report body
    for line in report_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        elif line.startswith("# "):
            doc.add_heading(line[2:], 1)
        elif line.startswith("## "):
            doc.add_heading(line[3:], 2)
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(line[2:], style='List Bullet')
        elif line[0].isdigit() and ". " in line[:4]:
            doc.add_paragraph(line, style='List Number')
        else:
            doc.add_paragraph(line)

    # Cleaning log appendix
    doc.add_heading("Appendix: Data Cleaning Log", 1)
    doc.add_paragraph("The following changes were made to the dataset automatically by the AI cleaning agent:")

    buf = BytesIO()
    doc.save(buf)
    return base64.b64encode(buf.getvalue()).decode()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6: POWERPOINT SLIDES
# ─────────────────────────────────────────────────────────────────────────────
def make_slides(insights, report_text):
    slide_prompt = f"""
Create a presentation slide plan. Return ONLY a JSON array, no markdown, no backticks.
Each slide object must have exactly these keys:
{{"title": "string", "bullets": ["string"], "headline_stat": "string", "notes": "string"}}

Generate exactly 7 slides:
1. Title slide
2. Global KPI overview
3. Trend Analysis
4. Peak & Timeline
5. Top Affected Countries
6. Anomalies & Risk Signals
7. Public Health Recommendations

Use this data: {json.dumps(insights)}
"""
    raw = call_claude(
        messages=[{"role": "user", "content": slide_prompt}],
        max_tokens=1500
    )
    slides_plan = json.loads(raw.replace("```json","").replace("```","").strip())

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for sd in slides_plan:
        slide = prs.slides.add_slide(blank)
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = PptxRGB(0x0f, 0x17, 0x2a)

        # Title
        txb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(1.2))
        tf = txb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = sd.get("title", "")
        if p.runs:
            r = p.runs[0]
            r.font.size  = Pt(28)
            r.font.bold  = True
            r.font.color.rgb = PptxRGB(0xff, 0xff, 0xff)

        # Headline stat
        stat = sd.get("headline_stat", "")
        if stat:
            stb = slide.shapes.add_textbox(Inches(9.5), Inches(1.5), Inches(3.3), Inches(1.5))
            sp = stb.text_frame.paragraphs[0]
            sp.text = stat
            if sp.runs:
                sp.runs[0].font.size  = Pt(28)
                sp.runs[0].font.bold  = True
                sp.runs[0].font.color.rgb = PptxRGB(0x4e, 0xcd, 0xc4)

        # Bullets
        bullets = sd.get("bullets", [])
        if bullets:
            bxb = slide.shapes.add_textbox(Inches(0.5), Inches(1.7), Inches(8.8), Inches(5.3))
            btf = bxb.text_frame
            btf.word_wrap = True
            for j, b in enumerate(bullets):
                bp = btf.paragraphs[0] if j == 0 else btf.add_paragraph()
                bp.text = f"→  {b}"
                if bp.runs:
                    bp.runs[0].font.size  = Pt(15)
                    bp.runs[0].font.color.rgb = PptxRGB(0xcc, 0xcc, 0xee)

        # Speaker notes
        notes = sd.get("notes", "")
        if notes:
            slide.notes_slide.notes_text_frame.text = notes

    buf = BytesIO()
    prs.save(buf)
    return base64.b64encode(buf.getvalue()).decode()

# ─────────────────────────────────────────────────────────────────────────────
# /run ENDPOINT — ties all steps together
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/run", methods=["POST"])
def run_pipeline():
    # Verify secret token — reject anyone who isn't Make.com
    token = request.headers.get("X-Pipeline-Secret", "")
    if token != os.environ.get("PIPELINE_SECRET", ""):
        return jsonify({"error": "Unauthorized"}), 401

    try:
        data     = request.get_json(force=True, silent=True) or {}
        csv_text = data.get("csv", "")
        company  = data.get("company", COMPANY_NAME)

        if not csv_text:
            return jsonify({"error": "No CSV provided"}), 400

        # ── Load full dataset ──────────────────────────────────────────────
        try:
            df_full = pd.read_csv(StringIO(csv_text))
        except Exception as e:
            return jsonify({"error": f"CSV parse failed: {str(e)}"}), 400

        # ── FIX #2: Pre-compute accurate global stats from FULL dataset ────
        global_stats = compute_global_stats(df_full)

        # ── FIX #1: Stratified sample — all countries represented ──────────
        df_sample = stratified_sample(df_full, max_rows=MAX_ROWS)

        # ── FIX #3: Clean the sample (agent inspects summary stats) ────────
        cleaning_result = agent_clean_csv(df_sample, max_iterations=MAX_ITERATIONS)
        df_clean        = cleaning_result["clean_df"]
        clean_csv_text  = df_clean.to_csv(index=False)

        # ── Steps 3–6 ──────────────────────────────────────────────────────
        insights       = analyse_with_claude(clean_csv_text, global_stats)
        dashboard_html = generate_dashboard(insights, company)
        report_text    = write_report(insights)
        docx_b64       = make_docx(report_text, insights)
        pptx_b64       = make_slides(insights, report_text)

        date_str = pd.Timestamp.today().strftime("%Y-%m-%d")

        return jsonify({
            "docx_b64":        docx_b64,
            "docx_filename":   f"Report_{date_str}.docx",
            "pptx_b64":        pptx_b64,
            "pptx_filename":   f"Slides_{date_str}.pptx",
            "dashboard_html":  dashboard_html,
            "dash_filename":   f"Dashboard_{date_str}.html",
            "insights":        insights,
            "global_stats":    global_stats,
            "cleaning_log":    cleaning_result["cleaning_log"],
            "flags_for_human": cleaning_result["flags_for_human"],
            "cleaning_stats": {
                "rows_in_full_dataset": len(df_full),
                "rows_sampled":         len(df_sample),
                "rows_after_cleaning":  cleaning_result["rows_after"]
            },
            "status": "success"
        })

    except Exception as e:
        return jsonify({"error": str(e), "status": "failed"}), 500

# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
