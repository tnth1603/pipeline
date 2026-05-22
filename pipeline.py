# ── IMPORTS ──────────────────────────────────────────────────────────────────
from flask import Flask, request, jsonify
import anthropic, json, os, base64
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

# ── CUSTOMISE THESE ───────────────────────────────────────────────────────────
COMPANY_NAME = "Global Monkeypox Tracker"

# ── STEP 2: AGENT CLEANING ────────────────────────────────────────────────────
ANALYSIS_SYSTEM_PROMPT = """
You are a senior epidemiological data analyst. The dataset is global Monkeypox tracking data from Our World in Data.

Columns:
- location: country or region name
- date: date of record (YYYY-MM-DD)
- iso_code: country ISO code
- total_cases: cumulative total cases
- total_deaths: cumulative total deaths
- new_cases: new cases on that date
- new_deaths: new deaths on that date
- new_cases_smoothed: 7-day smoothed new cases
- new_deaths_smoothed: 7-day smoothed new deaths
- new_cases_per_million: new cases per million population
- total_cases_per_million: total cases per million population
- new_cases_smoothed_per_million: smoothed new cases per million
- new_deaths_per_million: new deaths per million
- total_deaths_per_million: total deaths per million
- new_deaths_smoothed_per_million: smoothed new deaths per million

Analyse the data and return ONLY a valid JSON object with this exact structure:
{
  "period": "date range covered e.g. May 2022 - Dec 2023",
  "row_count": 0,
  "kpis": {
    "total_global_cases": 0,
    "total_global_deaths": 0,
    "case_fatality_rate_pct": 0.0,
    "peak_daily_new_cases": 0,
    "countries_affected": 0
  },
  "trends": [
    "trend observation 1",
    "trend observation 2",
    "trend observation 3"
  ],
  "anomalies": [
    "anomaly or unusual pattern description"
  ],
  "top_items": [
    {"name": "location name", "value": 0, "label": "total_cases"}
  ],
  "recommendations": [
    "public health recommendation 1",
    "public health recommendation 2",
    "public health recommendation 3"
  ]
}
Return ONLY the JSON object, no markdown, no explanation.
"""

def summarise_df(df, fixed_so_far):
    summary = {
        "shape": {"rows": int(df.shape[0]), "cols": int(df.shape[1])},
        "columns": {},
        "sample_rows": df.head(5).to_dict(orient='records'),
        "issues_already_fixed": [x['issue_type'] + ": " + x['description'] for x in fixed_so_far]
    }
    for col in df.columns:
        col_info = {
            "dtype": str(df[col].dtype),
            "nulls": int(df[col].isnull().sum()),
            "unique": int(df[col].nunique()),
            "sample": [str(x) for x in df[col].dropna().unique()[:8]]
        }
        if pd.api.types.is_numeric_dtype(df[col]):
            col_info["min"] = float(df[col].min()) if not df[col].isnull().all() else None
            col_info["max"] = float(df[col].max()) if not df[col].isnull().all() else None
            col_info["mean"] = float(df[col].mean()) if not df[col].isnull().all() else None
        summary["columns"][col] = col_info
    return json.dumps(summary, indent=2, default=str)

def safe_exec_fix(df, fix_code):
    allowed_globals = {
        "df": df.copy(), "pd": pd, "np": np,
        "__builtins__": {
            "len": len, "int": int, "float": float,
            "str": str, "list": list, "dict": dict,
            "True": True, "False": False, "None": None
        }
    }
    exec(fix_code, allowed_globals)
    return allowed_globals["df"]

def agent_clean_csv(csv_text, max_iterations=10, auto_fix_confidence=["high", "medium"]):
    try:
        df = pd.read_csv(StringIO(csv_text))
    except UnicodeDecodeError:
        df = pd.read_csv(StringIO(csv_text.encode('latin-1').decode('utf-8', errors='replace')))

    rows_before = len(df)
    cleaning_log = []
    flags_for_human = []
    iteration = 0

    while iteration < max_iterations:
        iteration += 1
        data_summary = summarise_df(df, cleaning_log)

        inspect_msg = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=800,
            system=CLEANING_AGENT_PROMPT,
            messages=[{"role": "user", "content": f"Inspect this dataset:\n\n{data_summary}"}]
        )

        raw = inspect_msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
        try:
            decision = json.loads(raw)
        except json.JSONDecodeError:
            break

        if not decision.get("issue_found", False):
            break

        issue_type = decision.get("issue_type", "unknown")
        description = decision.get("description", "")
        confidence = decision.get("confidence", "low")
        needs_review = decision.get("needs_human_review", False)
        fix_code = decision.get("fix_code", "")

        if needs_review or confidence not in auto_fix_confidence:
            flag_entry = {
                "iteration": iteration, "issue_type": issue_type,
                "column": decision.get("column"), "description": description,
                "reasoning": decision.get("reasoning"),
                "review_reason": decision.get("human_review_reason"),
                "action": "skipped — flagged for human review"
            }
            flags_for_human.append(flag_entry)
            cleaning_log.append(flag_entry)
            continue

        rows_before_fix = len(df)
        try:
            df = safe_exec_fix(df, fix_code)
            rows_after_fix = len(df)
        except Exception as e:
            cleaning_log.append({
                "iteration": iteration, "issue_type": issue_type,
                "description": description, "action": f"fix_failed: {str(e)}"
            })
            continue

        col_stats_after = summarise_df(df, [])
        verify_msg = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=200,
            system=VERIFY_PROMPT,
            messages=[{"role": "user", "content": f"Fix: {description}\nCode: {fix_code}\nAfter: {col_stats_after[:500]}"}]
        )
        verify_raw = verify_msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
        try:
            verify = json.loads(verify_raw)
        except:
            verify = {"verified": True, "note": "verification parse error"}

        log_entry = {
            "iteration": iteration, "issue_type": issue_type,
            "column": decision.get("column"), "description": description,
            "reasoning": decision.get("reasoning"), "confidence": confidence,
            "fix_code": fix_code, "rows_before": rows_before_fix,
            "rows_after": rows_after_fix, "verified": verify.get("verified"),
            "verify_note": verify.get("note"), "action": "applied"
        }
        cleaning_log.append(log_entry)

    return {
        "clean_csv": df.to_csv(index=False),
        "cleaning_log": cleaning_log,
        "flags_for_human": flags_for_human,
        "iterations": iteration,
        "rows_before": rows_before,
        "rows_after": len(df)
    }

# ── STEP 3: ANALYSIS ──────────────────────────────────────────────────────────
ANALYSIS_SYSTEM_PROMPT = """
You are a senior data analyst. Analyse the CSV data and return ONLY a valid JSON object with keys:
{
  "period": "date range covered",
  "row_count": 0,
  "kpis": {"metric_name": value},
  "trends": ["trend 1", "trend 2", "trend 3"],
  "anomalies": ["anomaly description"],
  "top_items": [{"name": "item", "value": 0, "label": "metric"}],
  "recommendations": ["action 1", "action 2", "action 3"]
}
Return ONLY the JSON object, no markdown, no explanation.
"""

def analyse_with_claude(clean_csv):
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1500,
        system=ANALYSIS_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"Analyse this dataset:\n\n{clean_csv[:8000]}"}]
    )
    raw = msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

# ── STEP 4: DASHBOARD ─────────────────────────────────────────────────────────
def generate_dashboard(insights, company=COMPANY_NAME):
    prompt = f"""
Create a complete self-contained HTML dashboard for {company}.
Use Chart.js from https://cdn.jsdelivr.net/npm/chart.js
Include:
- Header with company name and period: {insights.get('period','')}
- KPI cards: {json.dumps(insights.get('kpis',{}))}
- Bar chart for top items: {json.dumps(insights.get('top_items',[]))}
- Trends and anomalies sections
- Recommendations section
Use a dark theme (#0f172a background). Make it visually polished.
Return ONLY the complete HTML, no explanation, no markdown fences.
"""
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )
    html = msg.content[0].text.strip().replace("```html", "").replace("```", "").strip()
    return html

# ── STEP 5: REPORT ────────────────────────────────────────────────────────────
REPORT_PROMPT = """
Write a professional business analytics report based on these insights.
Use these exact section headings (start each with #):
# Executive Summary
# Key Performance Indicators
# Trend Analysis
# Anomalies & Risks
# Strategic Recommendations

Be concise, data-driven, and action-oriented.
Write for senior business managers. No jargon.
Insights: {insights}
"""

def write_report(insights):
    prompt = REPORT_PROMPT.format(insights=json.dumps(insights, indent=2))
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )
    return msg.content[0].text

def make_docx(report_text, insights):
    doc = Document()
    title = doc.add_heading(f"{COMPANY_NAME} — Analytics Report", 0)
    title.runs[0].font.color.rgb = RGBColor(0x1a, 0x18, 0x14)
    doc.add_paragraph(f"Period: {insights.get('period', 'N/A')} | Rows analysed: {insights.get('row_count', 'N/A')}")

    doc.add_heading("KPI Snapshot", 1)
    kpis = insights.get("kpis", {})
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = "Metric"; hdr[1].text = "Value"
    for k, v in kpis.items():
        row = table.add_row().cells
        row[0].text = str(k).replace("_", " ").title()
        row[1].text = str(v)

    for line in report_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        elif line.startswith("# "):
            doc.add_heading(line[2:], 1)
        elif line.startswith("## "):
            doc.add_heading(line[3:], 2)
        elif line.startswith("- ") or line.startswith("• "):
            doc.add_paragraph(line[2:], style='List Bullet')
        else:
            doc.add_paragraph(line)

    buf = BytesIO()
    doc.save(buf)
    return base64.b64encode(buf.getvalue()).decode()

# ── STEP 6: SLIDES ────────────────────────────────────────────────────────────
def make_slides(insights, report_text):
    slide_prompt = f"""
Create a slide deck plan for a business analytics presentation.
Return ONLY a JSON array, no markdown. Each slide object:
{{
  "title": "Slide title",
  "bullets": ["Point 1", "Point 2", "Point 3"],
  "headline_stat": "Key number to highlight",
  "notes": "Speaker notes"
}}
Generate 6-8 slides covering: title, KPIs, trends, anomalies, top items, recommendations, next steps.
Data: {json.dumps(insights, indent=2)}
"""
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": slide_prompt}]
    )
    raw = msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
    slides_plan = json.loads(raw)

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for slide_data in slides_plan:
        slide = prs.slides.add_slide(blank_layout)
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = PptxRGB(0x0f, 0x17, 0x2a)

        txb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(1))
        tf = txb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = slide_data.get("title", "")
        p.runs[0].font.size = Pt(28)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = PptxRGB(0xff, 0xff, 0xff)

        stat = slide_data.get("headline_stat", "")
        if stat:
            stb = slide.shapes.add_textbox(Inches(9), Inches(1.2), Inches(3.8), Inches(1.5))
            stf = stb.text_frame
            sp = stf.paragraphs[0]
            sp.text = stat
            sp.runs[0].font.size = Pt(36)
            sp.runs[0].font.bold = True
            sp.runs[0].font.color.rgb = PptxRGB(0x4e, 0xcd, 0xc4)

        bullets = slide_data.get("bullets", [])
        if bullets:
            bxb = slide.shapes.add_textbox(Inches(0.5), Inches(1.6), Inches(8), Inches(5))
            btf = bxb.text_frame
            btf.word_wrap = True
            for j, bullet in enumerate(bullets):
                bp = btf.paragraphs[0] if j == 0 else btf.add_paragraph()
                bp.text = f"→  {bullet}"
                bp.runs[0].font.size = Pt(16)
                bp.runs[0].font.color.rgb = PptxRGB(0xcc, 0xcc, 0xee)

        notes = slide_data.get("notes", "")
        if notes:
            slide.notes_slide.notes_text_frame.text = notes

    buf = BytesIO()
    prs.save(buf)
    return base64.b64encode(buf.getvalue()).decode()

# ── STEP 7: /run ENDPOINT ─────────────────────────────────────────────────────
@app.route("/run", methods=["POST"])
def run_pipeline():
    data = request.get_json()
    csv_text = data.get("csv", "")
    company = data.get("company", COMPANY_NAME)

    if not csv_text:
        return jsonify({"error": "No CSV provided"}), 400

    try:
        cleaning_result = agent_clean_csv(csv_text, max_iterations=10)
        clean_csv = cleaning_result["clean_csv"]
        cleaning_log = cleaning_result["cleaning_log"]
        flags = cleaning_result["flags_for_human"]

        insights = analyse_with_claude(clean_csv)
        dashboard_html = generate_dashboard(insights, company)
        report_text = write_report(insights)
        docx_b64 = make_docx(report_text, insights)
        pptx_b64 = make_slides(insights, report_text)

        date_str = pd.Timestamp.today().strftime("%Y-%m-%d")
        return jsonify({
            "docx_b64": docx_b64,
            "docx_filename": f"Report_{date_str}.docx",
            "pptx_b64": pptx_b64,
            "pptx_filename": f"Slides_{date_str}.pptx",
            "dashboard_html": dashboard_html,
            "dash_filename": f"Dashboard_{date_str}.html",
            "insights": insights,
            "cleaning_log": cleaning_log,
            "flags_for_human": flags,
            "cleaning_stats": {
                "rows_before": cleaning_result["rows_before"],
                "rows_after": cleaning_result["rows_after"],
                "iterations": cleaning_result["iterations"]
            },
            "status": "success"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
