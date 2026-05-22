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

# ── CUSTOMISE ─────────────────────────────────────────────────────────────────
COMPANY_NAME = "Global Monkeypox Tracker"
MAX_ROWS = 1000        # keep small for free tier 512MB RAM
MAX_ITERATIONS = 1     # 1 cleaning iteration to save credits + memory

# ── STEP 2: AGENT CLEANING ────────────────────────────────────────────────────
CLEANING_AGENT_PROMPT = """
You are an expert data cleaning agent. You will receive a summary of a dataset.
Identify ONE data quality issue that has not yet been fixed.
Return ONLY a JSON object:

{
  "issue_found": true,
  "issue_type": "null_values | duplicates | outlier | format_inconsistency | placeholder",
  "column": "column name affected, or 'multiple'",
  "description": "Plain English description of the issue",
  "reasoning": "Why you believe this is an error",
  "confidence": "high | medium | low",
  "fix_code": "Single Python line using df variable. Must assign back to df.",
  "needs_human_review": false,
  "human_review_reason": ""
}

If no issues remain, return: { "issue_found": false }

Rules:
- Only ONE issue per response
- fix_code must be a single line
- Never drop more than 5% of rows
"""

VERIFY_PROMPT = """
Verify this data fix was applied correctly.
Return ONLY JSON: { "verified": true, "note": "brief note" }
"""

def summarise_df(df, fixed_so_far):
    summary = {
        "shape": {"rows": int(df.shape[0]), "cols": int(df.shape[1])},
        "columns": {},
        "sample_rows": df.head(3).to_dict(orient='records'),
        "issues_already_fixed": [x.get('description','') for x in fixed_so_far]
    }
    for col in df.columns:
        col_info = {
            "dtype": str(df[col].dtype),
            "nulls": int(df[col].isnull().sum()),
            "unique": int(df[col].nunique()),
            "sample": [str(x) for x in df[col].dropna().unique()[:5]]
        }
        if pd.api.types.is_numeric_dtype(df[col]) and not df[col].isnull().all():
            col_info["min"] = float(df[col].min())
            col_info["max"] = float(df[col].max())
            col_info["mean"] = round(float(df[col].mean()), 2)
        summary["columns"][col] = col_info
    return json.dumps(summary, indent=2, default=str)[:4000]  # cap at 4000 chars

def safe_exec_fix(df, fix_code):
    allowed_globals = {
        "df": df.copy(), "pd": pd, "np": np,
        "__builtins__": {
            "len": len, "int": int, "float": float,
            "str": str, "list": list, "True": True,
            "False": False, "None": None
        }
    }
    exec(fix_code, allowed_globals)
    return allowed_globals["df"]

def agent_clean_csv(csv_text, max_iterations=1):
    try:
        df = pd.read_csv(StringIO(csv_text))
    except Exception:
        df = pd.read_csv(StringIO(csv_text.encode('latin-1').decode('utf-8', errors='replace')))

    rows_before = len(df)
    cleaning_log = []
    flags_for_human = []

    for iteration in range(1, max_iterations + 1):
        try:
            data_summary = summarise_df(df, cleaning_log)
            msg = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=500,
                system=CLEANING_AGENT_PROMPT,
                messages=[{"role": "user", "content": f"Inspect:\n{data_summary}"}]
            )
            raw = msg.content[0].text.strip().replace("```json","").replace("```","").strip()
            decision = json.loads(raw)
        except Exception as e:
            cleaning_log.append({"iteration": iteration, "action": f"error: {str(e)}"})
            break

        if not decision.get("issue_found", False):
            break

        confidence = decision.get("confidence", "low")
        needs_review = decision.get("needs_human_review", False)
        fix_code = decision.get("fix_code", "")
        description = decision.get("description", "")

        if needs_review or confidence == "low":
            flags_for_human.append({
                "iteration": iteration,
                "description": description,
                "action": "flagged for human review"
            })
            cleaning_log.append({"iteration": iteration, "description": description, "action": "flagged"})
            continue

        try:
            rows_before_fix = len(df)
            df = safe_exec_fix(df, fix_code)
            cleaning_log.append({
                "iteration": iteration,
                "issue_type": decision.get("issue_type"),
                "description": description,
                "fix_code": fix_code,
                "rows_before": rows_before_fix,
                "rows_after": len(df),
                "action": "applied"
            })
        except Exception as e:
            cleaning_log.append({"iteration": iteration, "description": description, "action": f"fix_failed: {str(e)}"})

    return {
        "clean_csv": df.to_csv(index=False),
        "cleaning_log": cleaning_log,
        "flags_for_human": flags_for_human,
        "rows_before": rows_before,
        "rows_after": len(df)
    }

# ── STEP 3: ANALYSIS ──────────────────────────────────────────────────────────
ANALYSIS_SYSTEM_PROMPT = """
You are a senior epidemiological data analyst. Dataset: global Monkeypox tracking data.
Columns: location, date, iso_code, total_cases, total_deaths, new_cases, new_deaths,
new_cases_smoothed, new_deaths_smoothed, new_cases_per_million, total_cases_per_million,
total_deaths_per_million.

Return ONLY a valid JSON object:
{
  "period": "date range e.g. May 2022 - Dec 2023",
  "row_count": 0,
  "kpis": {
    "total_global_cases": 0,
    "total_global_deaths": 0,
    "case_fatality_rate_pct": 0.0,
    "peak_daily_new_cases": 0,
    "countries_affected": 0
  },
  "trends": ["trend 1", "trend 2", "trend 3"],
  "anomalies": ["anomaly 1"],
  "top_items": [{"name": "location", "value": 0, "label": "total_cases"}],
  "recommendations": ["recommendation 1", "recommendation 2", "recommendation 3"]
}
No markdown. No explanation. JSON only.
"""

def analyse_with_claude(clean_csv):
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        system=ANALYSIS_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"Analyse:\n\n{clean_csv[:6000]}"}]
    )
    raw = msg.content[0].text.strip().replace("```json","").replace("```","").strip()
    return json.loads(raw)

# ── STEP 4: DASHBOARD ─────────────────────────────────────────────────────────
def generate_dashboard(insights, company=COMPANY_NAME):
    prompt = f"""
Create a complete self-contained HTML dashboard for {company}.
Use Chart.js from https://cdn.jsdelivr.net/npm/chart.js
Include:
- Header with title and period: {insights.get('period','')}
- KPI cards: {json.dumps(insights.get('kpis',{}))}
- Bar chart for top items: {json.dumps(insights.get('top_items',[])[:5])}
- Trends list and recommendations list
Dark theme (#0f172a). Clean and professional.
Return ONLY complete HTML. No markdown. No explanation.
"""
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}]
    )
    return msg.content[0].text.strip().replace("```html","").replace("```","").strip()

# ── STEP 5: REPORT ────────────────────────────────────────────────────────────
REPORT_PROMPT = """
Write a professional analytics report. Use these headings (prefix with #):
# Executive Summary
# Key Performance Indicators
# Trend Analysis
# Anomalies & Risks
# Strategic Recommendations

Concise, data-driven, for senior managers.
Insights: {insights}
"""

def write_report(insights):
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1500,
        messages=[{"role": "user", "content": REPORT_PROMPT.format(insights=json.dumps(insights))}]
    )
    return msg.content[0].text

def make_docx(report_text, insights):
    doc = Document()
    title = doc.add_heading(f"{COMPANY_NAME} — Analytics Report", 0)
    if title.runs:
        title.runs[0].font.color.rgb = RGBColor(0x1a, 0x18, 0x14)
    doc.add_paragraph(f"Period: {insights.get('period','N/A')} | Rows: {insights.get('row_count','N/A')}")

    doc.add_heading("KPI Snapshot", 1)
    kpis = insights.get("kpis", {})
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = "Metric"
    hdr[1].text = "Value"
    for k, v in kpis.items():
        row = table.add_row().cells
        row[0].text = str(k).replace("_"," ").title()
        row[1].text = str(v)

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
        else:
            doc.add_paragraph(line)

    buf = BytesIO()
    doc.save(buf)
    return base64.b64encode(buf.getvalue()).decode()

# ── STEP 6: SLIDES ────────────────────────────────────────────────────────────
def make_slides(insights, report_text):
    slide_prompt = f"""
Create a slide deck plan. Return ONLY a JSON array, no markdown.
Each slide: {{"title": "...", "bullets": ["..."], "headline_stat": "...", "notes": "..."}}
6 slides: title, KPIs, trends, anomalies, top locations, recommendations.
Data: {json.dumps(insights)}
"""
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1500,
        messages=[{"role": "user", "content": slide_prompt}]
    )
    raw = msg.content[0].text.strip().replace("```json","").replace("```","").strip()
    slides_plan = json.loads(raw)

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for sd in slides_plan:
        slide = prs.slides.add_slide(blank)
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = PptxRGB(0x0f, 0x17, 0x2a)

        # Title
        txb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(1))
        tf = txb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = sd.get("title", "")
        if p.runs:
            p.runs[0].font.size = Pt(28)
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = PptxRGB(0xff, 0xff, 0xff)

        # Stat
        stat = sd.get("headline_stat", "")
        if stat:
            stb = slide.shapes.add_textbox(Inches(9), Inches(1.2), Inches(3.8), Inches(1.5))
            sp = stb.text_frame.paragraphs[0]
            sp.text = stat
            if sp.runs:
                sp.runs[0].font.size = Pt(32)
                sp.runs[0].font.bold = True
                sp.runs[0].font.color.rgb = PptxRGB(0x4e, 0xcd, 0xc4)

        # Bullets
        bullets = sd.get("bullets", [])
        if bullets:
            bxb = slide.shapes.add_textbox(Inches(0.5), Inches(1.6), Inches(8), Inches(5))
            btf = bxb.text_frame
            btf.word_wrap = True
            for j, b in enumerate(bullets):
                bp = btf.paragraphs[0] if j == 0 else btf.add_paragraph()
                bp.text = f"->  {b}"
                if bp.runs:
                    bp.runs[0].font.size = Pt(15)
                    bp.runs[0].font.color.rgb = PptxRGB(0xcc, 0xcc, 0xee)

        notes = sd.get("notes", "")
        if notes:
            slide.notes_slide.notes_text_frame.text = notes

    buf = BytesIO()
    prs.save(buf)
    return base64.b64encode(buf.getvalue()).decode()

# ── /run ENDPOINT ─────────────────────────────────────────────────────────────
@app.route("/run", methods=["POST"])
def run_pipeline():
    try:
        data = request.get_json(force=True, silent=True) or {}
        csv_text = data.get("csv", "")
        company = data.get("company", COMPANY_NAME)

        if not csv_text:
            return jsonify({"error": "No CSV provided"}), 400

        # Sample down to MAX_ROWS to stay within free tier memory
        try:
            df_raw = pd.read_csv(StringIO(csv_text))
            if len(df_raw) > MAX_ROWS:
                df_raw = df_raw.sample(n=MAX_ROWS, random_state=42)
            csv_text = df_raw.to_csv(index=False)
        except Exception as e:
            return jsonify({"error": f"CSV parse error: {str(e)}"}), 400

        # Run pipeline steps
        cleaning_result = agent_clean_csv(csv_text, max_iterations=MAX_ITERATIONS)
        clean_csv = cleaning_result["clean_csv"]

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
            "cleaning_log": cleaning_result["cleaning_log"],
            "flags_for_human": cleaning_result["flags_for_human"],
            "cleaning_stats": {
                "rows_before": cleaning_result["rows_before"],
                "rows_after": cleaning_result["rows_after"]
            },
            "status": "success"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
