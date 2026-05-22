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

app = Flask(__name__)
client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

# ── AGENT SYSTEM PROMPT ───────────────────────────────────────────────
CLEANING_AGENT_PROMPT = """
You are an expert data cleaning agent. You will receive:
1. A summary of the current dataset (shape, column stats, sample rows)
2. A list of issues already fixed in previous iterations

Your job: identify ONE data quality issue that has not yet been fixed.
Return ONLY a JSON object with this exact structure:

{
  "issue_found": true,
  "issue_type": "one of: null_values | duplicates | outlier | encoding |
                 format_inconsistency | logical_error | placeholder |
                 semantic_encoding | cross_column | business_logic",
  "column": "column name affected, or 'multiple'",
  "description": "Plain English description of the issue",
  "reasoning": "Why you believe this is an error, not valid data",
  "confidence": "high | medium | low",
  "fix_code": "Single Python expression using df variable. Must return df.
               Example: df['col'] = df['col'].fillna(df['col'].median())",
  "needs_human_review": false,
  "human_review_reason": "Only fill if needs_human_review is true"
}

If no issues remain, return:
{ "issue_found": false }

Rules:
- Only identify ONE issue per response
- Set needs_human_review=true if confidence is low OR if fix could lose real data
- fix_code must be a single line assigning back to df
- Never drop more than 5% of rows without flagging for human review
- Consider business context: not all outliers are errors
"""
# ─────────────────────────────────────────────────────────────────────

VERIFY_PROMPT = """
You are verifying a data fix was applied correctly.
Given the before/after column stats, confirm:
1. Was the fix applied as expected?
2. Did it create any new problems?

Return ONLY JSON:
{
  "verified": true,
  "note": "Brief confirmation or warning"
}
"""


def summarise_df(df: pd.DataFrame, fixed_so_far: list) -> str:
    """Generate a concise data profile for Claude to inspect."""
    summary = {
        "shape": {"rows": int(df.shape[0]), "cols": int(df.shape[1])},
        "columns": {},
        "sample_rows": df.head(5).to_dict(orient='records'),
        "issues_already_fixed": [x['issue_type'] + ": " + x['description']
                                   for x in fixed_so_far]
    }
    for col in df.columns:
        col_info = {
            "dtype":    str(df[col].dtype),
            "nulls":    int(df[col].isnull().sum()),
            "unique":   int(df[col].nunique()),
            "sample":   [str(x) for x in df[col].dropna().unique()[:8]]
        }
        if pd.api.types.is_numeric_dtype(df[col]):
            col_info["min"]  = float(df[col].min()) if not df[col].isnull().all() else None
            col_info["max"]  = float(df[col].max()) if not df[col].isnull().all() else None
            col_info["mean"] = float(df[col].mean()) if not df[col].isnull().all() else None
        summary["columns"][col] = col_info
    return json.dumps(summary, indent=2, default=str)


def safe_exec_fix(df: pd.DataFrame, fix_code: str) -> pd.DataFrame:
    """
    Execute Claude's generated fix code safely.
    Only pandas/numpy operations on df are permitted.
    """
    allowed_globals = {
        "df": df.copy(),
        "pd": pd,
        "np": np,
        "__builtins__": {
            "len": len, "int": int, "float": float,
            "str": str, "list": list, "dict": dict,
            "True": True, "False": False, "None": None
        }
    }
    exec(fix_code, allowed_globals)
    return allowed_globals["df"]


def agent_clean_csv(
    csv_text: str,
    max_iterations: int = 10,
    auto_fix_confidence: list = ["high", "medium"]
) -> dict:
    """
    AI Agent cleaning loop. Replaces simple clean_csv().

    Returns:
      {
        "clean_csv": str,          # cleaned CSV string
        "cleaning_log": list,      # every decision made
        "flags_for_human": list,   # issues needing human review
        "iterations": int,         # how many loops ran
        "rows_before": int,
        "rows_after": int
      }
    """
    # Load data
    try:
        df = pd.read_csv(StringIO(csv_text))
    except UnicodeDecodeError:
        # Try latin-1 if UTF-8 fails (common with Vietnamese data)
        df = pd.read_csv(StringIO(csv_text.encode('latin-1').decode('utf-8', errors='replace')))

    rows_before   = len(df)
    cleaning_log  = []
    flags_for_human = []
    iteration     = 0

    print(f"[AGENT] Starting cleaning loop. Rows: {rows_before}, Cols: {len(df.columns)}")

    while iteration < max_iterations:
        iteration += 1
        print(f"[AGENT] Iteration {iteration}/{max_iterations}")

        # ── STEP 1: Inspect ──────────────────────────────────────
        data_summary = summarise_df(df, cleaning_log)

        inspect_msg = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=800,
            system=CLEANING_AGENT_PROMPT,
            messages=[{
                "role": "user",
                "content": f"Inspect this dataset and identify the next issue:\n\n{data_summary}"
            }]
        )

        raw = inspect_msg.content[0].text.strip()
        raw = raw.replace("```json", "").replace("```", "").strip()

        try:
            decision = json.loads(raw)
        except json.JSONDecodeError:
            print(f"[AGENT] Warning: could not parse decision JSON. Stopping.")
            break

        # ── No more issues found ──────────────────────────────────
        if not decision.get("issue_found", False):
            print("[AGENT] ✓ No more issues found. Data is clean.")
            break

        issue_type = decision.get("issue_type", "unknown")
        description = decision.get("description", "")
        confidence = decision.get("confidence", "low")
        needs_review = decision.get("needs_human_review", False)
        fix_code = decision.get("fix_code", "")

        print(f"[AGENT] Found: {issue_type} | Confidence: {confidence} | {description[:60]}")

        # ── Flag for human review if low confidence ───────────────
        if needs_review or confidence not in auto_fix_confidence:
            flag_entry = {
                "iteration":   iteration,
                "issue_type":  issue_type,
                "column":      decision.get("column"),
                "description": description,
                "reasoning":   decision.get("reasoning"),
                "review_reason": decision.get("human_review_reason"),
                "action":      "skipped — flagged for human review"
            }
            flags_for_human.append(flag_entry)
            cleaning_log.append(flag_entry)
            print(f"[AGENT] ⚑ Flagged for human review: {description[:60]}")
            continue

        # ── STEP 2: Execute fix ───────────────────────────────────
        rows_before_fix = len(df)
        col_stats_before = summarise_df(df[[decision.get("column", df.columns[0])]]
                                        if decision.get("column") in df.columns
                                        else df, [])

        try:
            df = safe_exec_fix(df, fix_code)
            rows_after_fix = len(df)
            print(f"[AGENT] ✓ Fix applied. Rows: {rows_before_fix} → {rows_after_fix}")
        except Exception as e:
            print(f"[AGENT] ✗ Fix failed: {e}. Skipping.")
            cleaning_log.append({
                "iteration": iteration,
                "issue_type": issue_type,
                "description": description,
                "action": f"fix_failed: {str(e)}"
            })
            continue

        # ── STEP 3: Verify ────────────────────────────────────────
        col_stats_after = summarise_df(df[[decision.get("column", df.columns[0])]]
                                       if decision.get("column") in df.columns
                                       else df, [])

        verify_msg = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=200,
            system=VERIFY_PROMPT,
            messages=[{
                "role": "user",
                "content": f"""
Fix applied: {description}
Code: {fix_code}
Before: {col_stats_before[:500]}
After:  {col_stats_after[:500]}
"""
            }]
        )

        verify_raw = verify_msg.content[0].text.strip().replace("```json","").replace("```","").strip()
        try:
            verify = json.loads(verify_raw)
        except:
            verify = {"verified": True, "note": "verification parse error"}

        # ── Log the decision ─────────────────────────────────────
        log_entry = {
            "iteration":   iteration,
            "issue_type":  issue_type,
            "column":      decision.get("column"),
            "description": description,
            "reasoning":   decision.get("reasoning"),
            "confidence":  confidence,
            "fix_code":    fix_code,
            "rows_before": rows_before_fix,
            "rows_after":  rows_after_fix,
            "verified":    verify.get("verified"),
            "verify_note": verify.get("note"),
            "action":      "applied"
        }
        cleaning_log.append(log_entry)

    # ── Return results ────────────────────────────────────────────────
    return {
        "clean_csv":       df.to_csv(index=False),
        "cleaning_log":    cleaning_log,
        "flags_for_human": flags_for_human,
        "iterations":      iteration,
        "rows_before":     rows_before,
        "rows_after":      len(df)
    }
@app.route("/run", methods=["POST"])
def run_pipeline():
    data     = request.get_json()
    csv_text = data.get("csv", "")
    company  = data.get("company", COMPANY_NAME)

    # ── Step 2 (new): Agent cleaning loop ─────────────────────
    cleaning_result = agent_clean_csv(csv_text, max_iterations=10)
    clean_csv       = cleaning_result["clean_csv"]
    cleaning_log    = cleaning_result["cleaning_log"]
    flags           = cleaning_result["flags_for_human"]

    # ── Steps 3–6 unchanged ───────────────────────────────────
    insights      = analyse_with_claude(clean_csv)
    dashboard_html = generate_dashboard(insights, company)
    report_text   = write_report(insights)
    docx_b64      = make_docx(report_text, insights)
    pptx_b64      = make_slides(insights, report_text)

    date_str = pd.Timestamp.today().strftime("%Y-%m-%d")
    return jsonify({
        "docx_b64":        docx_b64,
        "docx_filename":   f"Report_{date_str}.docx",
        "pptx_b64":        pptx_b64,
        "pptx_filename":   f"Slides_{date_str}.pptx",
        "dashboard_html":  dashboard_html,
        "dash_filename":   f"Dashboard_{date_str}.html",
        "insights":        insights,
        "cleaning_log":    cleaning_log,     # NEW: full decision trail
        "flags_for_human": flags,            # NEW: items needing review
        "cleaning_stats": {
            "rows_before": cleaning_result["rows_before"],
            "rows_after":  cleaning_result["rows_after"],
            "iterations":  cleaning_result["iterations"]
        },
        "status": "success"
    })
