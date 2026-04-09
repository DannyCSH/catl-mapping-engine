"""
CATL Four-Layer Mapping Engine - MVP (Product Site v2)
Flask Backend: Session management, Excel COM calculation, output extraction, ZIP export
"""
import os
import sys
import uuid
import json
import shutil
import zipfile
import pandas as pd
from datetime import datetime

# Force UTF-8 mode for Python 3 on Windows
if sys.platform == "win32":
    os.environ["PYTHONIOENCODING"] = "utf-8"
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

from flask import Flask, render_template, request, jsonify, send_file, abort

try:
    import pythoncom
    import win32com.client as win32
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

app = Flask(__name__)

# ============================================================
# PATH CONFIG
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUNTIME_DIR = os.path.join(BASE_DIR, 'runtime')
EXPORTS_DIR = os.path.join(BASE_DIR, 'exports')
LOG_DIR = os.path.join(BASE_DIR, 'logs')

TEMPLATE_XLSX = r"D:\VibeCoding\codex\reports\hku_catl_standards_mapping\catl_four_layer_mapping_engine_generic_template_2026-04-07.xlsx"
SHENXING_XLSX = r"D:\VibeCoding\codex\reports\hku_catl_standards_mapping\catl_four_layer_mapping_engine_shenxing_case_snapshot_2026-04-07.xlsx"
CSV_VALUES_DIR = r"D:\VibeCoding\codex\reports\hku_catl_standards_mapping\review_csv_generic_template_v1_values"

os.makedirs(RUNTIME_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)


def log_event(event_type, message, session_id=None):
    """Append a JSON line to the project coordination log."""
    log_path = r"D:\VibeCoding\codex\reports\hku_catl_standards_mapping\frontend_project_coordination_log_2026-04-08.jsonl"
    entry = {
        "timestamp": datetime.now().isoformat() + "+08:00",
        "type": event_type,
        "project": "hku_catl_standards_mapping_product_site_v2",
        "message": message
    }
    if session_id:
        entry["session_id"] = session_id
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass


def read_csv_auto(path, header=0):
    """Read CSV with automatic encoding detection. header: row index to use as column names (0-based)."""
    for enc in ["utf-8-sig", "gbk", "latin1"]:
        try:
            df = pd.read_csv(path, encoding=enc, header=header, on_bad_lines="skip")
            return df.fillna("")
        except Exception:
            continue
    return pd.DataFrame()


def read_xlsx_sheet(xlsx_path, sheet_name, header=2):
    """Read a sheet from XLSX, returning a DataFrame. header: 0-based row index for column names."""
    try:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=header, engine="openpyxl")
        return df.fillna("")
    except Exception:
        return pd.DataFrame()


# ============================================================
# PAGE ROUTES
# ============================================================
@app.route("/")
def index():
    log_event("note", "Homepage accessed")
    return render_template("index.html")


@app.route("/intake")
def intake():
    return render_template("intake.html")


@app.route("/outputs")
def outputs():
    return render_template("outputs.html")


@app.route("/outputs/eu/<session_id>")
def output_eu(session_id):
    return render_template("output_eu.html", session_id=session_id)


@app.route("/outputs/cbam/<session_id>")
def output_cbam(session_id):
    return render_template("output_cbam.html", session_id=session_id)


@app.route("/outputs/tuv/<session_id>")
def output_tuv(session_id):
    return render_template("output_tuv.html", session_id=session_id)


# ============================================================
# API: SCHEMA
# ============================================================
@app.route("/api/schema/intake")
def schema_intake():
    """Return the form schema derived from 01_手动输入.csv"""
    csv_path = os.path.join(CSV_VALUES_DIR, "01_手动输入.csv")
    try:
        # CSV row 0 = title, row 1 = description, row 2 = column headers, row 3+ = data
        df = read_csv_auto(csv_path, header=2)

        records = df.to_dict(orient="records")
        schema = {}
        for r in records:
            cat = str(r.get("类别", "未分类")).strip()
            if cat not in schema:
                schema[cat] = []
            schema[cat].append({
                "fieldCode":  str(r.get("字段代码", "")),
                "cnLabel":    str(r.get("手动输入项（中文）", "")),
                "enLabel":    str(r.get("英文对应", "")),
                "unit":       str(r.get("单位", "")),
                "exampleValue": str(r.get("示例值", "")),
                "actualValue": str(r.get("实际输入值", "")),
                "note":       str(r.get("说明", "")),
                "evidence":   str(r.get("常见证据", "")),
                "dataType":   str(r.get("字段数据类型", "")),
                "isRequired": str(r.get("是否必填", "")) == "是",
                "impact":     str(r.get("主要影响输出", "")),
                "confidentiality": str(r.get("保密等级", "")),
                "gapStatus":  str(r.get("缺口状态", "")),
            })
        return jsonify({"status": "success", "data": schema})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


# ============================================================
# API: SESSION MANAGEMENT
# ============================================================
@app.route("/api/session", methods=["POST"])
def create_session():
    """Create a new session: copy master workbook to runtime dir."""
    data = request.json or {}
    mode = data.get("mode", "generic")
    session_id = str(uuid.uuid4())[:8]  # short ID for readability
    session_dir = os.path.join(RUNTIME_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    src_file = SHENXING_XLSX if mode == "shenxing" else TEMPLATE_XLSX

    try:
        shutil.copy2(src_file, os.path.join(session_dir, "workbook.xlsx"))

        # Load default values for form pre-fill
        # Shenxing mode: read from Shenxing XLSX sheet; Generic mode: read from CSV
        if mode == "shenxing" and os.path.exists(SHENXING_XLSX):
            df = read_xlsx_sheet(SHENXING_XLSX, "01_手动输入", header=2)
        else:
            csv_path = os.path.join(CSV_VALUES_DIR, "01_手动输入.csv")
            df = read_csv_auto(csv_path, header=2)
        defaults = {}
        for _, row in df.iterrows():
            fc = str(row.get("字段代码", "")).strip()
            av = str(row.get("实际输入值", "")).strip()
            if fc:
                defaults[fc] = av
        with open(os.path.join(session_dir, "draft.json"), "w", encoding="utf-8") as f:
            json.dump({"inputs": defaults, "mode": mode, "created": datetime.now().isoformat()}, f, ensure_ascii=False, indent=2)

        log_event("session_created", f"Session {session_id} created in {mode} mode", session_id)
        return jsonify({"status": "success", "session_id": session_id})
    except Exception as e:
        log_event("error", f"Session creation failed: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/session/<session_id>")
def get_session(session_id):
    """Return current session state."""
    session_dir = os.path.join(RUNTIME_DIR, session_id)
    if not os.path.exists(session_dir):
        return jsonify({"status": "error", "message": "Session not found"}), 404

    draft_path = os.path.join(session_dir, "draft.json")
    if os.path.exists(draft_path):
        with open(draft_path, encoding="utf-8") as f:
            draft = json.load(f)
    else:
        draft = {}

    has_workbook = os.path.exists(os.path.join(session_dir, "workbook.xlsx"))
    has_outputs = any(
        os.path.exists(os.path.join(session_dir, f"{s}.csv"))
        for s in ["02_欧盟电池法输出", "03_CBAM输出"]
    )
    return jsonify({
        "status": "success",
        "session_id": session_id,
        "has_workbook": has_workbook,
        "has_outputs": has_outputs,
        "draft": draft
    })


@app.route("/api/session/<session_id>/save", methods=["POST"])
def save_session(session_id):
    """Save form inputs as draft JSON."""
    session_dir = os.path.join(RUNTIME_DIR, session_id)
    if not os.path.exists(session_dir):
        return jsonify({"status": "error", "message": "Session not found"}), 404

    data = request.json or {}
    inputs = data.get("inputs", {})
    draft_path = os.path.join(session_dir, "draft.json")

    # Merge with existing draft
    existing = {}
    if os.path.exists(draft_path):
        with open(draft_path, encoding="utf-8") as f:
            existing = json.load(f)

    existing["inputs"] = inputs
    existing["last_saved"] = datetime.now().isoformat()

    with open(draft_path, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    return jsonify({"status": "success", "message": "Draft saved"})


@app.route("/api/session/<session_id>/calculate", methods=["POST"])
def calculate_session(session_id):
    """
    计算引擎主入口。

    计算策略（按优先级）：
    1. 纯 Python 计算（默认，始终可用，不依赖 Excel）
       — 基于 calculator.py 中的公式和规则
    2. Excel COM 重算（可选，如果本地有 Excel 且 COM 可用）
       — 复制 workbook → 写入输入 → CalculateFullRebuild → 导出 CSV

    纯 Python 计算完全自包含，可部署到云端。
    """
    session_dir = os.path.join(RUNTIME_DIR, session_id)
    if not os.path.exists(session_dir):
        return jsonify({"status": "error", "message": "Session not found"}), 404

    inputs = request.json.get("inputs", {})

    # Always save draft first
    draft_path = os.path.join(session_dir, "draft.json")
    existing = {}
    if os.path.exists(draft_path):
        with open(draft_path, encoding="utf-8") as f:
            existing = json.load(f)
    existing["inputs"] = inputs
    existing["calculated_at"] = datetime.now().isoformat()
    with open(draft_path, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    # ── Step 1: Pure Python 计算（主路径，始终执行）─────────────
    try:
        from calculator import run_full_calculation
        calc_results = run_full_calculation(inputs)

        # Save all computed CSVs
        sheets_to_save = [
            "02_欧盟电池法输出",
            "03_CBAM输出",
            "04_字段映射过程",
            "05_证据映射过程",
            "06_核验映射过程",
            "07_权限映射过程",
            "09_EU中间表",
            "10_CBAM中间表",
            "13_证据台账",
            "14_校验规则",
            "15_字段缺口与优先级",
            "16_假设与限制",
            "17_责任链与版本",
            "18_前端导出视图",
        ]

        for sname in sheets_to_save:
            if sname in calc_results and isinstance(calc_results[sname], list):
                records = calc_results[sname]
                fpath = os.path.join(session_dir, f"{sname}.csv")
                _write_csv(fpath, records)

        log_event("calculation_complete",
                  f"Session {session_id} pure Python calculation done, {len(calc_results)} sheets generated",
                  session_id)
    except Exception as e:
        log_event("error", f"Python calculation error in session {session_id}: {str(e)}", session_id)
        return jsonify({"status": "error", "message": f"Python calculation failed: {str(e)}"}), 500

    # ── Step 2: Excel COM 重算（仅当 COM 可用且本地有 Excel 时）──
    # 如果 Excel COM 可用，额外执行 Excel 重算以确保与模板计算完全一致
    excel_status = "excel_com_skipped"
    if COM_AVAILABLE:
        excel_status = _try_excel_recalc(session_dir, inputs, session_id)

    return jsonify({
        "status": "success",
        "message": "Calculation complete via pure Python engine",
        "excel_com_status": excel_status,
        "sheets_generated": len(calc_results) - 1,  # -1 for _calculation_meta
    })


def _try_excel_recalc(session_dir, inputs, session_id):
    """尝试 Excel COM 重算（可选增强路径）。"""
    wb_path = os.path.join(session_dir, "workbook.xlsx")
    if not os.path.exists(wb_path):
        return "workbook_not_found"

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        abs_wb_path = os.path.abspath(wb_path)
        wb = excel.Workbooks.Open(abs_wb_path, ReadOnly=False)
        try:
            ws_input = wb.Sheets("01_手动输入")
        except Exception:
            wb.Close(False)
            return "sheet_not_found"
        # Write inputs to column G
        used_range = ws_input.UsedRange
        for row in range(4, used_range.Rows.Count + 1):
            field_code = str(ws_input.Cells(row, 2).Value or "").strip()
            if field_code in inputs:
                ws_input.Cells(row, 7).Value = inputs[field_code]
        excel.CalculateFullRebuild()
        wb.Save()
        wb.Close(False)
        log_event("excel_recalc", f"Session {session_id} Excel COM recalculation done", session_id)
        return "excel_com_success"
    except Exception as e:
        log_event("error", f"Excel COM warning for session {session_id}: {str(e)}", session_id)
        return f"excel_com_error:{str(e)[:50]}"
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _write_csv(path, records):
    """Write list of dicts to a UTF-8 CSV file."""
    if not records:
        return
    import csv
    keys = list(records[0].keys())
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        writer.writerows(records)


# ============================================================
# API: OUTPUT FETCHING
# ============================================================
def _get_output_records(session_id, target):
    """Return output records for EU / CBAM / TUV from session or fallback."""
    session_dir = os.path.join(RUNTIME_DIR, session_id)

    # Map target -> primary sheet + fallback file
    mapping = {
        "eu": {
            "sheets": ["02_欧盟电池法输出", "18_前端导出视图"],
            "fallback": os.path.join(CSV_VALUES_DIR, "02_欧盟电池法输出.csv"),
        },
        "cbam": {
            "sheets": ["03_CBAM输出", "18_前端导出视图"],
            "fallback": os.path.join(CSV_VALUES_DIR, "03_CBAM输出.csv"),
        },
        "tuv": {
            "sheets": ["18_前端导出视图"],
            "fallback": os.path.join(CSV_VALUES_DIR, "18_前端导出视图.csv"),
        },
        "evidence": {
            "sheets": ["05_证据映射过程"],
            "fallback": os.path.join(CSV_VALUES_DIR, "05_证据映射过程.csv"),
        },
        "assurance": {
            "sheets": ["06_核验映射过程"],
            "fallback": os.path.join(CSV_VALUES_DIR, "06_核验映射过程.csv"),
        },
        "access": {
            "sheets": ["07_权限映射过程"],
            "fallback": os.path.join(CSV_VALUES_DIR, "07_权限映射过程.csv"),
        },
        "gaps": {
            "sheets": ["15_字段缺口与优先级"],
            "fallback": os.path.join(CSV_VALUES_DIR, "15_字段缺口与优先级.csv"),
        },
        "assumptions": {
            "sheets": ["16_假设与限制"],
            "fallback": os.path.join(CSV_VALUES_DIR, "16_假设与限制.csv"),
        },
        "evidence_ledger": {
            "sheets": ["13_证据台账"],
            "fallback": os.path.join(CSV_VALUES_DIR, "13_证据台账.csv"),
        },
    }

    if target not in mapping:
        return []

    cfg = mapping[target]
    records = []
    for sname in cfg["sheets"]:
        fpath = os.path.join(session_dir, f"{sname}.csv")
        if not os.path.exists(fpath):
            fpath = cfg["fallback"]
        if os.path.exists(fpath):
            df = read_csv_auto(fpath)
            if target == "tuv" and sname == "18_前端导出视图":
                # Filter only EU rows for TUV overview
                df = df[df.iloc[:, 0].astype(str).str.contains("EU|CBAM", na=False)]
            if len(df) > 2:
                df = df.iloc[2:].reset_index(drop=True)
            records.extend(df.to_dict(orient="records"))
            break

    return records


@app.route("/api/session/<session_id>/outputs/<target>")
def get_outputs(session_id, target):
    session_dir = os.path.join(RUNTIME_DIR, session_id)
    if not os.path.exists(session_dir):
        return jsonify({"status": "error", "message": "Session not found"}), 404
    try:
        records = _get_output_records(session_id, target)
        return jsonify({"status": "success", "data": records})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/outputs/static/<target>")
def get_static_outputs(target):
    """
    Return prebuilt static output data from CSV_VALUES_DIR.
    Used when no active session exists so the UI still shows meaningful content.
    """
    static_mapping = {
        "eu":        "02_欧盟电池法输出.csv",
        "cbam":      "03_CBAM输出.csv",
        "tuv":       "18_前端导出视图.csv",
        "evidence":   "05_证据映射过程.csv",
        "assurance":  "06_核验映射过程.csv",
        "access":     "07_权限映射过程.csv",
        "gaps":       "15_字段缺口与优先级.csv",
        "assumptions":"16_假设与限制.csv",
        "evidence_ledger": "13_证据台账.csv",
    }
    fname = static_mapping.get(target)
    if not fname:
        return jsonify({"status": "error", "message": "Unknown target"}), 400

    fpath = os.path.join(CSV_VALUES_DIR, fname)
    if not os.path.exists(fpath):
        return jsonify({"status": "error", "message": "Static file not found"}), 404

    try:
        df = read_csv_auto(fpath)
        if len(df) > 2:
            df = df.iloc[2:].reset_index(drop=True)
        records = df.to_dict(orient="records")
        return jsonify({"status": "success", "data": records})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/outputs/static_summary")
def get_static_summary():
    """Return counts from static CSV files for the outputs hub page."""
    static_files = {
        "eu": "02_欧盟电池法输出.csv",
        "cbam": "03_CBAM输出.csv",
        "tuv": "18_前端导出视图.csv",
        "evidence": "05_证据映射过程.csv",
        "assurance": "06_核验映射过程.csv",
        "access": "07_权限映射过程.csv",
        "gaps": "15_字段缺口与优先级.csv",
        "assumptions": "16_假设与限制.csv",
        "evidence_ledger": "13_证据台账.csv",
    }
    counts = {}
    for key, fname in static_files.items():
        fpath = os.path.join(CSV_VALUES_DIR, fname)
        if os.path.exists(fpath):
            try:
                df = read_csv_auto(fpath)
                # subtract 2 for header rows
                counts[key + "_count"] = max(0, len(df) - 2)
            except Exception:
                counts[key + "_count"] = 0
        else:
            counts[key + "_count"] = 0
    return jsonify({"status": "success", **counts})


@app.route("/api/session/<session_id>/output_summary")
def get_output_summary(session_id):
    """Return a high-level summary for the outputs overview page."""
    session_dir = os.path.join(RUNTIME_DIR, session_id)
    if not os.path.exists(session_dir):
        return jsonify({"status": "error", "message": "Session not found"}), 404

    def count_records(target):
        return len(_get_output_records(session_id, target))

    return jsonify({
        "status": "success",
        "session_id": session_id,
        "eu_count": count_records("eu"),
        "cbam_count": count_records("cbam"),
        "tuv_count": count_records("tuv"),
        "evidence_count": count_records("evidence"),
        "assurance_count": count_records("assurance"),
        "access_count": count_records("access"),
        "gaps_count": count_records("gaps"),
        "assumptions_count": count_records("assumptions"),
        "evidence_ledger_count": count_records("evidence_ledger"),
    })


# ============================================================
# API: EXPORT
# ============================================================
@app.route("/api/session/<session_id>/export/<target>")
def export_pack(session_id, target):
    """
    Generate a ZIP pack for EU / CBAM / TUV audit-prep.
    ZIP contains: CSVs, manifest.json, README.txt
    """
    session_dir = os.path.join(RUNTIME_DIR, session_id)
    if not os.path.exists(session_dir):
        return jsonify({"status": "error", "message": "Session not found"}), 404

    # Sheet bundles per pack
    all_sheets = [
        "02_欧盟电池法输出",
        "03_CBAM输出",
        "05_证据映射过程",
        "06_核验映射过程",
        "07_权限映射过程",
        "13_证据台账",
        "14_校验规则",
        "15_字段缺口与优先级",
        "16_假设与限制",
        "17_责任链与版本",
    ]
    pack_sheets = {
        "eu": all_sheets,
        "cbam": all_sheets,
        "tuv": all_sheets,
        "all": all_sheets,
    }

    if target not in pack_sheets:
        return jsonify({"status": "error", "message": "Invalid pack type"}), 400

    export_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"CATL_{target.upper()}_pack_{session_id}_{export_ts}.zip"
    zip_path = os.path.join(EXPORTS_DIR, zip_name)

    try:
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            # Write CSVs
            for sname in pack_sheets[target]:
                fpath = os.path.join(session_dir, f"{sname}.csv")
                if os.path.exists(fpath):
                    zf.write(fpath, arcname=f"{sname}.csv")

            # Write manifest.json
            manifest = {
                "pack_type": target.upper(),
                "session_id": session_id,
                "generated_at": datetime.now().isoformat(),
                "includes": pack_sheets[target],
                "disclaimer": (
                    "This is a research prototype output pack. "
                    "Not a legal submission. Not an法定 audit report. "
                    "For internal review and audit preparation only."
                ),
            }
            zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))

            # Write README.txt
            readme = f"""CATL Four-Layer Mapping Engine - {target.upper()} Export Pack
================================================================
Session ID: {session_id}
Generated:  {datetime.now().isoformat()}

CONTENTS:
{chr(10).join(f"  - {s}.csv" for s in pack_sheets[target])}
  - manifest.json

DISCLAIMER:
This pack is generated by a research prototype system (HKU-CATL Joint Lab).
It is NOT:
  - A legal submission to any regulatory authority
  - A formal audit report signed by TÜV / NB / CAB
  - An official CBAM or EU Battery Regulation declaration

It is intended for:
  - Internal data organization
  - Evidence package preparation
  - Audit readiness review

For questions, contact the HKU-CATL research team.
"""
            zf.writestr("README.txt", readme)

        log_event("export", f"Session {session_id} exported {target} pack to {zip_name}", session_id)
        return send_file(zip_path, as_attachment=True, download_name=zip_name)

    except Exception as e:
        log_event("error", f"Export failed for session {session_id}: {str(e)}", session_id)
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/session/<session_id>/export/csv/<sheet_name>")
def export_single_csv(session_id, sheet_name):
    """Export a single sheet as CSV download."""
    session_dir = os.path.join(RUNTIME_DIR, session_id)
    if not os.path.exists(session_dir):
        return jsonify({"status": "error", "message": "Session not found"}), 404

    fpath = os.path.join(session_dir, f"{sheet_name}.csv")
    if not os.path.exists(fpath):
        # Fall back to generic template CSV
        fpath = os.path.join(CSV_VALUES_DIR, f"{sheet_name}.csv")

    if not os.path.exists(fpath):
        return jsonify({"status": "error", "message": f"Sheet {sheet_name} not found"}), 404

    return send_file(fpath, as_attachment=True, download_name=f"{sheet_name}.csv")


# ============================================================
# HEALTH CHECK
# ============================================================
@app.route("/api/health")
def health():
    return jsonify({
        "status": "ok",
        "com_available": COM_AVAILABLE,
        "template_xlsx": os.path.exists(TEMPLATE_XLSX),
        "shenxing_xlsx": os.path.exists(SHENXING_XLSX),
        "csv_values_dir": os.path.exists(CSV_VALUES_DIR),
    })


if __name__ == "__main__":
    print("=" * 60)
    print("CATL Four-Layer Mapping Engine - Product Site v2")
    print("=" * 60)
    print(f"COM available: {COM_AVAILABLE}")
    print(f"Template:      {TEMPLATE_XLSX}")
    print(f"Shenxing:     {SHENXING_XLSX}")
    print(f"CSV Values:    {CSV_VALUES_DIR}")
    print(f"Runtime dir:  {RUNTIME_DIR}")
    print(f"Exports dir:  {EXPORTS_DIR}")
    print("=" * 60)
    print("Starting Flask on http://127.0.0.1:5000")
    print("=" * 60)
    app.run(debug=True, port=5000, host="0.0.0.0")
