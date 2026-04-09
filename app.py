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
DATA_DIR = os.path.join(BASE_DIR, 'data')
RUNTIME_DIR = os.path.join(BASE_DIR, 'runtime')
EXPORTS_DIR = os.path.join(BASE_DIR, 'exports')
LOG_DIR = os.path.join(BASE_DIR, 'logs')

TEMPLATE_XLSX = os.path.join(DATA_DIR, 'generic_template.xlsx')
SHENXING_XLSX = os.path.join(DATA_DIR, 'shenxing_case.xlsx')
CSV_VALUES_DIR = os.path.join(DATA_DIR, 'csv_values')

os.makedirs(RUNTIME_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)


def log_event(event_type, message, session_id=None):
    """Append a JSON line to the project coordination log."""
    log_path = os.path.join(LOG_DIR, 'events.jsonl')
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
# AI PROVIDER CONFIG
# ============================================================
# In-memory API key store. Keys are provider names, values are dicts with api_key and optional base_url.
AI_PROVIDER_KEYS = {
    "openai":      {"api_key": "", "base_url": "https://api.openai.com/v1"},
    "anthropic":   {"api_key": "", "base_url": "https://api.anthropic.com"},
    "minimax":     {"api_key": "", "base_url": "https://api.minimax.chat/v1"},
    "zhipu":       {"api_key": "", "base_url": "https://open.bigmodel.cn/api/paas/v4"},
    "qwen":        {"api_key": "", "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1"},
    "deepseek":    {"api_key": "", "base_url": "https://api.deepseek.com/v1"},
    "gemini":      {"api_key": "", "base_url": "https://generativelanguage.googleapis.com/v1beta"},
}

AI_PROVIDER_KEYS_FILE = os.path.join(BASE_DIR, "ai_keys.json")


def load_ai_keys():
    """Load persisted API keys from disk."""
    if os.path.exists(AI_PROVIDER_KEYS_FILE):
        try:
            with open(AI_PROVIDER_KEYS_FILE, encoding="utf-8") as f:
                saved = json.load(f)
            for k, v in saved.items():
                if k in AI_PROVIDER_KEYS:
                    AI_PROVIDER_KEYS[k]["api_key"] = v.get("api_key", "")
                    AI_PROVIDER_KEYS[k]["base_url"] = v.get("base_url", AI_PROVIDER_KEYS[k]["base_url"])
        except Exception:
            pass


def save_ai_keys():
    """Persist API keys to disk."""
    try:
        with open(AI_PROVIDER_KEYS_FILE, "w", encoding="utf-8") as f:
            json.dump(AI_PROVIDER_KEYS, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


load_ai_keys()


def _build_openai_messages(provider, system, user_content, image_data=None):
    """Build provider-specific message payload for chat completion."""
    if provider == "anthropic":
        msgs = [{"role": "user", "content": [{"type": "text", "text": user_content}]}]
        if image_data:
            msgs[0]["content"].insert(0, {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": image_data}})
        return msgs
    elif provider in ("minimax", "qwen"):
        msgs = [{"role": "system", "content": system}, {"role": "user", "content": user_content}]
        return msgs
    else:
        content = user_content
        if image_data:
            content = [{"type": "text", "text": user_content}, {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_data}"}}]
        return [{"role": "system", "content": system}, {"role": "user", "content": content}]


def _call_ai(provider, model, system, user_content, image_data=None, timeout=60):
    """Unified AI chat completion across multiple providers."""
    import urllib.request
    import urllib.error

    cfg = AI_PROVIDER_KEYS.get(provider, {})
    api_key = cfg.get("api_key", "")
    base_url = cfg.get("base_url", "")

    if not api_key:
        raise Exception(f"API key not configured for provider: {provider}")

    messages = _build_openai_messages(provider, system, user_content, image_data)

    if provider == "anthropic":
        payload = {
            "model": model,
            "messages": messages,
            "max_tokens": 4096,
        }
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
        url = f"{base_url}/messages"
    elif provider == "gemini":
        payload = {
            "contents": [{"parts": [{"text": user_content}]}],
            "generationConfig": {"temperature": 0.1, "maxOutputTokens": 4096},
        }
        headers = {"content-type": "application/json", "x-goog-api-key": api_key}
        url = f"{base_url}/models/{model}:generateContent"
    elif provider == "minimax":
        # Minimax requires group_id as URL param; API key as Bearer token
        payload = {
            "model": model,
            "messages": messages,
            "temperature": 0.1,
        }
        headers = {"Authorization": f"Bearer {api_key}", "content-type": "application/json"}
        # group_id is the first segment of the API key (before the dot)
        group_id = api_key.split(".")[0] if "." in api_key else api_key[:8]
        url = f"{base_url}/text/chatcompletion_v2?group_id={group_id}"
    else:
        payload = {
            "model": model,
            "messages": messages,
            "temperature": 0.1,
            "max_tokens": 4096,
        }
        headers = {"Authorization": f"Bearer {api_key}", "content-type": "application/json"}
        url = f"{base_url}/chat/completions"

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        raise Exception(f"AI API error ({e.code}): {err_body[:500]}")

    if provider == "anthropic":
        return result.get("content", [{}])[0].get("text", "")
    elif provider == "gemini":
        return result.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "")
    else:
        return result.get("choices", [{}])[0].get("message", {}).get("content", "")


# ============================================================
# FILE UPLOAD & AI PARSING
# ============================================================
@app.route("/api/ai/keys", methods=["GET", "POST"])
def ai_keys():
    """Get or update AI provider API keys."""
    if request.method == "POST":
        data = request.json or {}
        for prov in AI_PROVIDER_KEYS:
            if prov in data:
                AI_PROVIDER_KEYS[prov]["api_key"] = str(data[prov].get("api_key", "")).strip()
                AI_PROVIDER_KEYS[prov]["base_url"] = str(data[prov].get("base_url", AI_PROVIDER_KEYS[prov]["base_url"])).strip()
        save_ai_keys()
        return jsonify({"status": "success", "message": "Keys saved"})
    else:
        return jsonify({
            "status": "success",
            "providers": {
                k: {"has_key": bool(v["api_key"]), "base_url": v["base_url"]}
                for k, v in AI_PROVIDER_KEYS.items()
            }
        })


@app.route("/api/ai/parse_files", methods=["POST"])
def ai_parse_files():
    """
    Accept uploaded files + session_id + provider/model.
    Returns parsed field values extracted from the files.
    """
    if "files" not in request.files and "file_urls" not in request.json:
        return jsonify({"status": "error", "message": "No files provided"}), 400

    data = request.json or {}
    session_id = data.get("session_id")
    provider = data.get("provider", "openai")
    model = data.get("model", "gpt-4o")
    field_schema = data.get("field_schema", {})  # {field_code: {label, note, exampleValue}}

    if not session_id:
        return jsonify({"status": "error", "message": "No session_id provided"}), 400
    if provider not in AI_PROVIDER_KEYS:
        return jsonify({"status": "error", "message": f"Unknown provider: {provider}"}), 400

    # Build field schema description for the AI
    schema_text = "字段说明如下（字段代码：标签 / 说明）：\n"
    for fc, info in field_schema.items():
        label = info.get("label", fc)
        note = info.get("note", "")
        example = info.get("exampleValue", "")
        schema_text += f"- {fc}：{label}。说明：{note}。示例值：{example}\n"

    system_prompt = (
        "你是一个专业的欧盟电池法规数据提取助手。用户的文件中包含电池产品的技术参数和碳足迹相关数据。"
        "你的任务是从上传的文件中提取字段值，并以JSON格式返回。"
        "返回格式：{\"字段代码\": \"提取的值\", ...}。"
        "只返回你确信能从文件中提取到的值，对于无法确定的字段不要返回。"
        "所有值都应该是字符串格式。"
    )

    results = {}

    # Process uploaded files
    files = request.files.getlist("files") or []
    for f in files:
        fname = f.filename or "unknown"
        ext = os.path.splitext(fname)[-1].lower()

        if ext in (".png", ".jpg", ".jpeg", ".gif", ".webp"):
            # Image: use base64 + vision-capable model
            import base64
            img_bytes = f.read()
            img_b64 = base64.b64encode(img_bytes).decode("utf-8")
            user_msg = (
                f"请从这张图片中提取电池产品的参数数据。\n{schema_text}\n"
                "只返回能从图片中明确看到的字段数据。"
            )
            try:
                text = _call_ai(provider, model, system_prompt, user_msg, image_data=img_b64)
                parsed = _extract_json_from_response(text)
                results.update(parsed)
            except Exception as e:
                results[f"_error_{fname}"] = str(e)

        elif ext in (".pdf", ".docx", ".txt", ".csv", ".xlsx", ".xls"):
            # Text/Excel: read content and send as text
            content = ""
            try:
                if ext == ".pdf":
                    try:
                        import pypdf
                        reader = pypdf.PdfReader(f)
                        for page in reader.pages:
                            content += page.extract_text() or ""
                    except ImportError:
                        content = "[PDF解析需要安装 pypdf: pip install pypdf]"
                elif ext in (".xlsx", ".xls"):
                    import io
                    df = pd.read_excel(io.BytesIO(f.read()), header=None)
                    content = df.to_csv(index=False, encoding="utf-8")
                elif ext == ".csv":
                    content = f.read().decode("utf-8-sig", errors="replace")
                else:
                    content = f.read().decode("utf-8", errors="replace")
            except Exception as e:
                content = f"[读取失败: {str(e)}]"

            user_msg = (
                f"请从以下文件内容中提取电池产品的参数数据。\n{schema_text}\n"
                "文件内容如下：\n" + content[:8000]
            )
            try:
                text = _call_ai(provider, model, system_prompt, user_msg)
                parsed = _extract_json_from_response(text)
                results.update(parsed)
            except Exception as e:
                results[f"_error_{fname}"] = str(e)

    return jsonify({"status": "success", "parsed": results, "provider": provider})


@app.route("/api/ai/compliance_check", methods=["POST"])
def ai_compliance_check():
    """
    Run EU Battery Regulation compliance pre-check on session inputs.
    Returns missing/uncertain fields with priority levels.
    """
    data = request.json or {}
    session_id = data.get("session_id")
    provider = data.get("provider", "openai")
    model = data.get("model", "gpt-4o")

    if not session_id:
        return jsonify({"status": "error", "message": "No session_id provided"}), 400

    session_dir = os.path.join(RUNTIME_DIR, session_id)
    draft_path = os.path.join(session_dir, "draft.json")
    if not os.path.exists(draft_path):
        return jsonify({"status": "error", "message": "Session not found"}), 404

    with open(draft_path, encoding="utf-8") as f:
        draft = json.load(f)

    inputs = draft.get("inputs", {})

    # Load field schema for labels
    csv_path = os.path.join(CSV_VALUES_DIR, "01_手动输入.csv")
    df = read_csv_auto(csv_path, header=2)
    field_labels = {}
    for _, row in df.iterrows():
        fc = str(row.get("字段代码", "")).strip()
        label = str(row.get("字段标签", fc)).strip()
        category = str(row.get("字段分组", "")).strip()
        if fc:
            field_labels[fc] = {"label": label, "category": category}

    # Build current inputs summary
    inputs_summary = "\n".join([
        f"- {fc}: {val} （{field_labels.get(fc, {}).get('label', fc)}）"
        for fc, val in inputs.items() if val and str(val).strip()
    ]) or "（暂无填写数据）"

    system_prompt = (
        "你是一个欧盟电池法规（EU Battery Regulation 2023/1542）合规审查专家。"
        "你的任务是审查电池产品的碳足迹申报字段，找出缺失或不符合要求的数据。"
        "根据EU Battery Regulation Annex II和Article 7的要求，判断哪些字段缺失、哪些值疑似错误。"
        "以JSON格式返回审查结果："
        '{"issues": [{"field": "字段代码", "label": "字段标签", "issue": "问题描述", "priority": "P1/P2/P3", "suggestion": "建议操作"}]}'
        "P1 = 缺失会导致申报被拒绝的必填字段；P2 = 重要但可后续补充；P3 = 建议优化。"
        "只报告你确信存在的问题，不要随意猜测。"
    )

    user_msg = (
        "请审查以下电池产品的申报字段合规情况：\n\n"
        f"当前已填写数据：\n{inputs_summary}\n\n"
        "请对照EU Battery Regulation要求，识别缺失或不符合规范的数据项。"
    )

    try:
        text = _call_ai(provider, model, system_prompt, user_msg)
        issues = _extract_json_from_response(text).get("issues", [])
        return jsonify({"status": "success", "issues": issues, "provider": provider})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


def _extract_json_from_response(text):
    """Try to extract JSON from AI response text."""
    import re
    text = text.strip()
    # Try direct JSON parse first
    try:
        return json.loads(text)
    except Exception:
        pass
    # Try to find JSON in markdown code blocks
    match = re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", text)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except Exception:
            pass
    # Try to find {...} pattern
    match = re.search(r"\{[\s\S]+\}", text)
    if match:
        try:
            return json.loads(match.group(0))
        except Exception:
            pass
    return {}


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

    if not COM_AVAILABLE:
        return "com_unavailable"

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
    import sys
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5001
    print("=" * 60)
    print("CATL Four-Layer Mapping Engine - Product Site v2 (AI)")
    print("=" * 60)
    print(f"COM available: {COM_AVAILABLE}")
    print(f"Template:      {TEMPLATE_XLSX}")
    print(f"Shenxing:     {SHENXING_XLSX}")
    print(f"CSV Values:    {CSV_VALUES_DIR}")
    print(f"Runtime dir:  {RUNTIME_DIR}")
    print(f"Exports dir:  {EXPORTS_DIR}")
    print("=" * 60)
    print(f"Starting Flask on http://127.0.0.1:{port}")
    print("=" * 60)
    app.run(debug=True, port=port, host="0.0.0.0")
