#!/usr/bin/env python3
"""
app.py -- a local browser UI for the slide-generation platform.

This does NOT reimplement the pipeline -- it wraps build_deck.py exactly
as-is, so any future generator/extraction changes made to that file are
picked up automatically here too.

Run once from a terminal:
    python3 app.py
Then open http://localhost:5001 in your browser. Everything after that
(uploading a lesson, generating a deck, downloading the result) happens
in the browser -- no further command-line use needed.

The API key you enter in the browser is kept in server memory only for
the duration of that one request, in a local environment variable --
never written to disk, never logged, and cleared immediately after the
build finishes (success or failure). Same in spirit as the in-browser
key-prompt built earlier for teacher.html's AI features.
"""
import os
import sys
import io
import json
import traceback
import contextlib
import uuid
import threading
import hmac
from functools import wraps
from flask import Flask, request, render_template_string, send_file, jsonify

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_deck

OUTPUT_DIR = os.path.join(HERE, "generated")
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB upload cap

UI_USERNAME = os.environ.get("UI_USERNAME")
UI_PASSWORD = os.environ.get("UI_PASSWORD")


def require_ui_auth(view):
    """
    Protects the browser UI (index/generate/download) with HTTP Basic
    Auth. Does NOT apply to /webhook/* -- those use their own bearer
    token, a separate mechanism for a separate caller (approve.php, not
    a person in a browser).

    If UI_USERNAME/UI_PASSWORD aren't set, the UI is unprotected -- fine
    for local testing on localhost, not fine once this is deployed
    anywhere reachable by anyone else.
    """
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not UI_USERNAME or not UI_PASSWORD:
            return view(*args, **kwargs)
        auth = request.authorization
        if not auth or not (hmac.compare_digest(auth.username or "", UI_USERNAME)
                             and hmac.compare_digest(auth.password or "", UI_PASSWORD)):
            return ("Authentication required.", 401,
                    {"WWW-Authenticate": 'Basic realm="RedThread Slide Deck Platform"'})
        return view(*args, **kwargs)
    return wrapped

PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>RedThread Slide Deck Platform</title>
<style>
  :root { --purple: #3928AA; --navy: #0E0142; --coral: #F19A65; --blue-pale: #E5EFF9; --border: #DCD6EE; }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, 'Segoe UI', Arial, sans-serif; background: #F6F7FB; color: var(--navy); margin: 0; }
  header { background: var(--purple); color: white; padding: 24px 32px; }
  header h1 { margin: 0; font-size: 22px; }
  header p { margin: 6px 0 0; opacity: .85; font-size: 13px; }
  main { max-width: 720px; margin: 32px auto; padding: 0 20px; }
  .card { background: white; border: 1px solid var(--border); border-radius: 10px; padding: 24px; margin-bottom: 20px; }
  label { display: block; font-size: 13px; font-weight: 600; color: var(--purple); margin-bottom: 6px; text-transform: uppercase; letter-spacing: .03em; }
  input[type=text], input[type=password], textarea, select {
    width: 100%; padding: 10px 12px; border: 1px solid var(--border); border-radius: 6px;
    font-size: 14px; font-family: inherit; margin-bottom: 16px; background: var(--blue-pale);
  }
  textarea { min-height: 160px; resize: vertical; }
  input[type=file] { margin-bottom: 16px; }
  .row { display: flex; gap: 20px; }
  .row > div { flex: 1; }
  button {
    background: var(--coral); color: var(--navy); border: none; border-radius: 8px;
    padding: 12px 24px; font-size: 15px; font-weight: 700; cursor: pointer;
  }
  button:hover { opacity: .9; }
  button:disabled { opacity: .5; cursor: not-allowed; }
  #log { background: var(--navy); color: #C9BEEB; font-family: monospace; font-size: 12.5px;
         padding: 16px; border-radius: 8px; white-space: pre-wrap; min-height: 60px; margin-top: 16px; display: none; }
  #result { margin-top: 16px; display: none; }
  #result a { display: inline-block; background: var(--purple); color: white; padding: 12px 20px;
              border-radius: 8px; text-decoration: none; font-weight: 700; }
  .hint { font-size: 12px; color: #6B6478; margin-top: -10px; margin-bottom: 16px; }
  .tabs { display: flex; gap: 8px; margin-bottom: 16px; }
  .tab { padding: 8px 16px; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 600; background: var(--blue-pale); }
  .tab.active { background: var(--purple); color: white; }
</style>
</head>
<body>
<header>
  <h1>RedThread Slide Deck Platform</h1>
  <p>Upload a lesson, get back a finished, on-brand PowerPoint deck.</p>
</header>
<main>
  <div class="card">
    <div class="tabs">
      <div class="tab active" data-mode="raw" onclick="setMode('raw')">Raw lesson text</div>
      <div class="tab" data-mode="json" onclick="setMode('json')">Pre-extracted JSON</div>
    </div>

    <div id="mode-raw">
      {% if server_has_key %}
      <div class="hint" style="margin-top:0;">A server-configured API key is already available -- you don't need to enter one. (You still can, to use your own key for this one build instead.)</div>
      {% endif %}
      <label>Anthropic API key {% if server_has_key %}(optional){% endif %}</label>
      <input type="password" id="apiKey" placeholder="sk-ant-...">
      <div class="hint">Only used in memory for this one build, never saved or logged.</div>
      <label>Lesson file (.pdf or .txt Teacher Guide excerpt)</label>
      <input type="file" id="rawFile" accept=".txt,.pdf">
      <div class="hint">PDFs are sent to Claude as-is (it reads the actual page layout, tables, and highlighting directly) -- or paste plain text below instead of uploading a file.</div>
      <textarea id="rawText" placeholder="Paste raw lesson text here..."></textarea>
    </div>

    <div id="mode-json" style="display:none">
      <label>Pre-extracted lesson JSON</label>
      <input type="file" id="jsonFile" accept=".json">
      <div class="hint">Skips the extraction step -- useful for testing without an API key.</div>
    </div>

    <button id="goBtn" onclick="generate()">Generate slide deck</button>
    <div id="log"></div>
    <div id="result"></div>
  </div>
</main>

<script>
function setMode(mode) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.mode === mode));
  document.getElementById('mode-raw').style.display = mode === 'raw' ? 'block' : 'none';
  document.getElementById('mode-json').style.display = mode === 'json' ? 'block' : 'none';
}

async function generate() {
  const btn = document.getElementById('goBtn');
  const log = document.getElementById('log');
  const result = document.getElementById('result');
  log.style.display = 'block';
  result.style.display = 'none';
  log.textContent = 'Starting...';
  btn.disabled = true;

  const mode = document.querySelector('.tab.active').dataset.mode;
  const form = new FormData();
  form.append('mode', mode);

  if (mode === 'raw') {
    const apiKey = document.getElementById('apiKey').value.trim();
    if (apiKey) form.append('apiKey', apiKey);
    const file = document.getElementById('rawFile').files[0];
    const text = document.getElementById('rawText').value.trim();
    if (file) form.append('rawFile', file);
    else if (text) form.append('rawText', text);
    else { log.textContent = 'Upload a lesson file or paste lesson text first.'; btn.disabled = false; return; }
  } else {
    const file = document.getElementById('jsonFile').files[0];
    if (!file) { log.textContent = 'Choose a lesson JSON file first.'; btn.disabled = false; return; }
    form.append('jsonFile', file);
  }

  try {
    const res = await fetch('/generate', { method: 'POST', body: form });
    const data = await res.json();
    log.textContent = data.log || '';
    if (data.ok) {
      result.style.display = 'block';
      result.innerHTML = `<a href="/download/${data.filename}">Download ${data.filename}</a>`;
      if (data.extracted_json) {
        result.innerHTML += ` &nbsp; <a href="/download/${data.extracted_json}" style="background:var(--navy);">See what Claude extracted (JSON)</a>`;
      }
      if (data.drive_link) {
        result.innerHTML += ` &nbsp; <a href="${data.drive_link}" target="_blank" style="background:#0F9D58;">Open in Google Drive</a>`;
      }
    }
  } catch (e) {
    log.textContent = 'Request failed: ' + e;
  }
  btn.disabled = false;
}
</script>
</body>
</html>
"""


@app.route("/healthz")
def healthz():
    # Deliberately unauthenticated -- platform health checks (Render,
    # Cloud Run, etc.) can't supply basic-auth credentials, so this can't
    # live behind require_ui_auth the way "/" does. Returns no real
    # information, just confirms the process is up and responding.
    return jsonify(status="ok"), 200


@app.route("/")
@require_ui_auth
def index():
    return render_template_string(PAGE, server_has_key=bool(os.environ.get("ANTHROPIC_API_KEY")))


@app.route("/generate", methods=["POST"])
@require_ui_auth
def generate():
    mode = request.form.get("mode")
    job_id = uuid.uuid4().hex[:8]
    work_dir = os.path.join(OUTPUT_DIR, job_id)
    os.makedirs(work_dir, exist_ok=True)

    log_buf = io.StringIO()
    api_key_was_set = False
    try:
        if mode == "raw":
            browser_api_key = request.form.get("apiKey", "").strip()
            server_has_key = bool(os.environ.get("ANTHROPIC_API_KEY"))
            if browser_api_key:
                os.environ["ANTHROPIC_API_KEY"] = browser_api_key
                api_key_was_set = True
            elif not server_has_key:
                return jsonify(ok=False, log="No API key provided, and none is configured on the server."), 400
            # else: a server-side ANTHROPIC_API_KEY is already set (e.g. via
            # this deployment's environment config) -- use it as-is, don't
            # require the browser to supply one too.

            if "rawFile" in request.files and request.files["rawFile"].filename:
                uploaded = request.files["rawFile"]
                ext = os.path.splitext(uploaded.filename)[1].lower()
                if ext not in (".txt", ".pdf"):
                    return jsonify(ok=False, log=f"Unsupported file type {ext!r} -- upload a .txt or .pdf lesson."), 400
                lesson_path = os.path.join(work_dir, f"lesson{ext}")
                uploaded.save(lesson_path)
            else:
                raw_text = request.form.get("rawText", "").strip()
                if not raw_text:
                    return jsonify(ok=False, log="No lesson text or file provided."), 400
                lesson_path = os.path.join(work_dir, "lesson.txt")
                with open(lesson_path, "w", encoding="utf-8") as f:
                    f.write(raw_text)
        else:
            if "jsonFile" not in request.files or not request.files["jsonFile"].filename:
                return jsonify(ok=False, log="No JSON file provided."), 400
            lesson_path = os.path.join(work_dir, "lesson.json")
            request.files["jsonFile"].save(lesson_path)

        out_pptx = os.path.join(work_dir, "deck.pptx")

        with contextlib.redirect_stderr(log_buf):
            build_result = build_deck.build(lesson_path, out_pptx)

        if not os.path.exists(out_pptx):
            return jsonify(ok=False, log=log_buf.getvalue() + "\nNo output file was produced."), 500

        try:
            lesson_data = json.load(open(lesson_path, encoding="utf-8")) if lesson_path.endswith(".json") else None
            lesson_num = (lesson_data or {}).get("lesson_number", job_id)
        except Exception:
            lesson_num = job_id
        final_name = f"Lesson_{lesson_num}_{job_id}.pptx"
        final_path = os.path.join(OUTPUT_DIR, final_name)
        os.replace(out_pptx, final_path)

        extracted_json_name = None
        extracted_json_src = out_pptx + ".extracted.json"
        if os.path.exists(extracted_json_src):
            extracted_json_name = final_name + ".extracted.json"
            os.replace(extracted_json_src, os.path.join(OUTPUT_DIR, extracted_json_name))

        return jsonify(ok=True, log=log_buf.getvalue(), filename=final_name, extracted_json=extracted_json_name,
                       drive_link=build_result.get("drive_link") if build_result else None)

    except SystemExit as e:
        return jsonify(ok=False, log=log_buf.getvalue() + f"\n{e}"), 400
    except Exception:
        return jsonify(ok=False, log=log_buf.getvalue() + "\n" + traceback.format_exc()), 500
    finally:
        if api_key_was_set:
            os.environ.pop("ANTHROPIC_API_KEY", None)


WEBHOOK_TOKEN = os.environ.get("SLIDE_DECK_SERVICE_TOKEN")
WEBHOOK_LOG_DIR = os.path.join(OUTPUT_DIR, "webhook_jobs")
os.makedirs(WEBHOOK_LOG_DIR, exist_ok=True)


def _process_lesson_approved(payload):
    """
    Runs in a background thread, well after the webhook HTTP response has
    already been sent -- approve.php never waits on this.
    """
    job_note_path = os.path.join(WEBHOOK_LOG_DIR, f"{payload.get('lesson_id', 'unknown')}.json")

    if not payload.get("lesson_text"):
        # Known gap: rt_ingestion_extraction_items doesn't carry lesson
        # prose yet, only title/page-range metadata. Recording the
        # notification so it's visible and retriable once that's
        # resolved, rather than silently dropping it or fabricating
        # content to extract from.
        payload["_status"] = "waiting_on_lesson_text"
        payload["_note"] = (
            "Received lesson-approved notification, but lesson_text was "
            "empty. This lesson cannot be extracted/generated until the "
            "ingestion side provides the actual lesson prose (see the "
            "known gap noted in approve.php). Re-run this lesson once "
            "that's available."
        )
        with open(job_note_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        print(f"[webhook] lesson_id={payload.get('lesson_id')}: {payload['_note']}", file=sys.stderr)
        return

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        payload["_status"] = "failed_no_api_key"
        with open(job_note_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        print("[webhook] ANTHROPIC_API_KEY not set on the slide-deck service -- cannot extract.", file=sys.stderr)
        return

    lesson_txt_path = os.path.join(WEBHOOK_LOG_DIR, f"{payload['lesson_id']}.txt")
    with open(lesson_txt_path, "w", encoding="utf-8") as f:
        f.write(payload["lesson_text"])

    out_pptx = os.path.join(OUTPUT_DIR, f"Lesson_{payload.get('lesson_number', payload['lesson_id'])}.pptx")
    try:
        build_result = build_deck.build(lesson_txt_path, out_pptx)
        payload["_status"] = "done"
        payload["_output"] = out_pptx
        payload["_drive_link"] = (build_result or {}).get("drive_link")
        print(f"[webhook] lesson_id={payload.get('lesson_id')}: generated {out_pptx}"
              + (f", uploaded to {payload['_drive_link']}" if payload["_drive_link"] else ""), file=sys.stderr)
    except SystemExit as e:
        payload["_status"] = "unsupported_type"
        payload["_note"] = str(e)
        print(f"[webhook] lesson_id={payload.get('lesson_id')}: {e}", file=sys.stderr)
    except Exception:
        payload["_status"] = "failed"
        payload["_error"] = traceback.format_exc()
        print(f"[webhook] lesson_id={payload.get('lesson_id')} failed:\n{payload['_error']}", file=sys.stderr)

    with open(job_note_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)


@app.route("/webhook/lesson-approved", methods=["POST"])
def webhook_lesson_approved():
    if not WEBHOOK_TOKEN:
        return jsonify(error="Slide-deck service has no SLIDE_DECK_SERVICE_TOKEN configured -- refusing all webhook calls until it does."), 503

    auth = request.headers.get("Authorization", "")
    provided = auth[7:] if auth.startswith("Bearer ") else ""
    if not provided or not hmac.compare_digest(provided, WEBHOOK_TOKEN):
        return jsonify(error="Invalid or missing bearer token"), 401

    payload = request.get_json(silent=True)
    if not payload or not payload.get("lesson_id"):
        return jsonify(error="Missing or invalid JSON body (lesson_id required)"), 400

    # Respond immediately -- approve.php is waiting with only a ~3s
    # timeout. All real work happens in the background after this.
    threading.Thread(target=_process_lesson_approved, args=(payload,), daemon=True).start()
    return jsonify(accepted=True, lesson_id=payload["lesson_id"]), 202


@app.route("/webhook/status/<lesson_id>")
def webhook_status(lesson_id):
    path = os.path.join(WEBHOOK_LOG_DIR, f"{os.path.basename(lesson_id)}.json")
    if not os.path.exists(path):
        return jsonify(status="not_found_or_still_processing"), 404
    return jsonify(json.load(open(path, encoding="utf-8")))


@app.route("/download/<filename>")
@require_ui_auth
def download(filename):
    safe_name = os.path.basename(filename)
    path = os.path.join(OUTPUT_DIR, safe_name)
    if not os.path.exists(path):
        return "File not found or already cleaned up.", 404
    return send_file(path, as_attachment=True, download_name=safe_name)


RETENTION_HOURS = float(os.environ.get("RETENTION_HOURS", "24"))


def _cleanup_old_files():
    """
    Runs once at startup and then every hour in the background. Deletes
    generated decks and webhook job logs older than RETENTION_HOURS.
    A shared server accumulates these indefinitely otherwise -- this was
    never a concern for the local single-user version.
    """
    import time
    while True:
        cutoff = time.time() - (RETENTION_HOURS * 3600)
        removed = 0
        for root, dirs, files in os.walk(OUTPUT_DIR):
            for name in files:
                path = os.path.join(root, name)
                try:
                    if os.path.getmtime(path) < cutoff:
                        os.remove(path)
                        removed += 1
                except OSError:
                    pass
        if removed:
            print(f"[cleanup] removed {removed} file(s) older than {RETENTION_HOURS}h", file=sys.stderr)
        time.sleep(3600)


threading.Thread(target=_cleanup_old_files, daemon=True).start()

if __name__ == "__main__":
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "5001"))
    print(f"Starting RedThread Slide Deck Platform on {host}:{port}...")
    app.run(host=host, port=port, debug=False)
