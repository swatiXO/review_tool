"""A small web page for the review: upload a zip, watch progress, download the results.

Runs on your own machine (127.0.0.1 by default). One review runs at a time; the rest wait in
a queue. Every job keeps its files in its own folder under the jobs directory.

  python -m course_review.cli serve --checklist Course-Review-Checklist.xlsx
"""
import html
import json
import os
import re
import shutil
import threading
import time
import traceback
import uuid
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from flask import Flask, Response, abort, jsonify, redirect, render_template_string, request, send_from_directory, url_for

from . import cli

JOB_ID = re.compile(r"^[0-9a-f]{12}$")
DOWNLOADS = {"report.html": "text/html; charset=utf-8", "Course-Review-Checklist-filled.xlsx":
             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "review.json": "application/json",
             cli.MARKED_UP: "application/zip"}

PAGE = """<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{ title }}</title>
<style>
:root{--bg:#f5f7f9;--fg:#14202b;--mut:#566472;--line:#d5dce3;--card:#fff;--accent:#0b6b73;--ok:#146c43;--bad:#b42318;--warn:#8a5a00}
@media (prefers-color-scheme:dark){:root{--bg:#0e1519;--fg:#e3eaf0;--mut:#93a2ae;--line:#27333b;--card:#151e24;--accent:#4cc3cc;--ok:#5fd3a0;--bad:#ff8a80;--warn:#e3a04a}}
*{box-sizing:border-box}body{background:var(--bg);color:var(--fg);font:15px/1.55 system-ui,"Segoe UI",sans-serif;margin:0;padding:24px 16px}
main{max-width:46rem;margin:auto;display:flex;flex-direction:column;gap:20px}
h1{font-size:1.5rem;margin:0}h2{font-size:1.05rem;margin:0 0 6px}p{margin:0}a{color:var(--accent)}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:16px;display:flex;flex-direction:column;gap:12px}
label{display:flex;flex-direction:column;gap:4px;font-weight:600}label span{font-weight:400;color:var(--mut);font-size:.88rem}
input[type=file],input[type=text],select{font:inherit;padding:8px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--fg);width:100%}
button,.btn{font:inherit;font-weight:600;padding:9px 16px;border-radius:6px;border:0;background:var(--accent);color:#fff;cursor:pointer;text-decoration:none;display:inline-block}
@media (prefers-color-scheme:dark){button,.btn{color:#04252a}}
button.link{background:none;color:var(--mut);font-weight:400;text-decoration:underline;padding:0}
.row{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.mut{color:var(--mut)}
table{border-collapse:collapse;width:100%}th,td{text-align:left;padding:7px 8px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-size:.75rem;text-transform:uppercase;letter-spacing:.05em;color:var(--mut)}
.bar{height:10px;background:var(--line);border-radius:5px;overflow:hidden}.bar i{display:block;height:100%;background:var(--accent);width:0;transition:width .3s}
.ok{color:var(--ok);font-weight:600}.bad{color:var(--bad);font-weight:600}.warn{color:var(--warn);font-weight:600}
details summary{cursor:pointer;color:var(--mut)}
@media (prefers-reduced-motion:reduce){.bar i{transition:none}}
</style></head><body><main>
<header><h1><a href="{{ url_for('home') }}" style="color:inherit;text-decoration:none">Course Review</a></h1>
<p class="mut">Reviews a course package against the checklist workbook, on this computer.</p></header>
{{ body|safe }}
</main></body></html>"""

HOME = """
<form class="card" method="post" action="{{ url_for('create_job') }}" enctype="multipart/form-data">
  <h2>Review a package</h2>
  <label>Course package (.zip)<input type="file" name="package" accept=".zip" required></label>
  <label>Checklist workbook<span>{% if default_checklist %}Leave empty to use <b>{{ default_checklist }}</b>.{% else %}Required: no default checklist is configured.{% endif %}</span>
    <input type="file" name="checklist" accept=".xlsx" {% if not default_checklist %}required{% endif %}></label>
  <label>Use a language model<span>Only for question formats the rules do not recognise, and for judgement checks. Results from it are always suggestions.</span>
    <select name="model_mode">
      <option value="off">No, rules only (fast)</option>
      <option value="fallback">For unrecognised question formats</option>
      <option value="checks">Also for content checks (slow)</option>
    </select></label>
  <details><summary>Model and textbook settings</summary>
    <div class="card" style="border:0;padding:12px 0 0">
      <label>Ollama address<span>For example your ngrok link.</span><input type="text" name="model_url" value="{{ model_url }}" placeholder="http://localhost:11434"></label>
      <label>Model name<input type="text" name="model_name" value="{{ model_name }}" placeholder="qwen3:14b"></label>
      <label>Textbook index<span>Made with the index-book command.</span>
        <select name="book_index"><option value="">None</option>{% for b in books %}<option>{{ b }}</option>{% endfor %}</select></label>
    </div></details>
  <div class="row"><button type="submit">Start review</button><span class="mut">A 126-file package takes about a minute without a model.</span></div>
</form>
{% if jobs %}<div class="card"><h2>Earlier reviews</h2><table><tr><th>Package<th>When<th>Result</tr>
{% for j in jobs %}<tr><td><a href="{{ url_for('job_page', job_id=j.id) }}">{{ j.name }}</a>
<td class="mut">{{ j.created }}<td>{{ j.summary }}</tr>{% endfor %}</table></div>{% endif %}
"""

JOB = """
<div class="card"><h2>{{ job.name }}</h2>
<p class="mut">Started {{ job.created }}{% if job.options %} · {{ job.options }}{% endif %}</p>
<div id="state">
{% if job.state in ('queued','running') %}
  <p id="stage">{{ job.stage or 'Waiting to start' }}</p><div class="bar"><i id="fill"></i></div>
  <p class="mut">This page updates by itself. You can leave it and come back from the home page.</p>
{% elif job.state == 'failed' %}
  <p class="bad">The review could not finish.</p><p>{{ job.error }}</p>
{% else %}
  <p class="ok">Done in {{ job.seconds }} s: {{ job.summary }}</p>
  <div class="row"><a class="btn" href="{{ url_for('download', job_id=job.id, name='Marked-up-documents.zip') }}">Download marked-up documents</a>
  <a href="{{ url_for('download', job_id=job.id, name='Course-Review-Checklist-filled.xlsx') }}">Filled workbook</a>
  <a href="{{ url_for('download', job_id=job.id, name='report.html') }}" target="_blank">Open the report</a>
  <a href="{{ url_for('download', job_id=job.id, name='review.json') }}">JSON</a></div>
  <p class="mut">In the marked-up documents: red = fails the checklist, turquoise = model suggestion for a reviewer, green = passes.
  Each highlight has a comment naming the checklist item. A blank cell in the workbook means the tool did not decide it.</p>
{% endif %}</div>
<form method="post" action="{{ url_for('delete_job', job_id=job.id) }}"><button class="link" type="submit">Delete this review and its files</button></form></div>
{% if job.state in ('queued','running') %}<script>
const bar=document.getElementById('fill'),stage=document.getElementById('stage');
async function tick(){try{const r=await fetch('{{ url_for("job_status", job_id=job.id) }}');const s=await r.json();
 if(s.state==='done'||s.state==='failed'){location.reload();return}
 stage.textContent=s.stage||'Working';if(s.total){bar.style.width=Math.round(100*s.done/s.total)+'%'}else{bar.style.width='100%';bar.style.opacity='.35'}}catch(e){}
 setTimeout(tick,1500)}
tick();</script>{% endif %}
"""


class SyncExecutor:
    """Runs a job immediately; used by the tests."""

    def submit(self, fn, *a, **k):
        fn(*a, **k)


def _read_meta(job_dir: Path):
    try:
        return json.loads((job_dir / "meta.json").read_text(encoding="utf8"))
    except (OSError, ValueError):
        return None


def _write_meta(job_dir: Path, meta):
    tmp = job_dir / "meta.json.tmp"
    tmp.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf8")
    os.replace(tmp, job_dir / "meta.json")


def create_app(jobs_dir="web_jobs", checklist=None, books_dir="book_indexes", executor=None, max_upload_mb=500, cache_dir=None):
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = max_upload_mb * 1024 * 1024
    jobs = Path(jobs_dir).resolve()      # absolute: Flask resolves relative download folders against the package, not the cwd
    jobs.mkdir(parents=True, exist_ok=True)
    books = Path(books_dir).resolve()
    cache_dir = cache_dir or str(jobs / ".model_cache")
    pool = executor or ThreadPoolExecutor(max_workers=1)     # one review at a time: it is CPU heavy
    lock = threading.Lock()
    default_checklist = str(checklist) if checklist else None

    def page(title, template, **ctx):
        body = render_template_string(template, **ctx)
        return Response(render_template_string(PAGE, title=title, body=body),
                        headers={"Content-Security-Policy": "default-src 'self'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; frame-ancestors 'none'",
                                 "X-Content-Type-Options": "nosniff"})

    def job_dir(job_id):
        if not JOB_ID.match(job_id):
            abort(404)
        d = jobs / job_id
        if not d.is_dir():
            abort(404)
        return d

    def update(job_dir_, **kw):
        with lock:
            meta = _read_meta(job_dir_) or {}
            meta.update(kw)
            _write_meta(job_dir_, meta)

    def run_job(job_id, opts):
        d = jobs / job_id
        started = time.time()
        update(d, state="running", stage="Starting")

        def progress(stage, done, total):
            update(d, stage=stage + (f" ({done} of {total})" if total else ""), done=done, total=total)
        try:
            model = None
            if opts["model_mode"] != "off":
                model = cli.make_model("suggest", opts["model_url"] or None, opts["model_name"] or None, cache_dir,
                                       judge=opts["model_mode"] == "checks")
            book = None
            if opts["book_index"]:
                from .book import BookIndex
                book = BookIndex.load(books / opts["book_index"])
            pkg, res, xlsx, secs = cli.review(str(d / "package.zip"), opts["checklist"], str(d / "out"), model=model, book=book, progress=progress)
            c = Counter(f.status for f in res.findings)
            for name in DOWNLOADS:
                src = d / "out" / name
                if src.exists():
                    shutil.copy2(src, d / name)
            shutil.rmtree(d / "out", ignore_errors=True)
            update(d, state="done", stage="Done", seconds=round(time.time() - started),
                   summary=f"{len(pkg.docs)} documents: {c['fail']} fails, {c['pass']} passes, {c['needs_review']} for a reviewer",
                   counts=dict(c))
        except SystemExit as e:                    # a model or book problem reported by the CLI helpers
            update(d, state="failed", error=str(e))
        except Exception as e:
            traceback.print_exc()
            update(d, state="failed", error=f"{type(e).__name__}: {e}")
        finally:
            (d / "package.zip").unlink(missing_ok=True)       # the upload is not kept after the review

    @app.get("/")
    def home():
        listed = []
        for d in sorted(jobs.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True)[:15]:
            m = _read_meta(d) if d.is_dir() and JOB_ID.match(d.name) else None
            if m:
                listed.append({"id": d.name, "name": m.get("name", d.name), "created": m.get("created", ""),
                               "summary": m.get("summary") or {"failed": "Failed", "running": "Running...", "queued": "Waiting..."}.get(m.get("state"), "")})
        names = sorted(p.name for p in books.iterdir() if (p / "pages").is_dir()) if books.is_dir() else []
        return page("Course Review", HOME, jobs=listed, books=names, default_checklist=Path(default_checklist).name if default_checklist else None,
                    model_url=os.environ.get("OLLAMA_URL", ""), model_name=os.environ.get("OLLAMA_MODEL", ""))

    @app.post("/jobs")
    def create_job():
        pkg = request.files.get("package")
        if pkg is None or not pkg.filename:
            return page("Course Review", "<div class='card'><p class='bad'>Choose a zip file to review.</p></div>"), 400
        if not pkg.filename.lower().endswith(".zip"):
            return page("Course Review", "<div class='card'><p class='bad'>The package must be a .zip file.</p></div>"), 400
        job_id = uuid.uuid4().hex[:12]
        d = jobs / job_id
        d.mkdir()
        pkg.save(d / "package.zip")
        up = request.files.get("checklist")
        if up is not None and up.filename:
            if not up.filename.lower().endswith(".xlsx"):
                shutil.rmtree(d, ignore_errors=True)
                return page("Course Review", "<div class='card'><p class='bad'>The checklist must be an .xlsx file.</p></div>"), 400
            up.save(d / "checklist.xlsx")
            checklist_path = str(d / "checklist.xlsx")
        elif default_checklist:
            checklist_path = default_checklist
        else:
            shutil.rmtree(d, ignore_errors=True)
            return page("Course Review", "<div class='card'><p class='bad'>Upload the checklist workbook.</p></div>"), 400
        mode = request.form.get("model_mode", "off")
        if mode not in ("off", "fallback", "checks"):
            mode = "off"
        book = request.form.get("book_index", "")
        if book and not (books / book / "pages").is_dir():
            book = ""
        opts = {"checklist": checklist_path, "model_mode": mode, "model_url": request.form.get("model_url", "").strip(),
                "model_name": request.form.get("model_name", "").strip(), "book_index": book}
        label = {"off": "", "fallback": "model for unrecognised formats", "checks": "model for formats and content checks"}[mode]
        _write_meta(d, {"id": job_id, "name": os.path.basename(pkg.filename), "created": time.strftime("%Y-%m-%d %H:%M"),
                        "state": "queued", "stage": "Waiting to start", "options": label})
        pool.submit(run_job, job_id, opts)
        return redirect(url_for("job_page", job_id=job_id))

    @app.get("/jobs/<job_id>")
    def job_page(job_id):
        d = job_dir(job_id)
        return page("Course Review", JOB, job=_Obj(_read_meta(d) or {}))

    @app.get("/jobs/<job_id>/status")
    def job_status(job_id):
        m = _read_meta(job_dir(job_id)) or {}
        return jsonify({k: m.get(k) for k in ("state", "stage", "done", "total", "error", "summary")})

    @app.get("/jobs/<job_id>/files/<name>")
    def download(job_id, name):
        d = job_dir(job_id)
        if name not in DOWNLOADS or not (d / name).exists():
            abort(404)
        resp = send_from_directory(d, name, mimetype=DOWNLOADS[name].split(";")[0], as_attachment=name.endswith(("xlsx", "json", "zip")))
        if name == "report.html":
            resp.headers["Content-Security-Policy"] = "default-src 'none'; style-src 'unsafe-inline'; frame-ancestors 'none'"
        return resp

    @app.post("/jobs/<job_id>/delete")
    def delete_job(job_id):
        shutil.rmtree(job_dir(job_id), ignore_errors=True)
        return redirect(url_for("home"))

    # jobs left 'running' by a crash must not look alive forever
    for d in jobs.iterdir():
        m = _read_meta(d) if d.is_dir() and JOB_ID.match(d.name) else None
        if m and m.get("state") in ("queued", "running"):
            m.update(state="failed", error="The server was restarted while this review was running. Start it again.")
            _write_meta(d, m)
    return app


class _Obj:
    """Attribute access over a dict for the template, with missing keys as empty."""

    def __init__(self, d):
        self.__dict__.update(d)

    def __getattr__(self, name):
        return ""


def serve(checklist, host="127.0.0.1", port=8080, jobs_dir="web_jobs", books_dir="book_indexes"):
    app = create_app(jobs_dir=jobs_dir, checklist=checklist, books_dir=books_dir)
    print(f"Course Review is running at http://{host}:{port}  (Ctrl+C to stop)")
    if host not in ("127.0.0.1", "localhost"):
        print("Warning: this page has no login. Anyone who can reach this address can upload packages.")
    app.run(host=host, port=port, threaded=True)
