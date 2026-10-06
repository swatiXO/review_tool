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
:root{--bg:#f4f6f8;--fg:#16212b;--mut:#5a6774;--line:#d9e0e6;--card:#fff;--soft:#eef2f5;--accent:#0b6b73;--accent-fg:#fff;
 --fail:#b42318;--fail-bg:#fdecea;--rev:#8a5a00;--rev-bg:#fdf3e1;--pass:#146c43;--pass-bg:#e6f4ec}
@media (prefers-color-scheme:dark){:root{--bg:#0e1519;--fg:#e3eaf0;--mut:#98a6b2;--line:#26323a;--card:#151e24;--soft:#1b252c;--accent:#4cc3cc;--accent-fg:#04252a;
 --fail:#ff8a80;--fail-bg:#3a1d1b;--rev:#e8b062;--rev-bg:#35290f;--pass:#5fd3a0;--pass-bg:#11301f}}
*{box-sizing:border-box}
body{background:var(--bg);color:var(--fg);font:15px/1.55 system-ui,"Segoe UI",sans-serif;margin:0}
.top{background:var(--card);border-bottom:1px solid var(--line)}
.top div{max-width:60rem;margin:auto;padding:14px 16px;display:flex;align-items:baseline;gap:12px;flex-wrap:wrap}
.top a.brand{font-weight:700;font-size:1.15rem;color:var(--fg);text-decoration:none}
.top span{color:var(--mut);font-size:.9rem}
main{max-width:60rem;margin:auto;padding:24px 16px 48px;display:flex;flex-direction:column;gap:20px}
h1{font-size:1.4rem;margin:0;overflow-wrap:anywhere}h2{font-size:1.05rem;margin:0}p{margin:0}a{color:var(--accent)}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px;display:flex;flex-direction:column;gap:14px;min-width:0}
.head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap}
label{display:flex;flex-direction:column;gap:4px;font-weight:600}label span{font-weight:400;color:var(--mut);font-size:.88rem}
input[type=text],select,input.plain{font:inherit;padding:8px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--fg);width:100%}
button,.btn{font:inherit;font-weight:600;padding:10px 18px;border-radius:7px;border:0;background:var(--accent);color:var(--accent-fg);cursor:pointer;text-decoration:none;display:inline-block}
button:focus-visible,.btn:focus-visible,a:focus-visible,.drop:focus-within{outline:3px solid var(--accent);outline-offset:2px}
button.quiet{background:none;color:var(--mut);font-weight:400;text-decoration:underline;padding:4px 0}
.drop{position:relative;border:2px dashed var(--line);border-radius:10px;padding:28px 16px;text-align:center;background:var(--soft);transition:border-color .15s}
.drop.over,.drop:hover{border-color:var(--accent)}
.drop input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%}
.drop b{display:block;font-size:1.05rem}.drop .mut{font-size:.9rem}
.drop .chosen{display:none;margin-top:8px;font-weight:600;color:var(--accent);overflow-wrap:anywhere}.drop.has .chosen{display:block}
.row{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.mut{color:var(--mut)}.small{font-size:.88rem}
.scroll{overflow-x:auto}
table{border-collapse:collapse;width:100%}th,td{text-align:left;padding:8px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-size:.74rem;text-transform:uppercase;letter-spacing:.05em;color:var(--mut);font-weight:600}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
tr:last-child td{border-bottom:0}
code{font:600 .85rem ui-monospace,Consolas,monospace;background:var(--soft);padding:1px 6px;border-radius:4px;white-space:nowrap}
.pill{display:inline-block;font-size:.8rem;font-weight:600;padding:1px 9px;border-radius:999px;white-space:nowrap;margin:1px 2px 1px 0}
.p-fail{background:var(--fail-bg);color:var(--fail)}.p-rev{background:var(--rev-bg);color:var(--rev)}.p-pass{background:var(--pass-bg);color:var(--pass)}
.p-run{background:var(--soft);color:var(--mut)}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(9.5rem,1fr));gap:12px}
.tile{background:var(--card);border:1px solid var(--line);border-left-width:5px;border-radius:10px;padding:12px 14px}
.tile b{display:block;font-size:1.8rem;line-height:1.2;font-variant-numeric:tabular-nums}.tile span{color:var(--mut);font-size:.88rem}
.t-fail{border-left-color:var(--fail)}.t-fail b{color:var(--fail)}.t-rev{border-left-color:var(--rev)}.t-rev b{color:var(--rev)}
.t-pass{border-left-color:var(--pass)}.t-pass b{color:var(--pass)}.t-doc{border-left-color:var(--accent)}
.dl{display:grid;grid-template-columns:repeat(auto-fit,minmax(13rem,1fr));gap:12px}
.dl a{display:flex;flex-direction:column;gap:2px;padding:12px 14px;border:1px solid var(--line);border-radius:8px;text-decoration:none;color:var(--fg)}
.dl a:hover{border-color:var(--accent)}.dl a b{color:var(--accent)}.dl a span{color:var(--mut);font-size:.86rem}
.dl a.main{background:var(--accent);border-color:var(--accent)}.dl a.main b,.dl a.main span{color:var(--accent-fg)}
.bar{height:10px;background:var(--line);border-radius:5px;overflow:hidden}.bar i{display:block;height:100%;background:var(--accent);width:0;transition:width .3s}
.items{display:flex;flex-direction:column}
.item{display:flex;flex-wrap:wrap;gap:6px 14px;padding:10px 0;border-bottom:1px solid var(--line);align-items:flex-start}
.item:last-child{border-bottom:0}.item>div{flex:1 1 18rem;min-width:0;overflow-wrap:anywhere}
.item .cat{display:block;color:var(--mut);font-size:.78rem;text-transform:uppercase;letter-spacing:.04em}
.item .nums{flex:0 0 auto;display:flex;gap:4px;align-items:center}
.alert{border-left:5px solid var(--fail);background:var(--fail-bg);padding:12px 14px;border-radius:8px}
.legend{display:flex;gap:16px;flex-wrap:wrap;font-size:.88rem;color:var(--mut)}
.sw{display:inline-block;width:.9em;height:.9em;border-radius:3px;vertical-align:-1px;margin-right:5px}
details summary{cursor:pointer;color:var(--mut)}
.steps{display:flex;gap:8px;flex-wrap:wrap;font-size:.88rem;color:var(--mut)}.steps b{color:var(--fg)}
@media (prefers-reduced-motion:reduce){.bar i,.drop{transition:none}}
</style></head><body>
<header class="top"><div><a class="brand" href="{{ url_for('home') }}">Course Review</a>
<span>Checks a course package against the Course Review Checklist and fixes the formatting.</span></div></header>
<main>{{ body|safe }}</main></body></html>"""

HOME = """
<form class="card" method="post" action="{{ url_for('create_job') }}" enctype="multipart/form-data">
  <div class="head"><h2>Review a package</h2>
  <div class="steps"><span><b>1</b> Upload the zip</span><span>&rarr; <b>2</b> Wait for the review</span><span>&rarr; <b>3</b> Download the results</span></div></div>
  <div class="drop" id="drop">
    <input type="file" name="package" accept=".zip" required aria-label="Course package (.zip)">
    <b>Drop the course package (.zip) here</b>
    <span class="mut">or click to choose it</span>
    <span class="chosen" id="chosen"></span>
  </div>
  <label>Use a language model<span>Results from the model are always suggestions for a reviewer, never a Pass or Fail.</span>
    <select name="model_mode">
      <option value="off">No, rules only (fast)</option>
      <option value="fallback">For question formats the rules do not recognise</option>
      <option value="checks">Also for content checks (slow)</option>
    </select></label>
  <details><summary>More options: checklist, model server, textbook</summary>
    <div class="card" style="border:0;padding:14px 0 0">
      <label>Checklist workbook<span>{% if default_checklist %}Leave empty to use the built-in <b>{{ default_checklist }}</b>.{% else %}Required: no default checklist is configured.{% endif %}</span>
        <input class="plain" type="file" name="checklist" accept=".xlsx" {% if not default_checklist %}required{% endif %}></label>
      <label>Ollama address<span>For example your ngrok link.</span><input type="text" name="model_url" value="{{ model_url }}" placeholder="http://localhost:11434"></label>
      <label>Model name<input type="text" name="model_name" value="{{ model_name }}" placeholder="qwen3:14b"></label>
      <label>Textbook index<span>Made with the index-book command.</span>
        <select name="book_index"><option value="">None</option>{% for b in books %}<option>{{ b }}</option>{% endfor %}</select></label>
    </div></details>
  <div class="row"><button type="submit">Start review</button>
  <span class="mut small">{% if busy %}{{ busy }} {% endif %}Without a model a full package takes about a minute.</span></div>
</form>
{% if jobs %}<div class="card"><h2>Earlier reviews</h2><div class="scroll"><table>
<tr><th>Package<th>Started<th>Result</tr>
{% for j in jobs %}<tr><td><a href="{{ url_for('job_page', job_id=j.id) }}">{{ j.name }}</a>
<td class="mut" style="white-space:nowrap">{{ j.created }}
<td>{% if j.state == 'done' and j.counts %}<span class="pill p-fail">{{ j.counts.get('fail', 0) }} fail</span><span class="pill p-rev">{{ j.counts.get('needs_review', 0) }} to check</span><span class="pill p-pass">{{ j.counts.get('pass', 0) }} pass</span>
{% elif j.state == 'failed' %}<span class="pill p-fail">Could not finish</span>
{% elif j.state in ('queued', 'running') %}<span class="pill p-run">{{ 'Running' if j.state == 'running' else 'Waiting' }}&hellip;</span>
{% else %}<span class="mut">{{ j.summary }}</span>{% endif %}</tr>{% endfor %}
</table></div></div>{% endif %}
<script>
const drop=document.getElementById('drop'),inp=drop.querySelector('input'),chosen=document.getElementById('chosen');
inp.addEventListener('change',()=>{const f=inp.files[0];chosen.textContent=f?f.name+' ('+(f.size/1048576).toFixed(1)+' MB)':'';drop.classList.toggle('has',!!f)});
['dragenter','dragover'].forEach(e=>drop.addEventListener(e,()=>drop.classList.add('over')));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,()=>drop.classList.remove('over')));
</script>
"""

JOB = """
<div class="card"><div class="head">
  <div><h1>{{ job.name }}</h1>
  <p class="mut small">Started {{ job.created }}{% if job.seconds != '' %} · took {{ job.seconds }} s{% endif %}{% if job.options %} · {{ job.options }}{% endif %}</p></div>
  <form method="post" action="{{ url_for('delete_job', job_id=job.id) }}" onsubmit="return confirm('Delete this review and all its files?')">
    <button class="quiet" type="submit">Delete this review</button></form>
</div>
{% if job.state in ('queued','running') %}
  <p id="stage"><b>{{ job.stage or 'Waiting to start' }}</b></p><div class="bar" role="progressbar" aria-label="Progress"><i id="fill"></i></div>
  <p class="mut small">This page updates by itself. You can leave it and come back later from the home page.</p>
{% elif job.state == 'failed' %}
  <div class="alert"><p><b>The review could not finish.</b></p><p>{{ job.error }}</p></div>
  <p><a href="{{ url_for('home') }}">Start a new review</a></p>
{% endif %}
</div>
{% if job.state == 'done' %}
<div class="tiles">
  <div class="tile t-fail"><b>{{ counts.get('fail', 0) }}</b><span>fail the checklist</span></div>
  <div class="tile t-rev"><b>{{ counts.get('needs_review', 0) }}</b><span>need a reviewer</span></div>
  <div class="tile t-pass"><b>{{ counts.get('pass', 0) }}</b><span>pass</span></div>
  <div class="tile t-doc"><b>{{ job.documents or '' }}</b><span>documents reviewed</span></div>
</div>
<div class="card"><h2>Downloads</h2>
  <div class="dl">
    <a class="main" href="{{ url_for('download', job_id=job.id, name='Marked-up-documents.zip') }}"><b>Download fixed documents</b>
      <span>The package with formatting fixed and comments on what still needs a person (.zip)</span></a>
    <a href="{{ url_for('download', job_id=job.id, name='Course-Review-Checklist-filled.xlsx') }}"><b>Filled workbook</b>
      <span>The checklist, one row per lesson and chapter (.xlsx)</span></a>
    <a href="{{ url_for('download', job_id=job.id, name='report.html') }}" target="_blank" rel="noopener"><b>Open the report</b>
      <span>Every result with its evidence, in a new tab</span></a>
    <a href="{{ url_for('download', job_id=job.id, name='review.json') }}"><b>JSON</b><span>All findings, for other tools</span></a>
  </div>
  <div class="legend"><span><i class="sw" style="background:#ff0000"></i>Red: fails the checklist</span>
    <span><i class="sw" style="background:#00e5e5"></i>Turquoise: model suggestion, a reviewer decides</span>
    <span><i class="sw" style="background:#00e000"></i>Green: passes</span></div>
  <p class="mut small">Formatting with one right answer (page setup, fonts, sizes, digits, heading numbers, captions, bullets, colour,
  footers, file names) is fixed in the copies; the first comment in each file lists the changes. A blank cell in the workbook means the tool did not decide it.</p>
</div>
{% if codes %}<div class="card"><h2>What to look at first</h2>
  <p class="mut small">Checklist rules with problems in this package, most failures first. Formatting fails are usually already fixed in the downloaded copies.</p>
  <div class="items">
  {% for c in codes %}<div class="item"><div><code>{{ c.code }}</code> {% if c.category %}<span class="cat">{{ c.category }}</span>{% endif %}{{ c.title }}
    {% if c.example %}<div class="mut small">e.g. {{ c.example }}</div>{% endif %}</div>
    <span class="nums">{% if c.fail %}<span class="pill p-fail">{{ c.fail }} fail</span>{% endif %}{% if c.review %}<span class="pill p-rev">{{ c.review }} to check</span>{% endif %}</span></div>{% endfor %}
  </div>
  {% if more_codes %}<p class="mut small">and {{ more_codes }} more rule(s); see the report for all of them.</p>{% endif %}
</div>{% endif %}
{% if docs %}<div class="card"><h2>By document</h2>
  <div class="items">
  {% for d in docs %}<div class="item"><div>{{ d.name }}<div class="mut small">{{ d.folder }}</div></div>
    <span class="nums">{% if d.fail %}<span class="pill p-fail">{{ d.fail }} fail</span>{% endif %}{% if d.review %}<span class="pill p-rev">{{ d.review }} to check</span>{% endif %}</span></div>{% endfor %}
  </div>
  <p class="mut small">Results that belong to no single file (a missing document, subject-wide coverage) are in the report and in REVIEW-NOTES.txt inside the zip.</p>
</div>{% endif %}
{% endif %}
{% if job.state in ('queued','running') %}<script>
const bar=document.getElementById('fill'),stage=document.getElementById('stage');
async function tick(){try{const r=await fetch('{{ url_for("job_status", job_id=job.id) }}');const s=await r.json();
 if(s.state==='done'||s.state==='failed'){location.reload();return}
 stage.firstChild.textContent=s.stage||'Working';if(s.total){bar.style.width=Math.round(100*s.done/s.total)+'%';bar.style.opacity=1}else{bar.style.width='100%';bar.style.opacity='.35'}}catch(e){}
 setTimeout(tick,1500)}
tick();</script>{% endif %}
"""


def _overview(res, rules):
    """What the job page shows: rules with problems (most fails first) and per-document counts."""
    from .checks.guidelines import GUIDE_RULES
    from .checks.registry import AUTOMATION
    from .models import FAIL, REVIEW
    by_code, by_doc = {}, {}
    for f in res.findings:
        if f.status not in (FAIL, REVIEW):
            continue
        c = by_code.setdefault(f.code, {"code": f.code, "fail": 0, "review": 0, "example": ""})
        c["fail" if f.status == FAIL else "review"] += 1
        if not c["example"] or (f.status == FAIL and c["fail"] == 1):
            c["example"] = f.message or ""
        if f.doc:
            d = by_doc.setdefault(f.doc, {"name": f.doc.rsplit("/", 1)[-1], "folder": f.doc.rsplit("/", 1)[0] if "/" in f.doc else "",
                                          "fail": 0, "review": 0})
            d["fail" if f.status == FAIL else "review"] += 1
    for c in by_code.values():
        r = rules.get(c["code"])
        title = (r.text if r is not None else "") or GUIDE_RULES.get(c["code"], "") or AUTOMATION.get(c["code"], ("", ""))[1]
        m = re.match(r"\[([^\]]+)\]\s*(.*)", title)          # '[Naming and versioning] File name follows ...'
        c["category"], title = (m.group(1), m.group(2)) if m else ("", title)
        c["title"] = title if len(title) <= 140 else title[:137].rstrip() + "..."
        if len(c["example"]) >= 140:
            c["example"] = c["example"][:137].rstrip() + "..."
    codes = sorted(by_code.values(), key=lambda c: (-c["fail"], -c["review"], c["code"]))
    docs = sorted(by_doc.values(), key=lambda d: (-d["fail"], -d["review"], d["name"]))
    return codes, docs


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
            try:
                from .workbook import load_rules
                codes, docs = _overview(res, load_rules(opts["checklist"])[0])
            except Exception:                      # the overview is a convenience; the downloads are the result
                codes, docs = [], []
            update(d, state="done", stage="Done", seconds=round(time.time() - started), documents=len(pkg.docs),
                   summary=f"{len(pkg.docs)} documents: {c['fail']} fails, {c['pass']} passes, {c['needs_review']} for a reviewer",
                   counts=dict(c), codes=codes, docs=docs)
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
                listed.append({"id": d.name, "name": m.get("name", d.name), "created": m.get("created", ""), "state": m.get("state"),
                               "counts": m.get("counts") or {}, "summary": m.get("summary") or ""})
        names = sorted(p.name for p in books.iterdir() if (p / "pages").is_dir()) if books.is_dir() else []
        running = sum(1 for j in listed if j["state"] == "running")
        waiting = sum(1 for j in listed if j["state"] == "queued")
        busy = (f"{running} review running" + (f", {waiting} waiting" if waiting else "") + ": yours will start after them.") if running or waiting else ""
        return page("Course Review", HOME, jobs=listed, busy=busy, books=names, default_checklist=Path(default_checklist).name if default_checklist else None,
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
        meta = _read_meta(d) or {}
        codes = meta.get("codes") or []
        return page(meta.get("name") or "Course Review", JOB, job=_Obj(meta), counts=meta.get("counts") or {},
                    codes=codes[:15], more_codes=max(0, len(codes) - 15), docs=meta.get("docs") or [])

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
