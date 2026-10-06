"""Command line.

  python -m course_review.cli review PACKAGE.zip --out out/      (the team's checklist is built in; --checklist overrides it)
  python -m course_review.cli review ... --model-fallback suggest --model-url https://abcd.ngrok-free.dev
  python -m course_review.cli check-model --model-url https://abcd.ngrok-free.dev
  python -m course_review.cli eval-model PACKAGE.zip --model-url https://abcd.ngrok-free.dev
"""
import argparse
import os
import shutil
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

from . import annotate, engine, fallback, ingest, report, workbook
from .llm import OllamaClient


def make_model(mode, url=None, model=None, cache_dir=".course_review_cache", judge=False, codes=None):
    """A ModelConfig, or None when mode is 'off' and no model checks are requested. Raises SystemExit
    with a clear message if the server cannot be reached, so a run never silently loses its model."""
    if judge and mode == "off":
        mode = "suggest"
    if mode == "off":
        return None
    client = OllamaClient(url=url, model=model)
    ok, msg = client.available()
    if not ok:
        raise SystemExit(f"Model fallback requested but the model server is not usable: {msg}")
    return fallback.ModelConfig(client=client, cache=fallback.Cache(cache_dir), mode=mode, judge=judge, codes=codes)


MARKED_UP = "Marked-up-documents.zip"


# The team's Course Review Checklist ships with the tool; --checklist (or COURSE_REVIEW_CHECKLIST) overrides it.
BUILTIN_CHECKLIST = str(Path(__file__).parent / "data" / "Course-Review-Checklist.xlsx")


def default_checklist():
    return os.environ.get("COURSE_REVIEW_CHECKLIST") or BUILTIN_CHECKLIST


def review(zip_path, checklist, out_dir, profile_path=None, keep=False, model=None, book=None, progress=None):
    t0 = time.time()
    profile = ingest.load_profile(profile_path)
    rules, layout = workbook.load_rules(checklist)
    work = Path(tempfile.mkdtemp(prefix="course-review-"))
    try:
        say = progress or (lambda *a: None)
        say("Unpacking the zip", 0, 0)
        dest = ingest.safe_extract(zip_path, work)
        say("Sorting the files", 0, 0)
        pkg = ingest.classify(ingest.find_root(dest), profile)
        res = engine.run(pkg, profile, rules, layout, model=model, book=book, progress=progress)
        say("Writing the workbook and report", 0, 0)
        xlsx = report.write_outputs(pkg, res, rules, layout, checklist, out_dir, Path(zip_path).name, model=model)
        say("Marking up the documents", 0, 0)
        annotate.annotate_package(pkg, res.findings, Path(out_dir) / MARKED_UP, rules)
    finally:
        if not keep:
            shutil.rmtree(work, ignore_errors=True)
    return pkg, res, xlsx, time.time() - t0


def _model_args(p):
    p.add_argument("--model-url", help="Ollama base URL, e.g. your ngrok link (or set OLLAMA_URL)")
    p.add_argument("--model", help="model name (or set OLLAMA_MODEL; default qwen3:14b)")
    p.add_argument("--cache-dir", default=".course_review_cache", help="where model answers are cached")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="course_review")
    sub = ap.add_subparsers(dest="cmd", required=True)

    r = sub.add_parser("review", help="review a course package zip against the checklist workbook")
    r.add_argument("package", help="the .zip to review")
    r.add_argument("--checklist", default=None, help="checklist workbook (default: the built-in Course-Review-Checklist.xlsx)")
    r.add_argument("--out", default="review-output", help="output folder")
    r.add_argument("--profile", help="package profile YAML (defaults to the built-in one)")
    r.add_argument("--model-fallback", choices=["off", "suggest", "decide"], default="off",
                   help="use a local model for question formats the parser does not recognise. "
                        "suggest: model answers are shown as needs-review suggestions; "
                        "decide: model-based Fails are also written to the workbook")
    r.add_argument("--model-checks", nargs="?", const="all", metavar="CODES",
                   help="also run the model-assisted judgement checks (LP1, LP2, LP6, LP7, LP8, LP9, FG1, FG3, FG6, CE2, CE3, "
                        "and SLO mapping for PQ5/WS5). Optionally a comma list, e.g. LP1,FG1. Results are suggestions only.")
    r.add_argument("--book-index", help="folder made by index-book; lets LP2, LP8, CE2 and CE3 consult the textbook")
    _model_args(r)

    b = sub.add_parser("index-book", help="OCR the textbook PDF into a searchable page index (resumable)")
    b.add_argument("pdf")
    b.add_argument("--out", default="book_index", help="folder for the index")
    b.add_argument("--engine", choices=["tesseract", "vlm", "auto"], default="tesseract")
    b.add_argument("--pages", help="e.g. 1-20,40 (default: all pages)")
    b.add_argument("--dpi", type=int, default=200)
    b.add_argument("--lang", default="urd+eng", help="Tesseract languages")
    _model_args(b)

    sv = sub.add_parser("serve", help="run the web page (upload a zip, download the results)")
    sv.add_argument("--checklist", default=None, help="checklist workbook (default: the built-in one, or COURSE_REVIEW_CHECKLIST)")
    sv.add_argument("--host", default="127.0.0.1", help="127.0.0.1 keeps it on this computer (default)")
    sv.add_argument("--port", type=int, default=8080)
    sv.add_argument("--jobs-dir", default="web_jobs")
    sv.add_argument("--books-dir", default="book_indexes", help="folder holding textbook indexes made with index-book")

    lf = sub.add_parser("learn-format", help="learn how a document numbers its questions, once, then use it without a model")
    lf.add_argument("document", help="a .docx whose questions the parser does not recognise")
    lf.add_argument("--save", required=True, help="profile file to add the learned pattern to (created if missing)")
    lf.add_argument("--pattern", help="give the pattern yourself instead of asking the model")
    lf.add_argument("--yes", action="store_true", help="save without asking for confirmation")
    _model_args(lf)

    c = sub.add_parser("check-model", help="check the model server is reachable and has the model")
    _model_args(c)

    e = sub.add_parser("eval-model", help="measure the model's question recognition against the rule-based parser")
    e.add_argument("package")
    e.add_argument("--limit", type=int, default=10, help="documents to evaluate")
    e.add_argument("--json", help="write the full evaluation to this file")
    e.add_argument("--labels", help="JSON of hand-counted question numbers: {\"path/in/zip.docx\": 7}")
    _model_args(e)

    a = ap.parse_args(argv)

    if a.cmd == "serve":
        from . import web
        web.serve(a.checklist or default_checklist(), a.host, a.port, a.jobs_dir, a.books_dir)
        return

    if a.cmd == "learn-format":
        from . import formats
        from .docx_model import parse_docx
        info = parse_docx(a.document)
        profile = ingest.load_profile(a.save if Path(a.save).exists() else None)
        try:
            if a.pattern:
                _, matches, problems = formats.validate(a.pattern, info)
                found = {"pattern": a.pattern, "explanation": "given on the command line", "matches": matches, "problems": problems}
            else:
                model = make_model("suggest", a.model_url, a.model, a.cache_dir)
                found = formats.propose(model.client, info)
        except formats.FormatError as e:
            raise SystemExit(f"No usable pattern: {e}")
        print(f"Pattern: {found['pattern']}\n  {found['explanation']}\nIt matches {len(found['matches'])} paragraph(s):")
        for idx, num, text in found["matches"][:8]:
            print(f"  question {num}: {text[:70]}")
        if found["problems"]:
            raise SystemExit("Not saved. Problems: " + "; ".join(found["problems"]))
        ok = a.yes or input("Save this pattern to the profile? [y/N] ").strip().lower().startswith("y")
        if ok:
            added = formats.save_pattern(a.save, found["pattern"], note=Path(a.document).name)
            print(("Saved to " if added else "Already in ") + a.save + ". Use it with: review ... --profile " + a.save)
        else:
            print("Not saved.")
        return

    if a.cmd == "index-book":
        from . import book as bk
        client = None
        if a.engine == "vlm":
            client = OllamaClient(url=a.model_url, model=a.model or "qwen3-vl:8b", timeout=1800)
            ok, msg = client.available()
            if not ok:
                raise SystemExit(f"The vision model server is not usable: {msg}")
        done = bk.build_index(a.pdf, a.out, a.engine, a.pages, a.dpi, a.lang, client,
                              progress=lambda p, n: print(f"  page {p} done ({n} requested)", flush=True))
        print(f"Indexed {len(done)} page(s) into {a.out}")
        return

    if a.cmd == "check-model":
        ok, msg = OllamaClient(url=a.model_url, model=a.model).available()
        print(("OK: " if ok else "NOT USABLE: ") + msg)
        sys.exit(0 if ok else 1)

    if a.cmd == "eval-model":
        from . import evaluate
        model = make_model("suggest", a.model_url, a.model, a.cache_dir)
        evaluate.run(a.package, model, a.limit, a.json, a.labels)
        return

    judge = a.model_checks is not None
    codes = None if (a.model_checks in (None, "all")) else {c.strip().upper() for c in a.model_checks.split(",") if c.strip()}
    model = make_model(a.model_fallback, a.model_url, a.model, a.cache_dir, judge=judge, codes=codes)
    book_index = None
    if a.book_index:
        from .book import BookIndex
        book_index = BookIndex.load(a.book_index)
        print(f"Textbook index: {len(book_index.pages)} page(s) from {a.book_index}")
    pkg, res, xlsx, secs = review(a.package, a.checklist or default_checklist(), a.out, a.profile, model=model, book=book_index)
    c = Counter(f.status for f in res.findings)
    print(f"Reviewed {len(pkg.docs)} documents in {secs:.0f}s: {c['fail']} fails, {c['pass']} passes, "
          f"{c['needs_review']} need review, {c['na']} n/a")
    if model is not None:
        used = [s for s in res.model_stats if s.get("usable")]
        print(f"Model fallback ({model.mode}): {len(used)} region(s) recovered, "
              f"{len(res.model_stats) - len(used)} not recovered, {model.client.calls} model call(s), {model.client.seconds:.0f}s")
    print(f"Wrote {xlsx}")
    print(f"      {Path(a.out) / 'report.html'}")
    print(f"      {Path(a.out) / 'review.json'}")
    print(f"      {Path(a.out) / MARKED_UP}  (the documents with problems highlighted and commented)")
    if res.errors:
        print(f"{len(res.errors)} file(s) could not be read; see the report", file=sys.stderr)


if __name__ == "__main__":
    main()
