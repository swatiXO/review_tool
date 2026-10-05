"""Command line.

  python -m course_review.cli review PACKAGE.zip --checklist Course-Review-Checklist.xlsx --out out/
  python -m course_review.cli review ... --model-fallback suggest --model-url https://abcd.ngrok-free.dev
  python -m course_review.cli check-model --model-url https://abcd.ngrok-free.dev
  python -m course_review.cli eval-model PACKAGE.zip --model-url https://abcd.ngrok-free.dev
"""
import argparse
import shutil
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

from . import engine, fallback, ingest, report, workbook
from .llm import OllamaClient


def make_model(mode, url=None, model=None, cache_dir=".course_review_cache"):
    """A ModelConfig, or None when mode is 'off'. Raises SystemExit with a clear message
    if the server cannot be reached, so a run never silently loses its model."""
    if mode == "off":
        return None
    client = OllamaClient(url=url, model=model)
    ok, msg = client.available()
    if not ok:
        raise SystemExit(f"Model fallback requested but the model server is not usable: {msg}")
    return fallback.ModelConfig(client=client, cache=fallback.Cache(cache_dir), mode=mode)


def review(zip_path, checklist, out_dir, profile_path=None, keep=False, model=None):
    t0 = time.time()
    profile = ingest.load_profile(profile_path)
    rules, layout = workbook.load_rules(checklist)
    work = Path(tempfile.mkdtemp(prefix="course-review-"))
    try:
        dest = ingest.safe_extract(zip_path, work)
        pkg = ingest.classify(ingest.find_root(dest), profile)
        res = engine.run(pkg, profile, rules, layout, model=model)
        xlsx = report.write_outputs(pkg, res, rules, layout, checklist, out_dir, Path(zip_path).name, model=model)
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
    r.add_argument("--checklist", required=True, help="Course-Review-Checklist.xlsx (the rules)")
    r.add_argument("--out", default="review-output", help="output folder")
    r.add_argument("--profile", help="package profile YAML (defaults to the built-in one)")
    r.add_argument("--model-fallback", choices=["off", "suggest", "decide"], default="off",
                   help="use a local model for question formats the parser does not recognise. "
                        "suggest: model answers are shown as needs-review suggestions; "
                        "decide: model-based Fails are also written to the workbook")
    _model_args(r)

    c = sub.add_parser("check-model", help="check the model server is reachable and has the model")
    _model_args(c)

    e = sub.add_parser("eval-model", help="measure the model's question recognition against the rule-based parser")
    e.add_argument("package")
    e.add_argument("--limit", type=int, default=10, help="documents to evaluate")
    e.add_argument("--json", help="write the full evaluation to this file")
    e.add_argument("--labels", help="JSON of hand-counted question numbers: {\"path/in/zip.docx\": 7}")
    _model_args(e)

    a = ap.parse_args(argv)

    if a.cmd == "check-model":
        ok, msg = OllamaClient(url=a.model_url, model=a.model).available()
        print(("OK: " if ok else "NOT USABLE: ") + msg)
        sys.exit(0 if ok else 1)

    if a.cmd == "eval-model":
        from . import evaluate
        model = make_model("suggest", a.model_url, a.model, a.cache_dir)
        evaluate.run(a.package, model, a.limit, a.json, a.labels)
        return

    model = make_model(a.model_fallback, a.model_url, a.model, a.cache_dir)
    pkg, res, xlsx, secs = review(a.package, a.checklist, a.out, a.profile, model=model)
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
    if res.errors:
        print(f"{len(res.errors)} file(s) could not be read; see the report", file=sys.stderr)


if __name__ == "__main__":
    main()
