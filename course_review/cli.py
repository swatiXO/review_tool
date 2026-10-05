"""Command line: python -m course_review.cli review PACKAGE.zip --checklist Course-Review-Checklist.xlsx --out out/"""
import argparse
import shutil
import sys
import tempfile
import time
from pathlib import Path

from . import engine, ingest, report, workbook


def review(zip_path, checklist, out_dir, profile_path=None, keep=False):
    t0 = time.time()
    profile = ingest.load_profile(profile_path)
    rules, layout = workbook.load_rules(checklist)
    work = Path(tempfile.mkdtemp(prefix="course-review-"))
    try:
        dest = ingest.safe_extract(zip_path, work)
        pkg = ingest.classify(ingest.find_root(dest), profile)
        res = engine.run(pkg, profile, rules, layout)
        xlsx = report.write_outputs(pkg, res, rules, layout, checklist, out_dir, Path(zip_path).name)
    finally:
        if not keep:
            shutil.rmtree(work, ignore_errors=True)
    return pkg, res, xlsx, time.time() - t0


def main(argv=None):
    ap = argparse.ArgumentParser(prog="course_review")
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("review", help="review a course package zip against the checklist workbook")
    r.add_argument("package", help="the .zip to review")
    r.add_argument("--checklist", required=True, help="Course-Review-Checklist.xlsx (the rules)")
    r.add_argument("--out", default="review-output", help="output folder")
    r.add_argument("--profile", help="package profile YAML (defaults to the built-in one)")
    a = ap.parse_args(argv)
    pkg, res, xlsx, secs = review(a.package, a.checklist, a.out, a.profile)
    from collections import Counter
    c = Counter(f.status for f in res.findings)
    print(f"Reviewed {len(pkg.docs)} documents in {secs:.0f}s: {c['fail']} fails, {c['pass']} passes, {c['needs_review']} need review, {c['na']} n/a")
    print(f"Wrote {xlsx}")
    print(f"      {Path(a.out) / 'report.html'}")
    print(f"      {Path(a.out) / 'review.json'}")
    if res.errors:
        print(f"{len(res.errors)} file(s) could not be read; see the report", file=sys.stderr)


if __name__ == "__main__":
    main()
