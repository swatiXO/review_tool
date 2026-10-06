"""Measure how well a model recognises questions, before anyone relies on it.

Two kinds of ground truth:
  * the rule-based parser's output on documents it handles cleanly ("silver" labels:
    no hand work, but only covers formats the parser already knows)
  * hand-counted question numbers per file, from a --labels JSON file such as
    {"Grade-6/Chapter-1/Chapter-1-Assessment.docx": 7}. This is the one that measures
    the documents the fallback is actually for.
"""
import json
import shutil
import tempfile
from pathlib import Path

from . import fallback, ingest
from .docx_model import parse_docx
from .questions import numbering_is_regular, parse_questions


def _spread(items, limit):
    if limit >= len(items):
        return items
    step = len(items) / limit
    return [items[int(i * step)] for i in range(limit)]


def run(zip_path, model, limit=10, json_path=None, labels_path=None):
    profile = ingest.load_profile()
    labels = json.loads(Path(labels_path).read_text(encoding="utf8")) if labels_path else {}
    work = Path(tempfile.mkdtemp(prefix="course-review-eval-"))
    rows = []
    try:
        dest = ingest.safe_extract(zip_path, work)
        pkg = ingest.classify(ingest.find_root(dest), profile)
        docs = [d for d in pkg.docs if d.ext == "docx" and d.doc_type in ("chapter_exam", "worksheet", "pop_quiz")]
        docs.sort(key=lambda d: d.rel)
        targets = []
        for d in docs:
            if d.rel in labels:
                targets.append((d, int(labels[d.rel]), "hand"))
                continue
            info = parse_docx(d.abs)
            qs = parse_questions(info, profile)
            if d.doc_type != "pop_quiz" and qs and numbering_is_regular(qs):
                targets.append((d, qs, "silver"))
        targets = _spread(targets, limit)
        print(f"Model {model.client.model} at {model.client.url}: evaluating {len(targets)} document(s)\n")
        print(f"{'document':58} {'truth':>6} {'model':>6} {'prec':>5} {'recall':>6} {'secs':>5}  rejected")
        for d, truth, kind in targets:
            info = parse_docx(d.abs)
            ex = fallback.extract(info, model.client, cache=None)
            got = set(ex.starts)
            if kind == "silver":
                # a model start is right if it falls inside a real question's block; each block counts once
                hit_blocks = {i for i, q in enumerate(truth) if any(q.start <= s < q.end for s in got)}
                good_starts = {s for s in got if any(q.start <= s < q.end for q in truth)}
                prec = len(good_starts) / len(got) if got else 0.0
                rec = len(hit_blocks) / len(truth) if truth else 0.0
                n_truth = len(truth)
            else:
                n_truth, prec, rec = truth, None, None
            rows.append({"doc": d.rel, "kind": kind, "truth": n_truth, "model": len(got), "precision": prec,
                         "recall": rec, "count_match": n_truth == len(got), "seconds": round(ex.seconds, 1),
                         "proposed": ex.proposed, "rejected": ex.rejected, "usable": ex.usable, "note": ex.note})
            p = f"{prec:.2f}" if prec is not None else "-"
            r = f"{rec:.2f}" if rec is not None else "-"
            print(f"{d.rel.split('/')[-1][:58]:58} {n_truth:>6} {len(got):>6} {p:>5} {r:>6} {ex.seconds:>5.0f}  "
                  f"{ex.rejected}/{ex.proposed}{'' if ex.usable else '  UNUSABLE: ' + ex.note}")
    finally:
        shutil.rmtree(work, ignore_errors=True)
    if rows:
        ok = sum(1 for r in rows if r["count_match"])
        usable = sum(1 for r in rows if r["usable"])
        sil = [r for r in rows if r["precision"] is not None]
        print(f"\nCount matches truth: {ok}/{len(rows)}   usable extractions: {usable}/{len(rows)}   "
              f"avg {sum(r['seconds'] for r in rows) / len(rows):.0f}s per document")
        if sil:
            print(f"Silver labels: mean precision {sum(r['precision'] for r in sil) / len(sil):.2f}, "
                  f"mean recall {sum(r['recall'] for r in sil) / len(sil):.2f}")
        print("Silver labels come from the rule-based parser, so they only cover formats it already knows. "
              "Use --labels with hand-counted files to measure the documents the fallback is for.")
    if json_path:
        Path(json_path).write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf8")
    return rows
