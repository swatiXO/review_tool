"""Run every check and assemble the results.

Flow: classify files -> per-document formatting checks -> structure checks ->
question-document checks -> derived rules (LP10, FG8) -> subject roll-up -> grids
for the workbook. Findings are never dropped: a code the tool does not decide shows
up as 'not automated' in the summary.
"""
import re
from collections import Counter, defaultdict

from .checks import assessments, coverage, formatting, judgement, lessonplan, slides
from .checks.common import result
from .checks.registry import AUTOMATION
from .docx_model import parse_docx
from .models import FAIL, NA, PASS, REVIEW, Finding, LessonKey
from .pptx_model import parse_pptx

LP10_COMPONENTS = ["WE1", "WE2", "WE3", "WE4", "WE6", "WE7", "WE8"]
FG8_COMPONENTS = ["WE5", "WE6", "WE7", "WE23"]
LESSON_SHEETS = {  # sheet name -> output type in the workbook Key sheet
    "Lesson Plan": "Lesson Plan", "Facilitators Guide": "Facilitator's Guide", "Pop Quiz": "Pop Quiz",
    "Chapter Exam": "Chapter Exam", "Data Bank": "Data Bank",
}


class Context:
    def __init__(self, pkg, profile, rules, model=None, book=None):
        self.pkg, self.profile, self.rules = pkg, profile, rules
        self.book = book            # book.BookIndex or None (no textbook index)
        self.model = model          # fallback.ModelConfig or None (model fallback off)
        self.th = profile["thresholds"]
        self.section_labels = [l for s in profile["lesson_plan_sections"] for l in s["labels"]]
        self._cache, self.errors = {}, {}

    def _load(self, doc, parser):
        if doc.rel in self._cache:
            return self._cache[doc.rel]
        try:
            info = parser(doc.abs)
        except Exception as e:  # a corrupt file is a finding, not a crash
            self.errors[doc.rel] = f"{type(e).__name__}: {e}"
            info = None
        self._cache[doc.rel] = info
        return info

    def docx(self, doc):
        return self._load(doc, parse_docx)

    def pptx(self, doc):
        return self._load(doc, parse_pptx)


def _collapse_lessonless_chapters(pkg):
    """A chapter with no lesson folders is one unit: its files must share one key."""
    has_folder = defaultdict(bool)
    for d in pkg.docs:
        if d.folder_label and d.chapter is not None:
            has_folder[d.chapter] = True
    for d in pkg.docs:
        if d.scope == "lesson" and d.chapter is not None and not has_folder[d.chapter]:
            d.lesson, d.variant = None, ""


def _clean_title(label):
    t = re.sub(r"(?i)^lesson[-_ ]*\d+[-_ ]*", "", label or "")
    t = re.sub(r"(?i)^(?:chapter|chpater)[-_ ]*\d+[-_ ]*", "", t)
    t = re.sub(r"(?i)^part[-_ ]*\d+[-_ ]*", "", t)
    return t.replace("-", " ").replace("_", " ").strip()


def lesson_label(pkg, key: LessonKey):
    if key.lesson is None:
        return f"Chapter {key.chapter} (no lesson folders)"
    base = f"Chapter {key.chapter} - Lesson {key.lesson}{key.variant}"
    title = _clean_title(pkg.lesson_labels.get(key, ""))
    return f"{base} ({title})" if title else base


def cell_status(findings):
    """Collapse findings for one cell into 'pass' | 'fail' | 'na' | None (undecided)."""
    if not findings:
        return None
    if any(f.status == FAIL for f in findings):
        return "fail"
    if all(f.status == NA for f in findings):
        return "na"
    if all(f.status in (PASS, NA) and not f.partial for f in findings):
        return "pass"
    return None


def cell_note(findings):
    parts = []
    for f in findings:
        if f.status == FAIL:
            parts.append(f"{f.code}: {f.message}")
        elif f.status == PASS and f.partial:
            parts.append(f"{f.code} (auto, partial): {f.message}. " + " ".join(f.evidence[:1]))
        elif f.status == REVIEW:
            parts.append(f"{f.code} (reviewer): {f.message}")
    return " | ".join(parts)


class Results:
    def __init__(self):
        self.findings = []
        self.subject = {}          # code -> (status|None, note, [per-doc findings])
        self.grid = {}             # sheet -> [rows]
        self.inventory = {}
        self.errors = {}
        self.slo_map = []
        self.model_stats = []


def run(pkg, profile, rules, layout, model=None, book=None):
    _collapse_lessonless_chapters(pkg)
    ctx = Context(pkg, profile, rules, model, book)
    res = Results()
    F = res.findings
    formatted = [d for d in pkg.docs if d.doc_type in profile["formatted_types"] and not d.superseded and d.ext in ("docx", "pptx")]

    # 1. per-document formatting + structure
    doc_findings = defaultdict(list)    # code -> [Finding]
    for d in formatted:
        info = ctx.docx(d) if d.ext == "docx" else ctx.pptx(d)
        if info is None:
            continue
        table = formatting.DOCX_CHECKS if d.ext == "docx" else formatting.PPTX_CHECKS
        for code, fn in table.items():
            f = fn(ctx, d, info)
            f.doc = d.rel
            f.lesson = d.key if d.scope == "lesson" and d.chapter is not None else None
            F.append(f)
            doc_findings[code].append(f)
        if d.doc_type == "facilitator_guide" and d.ext == "pptx":
            for code, fn in slides.DOC_CHECKS.items():
                f = fn(ctx, d, info)
                f.doc, f.lesson = d.rel, d.key
                F.append(f)
        if d.doc_type == "lesson_plan" and d.ext == "docx":
            for code, fn in lessonplan.DOC_CHECKS.items():
                f = fn(ctx, d, info)
                f.doc, f.lesson = d.rel, d.key
                F.append(f)
            lp3 = next(f for f in F[::-1] if f.code == "LP3" and f.doc == d.rel)
            w14 = Finding("WE14", lp3.status, lp3.message, list(lp3.evidence), lesson=d.key, doc=d.rel, partial=lp3.partial)
            F.append(w14)
            doc_findings["WE14"].append(w14)

    # 2. question documents
    F.extend(assessments.pop_quiz_checks(ctx))
    for d in pkg.docs:
        if d.superseded or d.ext != "docx" or d.doc_type not in ("chapter_exam", "worksheet"):
            continue
        info = ctx.docx(d)
        if info is not None:
            F.extend(assessments.exam_checks(ctx, d, info, d.doc_type))
    db_findings, bank_items = assessments.data_bank_checks(ctx)
    F.extend(db_findings)
    dbs = assessments.dbs1(ctx, bank_items)
    F.append(dbs)
    doc_findings["DBS1"].append(dbs)

    cov_findings, res.slo_map = coverage.coverage_findings(ctx)
    F.extend(cov_findings)
    for f in cov_findings:
        if f.code in ("ST1", "ST2"):
            doc_findings[f.code].append(f)
    judgement.augment_coverage(ctx, F)
    F.extend(judgement.judgement_findings(ctx))

    # 3. missing artifacts per lesson / chapter
    lesson_keys = sorted({d.key for d in pkg.docs if d.scope == "lesson" and d.chapter is not None}, key=lambda k: k.sort_key())
    chapters = sorted({d.chapter for d in pkg.docs if d.chapter is not None})
    have = defaultdict(set)
    for d in pkg.docs:
        if not d.superseded:
            have[d.doc_type].add(d.key if d.scope == "lesson" else d.chapter)
    missing = []
    for k in lesson_keys:
        if k not in have["lesson_plan"]:
            F.append(result("LP3", FAIL, "No Lesson Plan file found for this lesson", lesson=k))
            missing.append((lesson_label(pkg, k), "Lesson Plan"))
        if k not in have["facilitator_guide"]:
            missing.append((lesson_label(pkg, k), "Facilitator Guide"))
        if k not in have["chapter_exam"]:
            F.append(result("CE1", FAIL, "No Chapter Exam (per-lesson assessment) file found for this lesson", lesson=k))
            missing.append((lesson_label(pkg, k), "Chapter Exam (per-lesson assessment)"))
    for ch in chapters:
        if ch not in have["worksheet"]:
            F.append(result("WS1", FAIL, "No Worksheet (per-chapter exam) file found for this chapter", chapter=ch))
            missing.append((f"Chapter {ch}", "Worksheet (per-chapter exam)"))

    # 4. derived LP10 / FG8
    def derive(code, components, doc_type):
        for d in [x for x in formatted if x.doc_type == doc_type]:
            comp = [f for f in F if f.doc == d.rel and f.code in components]
            if not comp:
                continue
            fails = [f for f in comp if f.status == FAIL]
            if fails:
                msg = "; ".join(f"{f.code}: {f.message}" for f in fails[:3])
                F.append(result(code, FAIL, msg, [f"{f.code}: {f.message}" for f in fails], lesson=d.key, doc=d.rel))
            elif all(f.status in (PASS, NA) and not f.partial for f in comp):
                F.append(result(code, PASS, "All formatting rules this item covers pass", lesson=d.key, doc=d.rel))
            else:
                F.append(result(code, REVIEW, "No failures; some parts need a reviewer: " +
                                ", ".join(f"{f.code}" for f in comp if f.status == REVIEW or f.partial), lesson=d.key, doc=d.rel))
    derive("LP10", LP10_COMPONENTS, "lesson_plan")
    derive("FG8", FG8_COMPONENTS, "facilitator_guide")

    # 5. subject roll-up
    langs = {info.language for d in formatted for info in [ctx._cache.get(d.rel)] if info is not None and info.language}
    for code in rules:
        if rules[code].scope != "Per subject":
            continue
        fs = doc_findings.get(code, [])
        if code == "WE9":
            if len(langs) <= 1:
                res.subject[code] = (None, f"Only {', '.join(sorted(langs)) or 'one'} language present, so no Urdu/English counterpart to compare", [])
            else:
                res.subject[code] = (None, "Both languages present; mirrored structure needs a reviewer", [])
            continue
        if not fs:
            res.subject[code] = (None, "Not automated" if code not in AUTOMATION or AUTOMATION[code][0] == "no" else "No applicable document", [])
            continue
        fails = [f for f in fs if f.status == FAIL]
        ok = [f for f in fs if f.status == PASS]
        na = [f for f in fs if f.status == NA]
        rev = [f for f in fs if f.status == REVIEW]
        package_level = all(f.doc is None for f in fs)   # a rule about the whole package, not one document
        if fails:
            if package_level:
                note = fails[0].message
            else:
                ex = "; ".join(f"{(f.doc or 'package').split('/')[-1]}: {f.message}" for f in fails[:2])
                note = f"{len(fails)} of {len(fs)} documents fail. {ex}"
            res.subject[code] = ("fail", note, fs)
        elif len(na) == len(fs):
            res.subject[code] = ("na", fs[0].message, fs)
        elif not rev and all(not f.partial for f in ok):
            res.subject[code] = ("pass", ok[0].message if package_level else f"All {len(ok)} applicable documents pass", fs)
        else:
            why = rev[0].message if rev else (ok[0].evidence[0] if ok and ok[0].evidence else "partly checked")
            res.subject[code] = (None, f"No failures in {len(fs)} documents; reviewer to confirm: {why}", fs)

    # 6. grids
    by = defaultdict(list)
    for f in F:
        by[(f.code, f.lesson, f.chapter)].append(f)
    for sheet, spec in layout.matrix.items():
        rows = []
        if spec["scope"] == "chapter":
            for ch in chapters:
                cells, notes = {}, []
                for code in spec["codes"]:
                    fs = [f for f in F if f.code == code and f.chapter == ch]
                    cells[code] = cell_status(fs)
                    n = cell_note(fs)
                    if n:
                        notes.append(n)
                rows.append({"label": f"Chapter {ch}" + (f" ({_clean_title(pkg.chapter_titles.get(ch, ''))})" if pkg.chapter_titles.get(ch) else ""),
                             "cells": cells, "note": " | ".join(notes), "key": ch})
        else:
            for k in lesson_keys:
                cells, notes = {}, []
                for code in spec["codes"]:
                    fs = by.get((code, k, None), [])
                    cells[code] = cell_status(fs)
                    n = cell_note(fs)
                    if n:
                        notes.append(n)
                rows.append({"label": lesson_label(pkg, k), "cells": cells, "note": " | ".join(notes), "key": k})
        res.grid[sheet] = rows

    res.inventory = {
        "documents": [{"path": d.rel, "type": d.doc_type, "chapter": d.chapter,
                       "lesson": None if d.lesson is None else f"{d.lesson}{d.variant}", "scope": d.scope,
                       "superseded": d.superseded, "notes": d.notes} for d in pkg.docs],
        "unclassified": pkg.unclassified, "missing": missing, "parse_errors": ctx.errors,
        "counts": dict(Counter(d.doc_type for d in pkg.docs)),
    }
    res.errors = ctx.errors
    res.model_stats = list(model.stats) if model is not None else []
    res.lesson_keys, res.chapters = lesson_keys, chapters
    return res
