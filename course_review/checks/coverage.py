"""SLO coverage: ST1, ST2 (whole subject), PQ5 (per lesson), WS4, WS5 (per chapter).

Coverage is computed from SLO tags written in the questions (see slo.tags_in). When the
questions carry no tags the result is needs-reviewer, not Fail: whether an SLO is tested
cannot be decided by code, and the model-assisted mapping in judgement.py can suggest it.
ST2 is the one literal rule ("zero coverage across the assessment types"), so with no tags
at all it reports every SLO as a gap, exactly as the workbook words it.
"""
from collections import Counter, defaultdict

from ..models import FAIL, NA, PASS, REVIEW, LessonKey
from ..questions import load_bank_items
from ..slo import load_slos, slos_for, tags_in
from .assessments import align_pop_quiz, get_questions
from .common import result

DOC_TYPES = ("pop_quiz", "chapter_exam", "worksheet", "data_bank")


def get_slos(ctx):
    if not hasattr(ctx, "_slos"):
        spec = next((d for d in ctx.pkg.docs if d.doc_type == "specification" and not d.superseded and d.ext == "docx"), None)
        ctx._slos = None
        if spec is not None:
            try:
                ctx._slos = load_slos(spec.abs) or None
            except Exception as e:
                ctx.errors[spec.rel] = f"{type(e).__name__}: {e}"
    return ctx._slos


def collect(ctx):
    """Tag evidence per assessment type.

    Returns dict with:
      hits[(chapter, lesson, index)][doc_type] -> number of questions tagged to that SLO
      lesson_q[(doc_type, LessonKey)] -> (questions, tagged_question_count, set(indexes))
      chapter_q[chapter] -> (questions, tagged_count, set((lesson, index)), per_lesson_counts)
    """
    slos = get_slos(ctx)
    hits = defaultdict(Counter)
    lesson_q, chapter_q = {}, {}

    mapping, _, problem = align_pop_quiz(ctx)
    for key, qs in mapping.items():
        lesson_slos = slos_for(slos, key)
        tagged, idx = 0, set()
        for q in qs:
            t = tags_in(q.block, lesson_slos) if lesson_slos else set()
            if t:
                tagged += 1
                idx |= t
                for n in t:
                    hits[(key.chapter, key.lesson, n)]["pop_quiz"] += 1
        lesson_q[("pop_quiz", key)] = (qs, tagged, idx)

    for d in ctx.pkg.docs:
        if d.superseded or d.ext != "docx" or d.doc_type not in ("chapter_exam", "worksheet"):
            continue
        info = ctx.docx(d)
        if info is None:
            continue
        qs, _, _ = get_questions(ctx, d, info)
        if d.doc_type == "chapter_exam":
            lesson_slos = slos_for(slos, d.key)
            tagged, idx = 0, set()
            for q in qs:
                t = tags_in(q.block, lesson_slos) if lesson_slos else set()
                if t:
                    tagged += 1
                    idx |= t
                    for n in t:
                        hits[(d.chapter, d.lesson, n)]["chapter_exam"] += 1
            lesson_q[("chapter_exam", d.key)] = (qs, tagged, idx)
        else:
            chapter_slos = [s for (ch, _), v in slos.items() if ch == d.chapter for s in v]
            tagged, idx, per_lesson = 0, set(), Counter()
            for q in qs:
                t = tags_in(q.block, None, chapter_slos)
                if t:
                    tagged += 1
                    idx |= t
                    for les, n in t:
                        hits[(d.chapter, les, n)]["worksheet"] += 1
                    for les in {les for les, _ in t}:
                        per_lesson[les] += 1
            chapter_q[d.chapter] = (qs, tagged, idx, per_lesson, d)

    bank = next((d for d in ctx.pkg.docs if d.doc_type == "data_bank" and not d.superseded and d.ext == "docx"), None)
    if bank is not None:
        for it in load_bank_items(bank.abs, ctx.profile):
            if it.chapter is None or "slo" not in it.fields:
                continue
            chapter_slos = [s for (ch, _), v in slos.items() if ch == it.chapter for s in v]
            for les, n in tags_in(it.fields["slo"], None, chapter_slos):
                hits[(it.chapter, les, n)]["data_bank"] += 1
    return {"hits": hits, "lesson_q": lesson_q, "chapter_q": chapter_q, "pq_problem": problem}


def slo_map_rows(slos, cov):
    rows = []
    for (ch, les), lst in sorted(slos.items()):
        for s in lst:
            h = cov["hits"].get((ch, les, s.index), Counter())
            rows.append({"chapter": ch, "lesson": les, "slo": s.index, "text": s.text, "bloom": s.bloom,
                         **{t: h.get(t, 0) for t in DOC_TYPES}, "total": sum(h.values())})
    return rows


def coverage_findings(ctx):
    """(findings, slo_map_rows). Empty when the package has no specification."""
    slos = get_slos(ctx)
    if not slos:
        return [result("ST1", REVIEW, "No Document of Specifications with SLO tables was found, so the coverage map cannot be built"),
                result("ST2", REVIEW, "No SLO list to compare the assessments against")], []
    cov = collect(ctx)
    rows = slo_map_rows(slos, cov)
    out = []
    total_slos = len(rows)
    out.append(result("ST1", PASS, f"Coverage map built: {total_slos} SLOs in {len(slos)} lessons checked against the tags in the "
                                   f"Pop Quiz, Chapter Exams, Worksheets and Data Bank (see the SLO Coverage section of the report)"))
    tagged_total = sum(tq for _, tq, *_ in cov["lesson_q"].values()) + sum(v[1] for v in cov["chapter_q"].values())
    gaps = [r for r in rows if r["total"] == 0]
    if tagged_total == 0 and not any(r["data_bank"] for r in rows):
        out.append(result("ST2", FAIL, f"No question in any assessment carries an SLO tag, so all {total_slos} SLOs have zero tag coverage",
                          [f"Ch{r['chapter']} L{r['lesson']} SLO {r['slo']}" for r in gaps[:12]]))
    elif gaps:
        out.append(result("ST2", FAIL, f"{len(gaps)} of {total_slos} SLOs have zero coverage by tag across all assessments",
                          [f"Ch{r['chapter']} L{r['lesson']} SLO {r['slo']}: {r['text'][:50]}" for r in gaps[:12]]))
    else:
        out.append(result("ST2", PASS, f"Every one of the {total_slos} SLOs is covered by at least one tagged question"))

    keys = sorted({d.key for d in ctx.pkg.docs if d.scope == "lesson" and d.chapter is not None and d.lesson is not None},
                  key=lambda k: k.sort_key())
    for key in keys:
        lesson_slos = slos_for(slos, key)
        entry = cov["lesson_q"].get(("pop_quiz", key))
        if cov["pq_problem"] or entry is None:
            out.append(result("PQ5", REVIEW, cov["pq_problem"] or "No Pop Quiz questions for this lesson", lesson=key))
            continue
        qs, tagged, idx = entry
        if not lesson_slos:
            out.append(result("PQ5", REVIEW, "The specification lists no SLOs for this lesson", lesson=key))
        elif not qs:
            out.append(result("PQ5", NA, "No Pop Quiz questions for this lesson", lesson=key))
        elif tagged == 0:
            out.append(result("PQ5", REVIEW, f"No question carries an SLO tag, so which of the {len(lesson_slos)} SLOs are tested "
                                             f"cannot be decided by code", lesson=key))
        else:
            missing = [s for s in lesson_slos if s.index not in idx]
            if missing and tagged == len(qs):
                out.append(result("PQ5", FAIL, f"{len(missing)} SLO(s) have no question: " + "; ".join(f"SLO {s.index}" for s in missing),
                                  [s.text[:70] for s in missing], lesson=key))
            elif missing:
                out.append(result("PQ5", REVIEW, f"{len(qs) - tagged} question(s) are untagged, so the {len(missing)} SLO(s) "
                                                 f"without a tagged question may still be tested", lesson=key))
            else:
                out.append(result("PQ5", PASS, "Every SLO has at least one tagged question", lesson=key))

    for ch, (qs, tagged, idx, per_lesson, d) in sorted(cov["chapter_q"].items()):
        chapter_lessons = {les for (c, les) in slos if c == ch}
        all_idx = {(les, s.index) for (c, les), v in slos.items() if c == ch for s in v}
        if not qs or tagged == 0:
            msg = "No Worksheet question carries an SLO or lesson tag, so spread and coverage cannot be decided by code"
            out.append(result("WS4", REVIEW, msg, chapter=ch, doc=d.rel))
            out.append(result("WS5", REVIEW, msg, chapter=ch, doc=d.rel))
            continue
        skipped = sorted(chapter_lessons - set(per_lesson))
        counts = ", ".join(f"L{les}: {per_lesson.get(les, 0)}" for les in sorted(chapter_lessons))
        if skipped and tagged == len(qs):
            out.append(result("WS4", FAIL, f"Lesson(s) {skipped} have no question ({counts})", chapter=ch, doc=d.rel))
        elif skipped:
            out.append(result("WS4", REVIEW, f"{len(qs) - tagged} question(s) are untagged; lessons {skipped} have no tagged question ({counts})",
                              chapter=ch, doc=d.rel))
        else:
            out.append(result("WS4", PASS, f"Every lesson has a question ({counts})", chapter=ch, doc=d.rel, partial=True,
                              evidence=["'Over-represented' is not defined by the workbook; a reviewer judges the balance."]))
        missing = sorted(all_idx - idx)
        if missing and tagged == len(qs):
            out.append(result("WS5", FAIL, f"{len(missing)} SLO(s) have no Worksheet question", [f"L{l} SLO {n}" for l, n in missing[:12]],
                              chapter=ch, doc=d.rel))
        elif missing:
            out.append(result("WS5", REVIEW, f"{len(qs) - tagged} question(s) are untagged; {len(missing)} SLO(s) have no tagged question",
                              chapter=ch, doc=d.rel))
        else:
            out.append(result("WS5", PASS, "Every SLO in the chapter has a Worksheet question", chapter=ch, doc=d.rel))
    return out, rows
