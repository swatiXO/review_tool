"""Question-document checks: Pop Quiz (PQ), Chapter Exam (CE), Worksheet (WS), Data Bank (DB).

Where the workbook states a rule that code can only partly decide (for example
"split roughly 70/30"), the finding is a partial pass: the part that can be
decided is decided, and the rest is stated in the evidence for the reviewer.
"""
import re
from collections import Counter, defaultdict
from difflib import SequenceMatcher

from ..models import FAIL, NA, PASS, REVIEW, LessonKey
from ..questions import (BankItem, load_bank_items, numbering_is_regular, parse_questions, segment_lessons, para_is_highlighted)
from ..textutil import dominant_script, normalize, to_western_digits
from .common import result

# ------------------------------------------------------------------ shared


def _label_present(block, labels):
    nb = normalize(block).lower()
    return any(normalize(l).lower() in nb for l in labels if l)


def _vocab_gate(ctx, info, labels, code, what):
    """None if the label vocabulary covers this document's language, else a REVIEW finding."""
    if labels:
        return None
    return result(code, REVIEW, f"No {what} wording is configured in the profile for this document")


def _slo_tag_present(ctx, text):
    return any(re.search(p, text) for p in ctx.profile["vocab"]["slo_tag_patterns"])


# ----------------------------------------------------------------- Pop Quiz


def align_pop_quiz(ctx):
    """Map each Pop Quiz lesson section to a package LessonKey.

    Lesson numbers restart in each chapter, so sections form groups, and groups are
    matched in order to the package's chapters that have lesson folders.
    Returns ({LessonKey: [Question]}, info_or_None, problem_or_None)."""
    pq = next((d for d in ctx.pkg.docs if d.doc_type == "pop_quiz" and not d.superseded), None)
    if pq is None:
        return {}, None, "No Pop Quiz document in the package"
    info = ctx.docx(pq)
    if info is None:
        return {}, None, "The Pop Quiz could not be read"
    groups = segment_lessons(info, ctx.profile)
    chapters = sorted({d.chapter for d in ctx.pkg.docs if d.lesson is not None and d.chapter is not None})
    if len(groups) != len(chapters):
        return {}, info, (f"Pop Quiz has {len(groups)} chapter group(s) but the package has {len(chapters)} chapter(s) with "
                          f"lesson folders; sections could not be matched to lessons")
    pkg_keys = defaultdict(set)
    for d in ctx.pkg.docs:
        if d.lesson is not None and d.chapter is not None:
            pkg_keys[d.chapter].add((d.lesson, d.variant))
    mapping = {}
    for ch, grp in zip(chapters, groups):
        for h in grp:
            variant = h["variant"] if (h["num"], h["variant"]) in pkg_keys[ch] else ""
            key = LessonKey(ch, h["num"], variant)
            mapping[key] = parse_questions(info, ctx.profile, h["start"] + 1, h["end"])
    if not any(mapping.values()):
        return {}, info, ("No questions were recognised anywhere in the Pop Quiz, so its question format may not be supported "
                          "(recognised: '1.', '1)', 'Q1.', 'Q.1', 'Question 1:', '(1)', automatic numbering)")
    return mapping, info, None


def pop_quiz_checks(ctx):
    """Findings for PQ1-PQ4 for every lesson in the package."""
    out = []
    mapping, info, problem = align_pop_quiz(ctx)
    keys = sorted({d.key for d in ctx.pkg.docs if d.scope == "lesson" and d.chapter is not None}, key=lambda k: k.sort_key())
    vocab = ctx.profile["vocab"]
    for key in keys:
        if problem:
            for code in ("PQ1", "PQ2", "PQ3", "PQ4"):
                out.append(result(code, REVIEW, problem, lesson=key))
            continue
        qs = mapping.get(key, [])
        if not qs and key.lesson is None:
            out.append(result("PQ1", REVIEW, "This chapter has no lesson folders and the Pop Quiz has no section for it; "
                                             "a reviewer decides whether it is expected to have one", lesson=key, doc=_pq_rel(ctx)))
            for code in ("PQ2", "PQ3", "PQ4"):
                out.append(result(code, NA, "No Pop Quiz questions for this chapter", lesson=key))
            continue
        if not qs:
            out.append(result("PQ1", FAIL, "No Pop Quiz questions found for this lesson (expected 2-3)", lesson=key, doc=_pq_rel(ctx)))
            for code in ("PQ2", "PQ3", "PQ4"):
                out.append(result(code, NA, "No Pop Quiz questions for this lesson", lesson=key))
            continue
        ev, bad = [], False
        if not 2 <= len(qs) <= 3:
            ev.append(f"{len(qs)} questions (expected 2-3)")
            bad = True
        notmcq = [q.num for q in qs if q.qtype != "mcq"]
        if notmcq:
            ev.append(f"question(s) {notmcq} are not MCQ")
            bad = True
        out.append(result("PQ1", FAIL if bad else PASS, "; ".join(ev) or f"{len(qs)} MCQ questions", ev, lesson=key, doc=_pq_rel(ctx)))

        labels = vocab["lesson_plan_location_labels"]
        missing = [q.num for q in qs if not _label_present(q.block, labels)]
        if missing:
            out.append(result("PQ2", FAIL, f"Question(s) {missing} do not state a 'Lesson Plan Location' (label looked for: {', '.join(labels)})",
                              lesson=key, doc=_pq_rel(ctx)))
        else:
            out.append(result("PQ2", PASS, "Every question states a Lesson Plan Location", partial=True,
                              evidence=["Whether each location is precise needs a reviewer."], lesson=key, doc=_pq_rel(ctx)))

        fl = vocab["feedback_labels"]
        nofb = [q.num for q in qs if not _label_present(q.block, fl)]
        if nofb:
            out.append(result("PQ3", FAIL, f"Question(s) {nofb} have no correct/incorrect feedback (labels looked for: {', '.join(fl)})",
                              lesson=key, doc=_pq_rel(ctx)))
        else:
            out.append(result("PQ3", PASS, "Feedback present on every question", partial=True,
                              evidence=["Whether the feedback meets the standard needs a reviewer."], lesson=key, doc=_pq_rel(ctx)))

        nohl = [q.num for q in qs if not q.highlighted]
        if nohl:
            marked = [q.num for q in qs if q.check_marked]
            extra = f"; answers are marked with a check-mark symbol on question(s) {marked}, not a yellow highlight" if marked else ""
            out.append(result("PQ4", FAIL, f"Correct answer is not highlighted yellow on question(s) {nohl}{extra}", lesson=key, doc=_pq_rel(ctx)))
        else:
            out.append(result("PQ4", PASS, "Correct answer highlighted yellow on every question", lesson=key, doc=_pq_rel(ctx)))
    return out


def _pq_rel(ctx):
    d = next((d for d in ctx.pkg.docs if d.doc_type == "pop_quiz" and not d.superseded), None)
    return d.rel if d else None


# ---------------------------------------------------- Chapter Exam / Worksheet


def _count_check(code, qs, lo, hi, key=None, chapter=None, doc=None):
    n = len(qs)
    if n == 0:
        return result(code, REVIEW, "No questions were recognised in this document, so the count cannot be checked. Its question "
                                    "format may not be supported (recognised: '1.', '1)', 'Q1.', 'Q.1', 'Question 1:', '(1)', "
                                    "automatic numbering, questions in tables)", lesson=key, chapter=chapter, doc=doc)
    if n and not numbering_is_regular(qs):
        return result(code, REVIEW, f"Question numbers run {[q.num for q in qs]}, not 1..n, so the document nests or restarts "
                                    f"numbering and the question count ({lo}-{hi} expected) is not decidable by code",
                      lesson=key, chapter=chapter, doc=doc)
    if not lo <= n <= hi:
        return result(code, FAIL, f"{n} questions (expected {lo}-{hi})", lesson=key, chapter=chapter, doc=doc)
    return result(code, PASS, f"{n} questions", lesson=key, chapter=chapter, doc=doc, partial=True,
                  evidence=["The question count is within range; the level split is not checked by code."])


def _tag_check(ctx, code, qs, info, key=None, chapter=None, doc=None):
    whole = "\n".join(p.text for p in info.paras)
    if not _slo_tag_present(ctx, whole):
        return result(code, FAIL, "No SLO tags found anywhere in the document", lesson=key, chapter=chapter, doc=doc)
    if not qs or not numbering_is_regular(qs):
        return result(code, REVIEW, "Some SLO tags exist but the questions could not be matched one-to-one", lesson=key, chapter=chapter, doc=doc)
    untagged = [q.num for q in qs if not _slo_tag_present(ctx, q.block)]
    if untagged:
        return result(code, FAIL, f"Question(s) {untagged} carry no SLO tag", lesson=key, chapter=chapter, doc=doc)
    return result(code, PASS, "Every question is SLO-tagged", lesson=key, chapter=chapter, doc=doc)


def exam_checks(ctx, doc, info, doc_type):
    """CE1/CE4 on a per-lesson Chapter Exam; WS1/WS2/WS3/WS6 on a per-chapter Worksheet."""
    qs = parse_questions(info, ctx.profile)
    out = []
    if doc_type == "chapter_exam":
        out.append(_count_check("CE1", qs, 6, 8, key=doc.key, doc=doc.rel))
        out.append(_tag_check(ctx, "CE4", qs, info, key=doc.key, doc=doc.rel))
    else:
        out.append(_count_check("WS1", qs, 8, 10, chapter=doc.chapter, doc=doc.rel))
        mix = Counter(q.qtype for q in qs)
        out.append(result("WS2", REVIEW, "Question formats found: " + ", ".join(f"{k} {v}" for k, v in mix.most_common()) +
                          " (the workbook asks for a mix drawn from six formats; the reviewer judges the mix)",
                          chapter=doc.chapter, doc=doc.rel))
        out.append(_tag_check(ctx, "WS3", qs, info, chapter=doc.chapter, doc=doc.rel))
        out.append(_ws6(ctx, doc, info, qs))
    return out


def _ws6(ctx, doc, info, qs):
    labels = ctx.profile["vocab"]["answer_key_labels"]
    body = info.body_paras()
    pos = None
    for p in body:
        t = p.text.strip()
        if len(t) <= 40 and any(normalize(l).lower() in normalize(t).lower() for l in labels):
            pos = p.idx
    if pos is None:
        inline = any(_label_present(q.block, labels) for q in qs)
        why = "answers appear next to each question, but there is no Answer Key section" if inline else "no Answer Key found"
        return result("WS6", FAIL, f"No Answer Key at the end of the document: {why}", chapter=doc.chapter, doc=doc.rel)
    last_q_end = max((q.end for q in qs), default=0)
    if pos < max((q.start for q in qs), default=0):
        return result("WS6", FAIL, "An Answer Key heading exists but is not after the last question", chapter=doc.chapter, doc=doc.rel)
    trailing = [p for p in body if p.idx > pos and p.text.strip()]
    return result("WS6", PASS, "Answer Key is at the end of the document", partial=True, chapter=doc.chapter, doc=doc.rel,
                  evidence=[f"{len(trailing)} paragraph(s) follow the heading; completeness needs a reviewer."])


# ---------------------------------------------------------------- Data Bank


def data_bank_checks(ctx):
    bank = next((d for d in ctx.pkg.docs if d.doc_type == "data_bank" and not d.superseded), None)
    out = []
    keys = sorted({d.key for d in ctx.pkg.docs if d.scope == "lesson" and d.chapter is not None}, key=lambda k: k.sort_key())
    if bank is None:
        for k in keys:
            for code in ("DB1", "DB2", "DB3", "DB4", "DB5"):
                out.append(result(code, REVIEW, "No Data Bank document in the package", lesson=k))
        return out, None
    items = load_bank_items(bank.abs, ctx.profile)
    if not items:
        for k in keys:
            for code in ("DB1", "DB2", "DB3", "DB4", "DB5"):
                out.append(result(code, REVIEW, "No Data Bank items were recognised. Items are expected as two-column tables whose "
                                                "first column uses the field names in the profile (vocab.data_bank_fields)",
                                  lesson=k, doc=bank.rel))
        return out, (bank, items)
    # group items by (chapter, bank lesson number) in document order
    lessons = []
    for it in items:
        sig = (it.chapter, it.lesson_no)
        if not lessons or lessons[-1][0] != sig:
            lessons.append((sig, []))
        lessons[-1][1].append(it)
    per_chapter = defaultdict(list)
    for sig, its in lessons:
        per_chapter[sig[0]].append(its)
    pkg_by_chapter = defaultdict(list)
    for k in keys:
        pkg_by_chapter[k.chapter].append(k)
    mapping, unmatched = {}, {}
    for ch, ks in pkg_by_chapter.items():
        bl = per_chapter.get(ch, [])
        if not bl:
            continue
        if len(bl) == len(ks):
            for k, its in zip(ks, bl):
                mapping[k] = its
        else:
            unmatched[ch] = f"Data Bank has {len(bl)} lesson(s) for chapter {ch} but the package has {len(ks)}; items could not be matched to lessons"

    pq_map, _, _ = align_pop_quiz(ctx)
    pop_texts = [normalize(q.text) for qs in pq_map.values() for q in qs]
    vocab = ctx.profile["vocab"]
    mcq = {normalize(x).lower() for x in vocab["mcq_types"]}
    lower = {normalize(x).lower() for x in vocab["lower_order_levels"]}
    sim_th = ctx.th["duplicate_similarity"]

    for k in keys:
        if k.chapter in unmatched:
            for code in ("DB1", "DB2", "DB3", "DB4", "DB5"):
                out.append(result(code, REVIEW, unmatched[k.chapter], lesson=k, doc=bank.rel))
            continue
        its = mapping.get(k)
        if not its and k.lesson is None:
            out.append(result("DB1", REVIEW, "This chapter has no lesson folders and the Data Bank has no items for it; "
                                             "a reviewer decides whether it is expected to have any", lesson=k, doc=bank.rel))
            for code in ("DB2", "DB3", "DB4", "DB5"):
                out.append(result(code, NA, "No Data Bank items for this chapter", lesson=k, doc=bank.rel))
            continue
        if not its:
            out.append(result("DB1", FAIL, "No Data Bank items for this lesson (expected 4-6)", lesson=k, doc=bank.rel))
            for code in ("DB2", "DB3", "DB4", "DB5"):
                out.append(result(code, NA, "No Data Bank items for this lesson", lesson=k, doc=bank.rel))
            continue
        ev, bad = [], False
        if not 4 <= len(its) <= 6:
            ev.append(f"{len(its)} items (expected 4-6)")
            bad = True
        types = Counter(normalize(i.fields.get("question_type", "")).lower() for i in its)
        non_mcq = sum(v for t, v in types.items() if t not in mcq)
        if non_mcq:
            ev.append(f"{non_mcq} of {len(its)} items are not MCQ (" + ", ".join(f"{t or 'unset'} {v}" for t, v in types.most_common(4)) + ")")
            bad = True
        levels = Counter(normalize(i.fields.get("level", "")).lower() for i in its)
        higher = sum(v for lv, v in levels.items() if lv not in lower)
        if higher:
            ev.append(f"{higher} of {len(its)} items are above lower-order (" + ", ".join(f"{l or 'unset'} {v}" for l, v in levels.most_common(5)) + ")")
            bad = True
        out.append(result("DB1", FAIL if bad else PASS, "; ".join(ev) or f"{len(its)} lower-order MCQ items", ev, lesson=k, doc=bank.rel))

        need = ["subject", "slo"]
        has_cl = all("chapter_lesson" in i.present or ({"chapter", "lesson"} <= i.present) for i in its)
        missing = [f for f in need if not all(f in i.present for i in its)]
        if not has_cl:
            missing = ["chapter/lesson"] + missing
        if missing:
            out.append(result("DB2", FAIL, "Items are missing the tag field(s): " + ", ".join(missing), lesson=k, doc=bank.rel,
                              evidence=["Fields present on items: " + ", ".join(sorted(its[0].raw_keys))]))
        else:
            out.append(result("DB2", PASS, "Every item carries Subject, Chapter, Lesson and SLO tags", lesson=k, doc=bank.rel))

        nofb = [f for f in ("feedback_correct", "feedback_incorrect") if not all(f in i.present and i.fields.get(f) for i in its)]
        if nofb:
            out.append(result("DB3", FAIL, "No correct/incorrect feedback field on the items", lesson=k, doc=bank.rel,
                              evidence=["Fields present on items: " + ", ".join(sorted(its[0].raw_keys))]))
        else:
            out.append(result("DB3", PASS, "Correct and incorrect feedback present on every item", lesson=k, doc=bank.rel, partial=True,
                              evidence=["Whether the feedback meets the standard needs a reviewer."]))

        nohl = sum(1 for i in its if not i.highlighted)
        if nohl:
            out.append(result("DB4", FAIL, f"Correct answer is not highlighted yellow on {nohl} of {len(its)} items", lesson=k, doc=bank.rel))
        else:
            out.append(result("DB4", PASS, "Correct answer highlighted yellow on every item", lesson=k, doc=bank.rel))

        dups = []
        for i in its:
            qt = normalize(i.fields.get("question", ""))
            if not qt:
                continue
            for pt in pop_texts:
                if SequenceMatcher(None, qt, pt).ratio() >= sim_th:
                    dups.append(qt[:50])
                    break
        if dups:
            out.append(result("DB5", FAIL, f"{len(dups)} item(s) near-duplicate a Pop Quiz question", dups[:5], lesson=k, doc=bank.rel))
        else:
            out.append(result("DB5", REVIEW, "No near-identical wording found; whether any item is a reworded duplicate needs a reviewer",
                              lesson=k, doc=bank.rel))
    return out, (bank, items)


def dbs1(ctx, bank_items):
    if not bank_items:
        return result("DBS1", REVIEW, "No Data Bank document in the package")
    bank, items = bank_items
    seq = [(i.chapter, i.lesson_no) for i in items if i.chapter is not None]
    bad = [f"chapter {b[0]} lesson {b[1]} comes after chapter {a[0]} lesson {a[1]}" for a, b in zip(seq, seq[1:])
           if (b[0], b[1] or 0) < (a[0], a[1] or 0)]
    if bad:
        return result("DBS1", FAIL, "Items are not sorted by chapter and lesson", bad[:5])
    return result("DBS1", PASS, "Items are sorted by chapter and lesson", partial=True,
                  evidence=["Whether items sit in the right subject tab needs a reviewer."])
