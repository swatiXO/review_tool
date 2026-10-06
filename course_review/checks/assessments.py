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
from ..fallback import recover_questions
from ..textutil import dominant_script, normalize, to_western_digits
from .common import result, mark


def _policy_one(ctx, f, summary):
    """A finding built on model-recovered questions is never a silent result.
    'suggest': a Pass/Fail becomes a needs-review suggestion that says what it would be.
    'decide': a Fail is written as a Fail; a Pass stays a partial pass (blank in the workbook)."""
    if f.status not in (PASS, FAIL):
        return f
    f.method = "model"
    tag = f"[model-assisted] {summary}. "
    if ctx.model.mode == "decide":
        f.message = tag + f.message
        if f.status == PASS:
            f.partial = True
    else:
        f.message = tag + f"Would be {f.status}: " + f.message
        f.status, f.partial = REVIEW, False
    f.evidence = [summary] + list(f.evidence)
    return f


def _model_policy(ctx, findings, qs, summary):
    if ctx.model is None or not any(q.source == "model" for q in qs):
        return findings
    return [_policy_one(ctx, f, summary) for f in findings]

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


_TITLE_NOISE = re.compile(r"[\d\W_]+")


def _title_words(text, drop):
    words = {w for w in _TITLE_NOISE.split(to_western_digits(normalize(text)).lower()) if w}
    return words - drop


def _lesson_title_text(ctx, key):
    """The opening text of the lesson's own documents (Lesson Plan first), where the lesson's name is written."""
    docs = sorted((d for d in ctx.pkg.docs if d.key == key and d.ext in ("docx", "pptx") and not d.superseded),
                  key=lambda d: (d.doc_type != "lesson_plan", d.doc_type != "facilitator_guide"))
    texts = [docs[0].folder_label or ""] if docs else []
    for d in docs[:2]:
        info = ctx.docx(d) if d.ext == "docx" else ctx.pptx(d)
        if info is None:
            continue
        if d.ext == "docx":
            texts += [p.text for p in info.paras if p.text.strip()][:3]
        elif info.slides_text:
            st = info.slides_text[0]
            texts += [st.title or ""] + list(st.texts)[:2]
    return " ".join(texts)


_STOP = {"کی", "کا", "کے", "اور", "و", "میں", "سے", "کو", "پر", "ہے", "the", "of", "and", "a", "in", "to"}


def _match_groups_by_title(ctx, groups, pkg_keys):
    """{chapter: group index}. Each package lesson votes for the Pop Quiz group whose section with
    the same lesson number shares the most title words with the lesson's own opening text (a
    unique best only); a chapter goes to the group most of its lessons voted for. Works when the
    package holds only some chapters, or one lesson, and whatever way the title is written."""
    drop = {normalize(w).lower() for w in ["lesson", "سبق", "باب", "chapter", "part"]} | _STOP
    votes = defaultdict(Counter)
    for ch, lessons in pkg_keys.items():
        for num, var in lessons:
            words = _title_words(_lesson_title_text(ctx, LessonKey(ch, num, var)), set())
            scored = []
            for gi, grp in enumerate(groups):
                h = next((h for h in grp if h["num"] == num and (not var or h["variant"] == var)), None)
                if h:
                    tw = {w for w in _title_words(h["title"], drop) if len(w) > 1}
                    scored.append((len(tw & words), gi))
            scored.sort(reverse=True)
            if scored and scored[0][0] >= 1 and (len(scored) == 1 or scored[0][0] > scored[1][0]):
                votes[ch][scored[0][1]] += 1
    chosen, used = {}, set()
    for ch, c in sorted(votes.items(), key=lambda kv: -max(kv[1].values())):
        (gi, n), *rest = c.most_common()
        if gi not in used and (not rest or n > rest[0][1]):
            chosen[ch] = gi
            used.add(gi)
    return chosen


def align_pop_quiz(ctx):
    """Map each Pop Quiz lesson section to a package LessonKey.

    Lesson numbers restart in each chapter, so sections form groups of lessons. A package
    chapter is matched to the group whose lesson titles appear in that chapter's own lesson
    documents; when titles say nothing and the counts agree, groups are matched in order.
    Lessons that cannot be placed get a reason in the returned 'unplaced' dict.
    Returns ({LessonKey: [Question]}, info_or_None, problem_or_None); cached per run."""
    cache = ctx.__dict__.get("_pq_align")
    if cache is not None:
        return cache
    ctx._pq_unplaced = {}
    ctx._pq_align = out = _align_pop_quiz(ctx)
    return out


def _align_pop_quiz(ctx):
    pq = next((d for d in ctx.pkg.docs if d.doc_type == "pop_quiz" and not d.superseded), None)
    if pq is None:
        return {}, None, "No Pop Quiz document in the package"
    info = ctx.docx(pq)
    if info is None:
        return {}, None, "The Pop Quiz could not be read"
    groups = segment_lessons(info, ctx.profile)
    pkg_keys = defaultdict(set)
    for d in ctx.pkg.docs:
        if d.lesson is not None and d.chapter is not None:
            pkg_keys[d.chapter].add((d.lesson, d.variant))
    chapters = sorted(pkg_keys)
    chosen = _match_groups_by_title(ctx, groups, pkg_keys)
    if not chosen and len(groups) == len(chapters):
        chosen = {ch: gi for gi, ch in enumerate(chapters)}
    if not chosen:
        return {}, info, (f"Pop Quiz has {len(groups)} chapter group(s); none of their lesson titles matched this package's "
                          f"lesson documents, so its sections could not be matched to lessons")
    for ch in chapters:
        if ch not in chosen:
            for num, var in pkg_keys[ch]:
                ctx._pq_unplaced[LessonKey(ch, num, var)] = (
                    "No Pop Quiz section's lesson titles matched this chapter's lesson documents, so its questions could not be found")
    mapping = {}
    for ch, gi in chosen.items():
        for h in groups[gi]:
            variant = h["variant"] if (h["num"], h["variant"]) in pkg_keys[ch] else ""
            key = LessonKey(ch, h["num"], variant)
            if (h["num"], variant) not in pkg_keys[ch]:
                continue                  # a lesson the package does not contain
            qs = parse_questions(info, ctx.profile, h["start"] + 1, h["end"])
            if ctx.model is not None and (not qs or not numbering_is_regular(qs)):
                recovered, _ = recover_questions(info, ctx.model, pq, h["start"] + 1, h["end"])
                if recovered:
                    qs = recovered
            mapping[key] = qs
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
        if key in ctx._pq_unplaced:
            for code in ("PQ1", "PQ2", "PQ3", "PQ4"):
                out.append(result(code, REVIEW, ctx._pq_unplaced[key], lesson=key, doc=_pq_rel(ctx)))
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
        out.append(result("PQ1", FAIL if bad else PASS, "; ".join(ev) or f"{len(qs)} MCQ questions", ev, lesson=key, doc=_pq_rel(ctx),
                          marks=[mark(q.text, "not a multiple-choice question") for q in qs if q.qtype != "mcq"]))

        labels = vocab["lesson_plan_location_labels"]
        missing = [q.num for q in qs if not _label_present(q.block, labels)]
        if missing:
            out.append(result("PQ2", FAIL, f"Question(s) {missing} do not state a 'Lesson Plan Location' (label looked for: {', '.join(labels)})",
                              lesson=key, doc=_pq_rel(ctx),
                              marks=[mark(q.text, "no 'Lesson Plan Location' given") for q in qs if q.num in missing]))
        else:
            out.append(result("PQ2", PASS, "Every question states a Lesson Plan Location", partial=True,
                              evidence=["Whether each location is precise needs a reviewer."], lesson=key, doc=_pq_rel(ctx)))

        fl = vocab["feedback_labels"]
        nofb = [q.num for q in qs if not _label_present(q.block, fl)]
        if nofb:
            out.append(result("PQ3", FAIL, f"Question(s) {nofb} have no correct/incorrect feedback (labels looked for: {', '.join(fl)})",
                              lesson=key, doc=_pq_rel(ctx),
                              marks=[mark(q.text, "no correct/incorrect feedback") for q in qs if q.num in nofb]))
        else:
            out.append(result("PQ3", PASS, "Feedback present on every question", partial=True,
                              evidence=["Whether the feedback meets the standard needs a reviewer."], lesson=key, doc=_pq_rel(ctx)))

        nohl = [q.num for q in qs if not q.highlighted]
        if nohl:
            marked = [q.num for q in qs if q.check_marked]
            extra = f"; answers are marked with a check-mark symbol on question(s) {marked}, not a yellow highlight" if marked else ""
            out.append(result("PQ4", FAIL, f"Correct answer is not highlighted yellow on question(s) {nohl}{extra}", lesson=key, doc=_pq_rel(ctx),
                              marks=[mark(q.text, "correct answer is not highlighted yellow" + (" (a check mark is used instead)" if q.check_marked else ""))
                                     for q in qs if not q.highlighted]))
        else:
            out.append(result("PQ4", PASS, "Correct answer highlighted yellow on every question", lesson=key, doc=_pq_rel(ctx)))
    if ctx.model is not None:
        for key, qs in mapping.items():
            meta = next((q.meta for q in qs if q.source == "model" and q.meta), None)
            if meta:
                summary = (f"model {meta['model']}: {meta['recognised']} question(s) recognised, {meta['proposed']} proposed, "
                           f"{meta['rejected']} rejected by verification")
                for f in out:
                    if f.lesson == key:
                        _policy_one(ctx, f, summary)
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


_DEFAULT_WORDS = {"higher_order_words": ["higher", "analy", "evaluat", "creat"], "lower_order_words": ["lower", "remember", "understand", "apply"]}


def question_level(q, vocab=None):
    """'higher' / 'lower' / None from the level written with the question (its heading line, or the
    template's Order Level column)."""
    vocab = vocab or {}
    words = lambda k: vocab.get(k) or _DEFAULT_WORDS[k]
    head = normalize(q.text.split("\n")[0] + " " + " ".join(re.findall(r"(?:order level|آرڈر لیول)[^\n]*", q.block, re.I))).lower()
    hi = any(normalize(w).lower() in head for w in words("higher_order_words"))
    lo = any(normalize(w).lower() in head for w in words("lower_order_words"))
    return "higher" if hi and not lo else "lower" if lo and not hi else None


def _split_check(f, qs, want_higher, label, vocab=None):
    """Turn a count Pass into Fail when the questions' own level labels show the split is far off."""
    if f.status != PASS or not qs:
        return f
    levels = [question_level(q, vocab) for q in qs]
    known = [l for l in levels if l]
    if len(known) < max(2, len(qs) / 2):
        f.evidence.append(f"The {label} split was not checked: fewer than half the questions state their level")
        return f
    share = sum(1 for l in known if l == "higher") / len(known)
    if abs(share - want_higher) > 0.2:
        f.status, f.partial = FAIL, False
        f.message = (f"{len(qs)} questions, but {round(share * 100)}% of those that state a level are higher-order "
                     f"(expected about {round(want_higher * 100)}%, split {label})")
        f.marks = [mark(q.text, "higher-order question") for q, l in zip(qs, levels) if l == "higher"] if share > want_higher else \
                  [mark(q.text, "lower-order question") for q, l in zip(qs, levels) if l == "lower"]
    else:
        f.evidence.append(f"{round(share * 100)}% of questions that state a level are higher-order (split {label})")
    return f


def _answer_key(ctx, info, qs):
    """CE5: an answer key is optional for the per-lesson Assessment; present -> Pass, absent -> N/A."""
    labels = ctx.profile["vocab"]["answer_key_labels"] + ctx.profile["vocab"].get("answer_line_labels", [])
    with_answer = [q.num for q in qs if _label_present(q.block, [l for l in labels if len(l) > 3])]
    key_section = any(len(p.text.strip()) <= 40 and _label_present(p.text, ctx.profile["vocab"]["answer_key_labels"])
                      for p in info.body_paras())
    if key_section:
        return result("CE5", PASS, "An answer key is included", partial=True, evidence=["Completeness needs a reviewer."])
    if with_answer:
        return result("CE5", PASS, f"Answers are given under {len(with_answer)} of {len(qs)} questions", partial=True,
                      evidence=["The guideline recommends a key at the end of the document rather than under each question."])
    return result("CE5", NA, "The assessment includes no answer key")


def _tag_check(ctx, code, qs, info, key=None, chapter=None, doc=None):
    whole = "\n".join(p.text for p in info.paras)
    if not _slo_tag_present(ctx, whole):
        return result(code, FAIL, "No SLO tags found anywhere in the document", lesson=key, chapter=chapter, doc=doc)
    if not qs or not numbering_is_regular(qs):
        return result(code, REVIEW, "Some SLO tags exist but the questions could not be matched one-to-one", lesson=key, chapter=chapter, doc=doc)
    untagged = [q.num for q in qs if not _slo_tag_present(ctx, q.block)]
    if untagged:
        return result(code, FAIL, f"Question(s) {untagged} carry no SLO tag", lesson=key, chapter=chapter, doc=doc,
                      marks=[mark(q.text, "question has no SLO tag") for q in qs if q.num in untagged])
    return result(code, PASS, "Every question is SLO-tagged", lesson=key, chapter=chapter, doc=doc)


def get_questions(ctx, doc, info):
    """(questions, model_summary, why_model_did_not_help) for a question document, parsed once.
    The rule-based parser goes first; the model is asked only when the rules find nothing
    or cannot trust the numbering."""
    cache = ctx.__dict__.setdefault("_qcache", {})
    if id(info) not in cache:
        qs = parse_questions(info, ctx.profile)
        summary = why_not = ""
        if ctx.model is not None and (not qs or not numbering_is_regular(qs)):
            recovered, summary = recover_questions(info, ctx.model, doc)
            if recovered:
                qs = recovered
            else:
                why_not = summary
        cache[id(info)] = (info, qs, summary, why_not)   # holding info keeps its id from being reused
    return cache[id(info)][1:]


def exam_checks(ctx, doc, info, doc_type):
    """CE1/CE4 on a per-lesson Chapter Exam; WS1/WS2/WS3/WS6 on a per-chapter Worksheet."""
    qs, summary, why_not = get_questions(ctx, doc, info)
    out = []
    if doc_type == "chapter_exam":
        out.append(_split_check(_count_check("CE1", qs, 6, 8, key=doc.key, doc=doc.rel), qs, 0.3, "70/30", ctx.profile["vocab"]))
        out.append(_tag_check(ctx, "CE4", qs, info, key=doc.key, doc=doc.rel))
        f = _answer_key(ctx, info, qs)
        f.lesson, f.doc = doc.key, doc.rel
        out.append(f)
    else:
        out.append(_split_check(_count_check("WS1", qs, 8, 10, chapter=doc.chapter, doc=doc.rel), qs, 0.6, "40/60", ctx.profile["vocab"]))
        mix = Counter(q.qtype for q in qs)
        out.append(result("WS2", REVIEW, "Question formats found: " + ", ".join(f"{k} {v}" for k, v in mix.most_common()) +
                          " (the workbook asks for a mix drawn from six formats; the reviewer judges the mix)",
                          chapter=doc.chapter, doc=doc.rel))
        out.append(_tag_check(ctx, "WS3", qs, info, chapter=doc.chapter, doc=doc.rel))
        out.append(_ws6(ctx, doc, info, qs))
    if why_not:
        for f in out:
            if f.status == REVIEW:
                f.evidence.append("Model fallback did not help: " + why_not)
    return _model_policy(ctx, out, qs, summary)


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
            out.append(result("DB4", FAIL, f"Correct answer is not highlighted yellow on {nohl} of {len(its)} items", lesson=k, doc=bank.rel,
                              marks=[mark(i.fields.get("question", ""), "correct answer is not highlighted yellow") for i in its if not i.highlighted]))
        else:
            out.append(result("DB4", PASS, "Correct answer highlighted yellow on every item", lesson=k, doc=bank.rel))

        dups = []
        for i in its:
            qt = normalize(i.fields.get("question", ""))
            if not qt:
                continue
            for pt in pop_texts:
                if SequenceMatcher(None, qt, pt).ratio() >= sim_th:
                    dups.append(qt)
                    break
        if dups:
            out.append(result("DB5", FAIL, f"{len(dups)} item(s) near-duplicate a Pop Quiz question", [d[:50] for d in dups[:5]], lesson=k, doc=bank.rel,
                              marks=[mark(d, "nearly the same as a Pop Quiz question") for d in dups]))
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
