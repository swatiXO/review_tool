"""Model-assisted checks for rules that need reading comprehension.

Run only with --model-checks. Every result is a needs-review suggestion (judge.suggestion):
the model never produces a Pass or Fail in the workbook. Rule wording is taken from the
checklist workbook, so these checks follow it if it is edited.

LP1  SLO coverage in Concept Building        LP2  no scope creep (SLOs, and the book if indexed)
LP6  warm-up concise, tests/previews         LP7  story runs through and is resolved
LP8  story faithful to the book (needs book) LP9  hard concepts flagged for a visual
FG1  SLOs stated aloud at Session Overview   FG3  warm-up slide has a debrief
FG6  takeaways tie to SLOs, close checks in  CE2/CE3 exam vs the book's end-of-lesson test (needs book)
PQ5/WS4/WS5 coverage when questions carry no SLO tags (augment_coverage)
CE1/WS1 order-level split by Bloom's levels when questions do not state them (augment_levels)
PQ3/DB3 feedback quality against the Pop Quiz standard (augment_feedback)
"""
import re

from ..judge import Material, cached_chat, judge, material_budget, suggestion, verify_quotes
from ..llm import LLMError
from ..models import FAIL, PASS, REVIEW
from ..slo import slos_for
from ..textutil import normalize, starts_with_label
from .assessments import align_pop_quiz, get_questions
from .common import mark, result
from .lessonplan import find_sections


# ------------------------------------------------------------------ material
def lesson_plan_doc(ctx, key):
    return next((d for d in ctx.pkg.docs if d.doc_type == "lesson_plan" and not d.superseded and d.ext == "docx" and d.key == key), None)


def guide_doc(ctx, key):
    return next((d for d in ctx.pkg.docs if d.doc_type == "facilitator_guide" and not d.superseded and d.ext == "pptx" and d.key == key), None)


def section_texts(ctx, info):
    """{section key: text} from the Lesson Plan, each running to the next section heading."""
    found = {k: v for k, v in find_sections(ctx, info).items() if v is not None}
    order = sorted(found.items(), key=lambda kv: kv[1])
    out = {}
    for i, (key, start) in enumerate(order):
        end = order[i + 1][1] if i + 1 < len(order) else len(info.paras)
        out[key] = "\n".join(p.text.strip() for p in info.paras[start:end] if p.text.strip())
    return out


def get_slos_text_for(ctx, key):
    from .coverage import get_slos
    slos = get_slos(ctx)
    return [s.text for s in slos_for(slos, key)] if slos else []


def lesson_slo_texts(ctx, key, sections, info=None):
    """The lesson's SLOs: from the specifications when they are in the package, otherwise the Lesson Plan's own SLO
    list (only the SLO statements, not the lines that follow them in the section, such as prior knowledge or a word list)."""
    texts = get_slos_text_for(ctx, key)
    if texts:
        return texts
    if info is not None:
        from .lessonplan import slo_items
        items, _ = slo_items(ctx, info)
        if items:
            return [re.sub(r"^\s*(?:[•▪●◦\-–—*]\s*|\(?\d+[.)]\s*)", "", p.text.strip()) for p in items]
    block = sections.get("slos", "")
    lines = [l.strip() for l in block.splitlines()[1:] if l.strip() and not re.search(r"[:：]$", l.strip())]
    return lines


def numbered(items):
    return "\n".join(f"{i}. {t}" for i, t in enumerate(items, 1))


def book_material(ctx, query):
    if getattr(ctx, "book", None) is None:
        return None
    text = ctx.book.passages(query, k=2)
    return Material("Textbook passages (OCR text, may contain errors)", text) if text else None


def where(key, doc):
    return {"lesson": key, "doc": doc.rel if doc else None}


# ------------------------------------------------------------------ the checks
def split_parts(text, budget):
    """Concept Building in parts that each fit the model's context, split between paragraphs.
    Long lessons are sent in parts instead of being cut, so nothing is lost when the context is
    lowered to keep the model on the GPU."""
    parts, cur = [], ""
    for para in text.split("\n"):
        while len(para) > budget:                      # one paragraph longer than a whole part
            parts.append(cur) if cur else None
            cur, parts = "", parts + [para[:budget]]
            para = para[budget:]
        if cur and len(cur) + len(para) + 1 > budget:
            parts.append(cur)
            cur = ""
        cur = f"{cur}\n{para}" if cur else para
    if cur:
        parts.append(cur)
    return parts or [text]


def _cb_parts(ctx, secs, slos):
    """Concept Building split to fit next to the SLO list; one part when it fits."""
    from ..judge import material_budget
    room = material_budget(ctx.model) - len(numbered(slos)) - 200
    return split_parts(secs["concept_building"], max(1500, room))


def lp1(ctx, key, doc, secs, slos):
    if not slos or "concept_building" not in secs:
        return None
    rule = ctx.rules["LP1"].text if "LP1" in ctx.rules else "Every stated SLO is fully covered in Concept Building"
    parts = _cb_parts(ctx, secs, slos)
    if len(parts) == 1:
        mats = [Material("SLOs (numbered)", numbered(slos)), Material("Concept Building", parts[0])]
        j = judge(ctx.model, "LP1", rule, mats,
                  'Is every numbered SLO fully covered somewhere in the Concept Building text? List the numbers of any SLO that is not '
                  'fully covered in "uncovered_slos" (an empty list if all are covered).',
                  extra_keys=("uncovered_slos",),
                  validate_extra=lambda e: e.get("uncovered_slos") is None or (isinstance(e["uncovered_slos"], list)
                                           and all(isinstance(n, int) and 1 <= n <= len(slos) for n in e["uncovered_slos"])))
        if j.usable and j.verdict == "pass" and j.extra.get("uncovered_slos"):
            j.usable, j.note = False, "the model said pass but also listed SLOs it found uncovered"
        f = suggestion("LP1", j, **where(key, doc))
        if j.usable and j.extra.get("uncovered_slos"):
            f.evidence.append("SLO(s) not fully covered: " + ", ".join(f"{n} ({slos[n - 1][:50]})" for n in j.extra["uncovered_slos"]))
        return f
    # Long Concept Building: ask, part by part, which SLOs that part covers; an SLO is covered if any part covers it.
    from ..judge import Judgement
    covered, quotes, notes, n = set(), [], [], len(slos)
    for i, part in enumerate(parts, 1):
        mats = [Material("SLOs (numbered)", numbered(slos)), Material(f"Concept Building, part {i} of {len(parts)}", part)]
        j = judge(ctx.model, f"LP1-part{i}", rule, mats,
                  f'This is part {i} of {len(parts)} of the Concept Building text. List in "covered_slos" the numbers of the SLOs that '
                  'THIS part covers fully, with a quote for each. Set "verdict" to pass if it covers at least one SLO, otherwise unclear.',
                  extra_keys=("covered_slos",),
                  validate_extra=lambda e: isinstance(e.get("covered_slos"), list)
                  and all(isinstance(x, int) and 1 <= x <= n for x in e["covered_slos"]))
        if j.usable:
            covered |= set(j.extra["covered_slos"])
            quotes += j.quotes
        elif j.note:
            notes.append(f"part {i}: {j.note}")
    if not quotes:
        j = Judgement(note="; ".join(notes) or "the model found nothing usable in any part")
        return suggestion("LP1", j, **where(key, doc))
    missing = [x for x in range(1, n + 1) if x not in covered]
    j = Judgement(verdict="fail" if missing else "pass", quotes=quotes[:3], usable=True,
                  reason=("These SLOs do not seem to be taught in Concept Building: " +
                          "; ".join(f"'{slos[x - 1][:70]}'" for x in missing) + "." if missing
                          else "every SLO is taught somewhere in Concept Building."))
    f = suggestion("LP1", j, **where(key, doc))
    if missing:
        f.evidence.append("SLO(s) not fully covered: " + ", ".join(f"{x} ({slos[x - 1][:50]})" for x in missing))
    if notes:
        f.evidence.append("Parts the model could not answer: " + "; ".join(notes))
    return f


def lp2(ctx, key, doc, secs, slos):
    if "concept_building" not in secs:
        return None
    rule = ctx.rules["LP2"].text if "LP2" in ctx.rules else "No scope creep in Concept Building"
    book = book_material(ctx, " ".join(slos)[:300])
    parts = _cb_parts(ctx, secs, slos) if not book else [secs["concept_building"]]
    question = ("Does the Concept Building text contain anything that goes beyond what can be traced to a stated SLO"
                + (" or the textbook passages" if book else "") + "? Answer fail only if you can quote the passage that goes beyond them.")
    results = []
    for i, part in enumerate(parts, 1):
        label = "Concept Building" if len(parts) == 1 else f"Concept Building, part {i} of {len(parts)}"
        mats = [Material("SLOs (numbered)", numbered(slos)), Material(label, part)] + ([book] if book else [])
        results.append(judge(ctx.model, "LP2" if len(parts) == 1 else f"LP2-part{i}", rule, mats, question))
    usable = [j for j in results if j.usable]
    # scope creep anywhere is scope creep; a pass needs every part read
    j = next((j for j in usable if j.verdict == "fail"), None) or (usable[0] if len(usable) == len(results) else
                                                                 next((j for j in results if not j.usable), results[0]))
    f = suggestion("LP2", j, **where(key, doc))
    if len(parts) > 1:
        f.evidence.append(f"Concept Building was read in {len(parts)} parts")
    if not book:
        f.evidence.append("The textbook was not consulted (no --book-index), so this only compares against the SLOs")
    return f


def lp6(ctx, key, doc, secs, slos):
    text = secs.get("warm_up")
    if not text:
        return None
    words = len(text.split())
    j = judge(ctx.model, "LP6", ctx.rules["LP6"].text if "LP6" in ctx.rules else "Warm-up is concise and tests prior knowledge or previews",
              [Material("Warm-up", text)],
              f"The warm-up has {words} words. Is it concise, and does it test prior knowledge and/or preview what is coming?")
    return suggestion("LP6", j, **where(key, doc))


def cb_subtopics(ctx, info):
    """[(heading, text)] for the Concept Building sub-topics, split at its headings."""
    from .formatting import doc_headings
    from .guidelines import lp_ranges
    r = lp_ranges(ctx, info).get("concept_building")
    if not r:
        return []
    heads = [p.idx for p in doc_headings(ctx, info) if r[0] < p.idx < r[1]]
    out = []
    for i, h in enumerate(heads):
        stop = heads[i + 1] if i + 1 < len(heads) else r[1]
        body = "\n".join(p.text.strip() for p in info.paras[h + 1:stop] if p.text.strip())
        out.append((info.paras[h].text.strip(), body))
    return out


def lp7(ctx, key, doc, secs, slos):
    if "introduction" not in secs or "concept_building" not in secs or "key_takeaways" not in secs:
        return None
    rule = ctx.rules["LP7"].text if "LP7" in ctx.rules else "Story runs through every sub-topic and is resolved in Key Takeaways"
    info = ctx.docx(doc)
    subs = cb_subtopics(ctx, info) if info is not None else []
    if len(subs) < 2:
        mats = [Material("Introduction (the story or scenario)", secs["introduction"]),
                Material("Concept Building", secs["concept_building"]), Material("Key Takeaways", secs["key_takeaways"])]
        j = judge(ctx.model, "LP7", rule, mats,
                  "Does the story or scenario from the Introduction come back in every Concept Building sub-topic, and is it resolved in Key Takeaways?")
        return suggestion("LP7", j, **where(key, doc))
    # Small models answer "does the story run through every sub-topic" badly (they say yes). They can
    # name a story's characters, which is checked against the Introduction, and code then looks for
    # those names in each sub-topic.
    names = story_names(ctx, key, secs["introduction"])
    if not names:
        mats = [Material("Introduction (the story or scenario)", secs["introduction"]),
                Material("Concept Building", secs["concept_building"]), Material("Key Takeaways", secs["key_takeaways"])]
        j = judge(ctx.model, "LP7", rule, mats,
                  "Does the story or scenario from the Introduction come back in every Concept Building sub-topic, and is it resolved in Key Takeaways?")
        return suggestion("LP7", j, **where(key, doc))
    from ..judge import Judgement, _loose
    vocab = ctx.profile["vocab"]
    before = [_loose(h) for h in vocab.get("honorifics", [])]
    after = [_loose(h) for h in vocab.get("honorifics_after", [])]

    def named(text):
        """Whole words only (a short name must not match inside another word), and not a historical
        person who shares the character's name ('حضرت حمزہ رضی اللہ عنہ' is not the story's Hamza)."""
        words = _loose(text).split()
        for i, w in enumerate(words):
            for nm in names:
                parts = _loose(nm).split()
                if words[i:i + len(parts)] != parts:
                    continue
                prev = words[i - 1] if i else ""
                nxt = words[i + len(parts)] if i + len(parts) < len(words) else ""
                if prev not in before and not any(nxt.startswith(a) for a in after if a):
                    return True
        return False
    has = [i for i, (h, body) in enumerate(subs, 1) if named(h + " " + body)]
    missing = [i for i in range(1, len(subs) + 1) if i not in has]
    resolved = named(secs["key_takeaways"] + " " + subs[-1][1])
    parts = [f"the story's characters ({', '.join(names)}) appear in {len(has)} of {len(subs)} Concept Building sub-topics"]
    if missing:
        parts.append("missing from: " + "; ".join(f"{i}. {subs[i - 1][0][:40]}" for i in missing))
    parts.append("resolved at the end" if resolved else "not resolved at the end")
    j = Judgement(verdict="pass" if not missing and resolved else "fail", reason="; ".join(parts) + ".",
                  quotes=[subs[i - 1][0] for i in has][:3] or names, usable=True)
    f = suggestion("LP7", j, **where(key, doc))
    f.evidence.append("Character names were taken from the Introduction by the model and checked against it; "
                      "the sub-topics were searched by code. A story told without its characters' names would be missed.")
    if missing:
        # point at the sub-topics that lack the story, not at the ones that have it
        f.marks = [mark(subs[i - 1][0], f"model suggests FAIL: this sub-topic does not bring back the lesson's story "
                                         f"({', '.join(names)}); {j.reason}", exact=True) for i in missing]
    return f


STORY_SYSTEM = """You read the Introduction of a school lesson plan (Urdu or English). It tells a short story or scenario.
List the names of the story's characters (people or named animals) exactly as written in the text. Do not list
prophets, companions or other historical or religious figures that the lesson teaches about; only the story's own
characters. Answer with JSON only: {"names": ["...", "..."]}. If there is no story with named characters, {"names": []}."""


def story_names(ctx, key, intro):
    from ..judge import _loose
    try:
        reply = cached_chat(ctx.model, f"story-names-{key.short()}", STORY_SYSTEM, intro[:4000])
    except LLMError:
        return []
    names = reply.get("names") if isinstance(reply, dict) else None
    if not isinstance(names, list):
        return []
    li = _loose(intro)
    return [n.strip() for n in names if isinstance(n, str) and len(n.strip()) >= 2 and _loose(n) and _loose(n) in li][:6]


def lp8(ctx, key, doc, secs, slos):
    if "introduction" not in secs:
        return None
    book = book_material(ctx, secs["introduction"][:300])
    if not book:
        return None
    j = judge(ctx.model, "LP8", ctx.rules["LP8"].text if "LP8" in ctx.rules else "Story content stays faithful to the book",
              [Material("Introduction (the story or scenario)", secs["introduction"]), book],
              "Does the story or scenario stay faithful to the textbook passages, with no facts that contradict them?")
    return suggestion("LP8", j, **where(key, doc))


def lp9(ctx, key, doc, secs, slos):
    text = secs.get("concept_building")
    if not text:
        return None
    j = judge(ctx.model, "LP9", ctx.rules["LP9"].text if "LP9" in ctx.rules else "Hard-to-convey concepts are flagged for a visual",
              [Material("Concept Building", text)],
              "Is any concept that is genuinely hard to convey in static text flagged for a Lottie file, infographic, table or image? "
              "Answer pass if such concepts are flagged (or none needs one), fail if a hard concept has no flag.")
    return suggestion("LP9", j, **where(key, doc))


def _slide_by(info, labels):
    labs = [normalize(l).lower() for l in labels]
    return [s for s in info.slides_text if any(l in normalize(s.title).lower() for l in labs)]


def _slide_text(slides):
    return "\n\n".join(f"[Slide {s.index}: {s.title}]\n{s.body}" + (f"\n(notes) {s.notes}" if s.notes else "") for s in slides)


def slides_for(ctx, info, role):
    """The slides for one part of the lesson: the model's reading of the deck when it read it (structure.py),
    the slide titles' label words only when it did not."""
    from ..structure import of
    st = of(info)
    if st is not None and st.slide_roles:
        return [s for s in info.slides_text if st.slide_roles.get(s.index) == role]
    labels = section_labels(ctx, role) if role in {x["key"] for x in ctx.profile["lesson_plan_sections"]} else []
    return _slide_by(info, labels) if labels else []


def section_labels(ctx, key):
    return next(s["labels"] for s in ctx.profile["lesson_plan_sections"] if s["key"] == key)


def fg1(ctx, key, g, slos):
    info = ctx.pptx(g)
    if info is None or len(info.slides_text) < 2:
        return None
    slides = slides_for(ctx, info, "slos") or info.slides_text[:3]
    j = judge(ctx.model, "FG1", ctx.rules["FG1"].text if "FG1" in ctx.rules else "SLOs are stated aloud in plain language at Session Overview",
              [Material("SLOs (numbered)", numbered(slos)), Material("Slides", _slide_text(slides))],
              "Do these slides have the facilitator state the SLOs aloud, in plain language, at the Session Overview?")
    return suggestion("FG1", j, **where(key, g))


def fg3(ctx, key, g, slos):
    info = ctx.pptx(g)
    slides = slides_for(ctx, info, "warm_up") if info else []
    if not slides:
        return None
    j = judge(ctx.model, "FG3", ctx.rules["FG3"].text if "FG3" in ctx.rules else "The Warm-up slide includes a debrief step",
              [Material("Warm-up slide(s)", _slide_text(slides))], "Does the Warm-up slide include a debrief step, not just the activity?")
    return suggestion("FG3", j, **where(key, g))


def fg6(ctx, key, g, slos):
    info = ctx.pptx(g)
    slides = slides_for(ctx, info, "key_takeaways") if info else []
    if not slides:
        return None
    tail = [s for s in (slides_for(ctx, info, "close") or info.slides_text[-2:]) if s not in slides]
    j = judge(ctx.model, "FG6", ctx.rules["FG6"].text if "FG6" in ctx.rules else "Key Takeaways ties back to the SLOs; the Close includes a check-in question",
              [Material("SLOs (numbered)", numbered(slos)), Material("Key Takeaways and closing slides", _slide_text(slides + tail))],
              "Does the Key Takeaways slide tie back to the SLOs, and does the Close include a check-in question for the students?")
    return suggestion("FG6", j, **where(key, g))


def exam_vs_book(ctx, code, key, doc, slos, instruction):
    info = ctx.docx(doc)
    if info is None or getattr(ctx, "book", None) is None:
        return None
    qs, _, _ = get_questions(ctx, doc, info)
    if not qs:
        return None
    book = book_material(ctx, " ".join(slos)[:300] + " سوال نمبر درج ذیل")
    if not book:
        return None
    text = "\n".join(f"{q.num}. {q.text}" for q in qs)
    j = judge(ctx.model, code, ctx.rules[code].text if code in ctx.rules else code,
              [Material("SLOs (numbered)", numbered(slos)), Material("Exam questions", text), book], instruction)
    return suggestion(code, j, **where(key, doc))


LESSON_PLAN_JUDGES = {"LP1": lp1, "LP2": lp2, "LP6": lp6, "LP7": lp7, "LP8": lp8, "LP9": lp9}
GUIDE_JUDGES = {"FG1": fg1, "FG3": fg3, "FG6": fg6}


def judgement_findings(ctx):
    m = ctx.model
    if m is None or not m.judge:
        return []
    out = []
    keys = sorted({d.key for d in ctx.pkg.docs if d.scope == "lesson" and d.chapter is not None}, key=lambda k: k.sort_key())
    for key in keys:
        lp = lesson_plan_doc(ctx, key)
        secs, slos = {}, []
        if lp is not None:
            info = ctx.docx(lp)
            if info is not None:
                secs = section_texts(ctx, info)
                slos = lesson_slo_texts(ctx, key, secs, info)
        for code, fn in LESSON_PLAN_JUDGES.items():
            if m.wants(code) and lp is not None and secs:
                f = fn(ctx, key, lp, secs, slos)
                if f:
                    out.append(f)
        g = guide_doc(ctx, key)
        if g is not None:
            for code, fn in GUIDE_JUDGES.items():
                if m.wants(code) and (code, key) not in ctx.decided:
                    f = fn(ctx, key, g, slos)
                    if f:
                        out.append(f)
        for d in ctx.pkg.docs:
            if d.doc_type == "chapter_exam" and not d.superseded and d.ext == "docx" and d.key == key:
                if m.wants("CE2"):
                    f = exam_vs_book(ctx, "CE2", key, d, slos, "Do the questions match the format and number of the book's end-of-lesson "
                                     "test without copying its wording or scenarios?")
                    out.extend([f] if f else [])
                if m.wants("CE3"):
                    f = exam_vs_book(ctx, "CE3", key, d, slos, "Do the higher-order questions go beyond the book's own test while staying "
                                     "inside the lesson's SLOs?")
                    out.extend([f] if f else [])
    return out


# ------------------------------------------------------- SLO mapping for coverage
MAP_SYSTEM = """You match exam questions to the learning outcomes (SLOs) they test. The text may be Urdu or English.
You get a numbered list of SLOs and a numbered list of questions. For each question, say which SLO numbers it tests.
Answer with JSON only: {"mapping": [{"question": 1, "slos": [2], "quote": "words copied exactly from that question"}]}
Use only the numbers you were given. A question may test no SLO (use an empty list) or several."""


def map_questions(ctx, tag, slo_items, questions):
    """{question number: set(slo item numbers)} verified against the supplied lists, or None.
    slo_items: [(label, text)] numbered 1..n in order. questions: [(number, text)]."""
    user = "SLOs:\n" + "\n".join(f"{i}. {t}" for i, (_, t) in enumerate(slo_items, 1)) + "\n\nQuestions:\n" + \
           "\n".join(f"{n}. {t[:300]}" for n, t in questions)
    try:
        reply = cached_chat(ctx.model, "MAP-" + tag, MAP_SYSTEM, user[:material_budget(ctx.model)])
    except LLMError:
        return None
    rows = reply.get("mapping") if isinstance(reply, dict) else None
    if not isinstance(rows, list):
        return None
    qtext = {n: normalize(t) for n, t in questions}
    out = {}
    for r in rows:
        try:
            n = int(r["question"])
            slos = [int(x) for x in r.get("slos", [])]
        except (KeyError, TypeError, ValueError):
            continue
        quote = normalize(str(r.get("quote", "")))
        if n in qtext and len(quote) >= 8 and quote in qtext[n] and all(1 <= s <= len(slo_items) for s in slos):
            out[n] = set(slos)
    return out if len(out) >= max(1, len(questions) // 2) else None


def augment_coverage(ctx, findings):
    """Where PQ5/WS5 are 'needs reviewer' only because questions carry no SLO tags, add the model's suggestion."""
    m = ctx.model
    if m is None or not m.judge:
        return
    slos_all = ctx_slos(ctx)
    if not slos_all:
        return
    mapping, _, _ = align_pop_quiz(ctx)
    for f in findings:
        if f.status != REVIEW or "carries an SLO" not in f.message:
            continue                      # only the 'no tags at all' cases are mapped by the model
        if f.code == "PQ5" and m.wants("PQ5") and f.lesson is not None:
            qs = mapping.get(f.lesson, [])
            lesson_slos = slos_all.get((f.lesson.chapter, f.lesson.lesson), [])
            if not qs or not lesson_slos:
                continue
            res = map_questions(ctx, f"PQ5-{f.lesson.short()}", [(s.label(), s.text) for s in lesson_slos], [(q.num, q.text) for q in qs])
            if res is None:
                f.evidence.append("[model-assisted] The model could not map the questions to the SLOs")
                continue
            covered = set().union(*res.values()) if res else set()
            missing = [s for s in lesson_slos if s.index not in covered]
            would = "fail" if missing else "pass"
            f.method = "model"
            f.message = (f"[model-assisted] Would be {would}: " + (f"{len(missing)} SLO(s) have no question (" +
                         "; ".join(f"SLO {s.index}" for s in missing) + ")" if missing else "every SLO has a question") +
                         ". The questions carry no SLO tags, so this is the model's mapping")
            f.evidence = [f"Q{n} -> SLO {sorted(v)}" for n, v in sorted(res.items())] + f.evidence
        elif f.code == "WS5" and m.wants("WS5") and f.chapter is not None:
            doc = next((d for d in ctx.pkg.docs if d.doc_type == "worksheet" and d.chapter == f.chapter and not d.superseded), None)
            info = ctx.docx(doc) if doc else None
            if info is None:
                continue
            qs, _, _ = get_questions(ctx, doc, info)
            items = [(s.label(), s.text) for (c, _), v in sorted(slos_all.items()) if c == f.chapter for s in v]
            if not qs or not items:
                continue
            res = map_questions(ctx, f"WS5-{f.chapter}", items, [(q.num, q.text) for q in qs])
            if res is None:
                f.evidence.append("[model-assisted] The model could not map the questions to the SLOs")
                continue
            covered = set().union(*res.values()) if res else set()
            missing = [items[i - 1][0] for i in range(1, len(items) + 1) if i not in covered]
            f.method = "model"
            f.message = (f"[model-assisted] Would be {'fail' if missing else 'pass'}: " +
                         (f"{len(missing)} SLO(s) have no question (" + "; ".join(missing[:8]) + ")" if missing else "every SLO has a question") +
                         ". The questions carry no SLO tags, so this is the model's mapping")


def ctx_slos(ctx):
    from .coverage import get_slos
    return get_slos(ctx)


# ------------------------------------------------- question level (CE1 / WS1 split)
LEVEL_SYSTEM = """You classify exam questions by Bloom's Taxonomy, using this table from the team's assessment guidelines.
The questions may be in Urdu or English.

| Order  | Level      | What it asks the learner to do                         | Example stem |
| Lower  | Remember   | Recall facts, terms, basic concepts                    | "What is...", "List the...", "Identify..." |
| Lower  | Understand | Explain an idea in their own words                     | "Explain why...", "Describe how..." |
| Lower  | Apply      | Use a concept in a familiar situation                  | "Use... to solve", "What happens when..." |
| Higher | Analyze    | Break an idea into parts, compare, find relationships  | "Compare...", "What's the difference between..." |
| Higher | Evaluate   | Judge, justify, or argue a position                    | "Do you agree that...", "Which works better, and why" |
| Higher | Create     | Combine ideas into something new                       | "Design...", "Propose...", "Write a... that..." |

Judge by what the question asks the learner to do, not by its format or length.
Answer with JSON only:
{"levels": [{"question": 1, "level": "remember", "quote": "words copied exactly from that question"}]}
Give one entry for every question number. level is one of: remember, understand, apply, analyze, evaluate, create."""

LOWER_LEVELS = {"remember", "understand", "apply"}
HIGHER_LEVELS = {"analyze", "analyse", "evaluate", "create"}


def classify_levels(ctx, tag, questions):
    """{question number: level} for [(number, text)], each verified by a quote from that question; None when
    fewer than half the questions got a verified level."""
    user = "Questions:\n" + "\n".join(f"{n}. {' '.join(t.split())[:400]}" for n, t in questions)
    try:
        reply = cached_chat(ctx.model, "LEVEL-" + tag, LEVEL_SYSTEM, user[:material_budget(ctx.model)])
    except LLMError:
        return None
    rows = reply.get("levels") if isinstance(reply, dict) else None
    if not isinstance(rows, list):
        return None
    text = {n: normalize(" ".join(t.split())) for n, t in questions}
    out = {}
    for r in rows:
        try:
            n, level = int(r["question"]), str(r["level"]).strip().lower()
        except (KeyError, TypeError, ValueError):
            continue
        quote = normalize(" ".join(str(r.get("quote", "")).split()))
        if n in text and level in LOWER_LEVELS | HIGHER_LEVELS and len(quote) >= min(8, len(text[n])) and quote in text[n]:
            out[n] = "analyze" if level == "analyse" else level
    return out if len(out) >= max(1, len(questions) / 2) else None


def augment_levels(ctx, findings):
    """Where CE1 / WS1 could not check the order-level split because the questions do not state their level,
    add the model's Bloom's classification as a suggestion."""
    m = ctx.model
    if m is None or not m.judge:
        return
    want = {"CE1": (0.3, "70/30", "chapter_exam"), "WS1": (0.6, "40/60", "worksheet")}
    for f in findings:
        if f.code not in want or not m.wants(f.code) or not any("split was not checked" in e for e in f.evidence):
            continue
        share_want, label, dtype = want[f.code]
        doc = next((d for d in ctx.pkg.docs if d.rel == f.doc and d.doc_type == dtype), None)
        info = ctx.docx(doc) if doc else None
        if info is None:
            continue
        qs, _, _ = get_questions(ctx, doc, info)
        if not qs:
            continue
        levels = classify_levels(ctx, f"{f.code}-{doc.rel}", [(q.num, q.block or q.text) for q in qs])
        if levels is None:
            f.evidence.append("[model-assisted] The model could not classify the questions' levels")
            continue
        known = [q for q in qs if q.num in levels]
        higher = [q for q in known if levels[q.num] in HIGHER_LEVELS]
        share = len(higher) / len(known)
        ok = abs(share - share_want) <= 0.2
        f.status, f.method, f.partial = REVIEW, "model", False
        f.message = (f"[model-assisted] Would be {'pass' if ok else 'fail'}: {len(qs)} questions; by Bloom's levels {len(higher)} of "
                     f"{len(known)} are higher-order ({round(share * 100)}%, expected about {round(share_want * 100)}%, split {label}). "
                     "The questions do not state their level, so this is the model's reading")
        f.evidence = [f"Q{q.num}: {levels[q.num]} ({'higher' if levels[q.num] in HIGHER_LEVELS else 'lower'}-order)" for q in known] + \
                     [e for e in f.evidence if "split" not in e]
        if ok:
            f.marks = [mark(known[0].text, f"model suggests PASS: {len(higher)} of {len(known)} questions are higher-order, "
                                           f"close to the {label} split")]
        else:
            side = higher if share > share_want else [q for q in known if q not in higher]
            what = "higher" if share > share_want else "lower"
            f.marks = [mark(q.text, f"model suggests FAIL: a {what}-order question ({levels[q.num]}); there are too many of these for "
                                    f"the {label} split") for q in side]


# ------------------------------------------------- feedback quality (PQ3 / DB3)
FEEDBACK_SYSTEM = """You check the feedback messages of quiz questions against the team's standard. The text may be Urdu or English.

The standard:
- Correct feedback briefly reinforces WHY the answer is right, in one sentence: a small win, not just "Correct!" or "Well done".
- Incorrect feedback names the likely misconception but does NOT give away or point to the right answer, and is never just
  "Incorrect, try again". The learner must still rethink and pick the answer themselves.
- Both are short enough to read in a few seconds.

For each question you get the question with its options (a check mark may show the right one) and the two messages.
Answer with JSON only:
{"items": [{"question": 1, "correct_ok": true, "correct_why": "it says why the answer is right", "incorrect_ok": false,
            "incorrect_why": "it names the right answer, so the learner does not have to rethink"}]}
Give one entry for every question number. The "why" fields explain your judgement against the standard in one short
English sentence; never copy the message itself."""


def feedback_labels(vocab):
    fields = vocab.get("data_bank_fields", {})
    flat = vocab.get("feedback_labels", [])
    incorrect = list(fields.get("feedback_incorrect", [])) + [l for l in flat if re.match(r"(?i)incorrect|غلط", l)]
    correct = list(fields.get("feedback_correct", [])) + [l for l in flat if l not in incorrect]
    return correct, incorrect


def feedback_of(block, vocab):
    """(correct feedback, incorrect feedback) written under a question, each '' when absent."""
    correct, incorrect = feedback_labels(vocab)
    lines = [l.strip() for l in (block or "").split("\n")]
    found = {"correct": "", "incorrect": ""}
    for i, line in enumerate(lines):
        low = normalize(line).lower()
        for kind, labels in (("incorrect", incorrect), ("correct", correct)):     # 'Incorrect' contains 'correct'
            lab = next((l for l in labels if low.startswith(normalize(l).lower())), None)
            if lab is None:
                continue
            rest = line[len(lab):].lstrip(" :：-–—")
            if not rest and i + 1 < len(lines):
                rest = lines[i + 1]
            if not found[kind]:
                found[kind] = rest.strip()
            break
    return found["correct"], found["incorrect"]


def judge_feedback(ctx, tag, items):
    """items: [(number, question with options, correct feedback, incorrect feedback)] ->
    {number: (correct_ok, correct_why, incorrect_ok, incorrect_why)} or None."""
    user = "\n\n".join(f"Question {n}:\n{' '.join(q.split())[:500]}\nCorrect feedback: {c}\nIncorrect feedback: {w}" for n, q, c, w in items)
    try:
        reply = cached_chat(ctx.model, "FEEDBACK-" + tag, FEEDBACK_SYSTEM, user[:material_budget(ctx.model)])
    except LLMError:
        return None
    rows = reply.get("items") if isinstance(reply, dict) else None
    if not isinstance(rows, list):
        return None
    nums = {n for n, *_ in items}
    out = {}
    for r in rows:
        try:
            n = int(r["question"])
        except (KeyError, TypeError, ValueError):
            continue
        if n in nums and isinstance(r.get("correct_ok"), bool) and isinstance(r.get("incorrect_ok"), bool):
            out[n] = (r["correct_ok"], str(r.get("correct_why", "")).strip(), r["incorrect_ok"], str(r.get("incorrect_why", "")).strip())
    return out if len(out) >= max(1, len(items) / 2) else None


def _feedback_finding(f, items, verdicts):
    bad, marks = [], []
    for n, _, c, w in items:
        if n not in verdicts:
            continue
        cok, cwhy, iok, iwhy = verdicts[n]
        if not cok:
            bad.append(n)
            marks.append(mark(c, f"model suggests FAIL: correct feedback on question {n} falls short: {cwhy or 'it does not say why the answer is right'}",
                              exact=False))
        if not iok:
            bad.append(n)
            marks.append(mark(w, f"model suggests FAIL: incorrect feedback on question {n} falls short: {iwhy or 'see the standard'}", exact=False))
    bad = sorted(set(bad))
    f.status, f.method, f.partial = REVIEW, "model", False
    if bad:
        f.message = (f"[model-assisted] Would be fail: the feedback on question(s) {bad} does not meet the standard (correct: say briefly "
                     "why it is right; incorrect: name the misconception without giving the answer away)")
    else:
        f.message = "[model-assisted] Would be pass: every question's feedback meets the standard"
        first = items[0]
        marks = [mark(first[2], "model suggests PASS: the feedback reinforces the right answer and names misconceptions without "
                                "giving the answer away", exact=False)]
    f.evidence = [f"Q{n}: correct {'ok' if v[0] else 'falls short'} ({v[1]}); incorrect {'ok' if v[2] else 'falls short'} ({v[3]})"
                  for n, v in sorted(verdicts.items())] + f.evidence
    f.marks = [x for x in marks if len(x["text"]) >= 6]


def augment_feedback(ctx, findings):
    """Where PQ3 / DB3 found feedback on every question, add the model's reading of whether it meets the standard."""
    m = ctx.model
    if m is None or not m.judge:
        return
    vocab = ctx.profile["vocab"]
    mapping = None
    for f in findings:
        if f.code not in ("PQ3", "DB3") or f.status != PASS or f.lesson is None or not m.wants(f.code):
            continue
        items = []
        if f.code == "PQ3":
            if mapping is None:
                mapping, _, _ = align_pop_quiz(ctx)
            for q in mapping.get(f.lesson, []):
                c, w = feedback_of(q.block, vocab)
                if c and w:
                    body = "\n".join(l for l in q.block.split("\n") if l.strip() not in (c, w) and not any(
                        normalize(l).lower().startswith(normalize(x).lower()) for x in vocab.get("feedback_labels", [])))
                    items.append((q.num, body, c, w))
        else:
            bank = getattr(ctx, "_db_by_lesson", {}).get(f.lesson)
            for n, it in enumerate(bank or [], 1):
                c, w = it.fields.get("feedback_correct", ""), it.fields.get("feedback_incorrect", "")
                if c and w:
                    items.append((n, f"{it.fields.get('question', '')}\n{it.fields.get('options', '')}", c, w))
        if not items:
            continue
        verdicts = judge_feedback(ctx, f"{f.code}-{f.lesson.short()}", items)
        if verdicts is None:
            f.evidence.append("[model-assisted] The model could not judge the feedback")
            continue
        _feedback_finding(f, items, verdicts)
