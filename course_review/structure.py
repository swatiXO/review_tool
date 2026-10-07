"""The model reads a document's structure; code checks the rules on it.

The model is given the document as numbered paragraphs (or slides) and answers with JSON that
only POINTS at them by number: which paragraph starts each section, which are headings, which are
SLO items, and so on. It never writes document text, so the wording used by every check and by the
corrected copy is always the file's own, and a paragraph number that does not exist is discarded.

The model's reading decides. Plain code verifies it: every pointer must exist, a section cannot start
on the title, a list item, a label or a note, and a heading cannot be a long sentence, an option line or a
verse. A pointer that fails is dropped. The rule-based parser is used only when the model is unavailable,
and to fill in a section the model did not see whose label stands on its own line.
"""
import json
import re
import time
from dataclasses import dataclass, field

from .fallback import file_sha1
from .llm import LLMError
from .textutil import is_arabic_scripture, normalize

PROMPT_VERSION = "s3"
LINE_CHARS = 160              # each paragraph is shown up to this many characters; enough to recognise its role
OVERLAP = 3                   # paragraphs repeated between chunks
MAX_REJECTED_SHARE = 0.3
SECTION_KEYS = ["introduction", "slos", "warm_up", "concept_building", "key_takeaways"]
LIST_KEYS = ["headings", "slo_items", "warm_up_questions", "takeaway_items"]

LP_SYSTEM = """You map the structure of a school Lesson Plan. It may be in Urdu or English and may not follow any template.
You get the document as numbered paragraphs: "[12] text". A tag may come before the text:
{H} the paragraph uses a Heading style, {•} bulleted list item, {1.} numbered list item, {T2} inside table 2.
Tags are hints only: authors often misuse styles.

A Lesson Plan has five sections. They can have any name or none, in any language:
- introduction: opens the lesson, often with a story or scenario (e.g. Introduction, تعارف)
- slos: the Student Learning Outcomes, what students will be able to do (e.g. SLOs, حاصلات تعلیم)
- warm_up: a short opening activity or questions to start students thinking (e.g. Warm-up, تحریکی سرگرمی)
- concept_building: the main teaching content (e.g. Concept Building, تصوراتی تعمیر)
- key_takeaways: the closing recap or summary of the lesson (e.g. Key Takeaways, حاصل کلام)

Answer with JSON only, using only paragraph numbers that appear in the input:
{"sections": {"introduction": 3, "slos": 5, "warm_up": 9, "concept_building": 14, "key_takeaways": 24},
 "headings": [3, 5, 9, 14, 15, 18, 24],
 "slo_items": [6, 7],
 "warm_up_questions": [10, 11],
 "takeaway_items": [26, 27],
 "glossary": null}

- sections: the paragraph where each section begins (its heading, or its first paragraph if it has no heading).
  Use null if that section's content is not in the document under any name or in any form. Some documents are not
  built on these five sections at all; then null is the right answer for several of them. Never use the document's
  title, a label such as "Knowledge:", a note, an instruction or a list item as a section.
- headings: every paragraph that is a title for a part of the document (section headings and sub-headings).
  Not running text, not story sentences, not list items, not questions, not table cells.
- slo_items: each learning-outcome statement. Not the sentence that introduces them, not labels like "Knowledge:".
- warm_up_questions: each question the warm-up asks students.
- takeaway_items: each recap point of the key takeaways.
- glossary: the paragraph that heads a list of difficult words and their meanings, or null."""

LP_PART_NOTE = """This is part {n} of {total} of the document. Paragraph numbers continue across parts.
Answer only about the paragraphs in this part; use null for a section that does not begin in this part."""

HEAD_SYSTEM = """You decide which lines of a school document are HEADINGS. It may be in Urdu or English.
A heading is a short title for a part of the document (a section or sub-section), like "Concept Building",
"4.1 Solids" or "صبر و تحمل کی اہمیت". These are NOT headings: running text, story sentences, questions,
list items, answer options, instructions, labels followed by their content on the same line ("Note: ..."),
and Quran verses or other quotations.
You get candidate lines as "[12] the line" with the paragraph before and after it, for context.

Answer with JSON only: {"headings": [12, 40], "not_headings": [17]}
Put every candidate number in exactly one of the two lists."""

FG_ROLES = ["cover", "slos", "introduction", "warm_up", "concept_building", "key_takeaways", "close", "other"]
FG_SYSTEM = """You map the slides of a teacher's Facilitator's Guide (a slide deck) to the parts of a lesson.
It may be in Urdu or English and may not follow any template. You get each slide as "[S4] text on the slide".
Templates often leave instructions on slides; judge each slide by what part of the lesson it is for.

Parts:
- cover: the title slide (topic, grade, duration)
- slos: the session overview / learning outcomes for students
- introduction: the lesson's introduction, story or scenario
- warm_up: the warm-up activity or questions
- concept_building: the main teaching content (often several slides)
- key_takeaways: the recap or summary
- close: closing / conclusion / next steps
- other: anything else

Answer with JSON only: {"slides": {"1": "cover", "2": "slos", "3": "introduction"}}
Give one part for every slide number in the input."""


@dataclass
class Structure:
    """The reconciled structure of one document."""
    kind: str                                         # lesson_plan | facilitator_guide
    sections: dict = field(default_factory=dict)      # key -> paragraph index (or slide number) or None
    source: dict = field(default_factory=dict)        # key -> both | model | parser | disputed | none
    found_as: dict = field(default_factory=dict)      # key -> the text that begins the section, when only the model found it
    headings: set = field(default_factory=set)        # headings the checks use
    agreed_headings: set = field(default_factory=set) # headings both readings agree on: safe to edit
    disputed_headings: set = field(default_factory=set)
    run_in: set = field(default_factory=set)          # section starts with their text run on ('Warm-up: ...'): the fix splits them
    lists: dict = field(default_factory=dict)         # slo_items / warm_up_questions / takeaway_items -> [index]
    glossary: object = None
    slide_roles: dict = field(default_factory=dict)   # slide number -> role
    notes: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)

    def disputed(self, key):
        return self.source.get(key) in ("parser", "disputed")


# ------------------------------------------------------------------ prompts
def lp_lines(info):
    """[(paragraph index, prompt line)] for the non-empty paragraphs outside the table of contents."""
    out = []
    for p in info.paras:
        t = " ".join(p.text.split())
        if not t or p.in_toc:
            continue
        tags = []
        if p.heading_level:
            tags.append("{H}")
        if p.list_kind == "bullet":
            tags.append("{•}")
        elif p.list_kind == "decimal":
            tags.append("{1.}")
        if p.in_table:
            tags.append("{T%d}" % (p.table_idx + 1))
        short = t if len(t) <= LINE_CHARS else t[:LINE_CHARS] + " …"
        out.append((p.idx, f"[{p.idx}] {''.join(tags)}{' ' if tags else ''}{short}"))
    return out


def chunked(lines, budget):
    chunks, cur, size = [], [], 0
    for idx, line in lines:
        if cur and size + len(line) + 1 > budget:
            chunks.append(cur)
            cur, size = cur[-OVERLAP:], sum(len(x[1]) + 1 for x in cur[-OVERLAP:])
        cur.append((idx, line))
        size += len(line) + 1
    if cur:
        chunks.append(cur)
    return chunks


def _ints(v):
    out = []
    for x in v if isinstance(v, list) else []:
        try:
            out.append(int(str(x).lstrip("PpSs[").rstrip("]")))
        except (TypeError, ValueError):
            out.append(None)
    return out


def _int(v):
    if v is None or v == "null":
        return None
    try:
        return int(str(v).lstrip("PpSs[").rstrip("]"))
    except (TypeError, ValueError):
        return "bad"


# ------------------------------------------------------------------ the model's reading
def read_lesson_plan(ctx, info, client, budget):
    """The model's reading as {"sections", lists..., "glossary"}, plus counts; raises LLMError."""
    lines = lp_lines(info)
    chunks = chunked(lines, budget)
    valid_all = {i for i, _ in lines}
    out = {"sections": {k: None for k in SECTION_KEYS}, "glossary": None, **{k: [] for k in LIST_KEYS}}
    proposed = rejected = 0
    for n, chunk in enumerate(chunks, 1):
        valid = {i for i, _ in chunk}
        user = "\n".join(line for _, line in chunk)
        if len(chunks) > 1:
            user = LP_PART_NOTE.format(n=n, total=len(chunks)) + "\n\n" + user
        reply = client.chat_json(LP_SYSTEM, user)
        if not isinstance(reply, dict):
            raise LLMError("the model's answer was not a JSON object")
        if not isinstance(reply.get("sections"), dict):
            raise LLMError("the model's answer had no 'sections' object")
        secs = reply["sections"]
        for k in SECTION_KEYS:
            v = _int(secs.get(k))
            if v is None:
                continue
            proposed += 1
            if v == "bad" or v not in valid:
                rejected += 1
                continue
            if out["sections"][k] is None or v < out["sections"][k]:
                out["sections"][k] = v
        for k in LIST_KEYS:
            for v in _ints(reply.get(k)):
                proposed += 1
                if v is None or v not in valid:
                    rejected += 1
                elif v not in out[k]:
                    out[k].append(v)
        g = _int(reply.get("glossary"))
        if g not in (None, "bad") and g in valid_all and out["glossary"] is None:
            out["glossary"] = g
    for k in LIST_KEYS:
        out[k].sort()
    # a second, focused question: which candidate lines are headings (recall over a whole document is weak)
    cands = heading_candidates(ctx, info, out["headings"])
    verdicts, hp, hr = read_headings(info, client, cands, budget) if cands else ({}, 0, 0)
    out["heading_verdicts"] = {str(k): v for k, v in verdicts.items()}
    return out, {"chunks": len(chunks), "proposed": proposed + hp, "rejected": rejected + hr, "heading_candidates": len(cands)}


def heading_candidates(ctx, info, first_pass):
    """Lines that might be headings: the parser's, the model's first reading, anything in a Heading style,
    and short bold lines."""
    from .checks.formatting import parser_headings
    cands = {p.idx for p in parser_headings(ctx, info)} | set(first_pass)
    for p in info.paras:
        t = p.text.strip()
        if not t or p.in_table or p.in_toc or p.list_kind:
            continue
        if p.heading_level or (p.all_bold and len(t) <= 90):
            cands.add(p.idx)
    return sorted(i for i in cands if not info.paras[i].in_table and info.paras[i].text.strip())


def _neighbour(info, idx, step):
    i = idx + step
    while 0 <= i < len(info.paras):
        t = " ".join(info.paras[i].text.split())
        if t:
            return t[:60]
        i += step
    return ""


def read_headings(info, client, cands, budget):
    """{index: True/False} for the candidate lines; raises LLMError."""
    lines = [(i, f"[{i}] {' '.join(info.paras[i].text.split())[:LINE_CHARS]}\n     (before: {_neighbour(info, i, -1)} | after: {_neighbour(info, i, 1)})")
             for i in cands]
    out, proposed, rejected = {}, 0, 0
    for chunk in chunked(lines, budget):
        valid = {i for i, _ in chunk}
        reply = client.chat_json(HEAD_SYSTEM, "\n".join(l for _, l in chunk))
        if not isinstance(reply, dict):
            raise LLMError("the model's answer was not a JSON object")
        for key, val in (("headings", True), ("not_headings", False)):
            for v in _ints(reply.get(key)):
                proposed += 1
                if v is None or v not in valid:
                    rejected += 1
                elif v not in out:
                    out[v] = val
    return out, proposed, rejected


def read_slides(info, client):
    lines = []
    for s in info.slides_text:
        t = " | ".join(" ".join(x.split()) for x in s.texts if x.strip())
        lines.append((s.index, f"[S{s.index}] {t[:300]}"))
    reply = client.chat_json(FG_SYSTEM, "\n".join(l for _, l in lines))
    roles = reply.get("slides") if isinstance(reply, dict) else None
    if not isinstance(roles, dict):
        raise LLMError("the model's answer had no 'slides' object")
    valid = {i for i, _ in lines}
    out, rejected = {}, 0
    for k, v in roles.items():
        n = _int(k)
        if n in (None, "bad") or n not in valid or v not in FG_ROLES:
            rejected += 1
            continue
        out[n] = v
    return out, {"chunks": 1, "proposed": len(roles), "rejected": rejected}


# ------------------------------------------------------------------ reconciling with the parser
def _short_line(info, idx):
    return len(info.paras[idx].text.strip()) <= 60


def reconcile_lesson_plan(ctx, info, model):
    from .checks.formatting import own_number, parser_headings
    from .checks.lessonplan import parser_sections
    st = Structure(kind="lesson_plan")
    parsed = parser_sections(ctx, info)
    first = next((p.idx for p in info.paras if p.text.strip() and not p.in_table), None)
    sublabels = {normalize(x).lower() for x in ctx.profile["vocab"].get("slo_sublabels", [])}
    notes = [normalize(x).lower() for x in ctx.profile["vocab"].get("slo_block_end_labels", []) + ctx.profile["vocab"].get("note_labels", [])]
    verdicts = {int(k): v for k, v in (model.get("heading_verdicts") or {}).items()}
    if not verdicts:      # an older reading without the heading question: its list stands for the verdicts
        verdicts = {i: True for i in model["headings"]}

    def can_be_heading(i):
        p = info.paras[i]
        t = normalize(p.text).strip().lower()
        from .checks.formatting import OPTION_LIST
        if OPTION_LIST.search(t):
            return False
        return (not p.in_table and len(t) <= 150 and not is_arabic_scripture(p.text)
                and not (i == first and not p.heading_level) and not own_number(ctx, p.text)
                and t.rstrip(":：") not in sublabels and not t.endswith((":", "：")))

    def section_start_ok(i):
        # a section the parser did not find may begin at a heading or at a plain paragraph (a section without a
        # heading), never at the title, a list item, an SLO sub-label or a 'Note: ...' line
        p = info.paras[i]
        t = normalize(p.text).strip().lower()
        from .checks.lessonplan import is_bullet_text
        if p.list_kind or is_bullet_text(p.text) or i == first or t.rstrip(":：") in sublabels:
            return False
        if any(t.startswith(n) for n in notes):
            return False              # 'Note: ...', 'Instructions for students', 'For teachers: ...'
        from .textutil import starts_with_label
        if starts_with_label(p.text, ctx.section_labels):
            return True               # 'Warm-up: ...' is the section's own label with its text run on
        return not re.match(r"^[^:：]{1,30}[:：]", t)      # 'Duration: 40 minutes', 'Method: self-paced' are details

    parser_heads = {p.idx for p in parser_headings(ctx, info)}
    taken = {v for v in list(parsed.values()) + list(model["sections"].values()) if v is not None}

    def heading_above(i):
        """The model often points at a section's first line of content; its heading is the line just above.
        A line that is itself a heading stays where it is."""
        me = info.paras[i]
        if i in parser_heads or verdicts.get(i) or me.heading_level or (me.all_bold and len(me.text.strip()) <= 90):
            return i
        j = i - 1
        while j >= 0 and not info.paras[j].text.strip():
            j -= 1
        if j < 0 or j == first or j in taken or info.paras[j].in_table:
            return i
        q = info.paras[j]
        from .textutil import starts_with_label
        if (j in parser_heads or verdicts.get(j) or q.heading_level or starts_with_label(q.text, ctx.section_labels)) \
                and section_start_ok(j):
            return j                  # only to a heading line right above, never to a 'Duration: 40 min' line
        return i

    for k in SECTION_KEYS:
        p, m = parsed.get(k), model["sections"].get(k)
        if m is not None and (info.paras[m].in_table or is_arabic_scripture(info.paras[m].text)):
            m = None
        if m is not None:
            m = heading_above(m)
        if m is not None and section_start_ok(m):
            # the model decides; its pointer must pass the checks above
            st.sections[k], st.source[k] = m, ("both" if m == p else "model")
            if m != p:
                st.found_as[k] = info.paras[m].text.strip()[:80]
        else:
            if m is not None:
                st.notes.append(f"the model placed {k} at '{info.paras[m].text.strip()[:40]}', which cannot start a section; not used")
            # the model saw no such section (or pointed somewhere a section cannot start): its label at the start of a line counts
            use = p
            st.sections[k], st.source[k] = use, ("parser" if use is not None else "none")

    starts = {v for v in st.sections.values() if v is not None}
    model_heads = {i for i, v in verdicts.items() if v and can_be_heading(i)}
    # a heading line the model was not asked about (rare: it is not a candidate) keeps the parser's say
    unasked = {i for i in parser_heads if i not in verdicts}
    # 'Warm-up: "..." and the first line of the section' is a run-in heading; the fix splits it
    from .checks.formatting import run_in_split
    st.run_in = {i for i in starts if run_in_split(ctx, info.paras[i].text) is not None and not is_arabic_scripture(info.paras[i].text)}
    # a section with no heading starts at plain text, which is not a heading either
    starts = {i for i in starts if info.paras[i].heading_level or i in parser_heads or verdicts.get(i)
              or (can_be_heading(i) and len(info.paras[i].text.strip()) <= 90 and verdicts.get(i) is not False
                  and not re.search(r"[.?؟۔]$", info.paras[i].text.strip()))}
    # numbering starts at the first section: lines above it (title, duration, method) are the title block
    first_start = min((v for v in st.sections.values() if v is not None), default=0)
    st.headings = st.agreed_headings = {i for i in model_heads | unasked | starts if i >= first_start}
    st.disputed_headings = set()

    # items must lie inside their own section; anything else the model listed is dropped
    ranges = section_ranges(st.sections, len(info.paras))
    for key, sec in (("slo_items", "slos"), ("warm_up_questions", "warm_up"), ("takeaway_items", "key_takeaways")):
        r = ranges.get(sec)
        st.lists[key] = [i for i in model[key] if r and r[0] <= i < r[1] and i not in st.headings]
    st.glossary = model.get("glossary")
    return st


def section_ranges(sections, n_paras):
    found = sorted((v, k) for k, v in sections.items() if v is not None)
    return {k: (v, found[i + 1][0] if i + 1 < len(found) else n_paras) for i, (v, k) in enumerate(found)}


def parser_slide_roles(ctx, info):
    """{slide number: role} from the slide titles' label words; slides without a label are left out."""
    labels = {}
    for sec in ctx.profile["lesson_plan_sections"]:
        extra = ctx.profile["vocab"].get("fg_section_labels", {}).get(sec["key"], [])
        labels[sec["key"]] = [normalize(l).lower() for l in sec["labels"] + extra]
    out = {}
    for s in info.slides_text:
        t = normalize(s.title).lower()
        for key, labs in labels.items():
            if any(l in t for l in labs):
                out[s.index] = key
                break
    return out


def reconcile_slides(ctx, info, roles):
    """The model's role for each slide; the title's label words only for a slide the model gave no role."""
    st = Structure(kind="facilitator_guide")
    parsed = parser_slide_roles(ctx, info)
    for s in info.slides_text:
        p, m = parsed.get(s.index), roles.get(s.index)
        if m is not None:
            st.slide_roles[s.index], st.source[s.index] = m, ("both" if p == m else "model")
        elif p is not None:
            st.slide_roles[s.index], st.source[s.index] = p, "parser"
    for key in SECTION_KEYS:
        st.sections[key] = next((n for n in sorted(st.slide_roles) if st.slide_roles[n] == key), None)
    return st


def heading_findings(ctx, info):
    """A reviewer note on each line the two readings disagree about being a heading. The corrected copy
    leaves these lines as they are."""
    from .checks.common import mark, result
    from .models import REVIEW
    st = of(info)
    if st is None or st.kind != "lesson_plan" or not st.disputed_headings:
        return []
    marks = []
    for i in sorted(st.disputed_headings):
        t = info.paras[i].text
        if i in st.headings:
            marks.append(mark(t, "model suggests PASS: the model reads this line as a heading (it was not styled or numbered "
                                 "as one, so the tool left it as it is). If it is a heading, give it a Heading style and a number"))
        else:
            marks.append(mark(t, "model suggests FAIL: the model reads this line as text, not a heading, so the tool did not "
                                 "number it. If it is a heading, give it a Heading style and a number"))
    return [result("STR1", REVIEW, f"'{m['text'][:40]}' may or may not be a heading; the tool left it unchanged",
                   method="model", marks=[m]) for m in marks]


def transfer(st, old, new):
    """The structure of `old` carried over to `new`, a corrected copy of it (headings numbered, bullets added,
    captions and a Table of Contents inserted): paragraphs are matched by their text without numbers and bullets."""
    from difflib import SequenceMatcher
    if st is None:
        return None
    strip = re.compile(r"^\s*(?:[•▪●◦\-–—*]\s*)?(?:\(?[\d٠-٩۰-۹]+(?:\.[\d٠-٩۰-۹]+)*[.)]?\s+)?")

    def keys(info):
        ps = [p for p in info.paras if p.text.strip()]
        return ps, [strip.sub("", normalize(p.text).strip().lower()) for p in ps]
    a, ka = keys(old)
    b, kb = keys(new)
    m = {}
    for blk in SequenceMatcher(None, ka, kb, autojunk=False).get_matching_blocks():
        for k in range(blk.size):
            m[a[blk.a + k].idx] = b[blk.b + k].idx

    # a paragraph the fix split in two (a heading with its text run on) matches the new paragraph its text starts with
    ka_by, kb_by = {p.idx: k for p, k in zip(a, ka)}, {p.idx: k for p, k in zip(b, kb)}
    taken = set(m.values())
    for i in sorted(set(ka_by) - set(m)):
        prev = max((m[j] for j in m if j < i), default=-1)
        nxt = min((m[j] for j in m if j > i), default=10 ** 9)
        cand = [j for j in sorted(kb_by) if prev < j < nxt and j not in taken and len(kb_by[j]) >= 2
                and ka_by[i].startswith(kb_by[j].rstrip(":： "))]
        if cand:
            m[i] = cand[0]
            taken.add(cand[0])

    def one(i):
        if i is None:
            return None
        if i in m:
            return m[i]
        later = [m[j] for j in sorted(m) if j > i]
        return later[0] if later else None
    out = Structure(kind=st.kind, source=dict(st.source), found_as=dict(st.found_as), notes=list(st.notes), stats=dict(st.stats),
                    slide_roles=dict(st.slide_roles))
    if st.kind == "lesson_plan":
        out.sections = {k: one(v) for k, v in st.sections.items()}
        out.headings = {m[i] for i in st.headings if i in m}
        out.agreed_headings = {m[i] for i in st.agreed_headings if i in m}
        out.disputed_headings = {m[i] for i in st.disputed_headings if i in m}
        out.run_in = set()                               # split by the fix: the label line is now a heading of its own
        out.headings |= {m[i] for i in st.run_in if i in m}
        out.agreed_headings |= {m[i] for i in st.run_in if i in m}
        out.lists = {k: [m[i] for i in v if i in m] for k, v in st.lists.items()}
        out.glossary = one(st.glossary)
    else:
        out.sections = dict(st.sections)
    return out


# ------------------------------------------------------------------ entry point
def read(ctx, doc, info):
    """The reconciled Structure of one document, or None (no model, unsupported type, or the model failed;
    the checks then use the parser alone, as before). Cached per file, model and prompt version."""
    model = ctx.model
    if model is None or not getattr(model, "structure", True):
        return None
    from .ingest import GENERIC
    kind = doc.doc_type if doc.doc_type in ("lesson_plan", "facilitator_guide") else None
    if doc.doc_type == GENERIC and doc.ext == "docx":
        kind = "lesson_plan"          # a document nothing else could place: the model may still see a Lesson Plan in it
    if kind is None or (kind == "lesson_plan" and doc.ext != "docx") or (kind == "facilitator_guide" and doc.ext != "pptx"):
        return None
    from .judge import material_budget
    key = f"structure|{PROMPT_VERSION}|{model.client.name}|{file_sha1(doc.abs)}"
    hit = model.cache.get(key) if model.cache is not None else None
    t0 = time.time()
    if hit is None:
        try:
            if kind == "lesson_plan":
                reading, counts = read_lesson_plan(ctx, info, model.client, material_budget(model))
            else:
                roles, counts = read_slides(info, model.client)
                reading = {"slides": {str(k): v for k, v in roles.items()}}
        except LLMError as e:
            model.stats.append({"doc": doc.rel, "structure": True, "usable": False, "note": f"structure not read: {e}"})
            return None
        hit = {"reading": reading, "counts": counts}
        usable = not counts["proposed"] or counts["rejected"] / counts["proposed"] <= MAX_REJECTED_SHARE
        hit["usable"] = usable
        if model.cache is not None:
            model.cache.put(key, hit)
    counts, usable = hit["counts"], hit["usable"]
    model.stats.append({"doc": doc.rel, "structure": True, "usable": usable, "seconds": round(time.time() - t0, 1),
                        "cached": time.time() - t0 < 0.5, **counts,
                        "note": "" if usable else f"{counts['rejected']} of {counts['proposed']} pointers were not in the document; reading discarded"})
    if not usable:
        return None
    if kind == "lesson_plan":
        st = reconcile_lesson_plan(ctx, info, hit["reading"])
        if doc.doc_type == GENERIC:
            found = [k for k, v in st.sections.items() if v is not None]
            if len(found) < 3:
                return None           # not a Lesson Plan: it keeps the general checks
            doc.doc_type, doc.scope = "lesson_plan", "lesson"
            doc.chapter = doc.chapter if doc.chapter is not None else 1
            doc.lesson = doc.lesson if doc.lesson is not None else 1
            doc.notes = [n for n in doc.notes if not n.startswith(("Recognised as", "converted"))] + \
                [f"Recognised as lesson plan by the model: it found {len(found)} of the five sections under other names. "
                 "Lesson and chapter numbers were assumed as 1 where the file does not give them."]
    else:
        st = reconcile_slides(ctx, info, {int(k): v for k, v in hit["reading"]["slides"].items()})
    st.stats = counts
    return st


def attach(ctx, doc, info):
    """Read the structure once and keep it on the parsed document, where the checks look for it."""
    if info is not None and not hasattr(info, "structure"):
        info.structure = read(ctx, doc, info)
    return getattr(info, "structure", None)


def of(info):
    return getattr(info, "structure", None)


def dump(st):
    return json.dumps({"kind": st.kind, "sections": st.sections, "source": st.source, "headings": sorted(st.headings),
                       "disputed_headings": sorted(st.disputed_headings), "lists": st.lists, "slides": st.slide_roles},
                      ensure_ascii=False)
