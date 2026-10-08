"""Facilitator Guide checks decidable from the slides: FG2 (split into on-slide content and
facilitator notes, nothing copy-pasted) and FG7 (slide structure follows the Lesson Plan's
sections). Both are partial: they check structure, not whether the content is right."""
from difflib import SequenceMatcher

from ..models import FAIL, NA, PASS, REVIEW
from ..textutil import normalize
from .common import mark, result

COPY_MIN_CHARS = 60
COPY_SIMILARITY = 0.9


def _lesson_plan_paras(ctx, doc):
    lp = next((d for d in ctx.pkg.docs if d.doc_type == "lesson_plan" and not d.superseded and d.key == doc.key and d.ext == "docx"), None)
    if lp is None:
        return None
    info = ctx.docx(lp)
    if info is None:
        return None
    # the SLOs are meant to be restated on the slides, so they are not counted as copying (only the SLO statements:
    # other text in that section, such as prior knowledge, still is)
    from .lessonplan import slo_items
    items, _ = slo_items(ctx, info)
    skip = {p.idx for p in items or []}
    return [normalize(p.text) for p in info.paras if len(p.text.strip()) >= COPY_MIN_CHARS and p.idx not in skip]


def _has_facilitator_part(slide, labels):
    if slide.notes:
        return True
    return any(normalize(t).lower().startswith(l) or l in normalize(t).lower()[:60] for t in slide.texts for l in labels)


def fg2(ctx, doc, info):
    slides = info.slides_text[1:]            # the first slide is the cover
    if not slides:
        return result("FG2", NA, "The guide has no slides after the cover")
    labels = [normalize(x).lower() for x in ctx.profile["vocab"].get("facilitator_note_labels", [])]
    missing = [s.index for s in slides if not _has_facilitator_part(s, labels)]
    plan = _lesson_plan_paras(ctx, doc)
    copied, marks = [], []
    if plan:
        # a text box holds several paragraphs, so compare it line by line with the plan's paragraphs
        for s in slides:
            for t in s.texts + ([s.notes] if s.notes else []):
                for line in t.split("\n"):
                    nt = normalize(line).strip(" •-–—*")
                    if len(nt) < COPY_MIN_CHARS:
                        continue
                    for pt in plan:
                        if abs(len(pt) - len(nt)) <= 0.25 * len(nt) and SequenceMatcher(None, nt, pt).quick_ratio() >= COPY_SIMILARITY \
                                and SequenceMatcher(None, nt, pt).ratio() >= COPY_SIMILARITY:
                            copied.append(s.index)
                            marks.append(mark(line, "copied word for word from the Lesson Plan", slide=s.index))
                            break
    if copied:
        msg = f"slide(s) {sorted(set(copied))} repeat {len(copied)} Lesson Plan paragraph(s) almost word for word"
        return result("FG2", FAIL, msg, [msg], marks=marks)
    if missing:
        msg = (f"No labelled facilitator-notes part and no speaker notes on slide(s) {missing}. The guide may mark facilitator notes "
               f"another way (labels looked for: {', '.join(ctx.profile['vocab']['facilitator_note_labels'][:3])}...), so a reviewer should look")
        return result("FG2", REVIEW, msg)
    return result("FG2", PASS, "Every slide has a facilitator part and no Lesson Plan paragraph is copied", partial=True,
                  evidence=["Whether the on-slide content is cleanly separated from the notes needs a reviewer."]
                  + ([] if plan else ["The Lesson Plan was not found, so copy-pasting was not checked."]))


def fg7(ctx, doc, info):
    from ..structure import of, parser_slide_roles
    sections = ctx.profile["lesson_plan_sections"]
    st = of(info)
    if st is not None:
        found = dict(st.sections)
    else:
        roles = parser_slide_roles(ctx, info)
        found = {sec["key"]: next((n for n in sorted(roles) if roles[n] == sec["key"]), None) for sec in sections}
    keys = [s["key"] for s in sections]
    missing = [k for k in keys if found[k] is None]
    present = [(found[k], k) for k in keys if found[k] is not None and k != "slos"]
    out_of_order = [present[i][1] for i in range(1, len(present)) if present[i][0] < present[i - 1][0]]
    # the SLOs are stated at the Session Overview, so that slide may come first; it must precede Concept Building
    if found.get("slos") is not None and found.get("concept_building") is not None and found["slos"] > found["concept_building"]:
        out_of_order.append("slos")
    if missing or out_of_order:
        ev = []
        if missing:
            ev.append("no slide for: " + ", ".join(missing))
        if out_of_order:
            ev.append("out of order: " + ", ".join(out_of_order))
        return result("FG7", FAIL, "; ".join(ev), ev)
    return result("FG7", PASS, "Slides follow the Lesson Plan's five sections in order", partial=True,
                  evidence=["Slide titles were matched to section names; whether the content under each matches needs a reviewer."])


DOC_CHECKS = {"FG2": fg2, "FG7": fg7}
