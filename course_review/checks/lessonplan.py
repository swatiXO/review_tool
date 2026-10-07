"""Lesson Plan structure checks: LP3 (five mandatory sections, in order), LP4 (decimal
heading numbering) and LP5 (SLOs as bullets). WE14 is the same rule as LP3, so it is
reported from the same result."""
import re

from ..docx_model import DocxInfo
from ..models import FAIL, NA, PASS, REVIEW
from ..textutil import normalize, starts_with_label, to_western_digits
from .common import mark, result
from .formatting import we12


def find_sections(ctx, info: DocxInfo):
    """First paragraph index of each mandatory section, or None: the model's reading reconciled with the
    parser's when the structure was read (structure.py), the parser's alone otherwise."""
    from ..structure import of
    st = of(info)
    if st is not None:
        return dict(st.sections)
    return parser_sections(ctx, info)


def parser_sections(ctx, info: DocxInfo):
    """First paragraph index of each mandatory section by its label words, or None."""
    found = {}
    for sec in ctx.profile["lesson_plan_sections"]:
        for p in info.paras:
            if p.in_table or p.in_toc:
                continue
            if starts_with_label(p.text, sec["labels"]):
                found[sec["key"]] = p.idx
                break
        else:
            found[sec["key"]] = None
    return found


def lp3(ctx, doc, info):
    found = find_sections(ctx, info)
    keys = [s["key"] for s in ctx.profile["lesson_plan_sections"]]
    missing = [k for k in keys if found[k] is None]
    present = [(found[k], k) for k in keys if found[k] is not None]
    out_of_order = [present[i][1] for i in range(1, len(present)) if present[i][0] < present[i - 1][0]]
    if missing or out_of_order:
        ev = []
        if missing:
            ev.append("missing: " + ", ".join(missing))
        if out_of_order:
            ev.append("out of order: " + ", ".join(out_of_order))
        return result("LP3", FAIL, "; ".join(ev), ev)
    from ..structure import of
    st = of(info)
    if st is not None and st.found_as:
        names = {s["key"]: s["labels"][0] for s in ctx.profile["lesson_plan_sections"]}
        return result("LP3", PASS, "All five sections present in order", evidence=[
            f"{names[k]} recognised by its content: '{t[:50]}'" for k, t in st.found_as.items()])
    return result("LP3", PASS, "All five sections present in order")


def lp4(ctx, doc, info):
    r = we12(ctx, doc, info)
    r.code = "LP4"
    return r


def is_bullet_text(text):
    """A typed bullet: '• item' or '•item' (no space is common in Urdu files), '- item'."""
    return bool(re.match(r"^\s*(?:[•▪●◦]|[\-–—*]\s)", text or ""))


def slo_blocks(ctx, info: DocxInfo):
    """[(heading index, end index)] for every SLO heading in the document. A block ends at the next
    section heading, or at the next heading that is not an SLO sub-label (Knowledge / Skills), so a
    document that is not built on the five sections does not turn into one long SLO list."""
    from .formatting import doc_headings
    slo_labels = next(s["labels"] for s in ctx.profile["lesson_plan_sections"] if s["key"] == "slos")
    all_labels = [l for s in ctx.profile["lesson_plan_sections"] for l in s["labels"]]
    sub = [normalize(x).lower().rstrip(":：") for x in ctx.profile["vocab"].get("slo_sublabels", [])]
    heads = {p.idx for p in doc_headings(ctx, info)}
    starts = [p.idx for p in info.paras if not p.in_table and not p.in_toc and starts_with_label(p.text, slo_labels)]
    out = []
    for st in starts:
        end = len(info.paras)
        for p in info.paras[st + 1:]:
            t = normalize(p.text).strip().lower().rstrip(":：")
            if p.in_table or not t:
                continue
            if starts_with_label(p.text, all_labels) or (p.idx in heads and t not in sub):
                end = p.idx
                break
        out.append((st, end))
    return out


def slo_items(ctx, info: DocxInfo):
    """(the Lesson Plan's SLO statements as paragraphs, number of SLO sections) or (None, 0) when there is no
    SLO section. The model's list when it read the structure; otherwise the lines under the SLO heading, without
    the lead-in, the 'Knowledge:' labels and what follows the list (notes, instructions, other headings)."""
    from ..structure import of
    st = of(info)
    if st is not None and st.sections.get("slos") is not None:
        return [info.paras[i] for i in st.lists.get("slo_items", [])], 1
    blocks = slo_blocks(ctx, info)
    if not blocks:
        return None, 0
    sub = [normalize(x).lower().rstrip(":：") for x in ctx.profile["vocab"].get("slo_sublabels", [])]
    stops = [normalize(x).lower() for x in ctx.profile["vocab"].get("slo_block_end_labels", [])]
    leads = [normalize(x).lower() for x in ctx.profile["vocab"].get("slo_lead_ins", [])]
    items = []
    for start, end in blocks:
        for p in info.paras[start + 1:end]:
            t = p.text.strip()
            if not t or p.in_table:
                continue
            if re.search(r"[:：]$", t) or normalize(t).lower().rstrip(":：") in sub:
                continue  # a sub-label such as 'Knowledge:' or a lead-in such as 'By the end you will be able to:'
            head = normalize(t).lower()[:40]
            if not items and any(head.startswith(x) for x in leads):
                continue  # 'By the end of this lesson you will ...' introduces the SLOs
            if any(head.startswith(x) or (x in head and ":" in t[:45]) for x in stops):
                break     # 'Note:', 'طلبہ کے لیے ہدایت:', 'برائے اساتذہ' ... end the SLO list; what follows is not SLOs
            items.append(p)
    return items, len(blocks)


def lp5(ctx, doc, info: DocxInfo):
    items, n = slo_items(ctx, info)
    if items is None:
        return result("LP5", REVIEW, "The SLO section was not found, so the SLO list could not be checked")
    return _lp5_verdict(items, f" in {n} SLO sections" if n > 1 else "")


def _lp5_verdict(items, where):
    if not items:
        return result("LP5", REVIEW, "No SLO items found under the SLO heading")
    bad = []
    for p in items:
        numbered = p.list_kind == "decimal" or bool(re.match(r"^\s*\d+[.)]\s", to_western_digits(p.text)))
        if numbered:
            bad.append(("numbered", p))
        elif p.list_kind != "bullet" and not is_bullet_text(p.text):
            bad.append(("not a bullet", p))
    if bad:
        ev = [f"'{p.text.strip()[:40]}' is {why}" for why, p in bad[:6]]
        return result("LP5", FAIL, f"{len(bad)} of {len(items)} SLO items{where} are not bullets ({bad[0][0]})", ev,
                      marks=[mark(p.text, f"SLO item is {why}, should be a bullet") for why, p in bad])
    return result("LP5", PASS, f"All {len(items)} SLO items{where} are bullets")


DOC_CHECKS = {"LP3": lp3, "LP4": lp4, "LP5": lp5}
