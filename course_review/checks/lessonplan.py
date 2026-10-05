"""Lesson Plan structure checks: LP3 (five mandatory sections, in order), LP4 (decimal
heading numbering) and LP5 (SLOs as bullets). WE14 is the same rule as LP3, so it is
reported from the same result."""
import re

from ..docx_model import DocxInfo
from ..models import FAIL, NA, PASS, REVIEW
from ..textutil import normalize, starts_with_label, to_western_digits
from .common import result
from .formatting import we12


def find_sections(ctx, info: DocxInfo):
    """first paragraph index of each mandatory section, or None."""
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
    return result("LP3", PASS, "All five sections present in order")


def lp4(ctx, doc, info):
    r = we12(ctx, doc, info)
    r.code = "LP4"
    return r


def lp5(ctx, doc, info: DocxInfo):
    found = find_sections(ctx, info)
    start = found.get("slos")
    if start is None:
        return result("LP5", REVIEW, "The SLO section was not found, so the SLO list could not be checked")
    later = [v for k, v in found.items() if v is not None and v > start]
    end = min(later) if later else len(info.paras)
    items = []
    for p in info.paras[start + 1:end]:
        t = p.text.strip()
        if not t or p.in_table:
            continue
        if re.search(r"[:：]$", t) and len(t) < 40:
            continue  # a sub-label such as 'Knowledge:'
        items.append(p)
    if not items:
        return result("LP5", REVIEW, "No SLO items found under the SLO heading")
    bad = []
    for p in items:
        numbered = p.list_kind == "decimal" or bool(re.match(r"^\s*\d+[.)]\s", to_western_digits(p.text)))
        if numbered:
            bad.append(("numbered", p))
        elif p.list_kind != "bullet" and not re.match(r"^\s*[•\-–—*▪●◦]\s", p.text):
            bad.append(("not a bullet", p))
    if bad:
        ev = [f"'{p.text.strip()[:40]}' is {why}" for why, p in bad[:6]]
        return result("LP5", FAIL, f"{len(bad)} of {len(items)} SLO items are not bullets ({bad[0][0]})", ev)
    return result("LP5", PASS, f"All {len(items)} SLO items are bullets")


DOC_CHECKS = {"LP3": lp3, "LP4": lp4, "LP5": lp5}
