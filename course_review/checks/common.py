"""Helpers shared by the check modules."""
import re
from collections import Counter

from ..models import FAIL, NA, PASS, REVIEW, Finding


def result(code, status, message="", evidence=None, partial=False, doc=None, lesson=None, chapter=None, method="deterministic",
           marks=None):
    return Finding(code=code, status=status, message=message, evidence=list(evidence or []),
                   partial=partial, doc=doc, lesson=lesson, chapter=chapter, method=method, marks=list(marks or []))


def mark(text, note="", slide=None, exact=True):
    """A place in the document to highlight. exact: `text` is a whole paragraph (one paragraph is
    marked); otherwise it is a quote, and every paragraph it touches is marked."""
    m = {"text": (text or "").strip(), "note": note, "exact": exact}
    if slide is not None:
        m["slide"] = slide
    return m


def canon_font(name):
    return re.sub(r"[\s\-_]", "", (name or "").lower())


def is_grey(hex_rgb, tolerance):
    h = (hex_rgb or "").lstrip("#")
    if len(h) != 6:
        return True
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        return True
    return max(r, g, b) - min(r, g, b) <= tolerance


def top(counter: Counter, n=4, fmt=lambda k, v: f"{k}: {v}"):
    return ", ".join(fmt(k, v) for k, v in counter.most_common(n))


def pct(part, whole):
    return f"{100.0 * part / whole:.0f}%" if whole else "0%"
