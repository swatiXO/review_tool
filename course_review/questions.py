"""Extract questions from Pop Quiz / exam / worksheet documents and items from a Data Bank.

Everything language-specific (prefixes, option letters, field names) comes from the
profile vocabulary, so a new subject or language needs a profile change, not code.
"""
import re
from dataclasses import dataclass, field
from typing import Optional

from docx import Document
from docx.oxml.ns import qn

from .textutil import normalize, to_western_digits

CHECKS = "✅✔☑"
OPTION_START = re.compile(r"^\s*[✅✔☑]?\s*\(?(?:الف|ب|ج|د|[A-Da-d])\s*[).]")
OPTION_INLINE = re.compile(r"(?:^|\s)[✅✔☑]?\s*\(?(?:الف|ب|ج|د|[A-Da-d])\s*\)")
QUESTION_PAREN = re.compile(r"^\s*\((\d{1,3})\)\s*\S")           # (1) text
QUESTION_SHORT = re.compile(r"^\s*q\.?\s*(\d{1,3})(?!\d)\s*\S", re.I)  # Q1 text, Q.1 text


def is_yellow(hex_fill: Optional[str]) -> bool:
    if not hex_fill or hex_fill.lower() in ("auto", "none"):
        return False
    h = hex_fill.lstrip("#")
    if len(h) != 6:
        return False
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        return False
    return r >= 200 and g >= 190 and b <= 170


def run_is_highlighted(run) -> bool:
    return (run.highlight or "").lower() == "yellow" or is_yellow(run.shading)


def para_is_highlighted(p) -> bool:
    return any(run_is_highlighted(r) for r in p.runs) or is_yellow(p.shading)


@dataclass
class Option:
    text: str
    highlighted: bool
    check: bool


@dataclass
class Question:
    num: int
    text: str
    start: int
    end: int
    options: list = field(default_factory=list)
    qtype: str = "open"
    highlighted: bool = False
    check_marked: bool = False
    block: str = ""
    inline_options: int = 0
    source: str = "rules"          # 'rules' or 'model'
    meta: Optional[dict] = None    # verification stats when source == 'model'


def question_regex(profile, labelled=False):
    """A question start: optional label (سوال / Question) + number + separator.
    labelled=True requires the label, used inside tables where bare numbers are answers."""
    prefixes = "|".join(re.escape(normalize(p)) for p in profile["vocab"]["question_prefixes"])
    label = rf"(?:(?:{prefixes})\s*[.:\-–—]?\s*)"
    sep = r"(?:[-–—:.)\]۔]|\s-\s)" if not labelled else r"(?:[-–—:.)\]۔(]|\s-\s)"
    head = label if labelled else label + "?"
    return re.compile(rf"^\s*{head}(\d{{1,3}})\s*{sep}\s*(\S.*)?$", re.I | re.S)


def table_question_regex(profile):
    """Inside a table a question cell reads 'سوال 1  قسم: ...': label + number, no separator."""
    prefixes = "|".join(re.escape(normalize(p)) for p in profile["vocab"]["question_prefixes"])
    return re.compile(rf"^\s*(?:{prefixes})\s*[.:\-–—]?\s*(\d{{1,3}})\b", re.I)


def lesson_heading_regex(profile):
    prefixes = "|".join(re.escape(normalize(p)) for p in profile["vocab"]["lesson_heading_prefixes"])
    return re.compile(rf"^\s*(?:{prefixes})\s*(\d+)\s*(?:\(\s*([^)]*?)\s*\))?", re.I)


def classify_type(block: str, option_count: int) -> str:
    n = normalize(block).lower()
    if option_count >= 3:
        return "mcq"
    if re.search(r"درست\s*(?:یا|/)\s*غلط|true\s*/?\s*(?:or)?\s*false", n):
        return "true_false"
    if "____" in block or "خالی جگہ" in n or "fill in" in n:
        return "fill_blank"
    if "جوڑ" in n or "match" in n:
        return "matching"
    return "open"


ROMAN = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100}


def question_number(text: str):
    """An int from digits ('12') or a roman numeral ('xii'); None if it is neither."""
    t = to_western_digits(text.strip())
    if t.isdigit():
        return int(t)
    t = t.lower()
    if t and all(ch in ROMAN for ch in t):
        total = 0
        for i, ch in enumerate(t):
            v = ROMAN[ch]
            total += -v if i + 1 < len(t) and ROMAN[t[i + 1]] > v else v
        return total if 0 < total < 200 else None
    return None


def compile_extra_patterns(profile):
    out = []
    for entry in profile["vocab"].get("extra_question_patterns", []) or []:
        pattern = entry["pattern"] if isinstance(entry, dict) else entry
        try:
            out.append(re.compile(pattern, re.I))
        except re.error:
            continue
    return out


def parse_questions(info, profile, lo: int = 0, hi: Optional[int] = None):
    """Questions among paragraphs with index in [lo, hi). Table paragraphs only count
    when they carry an explicit question label."""
    qre = question_regex(profile)
    qre_lab = question_regex(profile, labelled=True)
    qre_tab = table_question_regex(profile)
    lre = lesson_heading_regex(profile)
    extra = compile_extra_patterns(profile)
    end_labels = [normalize(x).lower() for x in profile["vocab"].get("question_region_end_labels", [])]
    paras = [p for p in info.paras if p.idx >= lo and (hi is None or p.idx < hi)]
    starts, seen_q = [], False
    for p in paras:
        t = to_western_digits(normalize(p.text))
        if not t:
            continue
        # teacher guidance / answer sections after the questions are not questions
        if seen_q and not p.in_table and len(t) <= 90 and any(t.lower().startswith(l) for l in end_labels):
            hi = p.idx
            starts.append((p.idx, None))
            break
        if not p.in_table and lre.match(t) and not OPTION_START.match(t):
            starts.append((p.idx, None))  # a lesson heading ends the previous question
            continue
        if p.heading_level == 1 and not p.in_table:
            starts.append((p.idx, None))  # so does a top-level heading
            continue
        m = qre_lab.match(t)
        labelled = bool(m)
        if m is None and p.in_table:
            m = qre_tab.match(t)
            labelled = bool(m)
        if m is None and not p.in_table:
            m = qre.match(t) or QUESTION_PAREN.match(t) or QUESTION_SHORT.match(t)
        if m is None:
            for rx in extra:                      # formats learned with `learn-format`
                m = rx.match(t)
                if m:
                    break
        if m and not OPTION_START.match(t):
            number = question_number(m.group(1))
            if number is not None:
                starts.append((p.idx, number, labelled))
                seen_q = True
    # header lines such as '31: lesson title' before the first labelled question are not questions
    first_lab = next((st[0] for st in starts if len(st) == 3 and st[2]), None)
    if first_lab is not None:
        starts = [st for st in starts if len(st) == 2 or st[2] or st[0] > first_lab]
    starts = [(st[0], st[1]) for st in starts]
    if not any(n is not None for _, n in starts):
        # no typed numbers: fall back to automatic list numbering ("1." generated by Word)
        for p in paras:
            if not p.in_table and p.list_kind == "decimal" and p.num_label:
                m = re.match(r"\s*(\d+)", to_western_digits(p.num_label))
                if m:
                    starts.append((p.idx, int(m.group(1))))
        starts.sort()
    return questions_from_starts(info, starts, hi)


def questions_from_starts(info, starts, hi: Optional[int] = None, source: str = "rules"):
    """Build Question objects from [(paragraph index, number|None)] question starts.
    An entry with number None only marks where the previous question ends. Used by the
    rule-based parser and by the model fallback, so both produce identical objects."""
    questions = []
    for i, (idx, num) in enumerate(starts):
        if num is None:
            continue
        end = starts[i + 1][0] if i + 1 < len(starts) else (hi if hi is not None else len(info.paras))
        block_paras = [p for p in info.paras if idx <= p.idx < end]
        q = Question(num=num, text=block_paras[0].text.strip(), start=idx, end=end, source=source)
        q.block = "\n".join(p.text for p in block_paras)
        for p in block_paras[1:]:
            if not p.in_table and OPTION_START.match(p.text.strip()):
                q.options.append(Option(p.text.strip(), para_is_highlighted(p), any(c in p.text for c in CHECKS)))
        q.inline_options = len(OPTION_INLINE.findall(q.block)) if not q.options else 0
        count = len(q.options) or q.inline_options
        q.qtype = classify_type(q.block, count)
        q.highlighted = any(o.highlighted for o in q.options) or (
            q.inline_options >= 3 and any(para_is_highlighted(p) for p in block_paras))
        q.check_marked = any(c in q.block for c in CHECKS)
        questions.append(q)
    return questions


def numbering_is_regular(questions) -> bool:
    """Questions are numbered 1..n once each; anything else means the document
    nests numbered sub-questions or restarts numbering, and the count is ambiguous."""
    return [q.num for q in questions] == list(range(1, len(questions) + 1))


def segment_lessons(info, profile):
    """Split a multi-lesson document (Pop Quiz) into chapter groups of lessons.

    Lesson numbers restart in each chapter, so a number that goes down (or repeats
    without a part letter) starts a new group. Returns
    [ [ {num, variant, start, end, title}, ... ], ... ]"""
    lre = lesson_heading_regex(profile)
    letters = {normalize(k): v for k, v in profile["vocab"]["variant_letters"].items()}
    heads = []
    for p in info.paras:
        if p.in_table:
            continue
        t = to_western_digits(normalize(p.text))
        m = lre.match(t)
        if m and len(t) < 160 and not OPTION_START.match(t):
            raw = normalize(m.group(2) or "")
            heads.append({"num": int(m.group(1)), "variant": letters.get(raw, raw.upper()[:1] if raw.isascii() else ""),
                          "start": p.idx, "title": p.text.strip()})
    for i, h in enumerate(heads):
        h["end"] = heads[i + 1]["start"] if i + 1 < len(heads) else len(info.paras)
    groups, prev = [], None
    for h in heads:
        new = prev is None or h["num"] < prev["num"] or (h["num"] == prev["num"] and not h["variant"] and not prev["variant"])
        if new:
            groups.append([])
        groups[-1].append(h)
        prev = h
    return groups


def chapter_from_text(text: str, profile) -> Optional[int]:
    """'باب دوم: ...' or 'Chapter 3' -> chapter number."""
    t = to_western_digits(normalize(text))
    m = re.search(r"(?:chapter|باب)\s*(\d+)", t, re.I)
    if m:
        return int(m.group(1))
    for n, words in profile["vocab"]["chapter_ordinals"].items():
        for w in words:
            if re.search(rf"(?:chapter|باب)\s*{re.escape(normalize(w))}\b", t, re.I):
                return int(n)
    return None


# ----------------------------------------------------------------- data bank


@dataclass
class BankItem:
    chapter: Optional[int]
    lesson_no: Optional[int]
    fields: dict            # canonical field -> text
    present: set            # canonical fields that exist as keys
    highlighted: bool
    raw_keys: list
    order: int


def _canon_map(profile):
    out = {}
    for canon, names in profile["vocab"]["data_bank_fields"].items():
        for n in names:
            out[normalize(n).lower()] = canon
    return out


def _cell_highlighted(cell):
    tc = cell._tc
    return any((h.get(qn("w:val")) or "").lower() == "yellow" for h in tc.iter(qn("w:highlight"))) or \
        any(is_yellow(s.get(qn("w:fill"))) for s in tc.iter(qn("w:shd")))


def _bank_item(fields, present, raw_keys, hl, order, profile):
    chapter = lesson_no = None
    cl = fields.get("chapter_lesson", "")
    if cl:
        parts = re.split(r"\s*/\s*|\n", cl)
        chapter = chapter_from_text(parts[0], profile)
        m = re.match(r"\s*(\d+)", to_western_digits(normalize(parts[1] if len(parts) > 1 else "")))
        lesson_no = int(m.group(1)) if m else None
    else:
        if fields.get("chapter"):
            chapter = chapter_from_text("chapter " + fields["chapter"], profile) or chapter_from_text(fields["chapter"], profile)
        m = re.search(r"(\d+)", to_western_digits(normalize(fields.get("lesson", ""))))
        lesson_no = int(m.group(1)) if m else None
    return BankItem(chapter, lesson_no, fields, present, hl, raw_keys, order)


def load_bank_items(path, profile, reading=None):
    """Data Bank items. reading: (layout, {field name: canonical field}) from the model (layout.read_bank_fields);
    without it the profile's field names are used. Two layouts are read: one table per item with the field names
    in the first column, and one table with the field names in its header row and an item in each further row."""
    doc = Document(path)
    cmap = dict(reading[1]) if reading else _canon_map(profile)
    items = []
    # header-row layout: a table whose first row names at least three fields, one of them the question
    for ti, t in enumerate(doc.tables):
        rows = t.rows
        if len(rows) < 2 or len(rows[0].cells) < 4:
            continue
        head = [cmap.get(normalize(c.text).lower()) for c in rows[0].cells]
        if "question" not in head or sum(1 for h in head if h) < 3:
            continue
        raw = [c.text.strip() for c in rows[0].cells]
        for ri, row in enumerate(rows[1:], 1):
            cells = row.cells
            fields, present, hl = {}, set(), False
            for canon, cell in zip(head, cells):
                if canon is None or canon == "item":
                    continue
                present.add(canon)
                fields[canon] = cell.text.strip()
                if canon in ("options", "answer") and _cell_highlighted(cell):
                    hl = True
            if not fields.get("question"):
                continue
            items.append(_bank_item(fields, present, raw, hl, ti * 1000 + ri, profile))
    if items:
        return items
    for ti, t in enumerate(doc.tables):
        fields, present, raw_keys, hl = {}, set(), [], False
        for row in t.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            key = normalize(cells[0].text).lower()
            raw_keys.append(cells[0].text.strip())
            canon = cmap.get(key)
            if canon is None:
                continue
            if canon == "item":
                continue
            present.add(canon)
            fields[canon] = cells[1].text.strip()
            if canon in ("options", "answer") and _cell_highlighted(cells[1]):
                hl = True
        if not present:
            continue
        items.append(_bank_item(fields, present, raw_keys, hl, ti, profile))
    return items
