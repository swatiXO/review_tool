"""SLOs from the Document of Specifications, and SLO tags found in questions.

The specification lists, per chapter, a table with the columns Lesson #, Lesson, SLOs,
Bloom's level and content coverage. SLOs are sentences with no codes, so a question
can be tied to an SLO only by (a) an explicit marker such as "SLO 2" or "L3 SLO 2", or
(b) quoting the SLO sentence. When neither exists, coverage is not decidable by code.
"""
import re
from dataclasses import dataclass

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table

from .textutil import normalize, to_western_digits

CHAPTER_LABEL = re.compile(r"^\s*chapter\s*(\d+)", re.I)
SLO_CODE = re.compile(r"[A-Z]-\d{2}-[A-Z]-\d{2}")


@dataclass
class Slo:
    chapter: int
    lesson: int
    lesson_name: str
    index: int                 # 1-based position within the lesson
    text: str
    bloom: str = ""
    coverage_note: str = ""

    @property
    def key(self):
        return (self.chapter, self.lesson, self.index)

    def label(self) -> str:
        return f"Ch{self.chapter} L{self.lesson} SLO {self.index}"


def load_slos(spec_path, plan=None):
    """{(chapter, lesson): [Slo]}. plan: {table number: {chapter, lesson_no, lesson_name, slo, bloom}} from the model
    (layout.read_spec_tables); without it, every table whose header has an 'SLOs' column, under a 'Chapter N' line."""
    doc = Document(spec_path)
    if plan:
        return _slos_from_plan(doc, plan)
    out, chapter = {}, None
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            text = to_western_digits(normalize("".join(t.text or "" for t in child.iter(qn("w:t")))))
            m = CHAPTER_LABEL.match(text)
            if m:
                chapter = int(m.group(1))
        elif tag == "tbl":
            table = Table(child, doc)
            header = [normalize(c.text).lower() for c in table.rows[0].cells]
            col = {}
            for i, h in enumerate(header):
                if h.startswith("lesson #") or h == "lesson no":
                    col["no"] = i
                elif h == "lesson":
                    col["name"] = i
                elif h.startswith("slo"):
                    col["slo"] = i
                elif h.startswith("bloom"):
                    col["bloom"] = i
                elif "coverage" in h:
                    col["cov"] = i
            if "slo" not in col or "no" not in col or chapter is None:
                continue
            last_no, last_name = None, ""
            for row in table.rows[1:]:
                cells = row.cells
                m = re.match(r"\s*(\d+)", to_western_digits(cells[col["no"]].text))
                no = int(m.group(1)) if m else last_no
                if no is None:
                    continue
                name = cells[col["name"]].text.strip() if "name" in col else ""
                text = cells[col["slo"]].text.strip()
                if not text or normalize(text).lower().startswith("slo"):
                    continue
                last_no, last_name = no, name or last_name
                lst = out.setdefault((chapter, no), [])
                lst.append(Slo(chapter, no, last_name, len(lst) + 1, text,
                               cells[col["bloom"]].text.strip() if "bloom" in col else "",
                               cells[col["cov"]].text.strip() if "cov" in col else ""))
    return out


def _slos_from_plan(doc, plan):
    out, n = {}, 0
    for child in doc.element.body.iterchildren():
        if child.tag.split("}")[1] != "tbl":
            continue
        cols = plan.get(n)
        n += 1
        if not cols:
            continue
        table = Table(child, doc)
        chapter, last_no, last_name = cols["chapter"], None, ""
        for row in table.rows:
            cells = row.cells
            def cell(k):
                i = cols.get(k)
                return cells[i].text.strip() if i is not None and i < len(cells) else ""
            m = re.match(r"\s*(\d+)", to_western_digits(cell("lesson_no")))
            if m is None and last_no is None:
                continue                  # the header row, or rows before the first lesson
            no = int(m.group(1)) if m else last_no
            text = cell("slo")
            if not text:
                continue
            last_no, last_name = no, cell("lesson_name") or last_name
            lst = out.setdefault((chapter, no), [])
            # an SLO cell may hold several SLOs, one per line
            for line in [l.strip(" •-–—*") for l in text.split("\n") if l.strip(" •-–—*")]:
                lst.append(Slo(chapter, no, last_name, len(lst) + 1, line, cell("bloom")))
    return out


def slos_for(slos, key):
    """SLOs of a package lesson. A split lesson (2A, 2B) shares its parent lesson's SLOs."""
    return slos.get((key.chapter, key.lesson), [])


def _norm_slo(text: str) -> str:
    return re.sub(r"[\s\.۔\-–—؛:,;]+$", "", normalize(text)).lower()


def tags_in(block: str, lesson_slos, chapter_slos=None):
    """Indexes (1-based, within the lesson) of the SLOs a question block is tied to.

    lesson_slos: the lesson's SLOs, used for bare 'SLO 2' markers.
    chapter_slos: when given (chapter-level documents), SLOs are matched across all of the
    chapter's lessons by quoted text or 'L3 SLO 2'; returns a set of (lesson, index)."""
    text = to_western_digits(normalize(block))
    found = set()
    if chapter_slos is None:
        for m in re.finditer(r"(?:\bslo\b|حاصلات?\s*تعلیم)[\s\-:#]*(\d{1,2})", text, re.I):
            n = int(m.group(1))
            if 1 <= n <= len(lesson_slos):
                found.add(n)
        for s in lesson_slos:
            q = _norm_slo(s.text)
            if len(q) >= 15 and q in text.lower():
                found.add(s.index)
        return found
    for m in re.finditer(r"\bl(?:esson)?\s*(\d{1,2})[\s,\-:]*slo[\s\-:#]*(\d{1,2})", text, re.I):
        found.add((int(m.group(1)), int(m.group(2))))
    for s in chapter_slos:
        q = _norm_slo(s.text)
        if len(q) >= 15 and q in text.lower():
            found.add((s.lesson, s.index))
    return found


def has_any_tag_marker(block: str) -> bool:
    text = to_western_digits(normalize(block))
    return bool(re.search(r"\bslo\b|حاصلات?\s*تعلیم", text, re.I) or SLO_CODE.search(text))
