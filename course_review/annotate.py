"""Marked-up copy of the package: every checklist result written into the documents themselves.

Red highlight = fails a checklist rule, turquoise = the model's suggestion for a reviewer to
check, bright green = passes. Each highlighted paragraph carries a Word comment naming the
checklist code and the reason. Problems that belong to the whole document (fonts, page size,
file name) go into one comment on the first paragraph. Existing highlights are never replaced,
because yellow on a correct answer is itself something the checklist requires.

PowerPoint has no comments the libraries can write, so slides get the highlight plus a small
review box listing the notes for that slide. Results that belong to no file (a missing
document, subject-wide coverage) go into REVIEW-NOTES.txt at the top of the zip.
"""
import re
import shutil
import tempfile
import zipfile
from collections import defaultdict
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .models import FAIL, PASS, REVIEW
from .textutil import normalize, to_western_digits

AUTHOR, INITIALS = "Course Review", "CR"
RANK = {FAIL: 3, REVIEW: 2, PASS: 1}
WORD_COLOUR = {FAIL: WD_COLOR_INDEX.RED, REVIEW: WD_COLOR_INDEX.TURQUOISE, PASS: WD_COLOR_INDEX.BRIGHT_GREEN}
SLIDE_COLOUR = {FAIL: "FF0000", REVIEW: "00FFFF", PASS: "00FF00"}
LABEL = {FAIL: "FAIL", REVIEW: "CHECK", PASS: "PASS"}
LEGEND = "Red = fails the checklist. Turquoise = model suggestion, a reviewer decides. Green = passes."
MIN_MATCH = 6


def _loose(text):
    t = re.sub(r"[^\w\s]", " ", to_western_digits(normalize(text or "")))
    return re.sub(r"\s+", " ", t).strip().lower()


def _pieces(text):
    """A mark may be a quote that skips words with '...' or spans several paragraphs."""
    parts = re.split(r"\.\.\.|…|\[\.\.\.[^\]]*\]|\n", text or "")
    return [p for p in (_loose(x) for x in parts) if len(p) >= MIN_MATCH]


def _matches(pieces, para_loose):
    """Indexes of paragraphs a mark points at: the paragraph contains a piece, or (for a quote
    spanning paragraphs) a non-trivial paragraph is wholly inside a piece."""
    hits = []
    for i, t in enumerate(para_loose):
        if not t:
            continue
        if any(p in t for p in pieces) or (len(t) >= MIN_MATCH * 2 and any(t in p for p in pieces)):
            hits.append(i)
    return hits


def _locate(m, para_loose, used, allowed=None):
    """Paragraph indexes a mark points at. A whole-paragraph mark takes one paragraph: the first
    equal one this finding has not used yet, so repeated headings or table cells are marked once
    each, in order. A quote marks every paragraph it touches."""
    ok = (lambda i: True) if allowed is None else allowed
    if not m.get("exact", True):
        return [i for i in _matches(_pieces(m["text"]), para_loose) if ok(i)]
    key = _loose(m["text"])
    if not key:
        return []
    cands = [i for i, t in enumerate(para_loose) if t == key and ok(i)]
    if not cands and len(key) >= MIN_MATCH:
        cands = [i for i, t in enumerate(para_loose) if key in t and ok(i)]
    if not cands:
        return []
    pick = next((i for i in cands if i not in used), cands[0])
    used.add(pick)
    return [pick]


# The workbook's sheet names are swapped against the team's names for these outputs: its "Chapter
# Exam" sheet (CE) is the per-lesson Assessment, its "Worksheet" sheet (WS) is the per-chapter
# Chapter Exam. Comments use the team's names so a reviewer is not confused.
OUTPUT_NAME = {"CE": "Assessment", "WS": "Chapter Exam"}


def _code(code):
    name = OUTPUT_NAME.get(re.sub(r"\d+$", "", code))
    return f"{code} ({name})" if name else code


def _line(f, note=""):
    if not f.code:                        # a _Note already holds a finished line
        return f.message
    return f"{_code(f.code)} {LABEL.get(f.status, f.status.upper())}: {note or f.message}"


def _worst(statuses):
    return max(statuses, key=lambda s: RANK.get(s, 0))


def _relevant(findings):
    return [f for f in findings if f.status in RANK]


# ---------------------------------------------------------------------------- Word

def _all_paragraphs(doc):
    body = doc.element.body
    return [Paragraph(p, doc._body) for p in body.iter(qn("w:p"))]


def _text_runs(par):
    return [r for r in par.runs if r.text.strip()]


def annotate_docx(src, dst, findings, fixed=None):
    """Write a marked-up copy of one .docx; returns how many paragraphs were marked. `fixed` lists what
    the tool already corrected in this copy, for the summary comment."""
    doc = Document(src)
    paras = _all_paragraphs(doc)
    loose = [_loose(p.text) for p in paras]
    per_para = defaultdict(list)          # para index -> [(status, line)]
    unplaced = []
    for f in _relevant(findings):
        placed, used = False, set()
        for m in f.marks:
            for i in _locate(m, loose, used):
                # a model suggestion quotes several places: comment once, highlight the rest
                once = f.method == "model" and placed
                per_para[i].append((f.status, None if once else _line(f, m.get("note"))))
                placed = True
        if not placed and f.status != PASS:
            unplaced.append(f)

    for i, items in sorted(per_para.items()):
        runs = _text_runs(paras[i])
        if not runs:
            unplaced_lines = [l for _, l in items if l]
            per_para[i] = []
            unplaced.extend(_Note(l) for l in unplaced_lines)
            continue
        colour = WORD_COLOUR[_worst(s for s, _ in items)]
        for r in runs:
            if r.font.highlight_color is None and r._r.find(qn("w:rPr") + "/" + qn("w:shd")) is None:
                r.font.highlight_color = colour
        lines = list(dict.fromkeys(l for _, l in items if l))
        if lines:
            _comment(doc, runs, lines)

    first = next((p for p in paras if _text_runs(p)), None)
    if first is not None:
        problems = [f for f in unplaced if f.status in (FAIL, REVIEW)]
        lines = ["Course Review of this document. " + LEGEND]
        if fixed:
            lines.append("Fixed by the tool in this copy:")
            lines += [f"- {l}" for l in fixed]
        if problems:
            lines.append("Whole-document results:")
            lines += [f"- {_line(f)}" for f in problems]
        marked = sum(1 for v in per_para.values() if v)
        lines.append(f"{marked} paragraph(s) are highlighted in the text below for a person to check."
                     if marked else "Nothing in the text itself needs a person's attention.")
        _comment(doc, _text_runs(first)[:1], lines)
    doc.save(dst)
    return sum(1 for v in per_para.values() if v)


class _Note:
    """A message that could not be anchored to its paragraph, shown in the top comment."""

    def __init__(self, line):
        self.status = FAIL if " FAIL:" in line else REVIEW if " CHECK:" in line else PASS
        self.code, self.message, self.marks = "", line, []


def _comment(doc, runs, lines):
    c = doc.add_comment(runs, text=lines[0], author=AUTHOR, initials=INITIALS)
    for l in lines[1:]:
        c.add_paragraph(l)


# ---------------------------------------------------------------------- PowerPoint

def annotate_pptx(src, dst, findings, fixed=None):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Emu, Pt
    from lxml import etree

    prs = Presentation(src)
    slides = list(prs.slides)
    slide_paras = []                      # (slide number, paragraph, loose text)
    for n, s in enumerate(slides, start=1):
        for shape in _shapes(s.shapes):
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    slide_paras.append((n, p, _loose("".join(r.text for r in p.runs))))
    loose = [t for _, _, t in slide_paras]
    notes = defaultdict(list)             # slide number -> [(status, line)]
    marked_paras = defaultdict(list)
    unplaced = []
    for f in _relevant(findings):
        placed, used = False, set()
        for m in f.marks:
            on_slide = None if m.get("slide") is None else (lambda i, s=m["slide"]: slide_paras[i][0] == s)
            for i in _locate(m, loose, used, on_slide):
                n = slide_paras[i][0]
                marked_paras[i].append(f.status)
                notes[n].append((f.status, _line(f, m.get("note"))))
                placed = True
        if not placed and f.status != PASS:
            unplaced.append(f)

    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for i, statuses in marked_paras.items():
        colour = SLIDE_COLOUR[_worst(statuses)]
        for r in slide_paras[i][1].runs:
            rpr = r._r.get_or_add_rPr()
            if rpr.find(f"{{{a}}}highlight") is not None:
                continue
            hl = etree.SubElement(rpr, f"{{{a}}}highlight")
            etree.SubElement(hl, f"{{{a}}}srgbClr", val=colour)
            # schema order: highlight must come after fill/effect elements and before fonts
            for tag in ("latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"):
                for el in rpr.findall(f"{{{a}}}{tag}"):
                    rpr.remove(el)
                    rpr.append(el)

    if slides:
        problems = [f for f in unplaced if f.status in (FAIL, REVIEW)]
        top = [(FAIL, "Course Review. " + LEGEND)] + [(PASS, "Fixed by the tool: " + l) for l in (fixed or [])] + \
              [(f.status, _line(f)) for f in problems]
        notes[1] = top + notes.get(1, [])
    for n, items in notes.items():
        lines = list(dict.fromkeys(l for _, l in items))
        box = slides[n - 1].shapes.add_textbox(Emu(int(prs.slide_width * 0.02)), Emu(int(prs.slide_height * 0.02)),
                                              Emu(int(prs.slide_width * 0.96)), Emu(int(prs.slide_height * 0.12)))
        box.name = "Course Review notes"
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(0xFF, 0xF4, 0xF4)
        box.line.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        tf = box.text_frame
        tf.word_wrap = True
        for k, l in enumerate(lines):
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            p.text = l
            for r in p.runs:
                r.font.size = Pt(10)
                r.font.color.rgb = RGBColor(0x60, 0x00, 0x00)
    prs.save(dst)
    return len(marked_paras)


def _shapes(shapes):
    for sh in shapes:
        if sh.shape_type == 6:  # group
            yield from _shapes(sh.shapes)
        else:
            yield sh


# ------------------------------------------------------------------------ package

def annotate_package(pkg, findings, out_zip, rules=None, profile=None, fix=True):
    """Zip of the whole package with each .docx/.pptx fixed where the rules are mechanical and
    commented where a person must decide, plus REVIEW-NOTES.txt."""
    from . import autofix
    root = Path(pkg.root)
    by_doc = defaultdict(list)
    loose_ends = []
    for f in findings:
        (by_doc[f.doc] if f.doc else loose_ends).append(f)
    ctx = None
    if fix and profile is not None:
        from .engine import Context
        ctx = Context(pkg, profile, rules or {})
    grade_subject = re.sub(r"[^A-Za-z0-9]+", "-", pkg.subject or "").strip("-")
    work = Path(tempfile.mkdtemp(prefix="course-markup-"))
    stats = {"documents": 0, "paragraphs": 0, "errors": {}, "renamed": {}}
    try:
        tree = work / root.name
        shutil.copytree(root, tree)
        for d in pkg.docs:
            fs = by_doc.get(d.rel, [])
            if d.ext not in ("docx", "pptx") or not fs or d.superseded:
                continue
            dst = tree / d.rel
            try:
                fixed_lines, extra, src = [], {}, d.abs
                if ctx is not None:
                    info = ctx.docx(d) if d.ext == "docx" else ctx.pptx(d)
                    version, found_in = autofix.find_version(d, info)
                    name = autofix.new_name(d, grade_subject, version)
                    base = name.rsplit(".", 1)[0] if name else Path(d.rel).stem
                    if version:
                        base = base[: -len(f"-v{version}")] if base.endswith(f"-v{version}") else base
                    staged = work / ("fixed." + d.ext)
                    if d.ext == "docx":
                        if info is not None:
                            fixed_lines, extra = autofix.fix_docx(d.abs, staged, ctx, d, info)
                    else:
                        footer = f"{base}" + (f" | v{version}" if version else "") + f" | {date.today().isoformat()}"
                        fixed_lines = autofix.fix_pptx(d.abs, staged, ctx, footer if name else None)
                    if fixed_lines:
                        src = str(staged)
                    if name and name != dst.name:
                        dst.unlink(missing_ok=True)
                        dst = dst.with_name(name)
                        stats["renamed"][d.rel] = name
                        fixed_lines.append(f"File renamed to {name} (pattern [Type]-[Identifier]-[Chapter]-v[Version])")
                    if d.ext == "docx" and src == str(staged):
                        autofix.add_footer(src, base, version)
                        fixed_lines.append("Footer added: document name, " + ("version, " if version else "") + "date, page number")
                    if name and version and found_in != "the file name":
                        fixed_lines.append(f"Version v{version} taken from {found_in}")
                    rel = (d.rel.rsplit("/", 1)[0] + "/" if "/" in d.rel else "") + dst.name
                    fs = autofix.remaining(fs, d.ext, autofix.recheck(ctx, d, src, rel), extra)
                n = (annotate_docx if d.ext == "docx" else annotate_pptx)(src, dst, fs, fixed_lines)
                stats["documents"] += 1
                stats["paragraphs"] += n
            except Exception as e:   # a file the libraries cannot rewrite stays as it was, and says so
                shutil.copy2(d.abs, tree / d.rel)
                stats["errors"][d.rel] = f"{type(e).__name__}: {e}"
                loose_ends.append(_Note(f"{d.rel}: could not be fixed or marked up ({type(e).__name__}); its results are below"))
                loose_ends.extend(f for f in fs if f.status in (FAIL, REVIEW))
        (tree / "REVIEW-NOTES.txt").write_text(_notes_text(loose_ends, pkg, rules, stats["renamed"]), encoding="utf-8")
        out_zip = Path(out_zip)
        out_zip.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as z:
            for p in sorted(tree.rglob("*")):
                if p.is_file() and not p.name.startswith("fixed."):
                    z.write(p, p.relative_to(work).as_posix())
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return stats


def _notes_text(findings, pkg, rules, renamed=None):
    lines = ["Course Review notes", "",
             "Formatting rules with one right answer (page setup, fonts, sizes, digits, heading numbers and styles,",
             "table captions, bullets, colour, footers, file names) have been fixed in these copies; each file's first",
             "comment lists what was changed. What needs a person is highlighted with a comment. " + LEGEND, ""]
    if renamed:
        lines += ["Files renamed to the guideline pattern:"] + [f"- {old}  ->  {new}" for old, new in sorted(renamed.items())] + [""]
    lines += ["Results that do not belong to one file:", ""]
    groups = defaultdict(lambda: ([], []))   # (status, message) -> (codes, places)
    for f in findings:
        if f.status not in (FAIL, REVIEW):
            continue
        codes, places = groups[(f.status, f.message)]
        if f.code and _code(f.code) not in codes:
            codes.append(_code(f.code))
        where = f.lesson.short() if getattr(f, "lesson", None) else (f"Ch{f.chapter}" if getattr(f, "chapter", None) else "")
        if where and where not in places:
            places.append(where)
    if not groups:
        lines.append("None.")
    for (status, msg), (codes, places) in groups.items():
        head = f"{', '.join(codes)} {LABEL[status]}" if codes else ""
        lines.append(f"- {head + ' ' if head else ''}{'(' + ', '.join(places) + ') ' if places else ''}{': ' if head else ''}{msg}".replace(" : ", ": "))
    if pkg.unclassified:
        lines += ["", "Files the tool could not place in a lesson or chapter (not reviewed):"]
        lines += [f"- {u}" for u in pkg.unclassified]
    return "\n".join(lines) + "\n"
