"""Builders for tiny .docx / .pptx / .xlsx fixtures so each rule can be exercised alone."""
import io
import os
import zipfile

import docx
import openpyxl
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from course_review import ingest
from course_review.engine import Context
from course_review.models import DocRef, Package, Rule

PROFILE = ingest.load_profile()
A4 = (8.27, 11.69)


def new_doc(page=A4, margin=1.0, spacing=1.15, font="Poppins", size=14):
    d = docx.Document()
    s = d.sections[0]
    s.page_width, s.page_height = Inches(page[0]), Inches(page[1])
    for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(s, side, Inches(margin))
    normal = d.styles["Normal"]
    normal.font.name = font
    normal.font.size = Pt(size)
    normal.paragraph_format.line_spacing = spacing
    return d


def add(d, text, style=None, size=None, bold=None, italic=None, font=None, urdu_font=None, rtl=False, color=None, highlight=None):
    p = d.add_paragraph(style=style)
    r = p.add_run(text)
    if size:
        r.font.size = Pt(size)
        if rtl:
            szcs = OxmlElement("w:szCs")
            szcs.set(qn("w:val"), str(int(size * 2)))
            r._r.get_or_add_rPr().append(szcs)
    if bold is not None:
        r.font.bold = bold
    if italic is not None:
        r.font.italic = italic
    if font:
        r.font.name = font
    if urdu_font:
        rpr = r._r.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        rf.set(qn("w:cs"), urdu_font)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    if highlight:
        r.font.highlight_color = highlight
    if rtl:
        p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    return p


def save(d, tmp_path, name="x.docx"):
    path = os.path.join(str(tmp_path), name)
    d.save(path)
    return path


def make_ctx(rules=None, docs=None):
    pkg = Package(root="", subject="Test", docs=docs or [], unclassified=[], chapter_titles={}, lesson_labels={})
    return Context(pkg, PROFILE, rules or {})


def ref(rel="Pkg/Chapter-1/Lesson-1/Lesson-1-Chapter-1-Lesson-Plan.docx", doc_type="lesson_plan", **kw):
    return DocRef(rel=rel, abs=rel, ext=rel.rsplit(".", 1)[-1], doc_type=doc_type, chapter=kw.get("chapter", 1),
                  lesson=kw.get("lesson", 1), scope=kw.get("scope", "lesson"))


def checklist(tmp_path, codes_by_sheet=None):
    """A miniature checklist workbook with the same layout as the real one."""
    wb = openpyxl.Workbook()
    wb.active.title = "Index"
    key = wb.create_sheet("Key")
    key.append(["Code Key"])
    key.append([])
    key.append([])
    key.append(["Code", "Output Type", "Scope", "Checklist Item"])
    rules = [("LP3", "Lesson Plan", "Per lesson", "five sections"), ("LP4", "Lesson Plan", "Per lesson", "numbering"),
             ("LP5", "Lesson Plan", "Per lesson", "bullets"), ("LP10", "Lesson Plan", "Per lesson", "formatting"),
             ("PQ1", "Pop Quiz", "Per lesson", "count"), ("PQ4", "Pop Quiz", "Per lesson", "highlight"),
             ("WS1", "Worksheet", "Per chapter", "count"),
             ("WE1", "Writing & Editing Guidelines", "Per subject", "file name"),
             ("WE4", "Writing & Editing Guidelines", "Per subject", "page setup")]
    for r in rules:
        key.append(list(r))
    for title, codes, scope in (("Lesson Plan", ["LP3", "LP4", "LP5", "LP10"], "lesson"), ("Pop Quiz", ["PQ1", "PQ4"], "lesson"),
                                ("Worksheet", ["WS1"], "chapter")):
        ws = wb.create_sheet(title)
        ws.append([title])
        ws.append(["Per " + scope])
        ws.append([])
        ws.append(["#", "Lesson"] + codes + ["Fails", "Notes"])
        ws.append(["Ex.", "Example"] + ["Pass"] * len(codes) + ["x", "EXAMPLE ROW"])
        for i in range(1, 4):
            ws.append([float(i), None] + [None] * len(codes) + [f'=COUNTIF(C{5 + i}:{chr(66 + len(codes))}{5 + i},"Fail")', None])
        from openpyxl.worksheet.datavalidation import DataValidation
        dv = DataValidation(type="list", formula1='"Pass,Fail,N/A"')
        dv.add(f"C5:{chr(66 + len(codes))}8")
        ws.add_data_validation(dv)
    sj = wb.create_sheet("Subject-Level")
    sj.append(["Subject-Level Checklist"])
    sj.append([])
    sj.append([])
    sj.append(["Code", "Section", "Checklist Item", "Status", "Notes"])
    sj.append(["WE1", "Naming", "file name", None, None])
    sj.append(["WE4", "Page", "page setup", None, None])
    path = os.path.join(str(tmp_path), "checklist.xlsx")
    wb.save(path)
    return path


def zip_dir(root, zip_path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for dp, _, files in os.walk(root):
            for f in files:
                full = os.path.join(dp, f)
                z.write(full, os.path.relpath(full, os.path.dirname(root)))
    return zip_path
