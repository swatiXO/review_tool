"""Writing & Editing Guidelines (WE codes) decided from file contents.

Each check takes (ctx, doc, info) where info is a DocxInfo or PptxInfo and returns a
Finding for that one document. engine.py rolls the per-document findings up to the
Subject-Level sheet and into LP10 / FG8.
"""
import os
import re
from collections import Counter

from ..docx_model import DocxInfo
from ..models import FAIL, NA, PASS, REVIEW
from ..pptx_model import PptxInfo
from ..questions import is_yellow
from ..textutil import EASTERN_DIGITS, normalize, to_western_digits
from .common import canon_font, is_grey, mark, pct, result, top

A4 = (8.27, 11.69)


# ------------------------------------------------------------------ naming

def we1(ctx, doc, info):
    stem = os.path.splitext(os.path.basename(doc.rel))[0]
    if re.match(r"^[^-]+-[^-]+-.+-v\d+(?:\.\d+)*$", stem):
        return result("WE1", PASS, f"'{stem}' has the [Type]-[Identifier]-[Topic]-v[Version] shape", partial=True,
                      evidence=["Shape checked; whether each part names the right type/identifier/topic is for the reviewer."])
    return result("WE1", FAIL, f"'{stem}' does not follow [DocumentType]-[Identifier]-[Chapter/Topic]-v[Version]")


def we2(ctx, doc, info):
    name = os.path.basename(doc.rel)
    m = re.search(r"draft|final", name, re.I)
    if m:
        return result("WE2", FAIL, f"File name contains '{m.group(0)}': {name}")
    return result("WE2", PASS, "No 'Draft' or 'Final' in the file name")


def we3(ctx, doc, info):
    name = os.path.basename(doc.rel)
    m = re.search(r"-v(\d+(?:\.\d+)*)", name, re.I)
    if not m:
        return result("WE3", FAIL, f"No version number in the file name: {name}")
    return result("WE3", REVIEW, f"Version v{m.group(1)} is present; whether it matches the document's stage needs a reviewer")


# --------------------------------------------------------------- page setup

def we4(ctx, doc, info: DocxInfo):
    ev = []
    for i, (w, h, m, orient) in enumerate(info.sections, start=1):
        label = f"section {i}: " if len(info.sections) > 1 else ""
        if not (abs(w - A4[0]) <= 0.06 and abs(h - A4[1]) <= 0.06):
            ev.append(f"{label}page is {w:.2f} x {h:.2f} in, A4 is {A4[0]} x {A4[1]} in")
        if orient != "portrait":
            ev.append(f"{label}orientation is {orient}")
        bad = [f"{n} {x:.2f}" for n, x in zip(("top", "bottom", "left", "right"), m) if x is None or abs(x - 1.0) > 0.02]
        if bad:
            ev.append(f"{label}margins not 1 inch: " + ", ".join(bad))
    spacing = Counter()
    for p in info.paras:
        if p.text.strip():
            kind, val = p.line_spacing
            spacing[(kind, val)] += len(p.text)
    total = sum(spacing.values())
    off = sum(v for (k, val), v in spacing.items() if not (k != "points" and abs(val - 1.15) <= 0.01))
    if total and off / total > ctx.th["text_size_tolerance"]:
        shown = Counter()
        for (kind, val), n in spacing.items():
            shown[f"{val:g}{'pt' if kind == 'points' else 'x'}"] += n
        ev.append("line spacing is not 1.15 for " + pct(off, total) + " of text (" +
                  top(shown, 3, lambda k, v: f"{k}: {pct(v, total)}") + ")")
    if ev:
        return result("WE4", FAIL, "; ".join(ev[:3]), ev)
    return result("WE4", PASS, "A4, portrait, 1-inch margins, 1.15 line spacing")


def we5(ctx, doc, info: PptxInfo):
    ev = []
    ratio = info.width_in / info.height_in if info.height_in else 0
    if abs(ratio - 16 / 9) > 0.02:
        ev.append(f"slide size {info.width_in:.2f} x {info.height_in:.2f} in is not 16:9")
    spacing, inherited = Counter(), 0
    for r in info.runs:
        if r.line_spacing is None:
            inherited += len(r.text)
            spacing[("inherited", info.master_line_spacing or 1.0)] += len(r.text)
        else:
            spacing[("explicit", r.line_spacing)] += len(r.text)
    total = sum(spacing.values())
    off = sum(v for (k, val), v in spacing.items() if abs(val - 1.5) > 0.01)
    if total and off / total > ctx.th["text_size_tolerance"]:
        ev.append("line spacing is not 1.5 for " + pct(off, total) + " of slide text (" +
                  top(spacing, 3, lambda k, v: f"{k[1]}x {k[0]}: {pct(v, total)}") + ")")
    if ev:
        return result("WE5", FAIL, "; ".join(ev), ev)
    return result("WE5", PASS, "16:9 landscape, 1.5 line spacing", partial=inherited > 0,
                  evidence=["Some spacing is inherited from the layout and was read from the master."] if inherited else [])


# -------------------------------------------------------------- typography

def _runs(info):
    if isinstance(info, DocxInfo):
        for p in info.paras:
            for r in p.runs:
                yield r
    else:
        yield from info.runs


def we6(ctx, doc, info):
    fonts = Counter()
    for r in _runs(info):
        if r.script == "latin":
            fonts[r.font or "(unresolved)"] += len(r.text)
    if not fonts:
        return result("WE6", NA, "No English text in this document")
    allowed = {canon_font(f) for f in ctx.profile.get("house_rules", {}).get("english_fonts", ["Poppins", "Montserrat"])}
    families = {canon_font(f) for f in fonts}
    bad = {f: n for f, n in fonts.items() if canon_font(f) not in allowed}
    if bad or len(families) > 1:
        ev = ["English fonts used: " + top(fonts, 5, lambda k, v: f"{k} ({v} chars)")]
        msg = "English text is not set in one of Poppins / Montserrat" if bad else "Poppins and Montserrat are mixed in one document"
        return result("WE6", FAIL, f"{msg}: " + top(fonts, 3, lambda k, v: f"{k} ({v} chars)"), ev)
    return result("WE6", PASS, f"All English text is {next(iter(fonts))}")


def we7(ctx, doc, info):
    fonts = Counter()
    for r in _runs(info):
        if r.script == "arabic":
            fonts[r.font or "(unresolved)"] += len(r.text)
    if not fonts:
        return result("WE7", NA, "No Urdu text in this document")
    names = ctx.profile.get("house_rules", {}).get("urdu_fonts", ["Jamil Noori Nastaliq"])
    allowed = {canon_font(f) for f in names}
    bad = {f: n for f, n in fonts.items() if canon_font(f) not in allowed}
    if bad:
        total = sum(fonts.values())
        return result("WE7", FAIL, f"Urdu text is not {names[0]}: " + top(fonts, 3, lambda k, v: f"{k} ({pct(v, total)})"),
                      ["Urdu fonts used: " + top(fonts, 5, lambda k, v: f"{k} ({v} chars)")])
    return result("WE7", PASS, f"All Urdu text is {names[0]}")


_CAPTION_STYLE = re.compile(r"caption", re.I)


def _caption_re(ctx, kind):
    prefixes = "|".join(re.escape(normalize(p)) for p in ctx.profile["vocab"]["caption_prefixes"][kind])
    return re.compile(rf"^\s*(?:{prefixes})\s*(\d+)\s*[:：.\-–]", re.I)


def grade_sizes(ctx):
    """(grade, sizes) from the Writing & Editing Guidelines' table for this package's grade."""
    hr = ctx.profile.get("house_rules", {})
    m = re.search(r"grade[-_ ]*(\d+)", getattr(ctx.pkg, "subject", "") or "", re.I)
    grade = int(m.group(1)) if m else hr.get("default_grade", 6)
    for band in hr.get("sizes_by_grade", []):
        lo, hi = band["grades"]
        if lo <= grade <= hi:
            return grade, band, bool(m)
    return grade, {"title": 20, "section_heading": 16, "body": 14, "table_body": 14, "table_header": 14, "caption": 11}, bool(m)


def we8(ctx, doc, info: DocxInfo):
    tol = ctx.th["text_size_tolerance"]
    tcap, fcap = _caption_re(ctx, "table"), _caption_re(ctx, "figure")
    grade, sz, known = grade_sizes(ctx)
    T, H, B, TB, TH, C = (f"Title ({sz['title']} bold)", f"Section heading ({sz['section_heading']} bold)", f"Body ({sz['body']})",
                          f"Table text ({sz['table_body']})", f"Table header row ({sz['table_header']} bold)", f"Caption ({sz['caption']} italic)")
    cats = {T: (sz["title"], True, None), H: (sz["section_heading"], True, None), B: (sz["body"], None, None),
            TB: (sz["table_body"], None, None), TH: (sz["table_header"], True, None), C: (sz["caption"], None, True),
            "TOC (Arial)": (None, None, None)}
    heads = {id(p) for p in doc_headings(ctx, info)}
    first = next((p for p in info.paras if p.text.strip() and not p.in_table), None)
    stats = {k: Counter() for k in cats}
    wrong = {k: 0 for k in cats}
    totals = {k: 0 for k in cats}
    for p in info.paras:
        if not p.text.strip():
            continue
        norm = normalize(p.text)
        is_caption = bool(_CAPTION_STYLE.search(p.style)) or bool(tcap.match(norm) or fcap.match(norm))
        if p.is_title or (p is first and not p.heading_level):
            cat = T
        elif p.heading_level == 1 or (id(p) in heads and not p.heading_level):
            cat = H
        elif p.heading_level:
            continue  # lower heading levels are not specified in the size table
        elif p.in_toc:
            cat = "TOC (Arial)"
        elif is_caption:
            cat = C
        elif p.in_table and p.row_idx == 0:
            cat = TH
        elif p.in_table:
            cat = TB
        else:
            cat = B
        want_size, want_bold, want_italic = cats[cat]
        for r in p.runs:
            if not r.script:
                continue
            n = len(r.text)
            totals[cat] += n
            bad = False
            if want_size and (r.size is None or abs(r.size - want_size) > 0.01):
                bad = True
                stats[cat][f"{r.size:g}pt" if r.size else "unset"] += n
            if want_bold and not r.bold:
                bad = True
                stats[cat]["not bold"] += n
            if want_italic and not r.italic:
                bad = True
                stats[cat]["not italic"] += n
            if cat == "TOC (Arial)" and canon_font(r.font) != "arial":
                bad = True
                stats[cat][r.font or "unset"] += n
            wrong[cat] += n if bad else 0
    ev, checked = [], 0
    for cat in cats:
        if not totals[cat]:
            continue
        checked += 1
        if wrong[cat] / totals[cat] > tol:
            ev.append(f"{cat}: {pct(wrong[cat], totals[cat])} of text is wrong ({top(stats[cat], 4)})")
    if not checked:
        return result("WE8", NA, "No text to check")
    basis = f"Sizes for Grade {grade} from the Writing & Editing Guidelines" + ("" if known else " (grade not found in the package name; assumed)")
    if ev:
        return result("WE8", FAIL, "; ".join(ev[:3]), ev + [basis])
    return result("WE8", PASS, "Sizes and weights match the table", evidence=[basis])


# ----------------------------------------------------------------- bilingual

def we10(ctx, doc, info):
    hits, verse, marks = [], 0, []
    texts = [(p.text, None) for p in info.paras] if isinstance(info, DocxInfo) else [(r.text, r.slide) for r in info.runs]
    marker = re.compile("﴿[^﴾]*﴾")  # Quranic verse-number ornament
    for t, slide in texts:
        found = EASTERN_DIGITS.findall(t)
        if found:
            hits.append((t.strip()[:60], len(found)))
            marks.append(mark(t, "non-Western digits: " + " ".join(found[:5]), slide))
            verse += sum(len(EASTERN_DIGITS.findall(m)) for m in marker.findall(t))
    if hits:
        total = sum(n for _, n in hits)
        extra = f", {verse} of them inside Quranic verse-number markers" if verse else ""
        return result("WE10", FAIL, f"{total} non-Western digit(s) in {len(hits)} place(s){extra}",
                      [f"'{t}' ({n})" for t, n in hits[:6]], marks=marks)
    return result("WE10", PASS, "All numerals are Western (1, 2, 3)")


def we11(ctx, doc, info: DocxInfo):
    bad, marks = [], []
    for p in info.paras:
        if not p.runs:
            continue
        ar = sum(len(r.text) for r in p.runs if r.script == "arabic")
        la = sum(len(r.text) for r in p.runs if r.script == "latin")
        if ar == la == 0:
            continue
        want_rtl = ar >= la
        if p.bidi != want_rtl:
            bad.append(f"paragraph {p.idx + 1} is {'RTL' if want_rtl else 'LTR'} text but set {'LTR' if want_rtl else 'RTL'}: '{p.text.strip()[:40]}'")
            marks.append(mark(p.text, f"{'Urdu' if want_rtl else 'English'} text set {'left-to-right' if want_rtl else 'right-to-left'}"))
    if bad:
        return result("WE11", FAIL, f"{len(bad)} paragraph(s) have the wrong text direction", bad[:6], marks=marks)
    return result("WE11", PASS, "Paragraph direction matches the script of its text", partial=True,
                  evidence=["Mid-sentence direction switching was not checked."])


# ------------------------------------------------------------------ colour

def _neutral(hex_val, theme, th):
    if theme and theme.lower() in ("text1", "background1", "dark1", "light1", "tx1", "bg1"):
        return True
    return is_grey(hex_val, th["colour_grey_tolerance"])


def we23(ctx, doc, info):
    th = ctx.th
    found, marks = Counter(), []
    exempt_highlight = doc.doc_type in ("pop_quiz", "data_bank")
    if isinstance(info, DocxInfo):
        for p in info.paras:
            before = sum(found.values())
            for r in p.runs:
                if r.color and r.color.lower() != "auto" and not _neutral(r.color, r.theme_color, th):
                    found[f"text colour #{r.color}"] += len(r.text)
                elif not r.color and r.theme_color and not _neutral(None, r.theme_color, th) and r.theme_color != "":
                    found[f"text theme colour {r.theme_color}"] += len(r.text)
                if r.highlight and r.highlight != "none" and not (exempt_highlight and r.highlight.lower() == "yellow"):
                    found[f"highlight {r.highlight}"] += len(r.text)
                if r.shading and r.shading.lower() not in ("auto", "ffffff") and not is_grey(r.shading, th["colour_grey_tolerance"]) \
                        and not (exempt_highlight and is_yellow(r.shading)):
                    found[f"run shading #{r.shading}"] += len(r.text)
            if p.shading and p.shading.lower() not in ("auto", "ffffff") and not is_grey(p.shading, th["colour_grey_tolerance"]):
                found[f"paragraph shading #{p.shading}"] += 1
            if sum(found.values()) > before and p.text.strip():
                marks.append(mark(p.text, "colour used (only black, white and grey are allowed)"))
        for t in info.tables:
            for f in t.cell_shadings:
                if f and f.lower() not in ("auto", "ffffff") and not is_grey(f, th["colour_grey_tolerance"]) and not \
                        (exempt_highlight and is_yellow(f)):
                    found[f"table cell shading #{f}"] += 1
        # Style definitions are not checked on their own: a coloured style shows up through the
        # effective colour of the runs that use it, and an unused or overridden one is invisible.
    else:
        tmap = getattr(info, "theme", {})
        for r in info.runs:
            if r.color and not is_grey(r.color, th["colour_grey_tolerance"]):
                found[f"text colour #{r.color}"] += len(r.text)
                marks.append(mark(r.text, f"text colour #{r.color}", r.slide))
            elif r.scheme and r.scheme not in ("tx1", "bg1", "dk1", "lt1", "tx2", "bg2", "dk2", "lt2"):
                found[f"theme colour {r.scheme}"] += len(r.text)
                marks.append(mark(r.text, f"theme colour {r.scheme}", r.slide))
        for slide, kind, val in info.fills:
            hexv = val if val and len(val) == 6 and re.match(r"^[0-9A-Fa-f]{6}$", val) else tmap.get(val or "", None)
            if hexv and not is_grey(hexv, th["colour_grey_tolerance"]):
                found[f"{kind} fill #{hexv}"] += 1
    if found:
        return result("WE23", FAIL, "Colour used: " + top(found, 3), [f"{k} ({v})" for k, v in found.most_common(8)], marks=marks)
    return result("WE23", PASS, "Text, shading and styles are black, white or grey",
                  partial=not isinstance(info, DocxInfo),
                  evidence=["Colours inherited from the slide master/theme were not resolved."] if not isinstance(info, DocxInfo) else [])


# ------------------------------------------------------------------ headings

_NUM = re.compile(r"^\s*(\d+(?:\.\d+){0,5})[.)]?\s+\S")


def leading_number(text):
    m = _NUM.match(to_western_digits(text))
    return tuple(int(x) for x in m.group(1).split(".")) if m else None


def _is_question_label(ctx, t):
    from ..questions import question_regex
    rx = ctx.__dict__.setdefault("_qlabel_rx", question_regex(ctx.profile, labelled=True))
    return bool(rx.match(to_western_digits(normalize(t))))


def own_number(ctx, text):
    """'سبق 1: ...' / 'Lesson 3 - ...' / 'باب دوم: ...' / 'Chapter 2: ...' are lesson and chapter labels that carry
    their own number. They are not section headings to number 1, 1.1 (and numbering them would break how the
    Pop Quiz and Data Bank are split into lessons)."""
    t = to_western_digits(normalize(text)).strip().lower()
    labels = [normalize(x).lower() for x in ctx.profile["vocab"].get("lesson_heading_prefixes", []) + ["باب", "chapter"]]
    ordinals = [normalize(w).lower() for ws in ctx.profile["vocab"].get("chapter_ordinals", {}).values() for w in ws]
    alt = "|".join(map(re.escape, ordinals)) or "x^"
    return any(re.match(rf"^{re.escape(l)}\s*[:\-–]?\s*(?:\d+|{alt})(?![\w])", t) for l in labels)


def heading_like(ctx, p):
    if p.in_table or not p.text.strip() or p.in_toc:
        return False
    if p.is_title:
        return False
    t = p.text.strip()
    if _is_question_label(ctx, t):
        return False                  # 'سوال ١ (...)' / 'Question 3:' is a question, not a section heading
    if own_number(ctx, t):
        return False
    nt = normalize(t).lower()
    if any(nt.startswith(normalize(l).lower()) for l in ctx.profile["vocab"].get("answer_line_labels", [])):
        return False                  # 'Model answer: ...' inside a question
    if p.heading_level:
        return True
    if len(t) > 90 or p.list_kind:
        return False
    from ..textutil import starts_with_label
    if starts_with_label(t, ctx.section_labels):
        return True
    return p.all_bold and not re.search(r"[.!?؟۔:：]$", t)     # not a sentence, not a 'Knowledge:' sub-label


def doc_headings(ctx, info: DocxInfo):
    """Section headings. The document's opening line is its title, not a numbered heading,
    unless it is set in a Heading style."""
    first = next((p for p in info.paras if p.text.strip() and not p.in_table), None)
    return [p for p in info.paras if heading_like(ctx, p) and not (p is first and not p.heading_level)]


def we12(ctx, doc, info: DocxInfo):
    heads = doc_headings(ctx, info)
    if not heads:
        return result("WE12", NA, "No headings found")
    seqs, unnumbered, ev, marks = [], [], [], []
    for p in heads:
        if p.num_label:
            nums = tuple(int(x) for x in re.findall(r"\d+", p.num_label)) if p.list_kind == "decimal" else None
        else:
            nums = leading_number(p.text)
        if nums is None:
            unnumbered.append(p)
            marks.append(mark(p.text, "heading has no decimal number (1, 1.1, 1.1.1)"))
        else:
            seqs.append((p, nums))
    if unnumbered:
        ev.append(f"{len(unnumbered)} of {len(heads)} headings have no decimal number, e.g. " +
                  "; ".join(f"'{p.text.strip()[:30]}'" for p in unnumbered[:3]))
    prev = ()
    for p, nums in seqs:
        label = ".".join(map(str, nums))
        n_ev = len(ev)
        if len(nums) > 3:
            ev.append(f"'{label}' goes deeper than three levels")
        elif len(nums) > len(prev) + 1:
            ev.append(f"'{label}' skips a level (after {'.'.join(map(str, prev)) or 'start'})")
        elif len(nums) == len(prev) + 1 and nums[-1] != 1:
            ev.append(f"'{label}' should start at 1 under {'.'.join(map(str, prev))}")
        elif len(nums) <= len(prev) and nums[:-1] == prev[:len(nums) - 1] and nums[-1] != prev[len(nums) - 1] + 1:
            ev.append(f"'{label}' is out of order after {'.'.join(map(str, prev))}")
        elif len(nums) == 1 and not prev and nums[0] != 1:
            ev.append(f"numbering starts at {nums[0]}, not 1")
        if len(ev) > n_ev:
            marks.append(mark(p.text, ev[-1]))
        prev = nums
    if ev:
        return result("WE12", FAIL, ev[0], ev[:6], marks=marks)
    return result("WE12", PASS, f"{len(heads)} headings use correct decimal numbering")


def we13(ctx, doc, info: DocxInfo):
    heads = doc_headings(ctx, info)
    if not heads:
        return result("WE13", NA, "No headings found")
    fake = [p for p in heads if not p.heading_level]
    if fake:
        return result("WE13", FAIL, f"{len(fake)} of {len(heads)} headings use the '{fake[0].style}' style, not a built-in Heading style",
                      [f"'{p.text.strip()[:40]}' (style: {p.style})" for p in fake[:6]],
                      marks=[mark(p.text, f"looks like a heading but uses the '{p.style}' style, not a Heading style") for p in fake])
    return result("WE13", PASS, "All headings use built-in Heading styles")


def estimate_pages(info: DocxInfo):
    """A deliberately low page estimate from text length and size (on real Urdu lesson plans it gave
    55-65% of the true count), used only when the file does not record its page count."""
    import math
    if len(info.margins_in) == 4 and info.page_w_in and info.page_h_in:
        w = (info.page_w_in - info.margins_in[2] - info.margins_in[3]) * 72
        h = (info.page_h_in - info.margins_in[0] - info.margins_in[1]) * 72
    else:
        w, h = 6.5 * 72, 9 * 72
    total = 0.0
    for p in info.paras:
        if not p.text.strip():
            continue
        sz = max((r.size for r in p.runs if r.size), default=11)
        per_line = max(1.0, w / (sz * 0.45))
        nastaliq = any(r.script == "arabic" for r in p.runs)
        total += max(1, math.ceil(len(p.text) / per_line)) * sz * (1.8 if nastaliq else 1.2)
    return total / h if h > 0 else 0.0


def we15(ctx, doc, info: DocxInfo):
    if doc.doc_type != "lesson_plan":
        return result("WE15", NA, "Only applies to Lesson Plans")
    if info.has_toc:
        return result("WE15", PASS, "Table of Contents present")
    if info.pages is None:
        est = estimate_pages(info)
        if est >= 3:
            return result("WE15", FAIL, f"No Table of Contents, and the document runs to at least {int(est)} pages "
                                        "(estimated from its text; the file does not record a page count)")
        return result("WE15", REVIEW, "No Table of Contents, and the file does not record its page count, so 'longer than two pages' cannot be decided")
    if info.pages > 2:
        return result("WE15", FAIL, f"{info.pages} pages and no Table of Contents")
    return result("WE15", PASS, f"{info.pages} page(s), so no Table of Contents is needed")


def we16(ctx, doc, info: DocxInfo):
    if not info.tables:
        return result("WE16", NA, "No tables")
    cap = _caption_re(ctx, "table")
    bad, marks = [], []
    expected = 1
    for t in info.tables:
        prev = info.paras[t.prev_para] if t.prev_para is not None else None
        m = cap.match(normalize(prev.text)) if prev else None
        first = next((info.paras[i] for i in t.para_idx if info.paras[i].text.strip()), None)
        if not m:
            bad.append(f"table {t.idx + 1} has no 'Table N:' caption directly above")
            if first:
                marks.append(mark(first.text, bad[-1]))
        else:
            if int(m.group(1)) != expected:
                bad.append(f"table {t.idx + 1} caption is numbered {m.group(1)}, expected {expected}")
                marks.append(mark(prev.text, bad[-1]))
            expected = int(m.group(1)) + 1
    if bad:
        return result("WE16", FAIL, f"{len(bad)} of {len(info.tables)} tables fail: {bad[0]}", bad[:6], marks=marks)
    return result("WE16", PASS, f"All {len(info.tables)} tables have a numbered caption above")


def we17(ctx, doc, info: DocxInfo):
    figs = [p for p in info.paras if p.images and not p.in_table]
    if not info.images:
        return result("WE17", NA, "No figures")
    cap = _caption_re(ctx, "figure")
    bad, expected = [], 1
    for p in figs:
        nxt = next((q for q in info.paras if q.idx > p.idx and q.text.strip()), None)
        m = cap.match(normalize(nxt.text)) if nxt else None
        if not m:
            bad.append(f"figure in paragraph {p.idx + 1} has no 'Figure N:' caption below")
            continue
        if int(m.group(1)) != expected:
            bad.append(f"caption '{nxt.text.strip()[:30]}' is numbered {m.group(1)}, expected {expected}")
        expected = int(m.group(1)) + 1
        runs = [r for r in nxt.runs if r.script]
        if any(r.size is None or abs(r.size - 11) > 0.01 or not r.italic for r in runs):
            bad.append(f"caption {m.group(1)} is not 11pt italic")
    if bad:
        return result("WE17", FAIL, f"{len(bad)} figure caption problem(s): {bad[0]}", bad[:6])
    return result("WE17", PASS, f"All {len(figs)} figures have a numbered 11pt italic caption below")


def we18(ctx, doc, info):
    if not info.images:
        return result("WE18", NA, "No figures")
    unk = [i for i in info.images if i.colourful is None]
    colour = [i for i in info.images if i.colourful]
    if colour:
        return result("WE18", FAIL, f"{len(colour)} of {len(info.images)} images contain colour",
                      [f"{i.name or 'image'} ({i.fmt}) colour spread {i.max_chroma:.0f}" for i in colour[:6]])
    if unk:
        return result("WE18", REVIEW, f"{len(unk)} image(s) could not be analysed (vector or unreadable)")
    return result("WE18", PASS, f"All {len(info.images)} images are monochrome")


def we19(ctx, doc, info):
    if not info.images:
        return result("WE19", NA, "No figures")
    low, th = [], ctx.th["min_image_dpi"]
    for i in info.images:
        if i.fmt.lower() in ("emf", "wmf", "svg") or not i.px or not i.width_emu:
            continue
        dpi = i.px[0] / (i.width_emu / 914400)
        if dpi < th:
            low.append(f"{i.name or 'image'} is {dpi:.0f} dpi at its placed size")
    if low:
        return result("WE19", FAIL, f"{len(low)} image(s) below {th} dpi", low[:6])
    return result("WE19", PASS, "Resolution is acceptable", partial=True,
                  evidence=["Screenshots with browser or app chrome are not detectable by code."])


def we20(ctx, doc, info):
    if not info.images:
        return result("WE20", NA, "No figures")
    missing = [i for i in info.images if not i.descr]
    if missing:
        return result("WE20", FAIL, f"{len(missing)} of {len(info.images)} figures have no alt text",
                      [f"{i.name or 'image'}" for i in missing[:6]])
    return result("WE20", PASS, "Every figure has alt text")


def we25(ctx, doc, info):
    if not isinstance(info, DocxInfo):
        kinds = {k for _, k, _ in info.footers}
        if not info.footers:
            return result("WE25", FAIL, "No slide shows a footer (no footer, slide-number or date placeholder on any slide)")
        missing = [n for k, n in (("sldNum", "page (slide) number"), ("dt", "date"), ("ftr", "document name and version")) if k not in kinds]
        texts = " ".join(t for _, k, t in info.footers if k == "ftr")
        if "ftr" in kinds and not re.search(r"v\s?\d+(?:\.\d+)*", texts, re.I):
            missing.append("version in the footer text")
        if missing:
            return result("WE25", FAIL, "Slide footer is missing: " + ", ".join(missing))
        return result("WE25", PASS, "Slides carry a footer with number, date, name and version", partial=True)
    ev = []
    ft = info.footer_text.strip()
    if not ft and not info.footer_fields:
        return result("WE25", FAIL, "The document has no footer")
    if "PAGE" not in info.footer_fields and not re.search(r"\d", ft):
        ev.append("no page number")
    if not re.search(r"v\s?\d+(?:\.\d+)*", ft, re.I):
        ev.append("no version")
    if "DATE" not in info.footer_fields and "SAVEDATE" not in info.footer_fields and \
            not re.search(r"\d{1,4}[-/.]\d{1,2}[-/.]\d{2,4}|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}", to_western_digits(ft)):
        ev.append("no date")
    if len(re.sub(r"v\s?\d+(?:\.\d+)*|\d[\d\-/.]*", "", ft).strip()) < 3:
        ev.append("no document name")
    if ev:
        return result("WE25", FAIL, "Footer is missing: " + ", ".join(ev), [f"Footer text: '{ft[:80]}'"])
    return result("WE25", PASS, "Footer has page number, name, version and date", partial=True,
                  evidence=["The document name in the footer was not compared with the real name."])


_APA = re.compile(r"\((?:[A-Z][A-Za-z\-']+)(?:\s+(?:&|and)\s+[A-Z][A-Za-z\-']+|\s+et al\.)?,\s+(?:\d{4}[a-z]?|n\.d\.)\)")
_NOT_APA = [
    (re.compile(r"\[\d+(?:\s*[,–-]\s*\d+)*\]"), "numbered citation like [1]"),
    (re.compile(r"\((?:[A-Z][A-Za-z\-']+)(?:\s+(?:&|and)\s+[A-Z][A-Za-z\-']+|\s+et al\.)?\s+\d{4}[a-z]?\)"), "(Author Year) without the comma"),
]


def _citations(info, marks=None):
    good, bad = [], []
    for p in info.paras:
        t = p.text
        good += _APA.findall(t)
        for rx, why in _NOT_APA:
            for m in rx.findall(t):
                bad.append(f"{m} ({why})")
                if marks is not None:
                    marks.append(mark(t, f"citation {m} is not (Author, Year): {why}"))
    return good, bad


def we21(ctx, doc, info: DocxInfo):
    marks = []
    good, bad = _citations(info, marks)
    if not good and not bad:
        return result("WE21", NA, "No outside-research citations found (Quran and Hadith references are not APA citations)")
    if bad:
        return result("WE21", FAIL, f"{len(bad)} citation(s) are not (Author, Year)", bad[:6], marks=marks)
    return result("WE21", PASS, f"{len(good)} in-text citation(s), all (Author, Year)")


def we22(ctx, doc, info: DocxInfo):
    labels = [normalize(l).lower() for l in ctx.profile["vocab"]["reference_list_labels"]]
    head = next((p for p in info.paras if not p.in_table and len(p.text.strip()) <= 40
                 and normalize(p.text).lower().strip(" :") in labels), None)
    good, bad = _citations(info)
    if head is None:
        if good or bad:
            return result("WE22", FAIL, "The document cites outside research but has no reference list")
        return result("WE22", NA, "No outside-research citations, so no reference list is needed")
    entries = [p for p in info.paras if p.idx > head.idx and p.text.strip() and not p.in_table]
    if not entries:
        return result("WE22", FAIL, "The reference list heading has no entries")
    ev = []
    keys = [normalize(p.text).lower() for p in entries]
    if keys != sorted(keys):
        ev.append("entries are not in alphabetical order")
    no_hang = [p for p in entries if p.hanging <= 0]
    if no_hang:
        ev.append(f"{len(no_hang)} of {len(entries)} entries have no hanging indent")
    if ev:
        return result("WE22", FAIL, "; ".join(ev), ev)
    return result("WE22", PASS, f"{len(entries)} references, alphabetical, with a hanging indent")


DOCX_CHECKS = {"WE21": we21, "WE22": we22, "WE1": we1, "WE2": we2, "WE3": we3, "WE4": we4, "WE6": we6, "WE7": we7, "WE8": we8, "WE10": we10,
               "WE11": we11, "WE12": we12, "WE13": we13, "WE15": we15, "WE16": we16, "WE17": we17, "WE18": we18,
               "WE19": we19, "WE20": we20, "WE23": we23, "WE25": we25}
PPTX_CHECKS = {"WE1": we1, "WE2": we2, "WE3": we3, "WE5": we5, "WE6": we6, "WE7": we7, "WE10": we10,
               "WE18": we18, "WE19": we19, "WE20": we20, "WE23": we23, "WE25": we25}
