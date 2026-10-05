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
from .common import canon_font, is_grey, pct, result, top

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
    allowed = {"poppins", "montserrat"}
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
    bad = {f: n for f, n in fonts.items() if canon_font(f) != "jamilnoorinastaliq"}
    if bad:
        total = sum(fonts.values())
        return result("WE7", FAIL, f"Urdu text is not Jamil Noori Nastaliq: " + top(fonts, 3, lambda k, v: f"{k} ({pct(v, total)})"),
                      ["Urdu fonts used: " + top(fonts, 5, lambda k, v: f"{k} ({v} chars)")])
    return result("WE7", PASS, "All Urdu text is Jamil Noori Nastaliq")


_CAPTION_STYLE = re.compile(r"caption", re.I)


def _caption_re(ctx, kind):
    prefixes = "|".join(re.escape(normalize(p)) for p in ctx.profile["vocab"]["caption_prefixes"][kind])
    return re.compile(rf"^\s*(?:{prefixes})\s*(\d+)\s*[:：.\-–]", re.I)


def we8(ctx, doc, info: DocxInfo):
    tol = ctx.th["text_size_tolerance"]
    tcap, fcap = _caption_re(ctx, "table"), _caption_re(ctx, "figure")
    cats = {"Title (20 bold)": (20, True, None), "Section heading, Heading 1 (16 bold)": (16, True, None),
            "Body (14)": (14, None, None), "Table text (14)": (14, None, None), "Table header row (bold)": (None, True, None),
            "Caption (11 italic)": (11, None, True), "TOC (Arial)": (None, None, None)}
    stats = {k: Counter() for k in cats}
    wrong = {k: 0 for k in cats}
    totals = {k: 0 for k in cats}
    for p in info.paras:
        if not p.text.strip():
            continue
        norm = normalize(p.text)
        is_caption = bool(_CAPTION_STYLE.search(p.style)) or bool(tcap.match(norm) or fcap.match(norm))
        if p.is_title:
            cat = "Title (20 bold)"
        elif p.heading_level == 1:
            cat = "Section heading, Heading 1 (16 bold)"
        elif p.heading_level:
            continue  # lower heading levels are not specified by the workbook
        elif p.in_toc:
            cat = "TOC (Arial)"
        elif is_caption:
            cat = "Caption (11 italic)"
        elif p.in_table:
            cat = "Table text (14)"
        else:
            cat = "Body (14)"
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
        if p.in_table and p.row_idx == 0 and not p.is_title:
            for r in p.runs:
                if r.script:
                    totals["Table header row (bold)"] += len(r.text)
                    if not r.bold:
                        wrong["Table header row (bold)"] += len(r.text)
                        stats["Table header row (bold)"]["not bold"] += len(r.text)
    ev, checked = [], 0
    for cat in cats:
        if not totals[cat]:
            continue
        checked += 1
        if wrong[cat] / totals[cat] > tol:
            ev.append(f"{cat}: {pct(wrong[cat], totals[cat])} of text is wrong ({top(stats[cat], 4)})")
    if not checked:
        return result("WE8", NA, "No text to check")
    if ev:
        return result("WE8", FAIL, "; ".join(ev[:3]), ev)
    return result("WE8", PASS, "Sizes and weights match the table", partial=not totals["Title (20 bold)"],
                  evidence=["No Title-style paragraph, so the title size was not checked."] if not totals["Title (20 bold)"] else [])


# ----------------------------------------------------------------- bilingual

def we10(ctx, doc, info):
    hits, verse = [], 0
    texts = [p.text for p in info.paras] if isinstance(info, DocxInfo) else [r.text for r in info.runs]
    marker = re.compile("﴿[^﴾]*﴾")  # Quranic verse-number ornament
    for t in texts:
        found = EASTERN_DIGITS.findall(t)
        if found:
            hits.append((t.strip()[:60], len(found)))
            verse += sum(len(EASTERN_DIGITS.findall(m)) for m in marker.findall(t))
    if hits:
        total = sum(n for _, n in hits)
        extra = f", {verse} of them inside Quranic verse-number markers" if verse else ""
        return result("WE10", FAIL, f"{total} non-Western digit(s) in {len(hits)} place(s){extra}",
                      [f"'{t}' ({n})" for t, n in hits[:6]])
    return result("WE10", PASS, "All numerals are Western (1, 2, 3)")


def we11(ctx, doc, info: DocxInfo):
    bad = []
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
    if bad:
        return result("WE11", FAIL, f"{len(bad)} paragraph(s) have the wrong text direction", bad[:6])
    return result("WE11", PASS, "Paragraph direction matches the script of its text", partial=True,
                  evidence=["Mid-sentence direction switching was not checked."])


# ------------------------------------------------------------------ colour

def _neutral(hex_val, theme, th):
    if theme and theme.lower() in ("text1", "background1", "dark1", "light1", "tx1", "bg1"):
        return True
    return is_grey(hex_val, th["colour_grey_tolerance"])


def we23(ctx, doc, info):
    th = ctx.th
    found = Counter()
    exempt_highlight = doc.doc_type in ("pop_quiz", "data_bank")
    if isinstance(info, DocxInfo):
        for p in info.paras:
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
        for t in info.tables:
            for f in t.cell_shadings:
                if f and f.lower() not in ("auto", "ffffff") and not is_grey(f, th["colour_grey_tolerance"]) and not \
                        (exempt_highlight and is_yellow(f)):
                    found[f"table cell shading #{f}"] += 1
        for name, (val, theme) in info.style_colors.items():
            if val and val.lower() != "auto" and not _neutral(val, theme, th):
                found[f"style '{name}' colour #{val}"] += 1
    else:
        tmap = getattr(info, "theme", {})
        for r in info.runs:
            if r.color and not is_grey(r.color, th["colour_grey_tolerance"]):
                found[f"text colour #{r.color}"] += len(r.text)
            elif r.scheme and r.scheme not in ("tx1", "bg1", "dk1", "lt1", "tx2", "bg2", "dk2", "lt2"):
                found[f"theme colour {r.scheme}"] += len(r.text)
        for slide, kind, val in info.fills:
            hexv = val if val and len(val) == 6 and re.match(r"^[0-9A-Fa-f]{6}$", val) else tmap.get(val or "", None)
            if hexv and not is_grey(hexv, th["colour_grey_tolerance"]):
                found[f"{kind} fill #{hexv}"] += 1
    if found:
        return result("WE23", FAIL, "Colour used: " + top(found, 3), [f"{k} ({v})" for k, v in found.most_common(8)])
    return result("WE23", PASS, "Text, shading and styles are black, white or grey",
                  partial=not isinstance(info, DocxInfo),
                  evidence=["Colours inherited from the slide master/theme were not resolved."] if not isinstance(info, DocxInfo) else [])


# ------------------------------------------------------------------ headings

_NUM = re.compile(r"^\s*(\d+(?:\.\d+){0,5})[.)]?\s+\S")


def leading_number(text):
    m = _NUM.match(to_western_digits(text))
    return tuple(int(x) for x in m.group(1).split(".")) if m else None


def heading_like(ctx, p):
    if p.in_table or not p.text.strip() or p.in_toc:
        return False
    if p.is_title:
        return False
    if p.heading_level:
        return True
    t = p.text.strip()
    if len(t) > 90 or p.list_kind:
        return False
    from ..textutil import starts_with_label
    if starts_with_label(t, ctx.section_labels):
        return True
    return p.all_bold and not re.search(r"[.!?؟۔]$", t)


def we12(ctx, doc, info: DocxInfo):
    heads = [p for p in info.paras if heading_like(ctx, p)]
    if not heads:
        return result("WE12", NA, "No headings found")
    seqs, unnumbered, ev = [], [], []
    for p in heads:
        if p.num_label:
            nums = tuple(int(x) for x in re.findall(r"\d+", p.num_label)) if p.list_kind == "decimal" else None
        else:
            nums = leading_number(p.text)
        if nums is None:
            unnumbered.append(p)
        else:
            seqs.append((p, nums))
    if unnumbered:
        ev.append(f"{len(unnumbered)} of {len(heads)} headings have no decimal number, e.g. " +
                  "; ".join(f"'{p.text.strip()[:30]}'" for p in unnumbered[:3]))
    prev = ()
    for p, nums in seqs:
        label = ".".join(map(str, nums))
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
        prev = nums
    if ev:
        return result("WE12", FAIL, ev[0], ev[:6])
    return result("WE12", PASS, f"{len(heads)} headings use correct decimal numbering")


def we13(ctx, doc, info: DocxInfo):
    heads = [p for p in info.paras if heading_like(ctx, p)]
    if not heads:
        return result("WE13", NA, "No headings found")
    fake = [p for p in heads if not p.heading_level]
    if fake:
        return result("WE13", FAIL, f"{len(fake)} of {len(heads)} headings use the '{fake[0].style}' style, not a built-in Heading style",
                      [f"'{p.text.strip()[:40]}' (style: {p.style})" for p in fake[:6]])
    return result("WE13", PASS, "All headings use built-in Heading styles")


def we15(ctx, doc, info: DocxInfo):
    if doc.doc_type != "lesson_plan":
        return result("WE15", NA, "Only applies to Lesson Plans")
    if info.has_toc:
        return result("WE15", PASS, "Table of Contents present")
    if info.pages is None:
        return result("WE15", REVIEW, "No Table of Contents, and the file does not record its page count, so 'longer than two pages' cannot be decided")
    if info.pages > 2:
        return result("WE15", FAIL, f"{info.pages} pages and no Table of Contents")
    return result("WE15", PASS, f"{info.pages} page(s), so no Table of Contents is needed")


def we16(ctx, doc, info: DocxInfo):
    if not info.tables:
        return result("WE16", NA, "No tables")
    cap = _caption_re(ctx, "table")
    bad = []
    expected = 1
    for t in info.tables:
        prev = info.paras[t.prev_para] if t.prev_para is not None else None
        m = cap.match(normalize(prev.text)) if prev else None
        if not m:
            bad.append(f"table {t.idx + 1} has no 'Table N:' caption directly above")
        else:
            if int(m.group(1)) != expected:
                bad.append(f"table {t.idx + 1} caption is numbered {m.group(1)}, expected {expected}")
            expected = int(m.group(1)) + 1
    if bad:
        return result("WE16", FAIL, f"{len(bad)} of {len(info.tables)} tables fail: {bad[0]}", bad[:6])
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
        return result("WE25", REVIEW, "Slide footers are not checked yet")
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


DOCX_CHECKS = {"WE1": we1, "WE2": we2, "WE3": we3, "WE4": we4, "WE6": we6, "WE7": we7, "WE8": we8, "WE10": we10,
               "WE11": we11, "WE12": we12, "WE13": we13, "WE15": we15, "WE16": we16, "WE17": we17, "WE18": we18,
               "WE19": we19, "WE20": we20, "WE23": we23, "WE25": we25}
PPTX_CHECKS = {"WE1": we1, "WE2": we2, "WE3": we3, "WE5": we5, "WE6": we6, "WE7": we7, "WE10": we10,
               "WE18": we18, "WE19": we19, "WE20": we20, "WE23": we23, "WE25": we25}
