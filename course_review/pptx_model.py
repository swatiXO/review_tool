"""Parse a .pptx into the little the formatting checks need: slide size, fonts per
script, explicit colours, line spacing, pictures and their alt text."""
import io
import re
from dataclasses import dataclass, field
from typing import Optional

from pptx import Presentation
from pptx.util import Emu

from .docx_model import ImageInfo, _image_info
from .textutil import script_counts

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NEUTRAL_SCHEME = {"tx1", "bg1", "dk1", "lt1", "tx2", "bg2", "dk2", "lt2"}


@dataclass
class SlideRun:
    slide: int
    text: str
    script: Optional[str]
    font: Optional[str]
    size: Optional[float]
    color: Optional[str]           # explicit srgb hex
    scheme: Optional[str]          # explicit scheme colour name
    line_spacing: Optional[float]  # explicit multiple, None if inherited


@dataclass
class PptxInfo:
    path: str
    width_in: float = 0
    height_in: float = 0
    slides: int = 0
    runs: list = field(default_factory=list)
    fills: list = field(default_factory=list)       # (slide, kind, hex|scheme)
    images: list = field(default_factory=list)      # ImageInfo + slide
    master_line_spacing: Optional[float] = None
    language: Optional[str] = None
    char_count: int = 0
    notes_chars: int = 0


def _theme_color_map(prs):
    out = {}
    try:
        master = prs.slide_masters[0]
        for rel in master.part.rels.values():
            if rel.reltype.endswith("/theme"):
                from lxml import etree
                root = etree.fromstring(rel.target_part.blob)
                cs = root.find(f".//{{{A}}}clrScheme")
                for el in cs:
                    name = etree.QName(el).localname
                    child = el[0]
                    out[name] = child.get("val") or child.get("lastClr")
    except Exception:
        pass
    return out


def parse_pptx(path) -> PptxInfo:
    prs = Presentation(path)
    info = PptxInfo(path=str(path), width_in=Emu(prs.slide_width).inches, height_in=Emu(prs.slide_height).inches,
                    slides=len(prs.slides))
    # master body line spacing, if defined
    try:
        m = prs.slide_masters[0]._element
        sp = m.find(f".//{{{A}}}lvl1pPr/{{{A}}}lnSpc/{{{A}}}spcPct")
        if sp is not None:
            info.master_line_spacing = int(sp.get("val")) / 100000.0
    except Exception:
        pass

    def walk_shapes(shapes, idx):
        for sh in shapes:
            if sh.shape_type is not None and getattr(sh, "shapes", None) is not None and sh.shape_type == 6:  # group
                walk_shapes(sh.shapes, idx)
                continue
            el = sh._element
            # fills
            for sf in el.findall(f".//{{http://schemas.openxmlformats.org/presentationml/2006/main}}spPr/{{{A}}}solidFill"):
                c = sf[0]
                info.fills.append((idx, "shape", c.get("val")))
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    ppr = para._p.find(f"{{{A}}}pPr")
                    ls = None
                    if ppr is not None:
                        sp = ppr.find(f"{{{A}}}lnSpc/{{{A}}}spcPct")
                        if sp is not None:
                            ls = int(sp.get("val")) / 100000.0
                    for r in para.runs:
                        if not r.text:
                            continue
                        ar, la = script_counts(r.text)
                        script = "arabic" if ar > la else ("latin" if la > 0 else None)
                        rpr = r._r.find(f"{{{A}}}rPr")
                        latin = cs = color = scheme = None
                        size = None
                        if rpr is not None:
                            e = rpr.find(f"{{{A}}}latin")
                            latin = e.get("typeface") if e is not None else None
                            e = rpr.find(f"{{{A}}}cs")
                            cs = e.get("typeface") if e is not None else None
                            if rpr.get("sz"):
                                size = int(rpr.get("sz")) / 100
                            sf = rpr.find(f"{{{A}}}solidFill")
                            if sf is not None:
                                c = sf[0]
                                if c.tag.endswith("srgbClr"):
                                    color = c.get("val")
                                elif c.tag.endswith("schemeClr"):
                                    scheme = c.get("val")
                        info.runs.append(SlideRun(idx, r.text, script, cs if script == "arabic" else latin,
                                                  size, color, scheme, ls))
            if sh.has_table if hasattr(sh, "has_table") else False:
                for row in sh.table.rows:
                    for cell in row.cells:
                        tcpr = cell._tc.find(f"{{{A}}}tcPr")
                        if tcpr is not None:
                            sf = tcpr.find(f"{{{A}}}solidFill")
                            if sf is not None:
                                info.fills.append((idx, "table cell", sf[0].get("val")))
            if sh.shape_type == 13:  # picture
                try:
                    blob = sh.image.blob
                    fmt = sh.image.ext
                except Exception:
                    blob, fmt = b"", ""
                px, colourful, chroma = _image_info(blob, fmt) if blob else (None, None, 0.0)
                cnv = el.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")
                descr = (cnv.get("descr") or "").strip() if cnv is not None else ""
                ii = ImageInfo(descr=descr, name=cnv.get("name") if cnv is not None else "",
                               width_emu=int(sh.width), height_emu=int(sh.height), px=px, fmt=fmt,
                               colourful=colourful, max_chroma=chroma, blob_len=len(blob))
                ii.slide = idx
                info.images.append(ii)

    for i, slide in enumerate(prs.slides, start=1):
        walk_shapes(slide.shapes, i)
        if slide.has_notes_slide:
            info.notes_chars += len(slide.notes_slide.notes_text_frame.text or "")

    info.theme = _theme_color_map(prs)
    ar = sum(script_counts(r.text)[0] for r in info.runs)
    la = sum(script_counts(r.text)[1] for r in info.runs)
    info.language = None if ar + la == 0 else ("ur" if ar >= la else "en")
    info.char_count = sum(len(r.text) for r in info.runs)
    return info
