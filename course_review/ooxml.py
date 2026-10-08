"""Word's file format requires the formatting inside a run (w:rPr) in a fixed order. Word may ignore formatting
that is out of order (for example a highlight after an out-of-place bold), so every run is put back in order
before a file is saved."""
from docx.oxml.ns import qn

RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
             "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position",
             "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
             "eastAsianLayout", "specVanish", "oMath"]
_RANK = {qn("w:" + t): i for i, t in enumerate(RPR_ORDER)}


def order_run_properties(root):
    """Sort the children of every w:rPr under `root` into the order the format requires. Returns how many moved."""
    moved = 0
    for rpr in root.iter(qn("w:rPr")):
        kids = list(rpr)
        ranked = sorted(kids, key=lambda k: _RANK.get(k.tag, len(RPR_ORDER)))   # stable: unknown ones keep their order, last
        if ranked != kids:
            for k in kids:
                rpr.remove(k)
            for k in ranked:
                rpr.append(k)
            moved += 1
    return moved
