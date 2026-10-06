"""Textbook index: OCR the book once, keep the text per page, search it by keywords.

The textbook is usually a scan, so the text has to be read from page images. Engines:
  tesseract  Tesseract with the Urdu model (fast, free, fully local)
  vlm        a local vision model served by Ollama (slower, may read Nastaliq better)
  auto       use the PDF's own text layer when a page has one, otherwise tesseract
Pages are cached as text files, so an interrupted run resumes and a finished index is reused.
Search is plain BM25 over words; no embedding model is needed.
"""
import base64
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
from collections import Counter
from pathlib import Path

from .textutil import normalize

WORD = re.compile(r"[^\W\d_]+", re.UNICODE)
TESSERACT_DEFAULT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
PDFTOPPM_DEFAULT = r"C:\Program Files (x86)\poppler-25.07.0\Library\bin\pdftoppm.exe"


def tokens(text: str):
    return [w for w in WORD.findall(normalize(text).lower()) if len(w) > 1]


class BookIndex:
    def __init__(self, pages):
        self.pages = dict(pages)                      # page number -> text
        self.tf = {p: Counter(tokens(t)) for p, t in self.pages.items()}
        self.len = {p: sum(c.values()) for p, c in self.tf.items()}
        self.avg = (sum(self.len.values()) / len(self.len)) if self.len else 0
        self.df = Counter()
        for c in self.tf.values():
            self.df.update(c.keys())

    @classmethod
    def load(cls, directory):
        pages = {}
        folder = Path(directory) / "pages"
        for f in sorted(folder.glob("*.txt")):
            try:
                pages[int(f.stem)] = f.read_text(encoding="utf8")
            except (ValueError, OSError):
                continue
        if not pages:
            raise ValueError(f"no page text found in {folder}")
        return cls(pages)

    def search(self, query, k=3):
        """[(page, score)] best matches first (BM25, k1=1.5, b=0.75)."""
        q = tokens(query)
        n = len(self.pages)
        scores = {}
        for p, tf in self.tf.items():
            s = 0.0
            for w in q:
                f = tf.get(w, 0)
                if not f:
                    continue
                idf = math.log(1 + (n - self.df[w] + 0.5) / (self.df[w] + 0.5))
                s += idf * f * 2.5 / (f + 1.5 * (0.25 + 0.75 * self.len[p] / (self.avg or 1)))
            if s:
                scores[p] = s
        return sorted(scores.items(), key=lambda x: -x[1])[:k]

    def passages(self, query, k=2, max_chars=1400):
        """Text of the best-matching pages, labelled with page numbers, for use as material."""
        out = []
        for p, _ in self.search(query, k):
            out.append(f"[book page {p}] " + " ".join(self.pages[p].split())[:max_chars])
        return "\n".join(out)


# ------------------------------------------------------------------ building
def _exe(name, default):
    return shutil.which(name) or (default if os.path.exists(default) else None)


def render_page(pdf, page, dpi, workdir):
    exe = _exe("pdftoppm", PDFTOPPM_DEFAULT)
    if not exe:
        raise RuntimeError("pdftoppm (poppler) is not installed")
    prefix = os.path.join(workdir, f"p{page}")
    subprocess.run([exe, "-r", str(dpi), "-f", str(page), "-l", str(page), "-png", "-singlefile", str(pdf), prefix],
                   check=True, capture_output=True)
    return prefix + ".png"


def ocr_tesseract(png, lang="urd+eng"):
    exe = _exe("tesseract", TESSERACT_DEFAULT)
    if not exe:
        raise RuntimeError("tesseract is not installed")
    r = subprocess.run([exe, png, "stdout", "-l", lang, "--psm", "6"], capture_output=True, check=True)
    return r.stdout.decode("utf8", errors="replace")


VLM_PROMPT = ("Transcribe all the text on this book page exactly as written, in reading order. "
              "The page is in Urdu. Do not translate, summarise or add anything. Output only the text.")


def _downscaled_png(png, max_side):
    """PNG bytes with the longest side at most max_side pixels. Vision models cost roughly in
    proportion to the pixels they see, so a full 200 dpi page is far slower than it needs to be."""
    from io import BytesIO
    from PIL import Image
    with Image.open(png) as im:
        im = im.convert("RGB")
        if max(im.size) > max_side:
            im.thumbnail((max_side, max_side))
        buf = BytesIO()
        im.save(buf, "PNG")
        return buf.getvalue()


def ocr_vlm(png, client, max_side=1280):
    image = base64.b64encode(_downscaled_png(png, max_side)).decode("ascii")
    return client.chat_vision(VLM_PROMPT, image)


def text_layer(pdf, page):
    try:
        from pypdf import PdfReader
        t = PdfReader(str(pdf)).pages[page - 1].extract_text() or ""
    except Exception:
        return ""
    return t if len(t.strip()) > 200 else ""      # a watermark alone is not a text layer


def parse_pages(spec, total):
    """'1-5,9,20-22' -> [1..5, 9, 20..22]; None -> all pages."""
    if not spec:
        return list(range(1, total + 1))
    out = []
    for part in spec.split(","):
        a, _, b = part.strip().partition("-")
        out += list(range(int(a), int(b or a) + 1))
    return [p for p in out if 1 <= p <= total]


def build_index(pdf, out_dir, engine="tesseract", pages=None, dpi=200, lang="urd+eng", client=None, progress=None):
    """OCR the pages into out_dir/pages/NNNN.txt. Resumable. Returns the page numbers done."""
    from pypdf import PdfReader
    total = len(PdfReader(str(pdf)).pages)
    wanted = parse_pages(pages, total)
    folder = Path(out_dir) / "pages"
    folder.mkdir(parents=True, exist_ok=True)
    (Path(out_dir) / "meta.json").write_text(json.dumps({"pdf": str(pdf), "engine": engine, "dpi": dpi, "lang": lang,
                                                         "model": getattr(client, "model", None), "pages": total}), encoding="utf8")
    done = []
    work = tempfile.mkdtemp(prefix="book-ocr-")
    try:
        for p in wanted:
            target = folder / f"{p:04d}.txt"
            if target.exists() and target.stat().st_size > 0:
                done.append(p)
                continue
            text = text_layer(pdf, p) if engine == "auto" else ""
            if not text:
                png = render_page(pdf, p, dpi, work)
                text = ocr_vlm(png, client) if engine == "vlm" else ocr_tesseract(png, lang)
                os.remove(png)
            target.write_text(text, encoding="utf8")
            done.append(p)
            if progress:
                progress(p, len(wanted))
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return done
