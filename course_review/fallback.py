"""Model-assisted recovery of question structure.

The rule-based parser handles the formats it knows. When it finds no questions, or
numbering it cannot trust, this module asks a local model which paragraphs start a
question. The model only PROPOSES; every proposal is verified against the document in
plain code (the paragraph must exist, and the quote the model gives must appear in
it), unverifiable proposals are discarded, and an extraction with too many rejections
is thrown away entirely. Results are cached per file so a re-run gives the same answer.
"""
import hashlib
import json
import os
import time
from dataclasses import dataclass, field
from pathlib import Path

from .llm import LLMError
from .questions import OPTION_START, questions_from_starts
from .textutil import normalize

PROMPT_VERSION = "q1"
MAX_CHUNK_CHARS = 3500
MAX_PARA_CHARS = 240
MAX_REJECTED_SHARE = 0.3

SYSTEM = """You label the structure of a school exam or quiz document.
You are given the document as numbered paragraphs, like: [12] text of paragraph twelve.
Find the paragraphs where a NEW QUESTION that the student must answer begins.

Rules:
- A question starts at the paragraph that carries its number or its wording, not at its answer options.
- Answer options (a, b, c, d / A, B, C, D), hints, sample answers, instructions, headings and teacher
  guidance are NOT questions.
- Sub-parts of one question (a, b, c, or i, ii, iii) belong to that question; they are not new questions.
- If a numbered instruction such as "Question 2: answer the following" is followed by several
  separate unnumbered questions, list each separate question, not the instruction.
- Never invent a paragraph number. Copy "quote" exactly from the start of that paragraph (8 to 30 characters).

Answer with JSON only: {"questions": [{"para": 12, "quote": "first words of paragraph 12"}]}
If there are no questions, answer {"questions": []}."""


@dataclass
class Extraction:
    starts: list                       # verified paragraph indexes
    proposed: int = 0
    rejected: int = 0
    chunks: int = 0
    seconds: float = 0.0
    cached: bool = False
    model: str = ""
    usable: bool = True
    note: str = ""
    quotes: dict = field(default_factory=dict)

    def summary(self) -> str:
        if not self.usable:
            return self.note
        return (f"model {self.model}: {len(self.starts)} question(s) recognised, {self.proposed} proposed, "
                f"{self.rejected} rejected by verification" + (", from cache" if self.cached else ""))

    def stats(self) -> dict:
        return {"model": self.model, "recognised": len(self.starts), "proposed": self.proposed,
                "rejected": self.rejected, "chunks": self.chunks, "seconds": round(self.seconds, 1),
                "cached": self.cached, "usable": self.usable, "note": self.note}


class Cache:
    """One small JSON file per (document, region, model, prompt version)."""

    def __init__(self, directory):
        self.dir = Path(directory)

    def _path(self, key):
        return self.dir / (hashlib.sha1(key.encode("utf8")).hexdigest() + ".json")

    def get(self, key):
        try:
            return json.loads(self._path(key).read_text(encoding="utf8"))
        except (OSError, ValueError):
            return None

    def put(self, key, value):
        try:
            self.dir.mkdir(parents=True, exist_ok=True)
            self._path(key).write_text(json.dumps(value, ensure_ascii=False), encoding="utf8")
        except OSError:
            pass  # a cache that cannot be written is not an error


def file_sha1(path) -> str:
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def listing(info, lo=0, hi=None):
    """[(index, text)] for non-empty paragraphs in [lo, hi), text shortened for the prompt."""
    out = []
    for p in info.paras:
        if p.idx < lo or (hi is not None and p.idx >= hi):
            continue
        t = " ".join(p.text.split())
        if not t:
            continue
        out.append((p.idx, t))
    return out


def chunks_of(items, max_chars=MAX_CHUNK_CHARS):
    chunks, cur, size = [], [], 0
    for idx, t in items:
        line_len = min(len(t), MAX_PARA_CHARS) + 8
        if cur and size + line_len > max_chars:
            chunks.append(cur)
            cur, size = [], 0
        cur.append((idx, t))
        size += line_len
    if cur:
        chunks.append(cur)
    return chunks


def render(chunk) -> str:
    return "\n".join(f"[{idx}] {t[:MAX_PARA_CHARS]}" for idx, t in chunk)


def verify(proposals, chunk):
    """Keep only proposals whose paragraph exists in this chunk and whose quote appears in it.
    Returns (accepted: {idx: quote}, rejected_count)."""
    by_idx = {idx: normalize(t) for idx, t in chunk}
    accepted, rejected = {}, 0
    for item in proposals:
        try:
            idx = int(item.get("para"))
            quote = normalize(str(item.get("quote", "")))
        except (AttributeError, TypeError, ValueError):
            rejected += 1
            continue
        text = by_idx.get(idx)
        if text is None or len(quote) < 3 or quote not in text:
            rejected += 1
            continue
        if OPTION_START.match(text):   # an answer option is never a question start
            rejected += 1
            continue
        accepted[idx] = quote
    return accepted, rejected


def extract(info, client, cache=None, doc_hash="", lo=0, hi=None) -> Extraction:
    """Ask the model for question starts in paragraphs [lo, hi). Never raises."""
    items = listing(info, lo, hi)
    key = f"{PROMPT_VERSION}|{client.name}|{doc_hash}|{lo}|{hi}"
    if cache is not None and doc_hash:
        hit = cache.get(key)
        if hit is not None:
            return Extraction(starts=hit["starts"], proposed=hit["proposed"], rejected=hit["rejected"], chunks=hit["chunks"],
                              cached=True, model=client.name, usable=hit["usable"], note=hit.get("note", ""))
    if not items:
        return Extraction(starts=[], model=client.name, usable=False, note="the region has no text")
    ex = Extraction(starts=[], model=client.name)
    t0 = time.time()
    accepted_all = {}
    for chunk in chunks_of(items):
        ex.chunks += 1
        try:
            reply = client.chat_json(SYSTEM, render(chunk))
        except LLMError as e:
            return Extraction(starts=[], model=client.name, usable=False, chunks=ex.chunks, seconds=time.time() - t0,
                              note=f"model unavailable: {e}")
        proposals = reply.get("questions") if isinstance(reply, dict) else None
        if not isinstance(proposals, list):
            return Extraction(starts=[], model=client.name, usable=False, chunks=ex.chunks, seconds=time.time() - t0,
                              note="model reply did not have a 'questions' list")
        accepted, rejected = verify(proposals, chunk)
        ex.proposed += len(proposals)
        ex.rejected += rejected
        accepted_all.update(accepted)
    ex.seconds = time.time() - t0
    ex.starts = sorted(accepted_all)
    ex.quotes = accepted_all
    if ex.proposed and ex.rejected / ex.proposed > MAX_REJECTED_SHARE:
        ex.usable = False
        ex.note = f"{ex.rejected} of {ex.proposed} model proposals failed verification, so the extraction was discarded"
    elif not ex.starts:
        ex.usable = False
        ex.note = "the model recognised no questions"
    if cache is not None and doc_hash:
        cache.put(key, {"starts": ex.starts, "proposed": ex.proposed, "rejected": ex.rejected, "chunks": ex.chunks,
                        "usable": ex.usable, "note": ex.note})
    return ex


def recover_questions(info, model_cfg, doc, lo=0, hi=None):
    """Questions recovered by the model for one document region, or (None, reason).
    The Question objects are marked source='model' and carry the verification stats."""
    if model_cfg is None:
        return None, "model fallback is off"
    ex = extract(info, model_cfg.client, model_cfg.cache, file_sha1(doc.abs) if doc.abs and os.path.exists(doc.abs) else "", lo, hi)
    model_cfg.stats.append({"doc": doc.rel, "lo": lo, "hi": hi, **ex.stats()})
    if not ex.usable:
        return None, ex.summary()
    qs = questions_from_starts(info, [(s, n + 1) for n, s in enumerate(ex.starts)], hi, source="model")
    meta = ex.stats()
    for q in qs:
        q.meta = meta
    return qs, ex.summary()


@dataclass
class ModelConfig:
    client: object
    cache: object = None
    mode: str = "suggest"              # 'suggest': model answers need confirming; 'decide': Fails are written
    stats: list = field(default_factory=list)
