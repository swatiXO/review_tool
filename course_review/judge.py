"""Model-assisted judgement of one checklist rule at a time.

The model is given the rule's wording (from the workbook) and only the material needed to
decide it, and must answer with a verdict, a reason and quotes. A verdict counts only if at
least one quote can be found, word for word, in the material that was supplied: a verdict
the model cannot support from the text is discarded, not shown. The outcome is always a
needs-review suggestion, never a Pass or Fail in the workbook.
"""
import hashlib
import json
import re
from dataclasses import dataclass, field
from typing import Optional

from .llm import LLMError
from .textutil import normalize

PROMPT_VERSION = "j1"
MAX_MATERIAL_CHARS = 6000
MIN_QUOTE_CHARS = 8

SYSTEM = """You are a careful reviewer of school course material. You judge ONE checklist rule about
ONE piece of material. The material may be in Urdu or English.

Answer with JSON only:
{"verdict": "pass" | "fail" | "unclear", "reason": "one or two plain sentences", "evidence": [{"quote": "..."}]}

Rules for your answer:
- Judge only the rule you are given, using only the material you are given. Do not use outside knowledge.
- "pass" means the material satisfies the rule; "fail" means it clearly does not; "unclear" means the
  material does not let you decide. Prefer "unclear" to guessing.
- Every "quote" must be copied exactly from the material (8 to 80 characters) and must be what your
  verdict rests on. For "fail", quote the passage that is wrong or the SLO / item that is not covered.
- Write the reason in English."""


@dataclass
class Material:
    label: str
    text: str


@dataclass
class Judgement:
    verdict: str = "unclear"           # pass | fail | unclear
    reason: str = ""
    quotes: list = field(default_factory=list)      # verified quotes
    rejected_quotes: int = 0
    extra: dict = field(default_factory=dict)
    usable: bool = False
    note: str = ""
    cached: bool = False
    cut: bool = False                  # the material was longer than the limit and was shortened


def _material_block(materials, budget=MAX_MATERIAL_CHARS):
    parts, used, cut = [], 0, False
    for m in materials:
        room = max(budget - used, 0)
        text = m.text.strip()
        if len(text) > room:
            text, cut = text[:room] + " [...cut]", True
        parts.append(f"### {m.label}\n{text}")
        used += len(text)
    return "\n\n".join(parts), cut


def verify_quotes(evidence, materials):
    haystack = [normalize(m.text) for m in materials]
    ok, rejected = [], 0
    for item in evidence if isinstance(evidence, list) else []:
        q = normalize(str(item.get("quote", ""))) if isinstance(item, dict) else ""
        if len(q) >= MIN_QUOTE_CHARS and any(q in h for h in haystack):
            ok.append(str(item["quote"]).strip())
        else:
            rejected += 1
    return ok, rejected


def judge(model, code, rule_text, materials, instruction, extra_keys=(), validate_extra=None):
    """Ask the model; returns a Judgement (usable only when the verdict is supported by a verified quote)."""
    block, cut = _material_block(materials)
    user = f"Rule {code}: {rule_text}\n\nWhat to decide: {instruction}\n\n{block}"
    key = hashlib.sha1("|".join([PROMPT_VERSION, model.client.name, code, user]).encode("utf8")).hexdigest()
    if model.cache is not None:
        hit = model.cache.get(key)
        if hit is not None:
            j = Judgement(**hit)
            j.cached = True
            model.stats.append({"doc": code, "judge": True, "usable": j.usable, "cached": True, "note": j.note})
            return j
    try:
        reply = model.client.chat_json(SYSTEM, user)
    except LLMError as e:
        j = Judgement(note=f"model unavailable: {e}")
        model.stats.append({"doc": code, "judge": True, "usable": False, "cached": False, "note": j.note})
        return j
    j = Judgement(cut=cut)
    if not isinstance(reply, dict) or reply.get("verdict") not in ("pass", "fail", "unclear"):
        j.note = "the model reply was not in the expected form"
    else:
        j.verdict = reply["verdict"]
        j.reason = str(reply.get("reason", "")).strip()
        j.quotes, j.rejected_quotes = verify_quotes(reply.get("evidence"), materials)
        for k in extra_keys:
            j.extra[k] = reply.get(k)
        if j.verdict == "unclear":
            j.note = "the model could not decide from the material"
        elif not j.quotes:
            j.verdict = "unclear"
            j.note = "the model gave a verdict but no quote that appears in the material, so it was discarded"
        elif validate_extra is not None and not validate_extra(j.extra):
            j.verdict = "unclear"
            j.note = "the model's answer was inconsistent with the material"
        else:
            j.usable = True
    if model.cache is not None:
        model.cache.put(key, {"verdict": j.verdict, "reason": j.reason, "quotes": j.quotes,
                              "rejected_quotes": j.rejected_quotes, "extra": j.extra, "usable": j.usable, "note": j.note,
                              "cut": j.cut})
    model.stats.append({"doc": code, "judge": True, "usable": j.usable, "cached": False, "note": j.note})
    return j


def suggestion(code, j, **where):
    """A needs-review finding that carries the model's suggestion, or a note that it could not decide."""
    from .checks.common import result
    from .models import REVIEW
    if j.usable:
        msg = f"[model-assisted] Would be {j.verdict}: {j.reason}"
        ev = [f'Quote: "{q}"' for q in j.quotes[:3]]
        if j.rejected_quotes:
            ev.append(f"{j.rejected_quotes} quote(s) from the model were not in the material and were ignored")
        if j.cut:
            ev.append("The material was longer than the model's limit and was shortened, so this may miss text at the end")
    else:
        msg = f"[model-assisted] The model could not decide ({j.note})"
        ev = []
    return result(code, REVIEW, msg, ev, method="model", **where)


def cached_chat(model, tag, system, user):
    """A JSON reply for (tag, user) from the cache or the model. Raises LLMError."""
    key = hashlib.sha1("|".join([PROMPT_VERSION, model.client.name, tag, system, user]).encode("utf8")).hexdigest()
    if model.cache is not None:
        hit = model.cache.get(key)
        if hit is not None:
            return hit["reply"]
    reply = model.client.chat_json(system, user)
    if model.cache is not None:
        model.cache.put(key, {"reply": reply})
    return reply
