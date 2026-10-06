"""Learn a new question format once, then parse it with plain code.

A person points at a document the rule-based parser does not understand. A model proposes a
regular expression for the start of each question; code validates it against the document,
shows what it matches, and only after a person approves is it saved into a profile file,
from which it is used with no model at all.
"""
import re
from datetime import date
from pathlib import Path

import yaml

from .fallback import listing, render
from .llm import LLMError
from .questions import numbering_is_regular, question_number
from .textutil import normalize, to_western_digits

SYSTEM = """You help a program recognise how the questions in a school exam document are numbered.
You get the document as numbered paragraphs, like [12] text. Write ONE Python regular expression that matches the
START of every paragraph where a question begins, and of no other paragraph.
Rules for the pattern:
- It is applied with re.match to the paragraph text (so it is anchored at the start; do not add ^).
- It must have exactly one capture group, around the question number (digits or a roman numeral).
- Keep it simple and general; do not copy words from the questions themselves.
Answer with JSON only: {"pattern": "...", "explanation": "one sentence"}"""

DANGEROUS = re.compile(r"\((?:[^()]*[*+][^()]*)\)[*+{]")   # nested quantifiers such as (a+)+ can hang the regex engine
MAX_PATTERN = 160


class FormatError(Exception):
    pass


class _N:
    def __init__(self, num):
        self.num = num


def validate(pattern, info, profile=None):
    """Check a candidate against the document. Returns (compiled, matches[(para idx, number, text)], problems)."""
    problems = []
    if not pattern or len(pattern) > MAX_PATTERN or DANGEROUS.search(pattern):
        raise FormatError("the pattern is empty, too long, or has nested quantifiers")
    try:
        rx = re.compile(pattern, re.I)
    except re.error as e:
        raise FormatError(f"not a valid regular expression: {e}")
    if rx.groups != 1:
        raise FormatError("the pattern must have exactly one capture group (the question number)")
    matches = []
    for p in info.paras:
        if p.in_table or not p.text.strip():
            continue
        t = to_western_digits(normalize(p.text))
        m = rx.match(t)
        if m:
            n = question_number(m.group(1))
            if n is not None:
                matches.append((p.idx, n, p.text.strip()))
    if len(matches) < 2:
        problems.append(f"it matches only {len(matches)} paragraph(s)")
    elif not numbering_is_regular([_N(n) for _, n, _ in matches]):
        problems.append(f"the numbers it finds run {[n for _, n, _ in matches][:12]}, not 1, 2, 3...")
    return rx, matches, problems


def propose(client, info, profile=None):
    """Ask the model for a pattern and validate it. Returns dict(pattern, explanation, matches, problems)."""
    items = listing(info)[:80]
    try:
        reply = client.chat_json(SYSTEM, render(items))
    except LLMError as e:
        raise FormatError(f"model unavailable: {e}")
    pattern = reply.get("pattern") if isinstance(reply, dict) else None
    if not isinstance(pattern, str):
        raise FormatError("the model did not return a pattern")
    _, matches, problems = validate(pattern, info)
    return {"pattern": pattern, "explanation": str(reply.get("explanation", "")).strip(), "matches": matches, "problems": problems}


def save_pattern(profile_path, pattern, note=""):
    """Add the pattern to a profile file (created if missing) under vocab.extra_question_patterns.
    Returns False if it was already there."""
    path = Path(profile_path)
    data = (yaml.safe_load(path.read_text(encoding="utf8")) if path.exists() else {}) or {}
    entries = data.setdefault("vocab", {}).setdefault("extra_question_patterns", [])
    if any((e["pattern"] if isinstance(e, dict) else e) == pattern for e in entries):
        return False
    entries.append({"pattern": pattern, "note": f"{note} (learned {date.today().isoformat()})".strip()})
    path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf8")
    return True
