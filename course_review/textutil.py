"""Script detection and text normalisation shared by parsers and checks.

Urdu text in real documents mixes Arabic and Urdu code points for the same
letter (for example U+064A vs U+06CC for yeh), so every comparison against a
vocabulary goes through normalize().
"""
import re
import unicodedata

_ARABIC_RANGES = ((0x0600, 0x06FF), (0x0750, 0x077F), (0x08A0, 0x08FF), (0xFB50, 0xFDFF), (0xFE70, 0xFEFF))

_LETTER_MAP = {
    0x064A: 0x06CC,  # Arabic yeh -> Farsi/Urdu yeh
    0x0649: 0x06CC,  # alef maksura -> yeh
    0x0643: 0x06A9,  # Arabic kaf -> keheh
    0x0647: 0x06C1,  # Arabic heh -> heh goal
    0x06C2: 0x06C1,  # heh goal with hamza (غزوۂ) -> heh goal
    0x06C0: 0x06C1,  # heh with yeh above -> heh goal
}
_STRIP = re.compile(
    "[ً-ٰٟۖ-ۜ۟-۪ۨ-ۭـ"
    "​-‏‪-‮⁦-⁩﻿]"
)
EASTERN_DIGITS = re.compile("[٠-٩۰-۹]")
_DIGIT_MAP = {**{0x0660 + i: ord("0") + i for i in range(10)}, **{0x06F0 + i: ord("0") + i for i in range(10)}}


def is_arabic_letter(ch: str) -> bool:
    o = ord(ch)
    return ch.isalpha() and any(a <= o <= b for a, b in _ARABIC_RANGES)


def is_latin_letter(ch: str) -> bool:
    return ("A" <= ch <= "Z") or ("a" <= ch <= "z")


def script_counts(text: str):
    """(arabic_letters, latin_letters)"""
    ar = sum(1 for c in text if is_arabic_letter(c))
    la = sum(1 for c in text if is_latin_letter(c))
    return ar, la


def dominant_script(text: str):
    ar, la = script_counts(text)
    if ar == 0 and la == 0:
        return None
    return "arabic" if ar >= la else "latin"


def to_western_digits(text: str) -> str:
    return text.translate(_DIGIT_MAP)


def normalize(text: str) -> str:
    """Fold presentation forms, unify Urdu/Arabic letter variants, strip
    diacritics and invisible marks, collapse whitespace."""
    text = unicodedata.normalize("NFKC", text or "")
    text = text.translate(_LETTER_MAP)
    text = _STRIP.sub("", text)
    text = text.replace(" ", " ")
    return re.sub(r"\s+", " ", text).strip()


def starts_with_label(text: str, labels, max_plain_len: int = 80):
    """True if normalised text begins with one of the labels and looks like a
    heading (label alone, label then a separator, or a short line)."""
    t = normalize(text)
    t = re.sub(r"^[\s\d.()\-–—•*#]+", "", t)
    for lab in labels:
        n = normalize(lab)
        if not n:
            continue
        if t.lower().startswith(n.lower()):
            rest = t[len(n):].lstrip()
            if not rest or rest[0] in ":：—–-|·" or len(t) <= max_plain_len:
                return True
    return False
