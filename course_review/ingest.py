"""Safe unzip and classification of package files.

Classification is driven by the profile (patterns), not by hard-coded names, and
file names are only a hint: anything that cannot be placed is returned in
Package.unclassified so the report can show it.
"""
import os
import re
import zipfile
from pathlib import Path

import yaml

from .models import DocRef, LessonKey, Package

DEFAULT_PROFILE = Path(__file__).with_name("profile.default.yaml")


def load_profile(path=None) -> dict:
    with open(path or DEFAULT_PROFILE, encoding="utf8") as f:
        return yaml.safe_load(f)


class UnsafeArchive(Exception):
    pass


def safe_extract(zip_path, dest, max_files=20000, max_total=4 << 30, max_ratio=1000):
    """Extract zip_path into dest, refusing path traversal and zip bombs."""
    dest = Path(dest).resolve()
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as z:
        infos = z.infolist()
        if len(infos) > max_files:
            raise UnsafeArchive(f"archive has {len(infos)} entries (limit {max_files})")
        total = sum(i.file_size for i in infos)
        if total > max_total:
            raise UnsafeArchive(f"archive expands to {total} bytes (limit {max_total})")
        for i in infos:
            if i.compress_size and i.file_size / max(i.compress_size, 1) > max_ratio and i.file_size > (50 << 20):
                raise UnsafeArchive(f"suspicious compression ratio on {i.filename}")
            target = (dest / i.filename).resolve()
            if dest != target and dest not in target.parents:
                raise UnsafeArchive(f"entry escapes the extraction folder: {i.filename}")
        z.extractall(dest)
    return dest


def find_root(dest) -> Path:
    """If the archive has a single top-level folder, that folder is the package root."""
    entries = [p for p in Path(dest).iterdir() if not p.name.startswith(("__MACOSX", "."))]
    if len(entries) == 1 and entries[0].is_dir():
        return entries[0]
    return Path(dest)


_VERSION = re.compile(r"v(\d+(?:\.\d+)*)", re.I)


def _version_key(name: str):
    m = _VERSION.search(name)
    return tuple(int(x) for x in m.group(1).split(".")) if m else ()


def _letter_for_part(n: int) -> str:
    return chr(ord("A") + n - 1) if 1 <= n <= 26 else ""


def _parse_location(parts, profile):
    """parts: path components below the root, last one is the file name."""
    pp = profile["path_patterns"]
    chapter = lesson = None
    variant = ""
    folder_label = ""
    folders, fname = parts[:-1], parts[-1]
    for p in folders:
        m = re.search(pp["chapter"], p)
        if m and chapter is None:
            chapter = int(m.group(1))
    for p in folders:
        m = re.search(pp["lesson"], p)
        if m and lesson is None:
            lesson = int(m.group(1))
            folder_label = p
            part = re.search(r"(?i)part[-_ ]*(\d+)(?:[-_ ]*([A-Za-z])(?=[-_ .]))?", p)
            if part:
                variant = (part.group(2) or _letter_for_part(int(part.group(1)))).upper()
    stem = os.path.splitext(fname)[0]
    if chapter is None:
        m = re.search(pp["chapter"], stem)
        if m:
            chapter = int(m.group(1))
    if lesson is None:
        m = re.search(pp["lesson"], stem)
        if m:
            lesson = int(m.group(1))
    if lesson is not None and not variant:
        m = re.search(pp["lesson"] + r"[-_ ]*([A-Za-z])(?=[-_ .])", stem)
        if m:
            variant = m.group(2).upper()
    return chapter, lesson, variant, folder_label


def classify(root, profile) -> Package:
    root = Path(root)
    docs, unclassified = [], []
    chapter_titles, lesson_labels = {}, {}
    rules = profile["doc_types"]
    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.name.startswith(("~$", ".")):
            continue
        rel = path.relative_to(root).as_posix()
        parts = rel.split("/")
        ext = path.suffix.lower().lstrip(".")
        stem = path.stem
        chapter, lesson, variant, folder_label = _parse_location(parts, profile)
        matched = None
        for r in rules:
            if ext in r.get("ext", [ext]) and re.search(r["pattern"], stem, re.I):
                matched = r
                break
        if matched is None:
            unclassified.append(rel)
            continue
        ref = DocRef(rel=rel, abs=str(path), ext=ext, doc_type=matched["type"], chapter=chapter,
                     lesson=lesson, variant=variant, scope=matched.get("scope", "lesson"),
                     folder_label=folder_label)
        if matched.get("note"):
            ref.notes.append(matched["note"])
        # subject-scope files have no chapter/lesson of their own
        if ref.scope == "subject":
            ref.chapter = ref.lesson = None
            ref.variant = ""
        if ref.scope == "chapter":
            ref.lesson = None
            ref.variant = ""
        docs.append(ref)
        for p in parts[:-1]:
            m = re.search(profile["path_patterns"]["chapter"], p)
            if m and ref.chapter is not None and int(m.group(1)) == ref.chapter and re.match(r"(?i)(chapter|chpater)", p):
                chapter_titles.setdefault(ref.chapter, p)
        if ref.lesson is not None:
            lesson_labels.setdefault(ref.key, folder_label)

    # newest version of each subject-scope document wins; older ones are superseded
    by_type = {}
    for d in docs:
        if d.doc_type in ("data_bank", "pop_quiz", "specification"):
            by_type.setdefault(d.doc_type, []).append(d)
    for dtype, group in by_type.items():
        if len(group) > 1:
            group.sort(key=lambda d: _version_key(os.path.basename(d.rel)))
            for old in group[:-1]:
                old.superseded = True
                old.notes.append(f"Superseded by {os.path.basename(group[-1].rel)}")

    subject = root.name
    return Package(root=str(root), subject=subject, docs=docs, unclassified=unclassified,
                   chapter_titles=chapter_titles, lesson_labels=lesson_labels)
