"""Plain data types shared across the tool."""
from dataclasses import dataclass, field
from typing import Optional

PASS, FAIL, NA, REVIEW = "pass", "fail", "na", "needs_review"
STATUSES = (PASS, FAIL, NA, REVIEW)


@dataclass(frozen=True)
class LessonKey:
    """A lesson inside a chapter. lesson is None for a chapter that has no
    lesson folders; variant is 'A', 'B'... when a lesson is split in parts."""
    chapter: Optional[int]
    lesson: Optional[int]
    variant: str = ""

    def sort_key(self):
        return (self.chapter or 0, self.lesson if self.lesson is not None else -1, self.variant)

    def short(self) -> str:
        if self.lesson is None:
            return f"Ch{self.chapter}"
        return f"Ch{self.chapter}-L{self.lesson}{self.variant}"


@dataclass
class DocRef:
    rel: str                      # path relative to the package root, '/' separated
    abs: str
    ext: str                      # 'docx' | 'pptx' | 'pdf' | 'xlsx' | other
    doc_type: Optional[str]
    chapter: Optional[int]
    lesson: Optional[int]
    variant: str = ""
    scope: str = "lesson"         # lesson | chapter | subject
    notes: list = field(default_factory=list)
    superseded: bool = False
    folder_label: str = ""

    @property
    def key(self) -> LessonKey:
        return LessonKey(self.chapter, self.lesson, self.variant)


@dataclass
class Rule:
    code: str
    output_type: str
    scope: str                    # 'Per lesson' | 'Per chapter' | 'Per subject'
    text: str


@dataclass
class Finding:
    code: str
    status: str                   # pass | fail | na | needs_review
    message: str = ""
    evidence: list = field(default_factory=list)
    lesson: Optional[LessonKey] = None
    chapter: Optional[int] = None
    doc: Optional[str] = None
    partial: bool = False         # a pass that covers only part of the rule text
    method: str = "deterministic"

    def to_dict(self):
        return {
            "code": self.code, "status": self.status, "message": self.message,
            "evidence": self.evidence, "lesson": self.lesson.short() if self.lesson else None,
            "chapter": self.chapter, "doc": self.doc, "partial": self.partial, "method": self.method,
        }


@dataclass
class Package:
    root: str
    subject: str
    docs: list
    unclassified: list
    chapter_titles: dict          # chapter number -> folder label
    lesson_labels: dict           # LessonKey -> folder label
