"""Learning a question format once: validation, saving, and parsing it afterwards without a model."""
import os

import pytest
import yaml

from course_review import cli, formats, ingest
from course_review.docx_model import parse_docx
from course_review.questions import numbering_is_regular, parse_questions, question_number

from fakes import _Base
from helpers import PROFILE, add, new_doc, save

TASK_LINES = [f"Task {n} >> Describe the idea number {n}." for n in range(1, 6)]


def doc(tmp_path, lines=TASK_LINES, name="x.docx"):
    d = new_doc()
    add(d, "Exam")
    for l in lines:
        add(d, l)
    return parse_docx(save(d, tmp_path, name))


class PatternModel(_Base):
    def __init__(self, reply):
        super().__init__()
        self.reply = reply

    def chat_json(self, system, user):
        self.calls += 1
        return self.reply


# ------------------------------------------------------------------ numbers
@pytest.mark.parametrize("text,value", [("7", 7), ("٣", 3), ("iv", 4), ("XII", 12), ("ix", 9), ("abc", None), ("", None)])
def test_question_numbers_from_digits_and_roman_numerals(text, value):
    assert question_number(text) == value


# --------------------------------------------------------------- validation
def test_a_good_pattern_matches_every_question_in_order(tmp_path):
    rx, matches, problems = formats.validate(r"Task\s*(\d+)\s*>>", doc(tmp_path), PROFILE)
    assert [n for _, n, _ in matches] == [1, 2, 3, 4, 5] and not problems


@pytest.mark.parametrize("pattern", ["", "(a+)+b", "Task (\\d+) (", "Task \\d+", "x" * 400, "(a)(b)"])
def test_unsafe_or_wrong_shaped_patterns_are_refused(tmp_path, pattern):
    with pytest.raises(formats.FormatError):
        formats.validate(pattern, doc(tmp_path), PROFILE)


def test_a_pattern_that_matches_too_little_or_in_the_wrong_order_is_flagged(tmp_path):
    _, matches, problems = formats.validate(r"Task\s*(1)\s*>>", doc(tmp_path), PROFILE)
    assert problems and "only 1" in problems[0]
    lines = ["Task 3 >> a", "Task 1 >> b", "Task 2 >> c"]
    _, _, problems = formats.validate(r"Task\s*(\d+)", doc(tmp_path, lines, "y.docx"), PROFILE)
    assert problems and "not 1, 2, 3" in problems[0]


def test_propose_validates_what_the_model_suggests(tmp_path):
    found = formats.propose(PatternModel({"pattern": r"Task\s*(\d+)\s*>>", "explanation": "Task N >>"}), doc(tmp_path), PROFILE)
    assert len(found["matches"]) == 5 and not found["problems"] and found["explanation"] == "Task N >>"
    with pytest.raises(formats.FormatError):
        formats.propose(PatternModel({"nothing": 1}), doc(tmp_path), PROFILE)
    with pytest.raises(formats.FormatError):
        formats.propose(PatternModel({"pattern": "(a+)+b"}), doc(tmp_path), PROFILE)


# --------------------------------------------------- saving and using it later
def test_saved_pattern_is_used_with_no_model_and_the_default_profile_still_applies(tmp_path):
    profile_path = str(tmp_path / "mine.yaml")
    assert formats.save_pattern(profile_path, r"Task\s*(\d+)\s*>>", note="x.docx")
    assert not formats.save_pattern(profile_path, r"Task\s*(\d+)\s*>>")          # saving twice adds nothing
    saved = yaml.safe_load(open(profile_path, encoding="utf8"))
    assert saved["vocab"]["extra_question_patterns"][0]["pattern"] == r"Task\s*(\d+)\s*>>"
    info = doc(tmp_path)
    assert parse_questions(info, PROFILE) == []                                    # unknown to the built-in rules
    profile = ingest.load_profile(profile_path)
    qs = parse_questions(info, profile)
    assert [q.num for q in qs] == [1, 2, 3, 4, 5] and numbering_is_regular(qs) and qs[0].source == "rules"
    assert profile["lesson_plan_sections"] == PROFILE["lesson_plan_sections"]       # the rest of the built-in profile is untouched
    known = doc(tmp_path, ["Question 1: a", "Question 2: b"], "z.docx")
    assert len(parse_questions(known, profile)) == 2                               # built-in formats keep working


def test_roman_numerals_work_once_the_pattern_is_learned(tmp_path):
    profile_path = str(tmp_path / "roman.yaml")
    formats.save_pattern(profile_path, r"([ivx]+)\.")
    lines = [f"{r}. What is the capital of country {n}?" for n, r in enumerate(("i", "ii", "iii", "iv", "v"), 1)]
    qs = parse_questions(doc(tmp_path, lines), ingest.load_profile(profile_path))
    assert [q.num for q in qs] == [1, 2, 3, 4, 5]


def test_profile_overlay_merges_dicts_and_replaces_other_values(tmp_path):
    p = tmp_path / "p.yaml"
    p.write_text(yaml.safe_dump({"thresholds": {"min_image_dpi": 300}, "formatted_types": ["lesson_plan"]}), encoding="utf8")
    prof = ingest.load_profile(str(p))
    assert prof["thresholds"]["min_image_dpi"] == 300 and prof["thresholds"]["text_size_tolerance"] == PROFILE["thresholds"]["text_size_tolerance"]
    assert prof["formatted_types"] == ["lesson_plan"]


# ------------------------------------------------------------------- the CLI
def test_cli_learn_format_with_an_explicit_pattern(tmp_path, capsys):
    path = str(tmp_path / "task.docx")
    d = new_doc()
    for l in TASK_LINES:
        add(d, l)
    d.save(path)
    prof = str(tmp_path / "profile.yaml")
    cli.main(["learn-format", path, "--save", prof, "--pattern", r"Task\s*(\d+)\s*>>", "--yes"])
    out = capsys.readouterr().out
    assert "matches 5 paragraph" in out and "Saved to" in out and os.path.exists(prof)


def test_cli_learn_format_refuses_a_pattern_with_problems(tmp_path):
    path = str(tmp_path / "task.docx")
    d = new_doc()
    for l in TASK_LINES:
        add(d, l)
    d.save(path)
    prof = str(tmp_path / "profile.yaml")
    with pytest.raises(SystemExit) as e:
        cli.main(["learn-format", path, "--save", prof, "--pattern", r"Task\s*(1)", "--yes"])
    assert "Not saved" in str(e.value) and not os.path.exists(prof)
