"""Tests against a REAL model. They skip unless OLLAMA_URL is set.

  set OLLAMA_URL=https://your-link.ngrok-free.dev        (PowerShell: $env:OLLAMA_URL = "...")
  set OLLAMA_MODEL=qwen3:14b                             (optional)
  set SAMPLE_ZIP=C:\\path\\to\\Grade-6-Islamiat.zip          (optional, enables the sample-package evaluation)
  python -m pytest tests/test_live_model.py -v -s

Models are not deterministic across versions, so these assert accuracy thresholds, not exact
behaviour, and print what they found. If a threshold fails, read the printed table first: the
point of these tests is to measure whether the chosen model is good enough to rely on.
"""
import os

import pytest

from course_review import evaluate, fallback
from course_review.docx_model import parse_docx
from course_review.llm import OllamaClient
from course_review.questions import parse_questions

from helpers import PROFILE, add, checklist, new_doc, save

pytestmark = [pytest.mark.live, pytest.mark.skipif(not os.environ.get("OLLAMA_URL"),
                                                    reason="set OLLAMA_URL (your ngrok link) to run the live model tests")]


@pytest.fixture(scope="module")
def client():
    c = OllamaClient()
    ok, msg = c.available()
    if not ok:
        pytest.skip(msg)
    print(f"\nusing {msg}")
    return c


def doc_from(tmp_path, lines, name="x.docx"):
    d = new_doc()
    for l in lines:
        add(d, l)
    return parse_docx(save(d, tmp_path, name))


def run(client, info):
    return fallback.extract(info, client)


# ------------------------------------------------------------------ reachability
def test_model_answers_json(client):
    reply = client.chat_json("Reply with JSON only.", 'Return {"ok": true}')
    assert isinstance(reply, dict)


# ---------------------------------------------------------------- English formats
ENGLISH_FORMATS = ["{}.", "{})", "Q{}.", "Q.{}", "Q {}:", "Question {})", "({})", "{} -", "Q{}", "Q{}:", "Question {}.", "{}:"]


def test_common_english_formats_give_the_right_count(client, tmp_path):
    failures = []
    for fmt in ENGLISH_FORMATS:
        lines = [fmt.format(i) + f" Name one thing about topic {i}?" for i in range(1, 7)]
        ex = run(client, doc_from(tmp_path, lines))
        ok = ex.usable and len(ex.starts) == 6
        print(f"  {fmt:14} -> {len(ex.starts)} question(s), {ex.rejected}/{ex.proposed} rejected, {ex.seconds:.0f}s")
        if not ok:
            failures.append(fmt)
    assert len(failures) <= 3, f"count wrong for formats: {failures}"


def test_roman_numerals_which_the_rules_do_not_know(client, tmp_path):
    lines = [f"{r}. What is the capital of country {n}?" for n, r in enumerate(("i", "ii", "iii", "iv", "v", "vi"), 1)]
    ex = run(client, doc_from(tmp_path, lines))
    assert ex.usable and len(ex.starts) == 6


# ------------------------------------------------------------------------- Urdu
URDU_DOCS = {
    "words": ["\u067e\u06c1\u0644\u0627 \u0633\u0648\u0627\u0644: \u062a\u0648\u062d\u06cc\u062f \u06a9\u06cc\u0627 \u06c1\u06d2\u061f",
              "\u062f\u0648\u0633\u0631\u0627 \u0633\u0648\u0627\u0644: \u0634\u0631\u06a9 \u06a9\u06cc\u0627 \u06c1\u06d2\u061f",
              "\u062a\u06cc\u0633\u0631\u0627 \u0633\u0648\u0627\u0644: \u0646\u0628\u0648\u062a \u06a9\u06cc\u0627 \u06c1\u06d2\u061f",
              "\u0686\u0648\u062a\u06be\u0627 \u0633\u0648\u0627\u0644: \u0631\u0633\u0627\u0644\u062a \u06a9\u06cc\u0627 \u06c1\u06d2\u061f",
              "\u067e\u0627\u0646\u0686\u0648\u0627\u06ba \u0633\u0648\u0627\u0644: \u0639\u0628\u0627\u062f\u062a \u06a9\u06cc\u0627 \u06c1\u06d2\u061f"],
    "eastern digits and full stop": [f"{d}\u06d4 \u0633\u0628\u0642 \u0645\u06cc\u06ba \u062f\u06cc\u06d2 \u06af\u0626\u06d2 \u062a\u0635\u0648\u0631 \u06a9\u0627 \u062c\u0648\u0627\u0628 \u062f\u06cc\u06ba\u06d4"
                                     for d in "\u06f1\u06f2\u06f3\u06f4\u06f5"],
    "label and number": [f"\u0633\u0648\u0627\u0644 {d}: \u0627\u06cc\u06a9 \u0645\u062b\u0627\u0644 \u062f\u06cc\u06ba\u06d4" for d in "\u0661\u0662\u0663\u0664\u0665"],
}


def test_urdu_question_formats(client, tmp_path):
    right = 0
    for name, lines in URDU_DOCS.items():
        ex = run(client, doc_from(tmp_path, lines))
        print(f"  Urdu {name:30} -> {len(ex.starts)} question(s)")
        right += ex.usable and len(ex.starts) == 5
    assert right >= 2, "the model should get at least 2 of 3 Urdu formats exactly right"


# ------------------------------------------------------------------ hard cases
def test_options_hints_and_sample_answers_are_not_questions(client, tmp_path):
    lines = []
    for i in range(1, 5):
        lines += [f"Question {i}: Which is the capital of country {i}?", "A) one", "B) two", "C) three", "D) four",
                  "Hint: think about the map.", "Sample answer: B"]
    ex = run(client, doc_from(tmp_path, lines))
    assert ex.usable and len(ex.starts) == 4


def test_an_instruction_followed_by_separate_questions_counts_the_questions(client, tmp_path):
    lines = ["Question 1: Answer the following briefly.", "Who was the first caliph?", "What is zakat?", "Name the five pillars.",
             "When was the Hijrah?", "Question 2: Choose the correct option.", "Fasting is in which month?", "A) Rajab", "B) Ramadan",
             "C) Shaban", "Salah is performed how many times a day?", "A) 3", "B) 4", "C) 5"]
    ex = run(client, doc_from(tmp_path, lines))
    print(f"  nested: {len(ex.starts)} questions (6 separate questions expected)")
    assert abs(len(ex.starts) - 6) <= 1


def test_a_document_with_no_questions_does_not_invent_any(client, tmp_path):
    lines = ["Lesson summary", "This lesson introduced the idea of roots.", "Key points", "Roots undo powers.",
             "Teacher notes", "Review the examples before the next class."]
    ex = run(client, doc_from(tmp_path, lines))
    assert not ex.starts


def test_same_input_gives_the_same_answer(client, tmp_path):
    lines = [f"Q{i}. Describe topic {i}." for i in range(1, 7)]
    info = doc_from(tmp_path, lines)
    assert run(client, info).starts == run(client, info).starts


def test_the_models_proposals_survive_verification(client, tmp_path):
    lines = [f"{i}) Explain idea number {i} in your own words." for i in range(1, 9)]
    ex = run(client, doc_from(tmp_path, lines))
    assert ex.proposed and ex.rejected / ex.proposed <= 0.2


# -------------------------------------------------------- the whole pipeline
def test_end_to_end_review_with_the_real_model(client, tmp_path):
    from test_fallback import package_with_roman_pop_quiz
    from course_review import cli
    zip_path = package_with_roman_pop_quiz(tmp_path)
    model = fallback.ModelConfig(client=client, cache=fallback.Cache(tmp_path / "cache"), mode="suggest")
    pkg, res, xlsx, _ = cli.review(zip_path, checklist(tmp_path), str(tmp_path / "out"), model=model)
    pq1 = next(f for f in res.findings if f.code == "PQ1" and f.lesson is not None)
    print("  PQ1:", pq1.message)
    assert pq1.method == "model" and pq1.status == "needs_review"
    assert "2 question(s) recognised" in pq1.message


# ----------------------------------------------------- the real sample package
@pytest.mark.skipif(not os.environ.get("SAMPLE_ZIP"), reason="set SAMPLE_ZIP to the Grade 6 Islamiat zip")
def test_model_agrees_with_the_rules_on_the_sample_package(client):
    rows = evaluate.run(os.environ["SAMPLE_ZIP"], fallback.ModelConfig(client=client), limit=6)
    assert rows
    sil = [r for r in rows if r["precision"] is not None]
    mean_p = sum(r["precision"] for r in sil) / len(sil)
    mean_r = sum(r["recall"] for r in sil) / len(sil)
    assert mean_p >= 0.8 and mean_r >= 0.8, f"precision {mean_p:.2f}, recall {mean_r:.2f}"
