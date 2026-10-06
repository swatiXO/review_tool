"""A package with one lesson (or a few chapters) still finds that lesson's Pop Quiz questions, by
matching the quiz's lesson titles to the lesson's own documents."""
import json

from course_review import cli

from helpers import add, checklist, new_doc, zip_dir
from test_fallback import package_with_roman_pop_quiz


def test_one_lesson_package_finds_its_pop_quiz_section(tmp_path):
    root = tmp_path / "Grade-6-Islamiat"
    lesson = root / "Chapter-4-Akhlaq" / "Lesson-2-Sabr"
    lesson.mkdir(parents=True)
    lp = new_doc()
    add(lp, "سبق: صبر و تحمل | باب چہارم: اخلاق و آداب", bold=True)
    for s in ("تعارف", "حاصلات تعلیم", "تحریکی سرگرمی", "تصوراتی تعمیر", "حاصل کلام"):
        add(lp, s, bold=True)
        add(lp, "کچھ متن یہاں ہے۔")
    lp.save(str(lesson / "Lesson-2-Chapter-4-Lesson-Plan.docx"))
    pq = new_doc()
    add(pq, "Pop Quiz")
    groups = [["توحید کی اہمیت", "نبوت اور رسالت"], ["نبی کریم", "ریاست مدینہ"], ["مشاورت", "صبر و تحمل"]]
    for g, titles in enumerate(groups):
        for n, t in enumerate(titles, 1):
            add(pq, f"سبق {n}: {t}", bold=True)
            for k in (1, 2):
                add(pq, f"سوال {k}: {t} سوال {g}{n}{k}؟")
                for opt in ("الف) ایک", "ب) دو", "ج) تین"):
                    add(pq, opt)
    (root / "Pop Quiz").mkdir()
    pq.save(str(root / "Pop Quiz" / "Pop Quiz.docx"))
    z = zip_dir(str(root), str(tmp_path / "one.zip"))
    out = tmp_path / "out"
    cli.review(z, str(checklist(tmp_path)), str(out))
    found = {f["code"]: f for f in json.load(open(out / "review.json", encoding="utf-8"))["findings"] if f["lesson"] == "Ch4-L2"}
    assert found["PQ1"]["status"] == "pass" and "2 MCQ" in found["PQ1"]["message"]


def test_loose_lesson_files_without_folders_keep_their_lesson_and_no_chapter_exam_fail(tmp_path):
    root = tmp_path / "upload"
    root.mkdir()
    lp = new_doc()
    add(lp, "Lesson Plan", bold=True)
    for s in ("Introduction", "SLOs", "Warm-up", "Concept Building", "Key Takeaways"):
        add(lp, s, bold=True)
        add(lp, "Some text here.")
    lp.save(str(root / "Lesson-2-Chapter-4-Lesson-Plan.docx"))
    z = zip_dir(str(root), str(tmp_path / "loose.zip"))
    out = tmp_path / "out"
    cli.review(z, str(checklist(tmp_path)), str(out))
    r = json.load(open(out / "review.json", encoding="utf-8"))
    assert [d["lesson"] for d in r["inventory"]["documents"]] == ["2"]
    ws1 = [f for f in r["findings"] if f["code"] == "WS1"]
    assert ws1 and all(f["status"] == "needs_review" for f in ws1)


def test_the_teams_checklist_is_built_in(tmp_path):
    import os
    assert os.path.exists(cli.BUILTIN_CHECKLIST)
    out = tmp_path / "out"
    cli.review(str(package_with_roman_pop_quiz(tmp_path)), cli.default_checklist(), str(out))
    assert (out / "Course-Review-Checklist-filled.xlsx").exists()
