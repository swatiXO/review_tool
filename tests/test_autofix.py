"""The fixed copies: reviewing them again, the mechanical Writing & Editing rules pass."""
import json
import zipfile

from docx.enum.text import WD_COLOR_INDEX

from course_review import cli

from helpers import A4, add, checklist, new_doc, zip_dir

URDU = "صبر کا مطلب ہے رکنا اور برداشت کرنا۔"


def package(tmp_path):
    root = tmp_path / "Grade-6-Islamiat"
    lesson = root / "Chapter-4-Akhlaq" / "Lesson-2-Sabr"
    lesson.mkdir(parents=True)
    d = new_doc(page=(8.5, 11), margin=0.75, spacing=1.0, size=11)
    add(d, "سبق: صبر و تحمل", bold=True, rtl=True)
    for s in ("تعارف", "حاصلات تعلیم", "تحریکی سرگرمی", "تصوراتی تعمیر", "حاصل کلام"):
        add(d, s, bold=True, size=14, rtl=True)
        add(d, URDU + " آیت ١٥٣", rtl=True, color="2E74B5")
    d.add_paragraph("پہلا نکتہ", style="List Number")
    d.add_paragraph("دوسرا نکتہ", style="List Number")
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text, t.cell(1, 0).text = "عنوان", "متن"
    d.save(str(lesson / "Lesson-2-Chapter-4-Lesson-Plan.docx"))
    pq = new_doc()
    add(pq, "Pop Quiz")
    add(pq, "سبق 2: صبر و تحمل", bold=True)
    add(pq, "سوال 1: صبر کیا ہے؟")
    add(pq, "✅ الف) رکنا")
    add(pq, "ب) بھاگنا")
    (root / "Pop Quiz").mkdir()
    pq.save(str(root / "Pop Quiz" / "Pop Quiz.docx"))
    return zip_dir(str(root), str(tmp_path / "pkg.zip"))


def statuses(out, name):
    r = json.load(open(out / "review.json", encoding="utf-8"))
    return {f["code"]: f["status"] for f in r["findings"] if (f["doc"] or "").endswith(name)}


def test_mechanical_rules_pass_after_fixing(tmp_path):
    out1, out2 = tmp_path / "o1", tmp_path / "o2"
    cl = str(checklist(tmp_path))
    cli.review(package(tmp_path), cl, str(out1))
    before = statuses(out1, "Lesson-2-Chapter-4-Lesson-Plan.docx")
    for code in ("WE1", "WE4", "WE8", "WE10", "WE13", "WE16", "WE23", "WE25", "WEG1", "WEG2"):
        assert before[code] == "fail", code
    fixed = out1 / cli.MARKED_UP
    names = zipfile.ZipFile(fixed).namelist()
    assert any(n.endswith("Lesson-Plan-Lesson-2-Chapter-4-v0.1.docx") for n in names)
    cli.review(str(fixed), cl, str(out2))
    after = statuses(out2, "Lesson-Plan-Lesson-2-Chapter-4-v0.1.docx")
    for code in ("WE1", "WE4", "WE7", "WE8", "WE10", "WE11", "WE12", "WE13", "WE16", "WE25", "WEG1", "WEG2"):
        assert after[code] == "pass", (code, after[code])
    assert after["WE23"] == "pass" or "highlight" in json.dumps(json.load(open(out2 / "review.json", encoding="utf-8")))


def test_check_mark_becomes_a_yellow_highlight_and_lesson_sections_still_split(tmp_path):
    out1, out2 = tmp_path / "o1", tmp_path / "o2"
    cl = str(checklist(tmp_path))
    cli.review(package(tmp_path), cl, str(out1))
    cli.review(str(out1 / cli.MARKED_UP), cl, str(out2))
    r = json.load(open(out2 / "review.json", encoding="utf-8"))
    pq = {f["code"]: f for f in r["findings"] if f["lesson"] == "Ch4-L2" and f["code"].startswith("PQ")}
    assert pq["PQ1"]["status"] in ("pass", "fail") and "Pop Quiz has" not in pq["PQ1"]["message"]
    assert pq["PQ4"]["status"] == "pass"


def test_quran_hadith_and_dua_text_is_never_edited(tmp_path):
    from docx import Document
    root = tmp_path / "Grade-6-Islamiat"
    ch = root / "Chapter-1-Quran"
    ch.mkdir(parents=True)
    d = new_doc(size=11)
    add(d, "باب اول: قرآن مجید", bold=True, rtl=True)
    add(d, "سورۃ العصر", bold=True, rtl=True)
    verse = "وَالْعَصْرِ ﴿١﴾ إِنَّ الْإِنسَانَ لَفِي خُسْرٍ ﴿٢﴾"
    add(d, "بِسْمِ اللَّهِ الرَّحْمَٰنِ الرَّحِيمِ", bold=True, rtl=True, urdu_font="Traditional Arabic")
    add(d, verse, rtl=True, urdu_font="Traditional Arabic")
    add(d, "ترجمہ: زمانے کی قسم ﴿١﴾ آیت ٢ میں", rtl=True)
    d.save(str(ch / "Chapter-1-Lesson-Plan.docx"))
    out = tmp_path / "out"
    cli.review(zip_dir(str(root), str(tmp_path / "p.zip")), str(checklist(tmp_path)), str(out))
    z = zipfile.ZipFile(out / cli.MARKED_UP)
    name = next(n for n in z.namelist() if n.endswith(".docx"))
    (tmp_path / "f.docx").write_bytes(z.read(name))
    texts = [p.text for p in Document(str(tmp_path / "f.docx")).paragraphs]
    assert "بِسْمِ اللَّهِ الرَّحْمَٰنِ الرَّحِيمِ" in texts            # not numbered as a heading
    assert verse in texts                                       # ayah numbers kept as written
    assert "ترجمہ: زمانے کی قسم ﴿١﴾ آیت 2 میں" in texts           # Urdu digit fixed, the verse mark kept
    fonts = {r.font.element.rPr.rFonts.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs")
             for p in Document(str(tmp_path / "f.docx")).paragraphs if p.text == verse for r in p.runs}
    assert fonts == {"Traditional Arabic"}
