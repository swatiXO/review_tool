# Course Review Tool

Reviews a course-package zip against the **Course Review Checklist** workbook and writes a
filled-in copy of that workbook, a JSON file and an HTML report. It runs on your own machine.
The only network call it can make is to an Ollama server you point it at (your ngrok link is fine).

The workbook is the authority. Rule codes, wording and scope are read from it at start-up, and
nothing is overridden. The Grade 6 Islamiat zip is only an example of a package.

## Quick start

The team's Course Review Checklist is built in (`course_review/data/Course-Review-Checklist.xlsx`). Pass `--checklist` (or set `COURSE_REVIEW_CHECKLIST`) only to use a different workbook; replace the built-in file when the checklist is updated.


```bash
pip install -r requirements.txt

# the web page: upload a zip, watch progress, download the results
python -m course_review.cli serve       # http://127.0.0.1:8080

# or the command line
python -m course_review.cli review PACKAGE.zip  --out review-output
```

| Output | What it is |
|---|---|
| `Marked-up-documents.zip` | **The package itself, fixed and marked up.** Rules with one right answer are applied to the copies: A4, 1-inch margins, 1.15 spacing and 8 pt after; Noto Nastaliq / Poppins; the guideline's sizes for the grade; Western digits; text direction; Heading styles and decimal numbers; table captions; bullets instead of numbered lists; no colour (yellow correct answers kept, check marks turned into yellow); footer with name, version, date and page; a Table of Contents in long Lesson Plans; the file renamed to the guideline pattern, keeping the document's own version (from the file name, footer or file properties; a version is never made up, so a file without one keeps its WE3 comment). On slides: 1.5 spacing, fonts, digits, colour and a footer. The first comment in each file lists what was fixed. **What needs a person stays highlighted with a comment**: each fixed copy is checked again, and a Fail is left out only if the copy now passes; needs-reviewer results are always kept. Every Word file has its problems highlighted in the text: red = fails the checklist, turquoise = a model suggestion for a reviewer to decide, green = passes. Each highlight carries a Word comment naming the checklist code and the reason, and a comment on the first paragraph lists the whole-document results (fonts, page size, file name). Existing highlights are never replaced, so yellow correct answers stay yellow. Slides get the highlight plus a red-bordered "Course Review notes" box. Results that belong to no single file (a missing Data Bank, subject-wide coverage) are in `REVIEW-NOTES.txt`. |
| `Course-Review-Checklist-filled.xlsx` | Your own workbook, one row per lesson / chapter. **A blank cell means the tool did not decide it**; the Notes column says why. Extra sheets: `Review Summary` (what the tool decides, per code) and `SLO Coverage`. |
| `report.html` | Subject-level results with per-document evidence, per-lesson grids, SLO coverage, package inventory. |
| `review.json` | Every finding with status, message, evidence, the document it came from, and whether a model helped. |

## What decides what

Open the `Review Summary` sheet, or read `course_review/checks/registry.py`. Each code is one of:

* **Yes** (decided by code): page setup, fonts (incl. the Urdu complex-script font and size), colour,
  numerals, heading numbering and styles, captions, alt text, footers, file naming, citations,
  Lesson Plan sections and SLO bullets, question counts and tags, yellow highlight on correct
  answers, SLO coverage (ST1, ST2).
* **Partly**: the part code can judge is judged; the rest is stated in the Notes. A partial pass is
  written as a blank cell, never as Pass.
* **Suggestion**: rules that need reading comprehension (LP1, LP2, LP6, LP7, LP8, LP9, FG1, FG3,
  FG6, CE2, CE3, and SLO mapping for PQ5/WS5). Only with `--model-checks`. Always shown as
  needs-review, never as Pass or Fail.
* **No**: left to a person (logo placement, FG4's undefined "five items", FG5 video cues, WE19
  screenshots, WE24, ...).

## Beyond the checklist: each output's own guidelines

The team reviews every output against the checklist, that output's guideline document and template, and the
Writing & Editing Guidelines. `checks/guidelines.py` adds the guideline rules code can check. They appear in the
marked-up documents and the report (not in the workbook, which only has checklist codes):

| Code | Output | Rule (source) |
|---|---|---|
| LPG1-LPG4 | Lesson Plan | Warm-up 1-3 questions; Key Takeaways in bullets; Book and SLO Coverage table; Glossary Terms (Lesson Plan Guidelines, template) |
| FGG1-FGG3 | Facilitator's Guide | Notes in the speaker-notes pane; a time per slide adding up to the planned duration; duration on the Cover |
| FG1, FG4 | Facilitator's Guide | Decided by code where the slides make it clear: SLOs pasted from the Lesson Plan; Concept Building slides without notes |
| ASG1 | Pop Quiz, Assessment, Chapter Exam, Data Bank | No "all / none of the above" |
| SB1-SB5 | Storyboard | Template columns; 2:00 cap; narration about 260-300 words; "New" / "Edited from"; short on-screen text |
| WEG1-WEG2 | Word documents | 8 pt after each paragraph; bullets (not numbering) for lists inside a section |

Decisions taken where the documents disagree (see `house_rules` in the profile): Urdu font is Noto Nastaliq; font
sizes follow the Writing & Editing Guidelines' table for the package's grade; the per-lesson Assessment is 6-8
questions at 70/30 and the per-chapter Chapter Exam 8-10 at 40/60. The workbook's sheet names are swapped against
the team's names for those two, so comments in the marked-up documents say "CE1 (Assessment)" and "WS1 (Chapter Exam)".

## How files are matched to checklist sheets

By **scope**, as the workbook defines it, not by file name: per-lesson `Assessment` files feed the
*Chapter Exam* sheet; per-chapter `Chapter-Exam-Chapter-N` files feed the *Worksheet* sheet (the
report notes the name mismatch). Patterns live in `course_review/profile.default.yaml`. Pass
`--profile your.yaml` for another package's conventions; the file is laid over the built-in
profile, so it only needs the settings it changes.

## Using a language model (optional)

```bash
set OLLAMA_URL=https://your-link.ngrok-free.dev        # PowerShell: $env:OLLAMA_URL = "..."
set OLLAMA_MODEL=qwen3.5:9b
python -m course_review.cli check-model

# 1. question formats the rules do not recognise
python -m course_review.cli review PKG.zip --checklist X.xlsx --model-fallback suggest
# 2. content checks (slow: about 10 model calls per lesson); optionally choose codes
python -m course_review.cli review PKG.zip --checklist X.xlsx --model-checks LP1,LP6,FG1
# 3. let the checks consult the textbook
python -m course_review.cli review PKG.zip --checklist X.xlsx --model-checks --book-index book_index
```

The rule that never changes: **the model proposes, code verifies, and a model answer is never a Pass
or Fail in the workbook.** A proposal counts only if its quoted words appear, word for word, in the
material it was given; otherwise it is thrown away. With the model off (the default) nothing
changes. If the link is down the run stops with a clear message instead of quietly skipping it.
Answers are cached per file and model, so a re-run gives the same result.

`--model-fallback decide` additionally writes model-based **Fails** (never Passes) to the workbook.

### Learn a new question format once

```bash
python -m course_review.cli learn-format odd-exam.docx --save my-profile.yaml    # model proposes a pattern
python -m course_review.cli learn-format odd-exam.docx --save my-profile.yaml --pattern "Task\s*(\d+)\s*>>" --yes
python -m course_review.cli review PKG.zip --checklist X.xlsx --profile my-profile.yaml
```

The pattern is validated against the document, shown to you, and saved only after you approve.
From then on that format is parsed by plain code, with no model.

### The textbook

The book is usually a scan, so it has to be read by OCR once:

```bash
python -m course_review.cli index-book book.pdf --out book_index --engine tesseract --pages 1-20
python -m course_review.cli index-book book.pdf --out book_index --engine vlm --model qwen3.5:9b   # vision model on your Ollama host
```

Pages are cached, so an interrupted run resumes. **Check the OCR yourself before relying on it**:
on the sample book (Urdu Nastaliq, noisy scan with a watermark) Tesseract's Urdu model produced
unusable text, and a vision model running on this CPU-only PC was far too slow. A GPU host is needed
for whole-book OCR.

### Measure a model before relying on it

```bash
python -m course_review.cli eval-model PACKAGE.zip --limit 8 [--labels hand-counts.json]
python -m pytest tests/test_live_model.py -v -s        # needs OLLAMA_URL; skipped otherwise
```

## Layout

```
course_review/
  ingest.py       safe unzip (zip-slip, size and ratio limits), classify files, newest Data Bank wins
  docx_model.py   Word -> model with effective formatting (style chain, theme fonts, complex-script slots)
  pptx_model.py   PowerPoint -> model (fonts, colours, spacing, pictures, slide titles and text)
  questions.py    question and Data Bank extraction (tables, Urdu digits, roman numerals, learned formats)
  slo.py          SLOs from the Document of Specifications; SLO tags in questions
  workbook.py     rules in, filled workbook out (extends rows, keeps validation and formulas)
  checks/         formatting (WE), lessonplan (LP), slides (FG), assessments (PQ CE WS DB), coverage (ST PQ5 WS4 WS5),
                  judgement (model-assisted), registry
  engine.py       runs everything, rolls up to subject level, derives LP10 / FG8, builds the grids
  fallback.py llm.py judge.py formats.py book.py evaluate.py     the model-assisted parts
  autofix.py      applies the mechanical Writing & Editing rules to the copies
  annotate.py     highlights and comments for what still needs a person
  report.py web.py cli.py
tests/            offline tests on generated documents and stand-in models; live tests skip without OLLAMA_URL
```

## Known limits

* Data Bank lessons are matched to package lessons per chapter by order; when the counts differ (the
  sample's Chapter 6 splits Lesson 2 in two) the chapter is reported as "needs reviewer".
* SLOs in the sample specification have no codes, so questions can be tied to them only by an explicit
  marker ("SLO 2", "L3 SLO 2") or by quoting the SLO sentence. With no tags, coverage rules say
  "needs reviewer"; ST2 stays literal (no tags = every SLO is a gap), as the workbook words it.
* "Lower-order" (DB1) and "high-resolution" (WE19) are not defined in the workbook; the profile uses
  Remember/Understand/Apply and 150 dpi. Change them in `profile.default.yaml`.
* Pop Quiz / Data Bank labels such as "Lesson Plan Location" and "Correct feedback" are English in the
  profile. Add the Urdu wording to `vocab` to avoid false Fails on Urdu packages.
* Colours inherited from a PowerPoint theme or master are not resolved, so a slide-colour Pass is partial.
* The web page has no login. Keep it on 127.0.0.1 (the default) unless you add one.
