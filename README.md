# Course Review Tool

Reviews a course-package zip against the **Course Review Checklist** workbook and writes a
filled-in copy of that workbook, a JSON file and an HTML report. It runs fully offline.

The workbook is the authority. Rule codes, wording and scope are read from it at start-up,
and nothing is overridden. The Grade 6 Islamiat zip is only an example of a package.

## Use

```bash
pip install -r requirements.txt
python -m course_review.cli review PACKAGE.zip --checklist Course-Review-Checklist.xlsx --out review-output
```

Outputs in `review-output/`:

| File | What it is |
|---|---|
| `Course-Review-Checklist-filled.xlsx` | The reviewers' own workbook, one row per lesson / chapter. A **blank cell means the tool did not decide it**; the Notes column says why. A `Review Summary` sheet lists, for every code, whether the tool decides it. |
| `report.html` | Subject-level results with per-document evidence, per-lesson grids, package inventory. |
| `review.json` | Every finding with status, message, evidence and the document it came from. |

## What it decides

Run `python -m course_review.cli review ...` and open the `Review Summary` sheet, or read
`course_review/checks/registry.py`. In short:

* **Decided by code** (phase 1): page setup, fonts (including the Urdu complex-script font and
  size), colour, numerals, heading numbering and styles, captions, alt text, footers, file
  naming, Lesson Plan sections and SLO bullets, Pop Quiz / Chapter Exam / Worksheet / Data Bank
  counts and tags, yellow highlighting of correct answers.
* **Partly decided**: where the workbook rule has a part code cannot judge (for example "feedback
  meets the standard"), the result is a *partial pass*. It is written to the workbook as a blank
  cell with the reason in Notes, never as a Pass.
* **Not decided**: rules that need reading comprehension (SLO coverage, scope creep, story
  continuity, warm-up quality, Facilitator Guide notes) and rules that need a human (logo
  placement). These are planned for a later phase using a local model; until then they are left
  blank and listed as "not automated".

A Pass in the workbook means code verified the whole rule. A Fail always carries evidence.

## How files are matched to checklist sheets

By **scope**, as the workbook defines it, not by file name: the per-lesson `Assessment` files
feed the *Chapter Exam* sheet and the per-chapter `Chapter-Exam-Chapter-N` files feed the
*Worksheet* sheet (the report notes the name mismatch). Patterns live in
`course_review/profile.default.yaml`; pass `--profile your.yaml` to use another package's
conventions (other languages, other folder names, other section words).

## Layout

```
course_review/
  ingest.py       safe unzip (zip-slip, size and ratio limits), classify files, newest Data Bank wins
  docx_model.py   Word -> model with effective formatting (style chain, theme fonts, complex-script slots)
  pptx_model.py   PowerPoint -> model (fonts, explicit colours, line spacing, pictures)
  questions.py    question and Data Bank extraction (tables, Urdu digits and separators, guidance cut-off)
  workbook.py     rules in, filled workbook out (extends rows, keeps validation and formulas)
  checks/         formatting.py (WE), lessonplan.py (LP), assessments.py (PQ CE WS DB), registry.py
  engine.py       runs everything, rolls up to subject level, derives LP10 / FG8, builds the grids
  report.py       xlsx + json + html
tests/            48 tests on generated documents plus one end-to-end review
```

## Model fallback (optional)

The rule-based parser recognises the common question formats. For a document whose questions
it does not recognise (or whose numbering it cannot trust), a local model can propose where
the questions start. The model only **proposes**: every proposal is checked in code against
the document (the paragraph must exist and the quoted words must be in it), proposals that
fail are discarded, and an extraction with more than 30% rejected is thrown away. Answers are
cached per file and model, so a re-run gives the same result.

```bash
# the model is reached through your Ollama link (ngrok is fine); nothing else leaves the machine
set OLLAMA_URL=https://your-link.ngrok-free.dev        # PowerShell: $env:OLLAMA_URL = "..."
set OLLAMA_MODEL=qwen3:14b

python -m course_review.cli check-model
python -m course_review.cli review PACKAGE.zip --checklist X.xlsx --model-fallback suggest
```

* `suggest` (recommended): results that depend on the model are written as needs-review
  suggestions ("Would be pass: ..."), so a model answer is never a Pass or Fail in the workbook.
* `decide`: a Fail found with model help is written as a Fail; a Pass stays a partial pass
  (blank in the workbook). Everything is tagged `[model-assisted]`.
* With the model off (the default), behaviour is exactly as before.
* If the link is down the run stops with a clear message, rather than quietly reviewing without it.

Measure a model before relying on it:

```bash
python -m course_review.cli eval-model PACKAGE.zip --limit 8 [--labels hand-counts.json]
python -m pytest tests/test_live_model.py -v -s          # needs OLLAMA_URL; skipped otherwise
```

Small models are not good enough: in a quick check on one document, a 3B model found nothing,
while 8B and 14B models found all questions with none invented. They also took 1-2 minutes per
document on a CPU-only machine, so a GPU host matters.

## Known limits

* Phase 1 only. The AI-assisted checks are not built.
* Data Bank lessons are matched to package lessons per chapter by order; when the counts differ
  (the sample's Chapter 6 splits Lesson 2 in two) the chapter is reported as "needs reviewer".
* "Lower-order" for DB1 is not defined in the workbook; the profile treats Remember /
  Understand / Apply as lower-order. Change `lower_order_levels` if the team means otherwise.
* "High-resolution" (WE19) is not quantified in the workbook; the profile uses 150 dpi.
* Pop Quiz / Data Bank labels such as "Lesson Plan Location" and "Correct feedback" are English
  in the profile. Add the Urdu wording to `vocab` in the profile to avoid false Fails on Urdu
  packages.
* Colours inherited from a PowerPoint theme or master are not resolved, so a slide-colour Pass is
  reported as partial.
