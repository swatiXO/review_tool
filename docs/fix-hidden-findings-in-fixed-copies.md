# Fix: the fixed copies hid findings the fix did not settle

Date: 2026-10-06 · Branch: `claude/autofix-keep-unfixed-findings` (from `master` at `0c528cd`)

## The problem

Since `0a2d3a2` ("Fix formatting in the documents instead of commenting on it"), the tool corrects
formatting in the copies inside `Marked-up-documents.zip` and leaves comments only for what still
needs a person. To decide which comments to leave out, `autofix.settled()` used a fixed list of
about 20 codes (WE1-WE4, WE6-WE8, WE10-WE13, WE16, WE23, WE25, WEG1, WEG2, LP4, LP5, LP10 for Word;
a shorter list for slides). Any finding with one of those codes was removed from the document,
whatever it said. That had three effects:

1. **"Needs reviewer" results disappeared.** A needs-review result is not something a formatting
   fix can settle, but it was hidden all the same. Examples:
   * WE3: "Version v0.1 is present; whether it matches the document's stage needs a reviewer".
     The tool writes `v0.1` into the file name when there is none, and then hid the one comment
     asking a person to confirm it.
   * LP5: "The SLO section was not found, so the SLO list could not be checked".
2. **Fails were hidden even when the fix did not run.** When `fix_docx` cannot map a document's
   paragraphs safely it returns without changing anything, but the filter still ran, so the
   document's page size, font, spacing and heading Fails vanished from an unchanged file.
3. **Fails were hidden when the fix covered only part of the rule.** For example, WE23 also fails
   on colours set in styles other than the Heading styles, which the fixer does not change; WE5 also covers slide size, which
   the fixer does not change.

The workbook and report were not affected; this was only about the comments in the marked-up
documents.

## The fix

The fixed copy is checked again, and a Fail is left out only if that copy no longer has it.

* `autofix.recheck(ctx, doc, path, rel)` parses the copy that will go into the zip (after fixing,
  renaming and adding the footer) and runs the same per-document checks on it: the formatting
  checks, the guideline checks and, for Lesson Plans, LP3-LP5. It returns `{}` if the copy cannot
  be read, and then nothing counts as fixed.
* `autofix.remaining(findings, ext, after, extra)` replaces `settled()`:
  * a finding that is not a Fail is always kept;
  * a Fail on a code the fixer works on is left out only if the copy now passes (or the rule no
    longer applies). If the copy still fails, or now needs a reviewer, the copy's own result is
    shown instead, so the comment describes the file the reader actually opens;
  * LP10 and FG8, which are derived from other codes, are left out only if none of their
    component codes fail on the copy;
  * PQ4 / DB4 keep their earlier rule (left out only when check marks were really turned into
    yellow highlights), because they are checked across the package, not per file;
  * every other Fail is kept.
* `annotate.annotate_package` calls `recheck` and `remaining` in place of `settled`.

## Before and after

On the test package in `tests/test_autofix.py`, the fixed Lesson Plan's summary comment:

| | Before | After |
|---|---|---|
| WE3 "v0.1 is present; does it match the stage?" | hidden | **shown** as CHECK |
| WE4, WE8 and the other formatting Fails the fix cleared | hidden | hidden (the copy passes) |
| LPG1, LPG3, LPG4 | shown | shown |

## Tests

Added to `tests/test_autofix.py`:

* `test_only_a_fail_the_fixed_copy_no_longer_shows_is_left_out`: needs-review kept; cleared Fail
  left out; Fail turned needs-review shown with the copy's message; Fail still failing kept;
  untouched code kept; LP10 kept while a component still fails; nothing cleared when the copy
  could not be re-checked.
* `test_fixed_copy_still_asks_about_what_the_fix_cannot_decide`: end to end, the fixed Lesson
  Plan carries the WE3 CHECK comment and no WE4 / WE8 FAIL.

Both fail on the previous code and pass with the fix. Full suite: 199 passed, 11 skipped
(the skipped tests need a live model server).

## Cost

Each fixed document is parsed and checked once more. That is the same work the first pass does
per document, and is small next to the fixing itself; no model is called.

## Not changed here

From the same review of `0a2d3a2`, still open:

* the footer fixer deletes any existing footer (logos, template text) and writes the file name as
  the "name";
* ~~a file without a version is renamed to `v0.1`~~: fixed next, see
  [fix-version-and-web-page.md](fix-version-and-web-page.md).
