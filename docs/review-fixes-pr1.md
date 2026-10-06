# Fixes from the review of pull request #1

Date: 2026-10-06 · Branch: `claude/autofix-keep-unfixed-findings`

The review found nine problems. Each one was reproduced first, then fixed, and each has a test
that fails on the code before the fix and passes after it.

## Problems that could make the fixed documents wrong

### 1. A version was invented from the footer
`Class V 2024` in a footer was read as version 2024 (the pattern allowed a space between "v" and
the number, and any number of digits).

**Fix:** a version is now either "v" directly followed by the number (`v2`, `v1.0`) or the word
Version / ver. / ورژن before it. The first number has at most two digits, so a year is never a
version. `Class V 2024`, `Class V 2` and `v2024` are not versions; `Version 2.1`, `ورژن ٣` and
`v1.0` are.

### 2. Versions in the middle of a file name were lost
Only a version at the very end of the name was recognised, so `Lesson-Plan-v2-Final.docx` and
`LessonPlan v1.0 (1).docx` lost theirs when renamed. (Before this pull request they were renamed
to `v0.1`, so the real version was lost then too.)

**Fix:** the version is found anywhere in the name (the last one, if there are several).

### 3. Out-of-date "needs reviewer" comments stayed in the fixed copy
Only Fails were re-checked. A needs-reviewer item the fix settles, such as "no page count, cannot
decide whether a Table of Contents is needed" after the fixer added one, kept its comment.

**Fix:** Fails and needs-reviewer results on the rules the fixer works on are both re-checked.
Each is left out if the fixed copy passes, and otherwise replaced by what the copy shows.
Needs-reviewer results on any other rule (content, SLOs, model suggestions) are never touched.

### 4. LP10 and FG8 were worked out wrongly
* FG8 on a Word Facilitator Guide could never clear: its list of parts included WE5, a
  slides-only check that never runs on Word files.
* LP10 was dropped whenever none of its parts failed, even when the engine would say "needs
  review" (for example WE3, whose version a person must confirm).

**Fix:** the engine's rule moved into `engine.derived_finding()`, and both the engine and the
fixer use it. It uses only the parts that actually ran on the document, and gives Pass, Fail or
needs-reviewer exactly as the review does.

## Problems that confused users

### 5. The web page and the download disagreed
The lists on the result page came from the review of the files as submitted, under their old
names, so they showed problems the downloaded copies no longer had.

**Fix:** `annotate_package` now records what it actually left open in each copy, under the copy's
new name (`stats["open"]`). The lists are built from that, and are titled "Still open in the
fixed copies" and "By fixed document". The count tiles still describe the package as submitted,
like the workbook, and say so.

### 6. Upload errors lost their red style
The new page styling had dropped the `bad` class used by error messages.

**Fix:** the `bad`, `ok` and `warn` styles are back; an error message is shown in a red box.

## Code quality

### 7. A copied list
The LP10 / FG8 parts list in `autofix.py` was a copy of the engine's. It is gone; the fixer uses
`engine.DERIVED` and `engine.derived_finding()`.

### 8. Each document was opened twice to read its version property
The Word and PowerPoint parsers now read the Version property while they have the file open
(`core_version` on the parsed document), and `find_version` uses that.

### 9. The checklist was loaded a second time, and errors were hidden
The review now keeps the rules it already loaded (`res.rules`), and the web page uses them. If the
lists cannot be made, the error is logged and the page says so (the downloads are unaffected),
instead of quietly showing nothing.

## Tests

* `tests/test_autofix.py`
  * `test_version_patterns_do_not_invent_or_lose_versions` (1, 2)
  * `test_review_items_and_summary_items_follow_the_fixed_copy` (3, 4)
  * `test_only_a_fail_the_fixed_copy_no_longer_shows_is_left_out`: LP10 now describes the copy
* `tests/test_web.py`
  * `test_lists_describe_the_fixed_copies_not_the_submitted_files` (5): the page lists only file
    names found in the download, and not rules that failed as submitted but are fixed in the
    copies (WE8, WEG1)
  * `test_upload_errors_are_shown_in_red` (6)

Full suite: 209 passed, 11 skipped (the skipped tests need a live model server). Run against the
code before these fixes, the six new or changed tests fail.

The branch also merges the latest `master` (`ebc6b0f`, turquoise notes); there were no conflicts.
