"""Which checklist codes the tool decides, and why the others are left to a person.

The wording of every rule comes from the workbook; this only records the automation
status so the report can state it. A code missing from here is reported as
'not automated' rather than skipped.
"""

AI = "Needs reading comprehension of the content (planned: local-model assisted checks)"
HUMAN = "Needs a reviewer; not reliably decidable by code"
PHASE2 = "Needs SLO extraction and question tagging across documents (planned)"
SUGGEST = "Model-assisted suggestion (run with --model-checks and a local model). Shown as needs-review, never as Pass or Fail"
BOOK = " Consults the textbook when --book-index is given."

# code -> (level, how)   level: 'yes' | 'partly' | 'no'
AUTOMATION = {
    "LP1": ("model", SUGGEST), "LP2": ("model", SUGGEST + BOOK),
    "LP3": ("yes", "Section headings located by profile vocabulary and checked for presence and order"),
    "LP4": ("yes", "Heading numbers read (typed or automatic) and checked for levels, order and skipped levels"),
    "LP5": ("yes", "Items under the SLO heading checked for bullet list formatting"),
    "LP6": ("model", SUGGEST), "LP7": ("model", "The model names the story's characters from the Introduction (checked against it); code looks for them in every Concept Building sub-topic. Shown as needs-review"), "LP8": ("model", SUGGEST + " Needs --book-index, otherwise left to a reviewer."), "LP9": ("model", SUGGEST),
    "LP10": ("yes", "Derived from WE1-WE4 and WE6-WE8 on this lesson's Lesson Plan file"),
    "FG1": ("partly", "Fail when the Session Overview pastes the Lesson Plan's SLO wording; otherwise a model suggestion"), "FG2": ("partly", "Facilitator-notes part on each slide and no text copied from the Lesson Plan; the split itself needs a reviewer"), "FG3": ("model", SUGGEST), "FG4": ("partly", "Concept Building slides without Facilitator Notes fail; the five items (Facilitator's Guide Guidelines) are looked for by wording"),
    "FG5": ("no", HUMAN), "FG6": ("model", SUGGEST), "FG7": ("partly", "Slide titles matched to the Lesson Plan sections; content under each needs a reviewer"),
    "FG8": ("yes", "Derived from WE5-WE7 and WE23 on this lesson's Facilitator Guide file"),
    "PQ1": ("yes", "Questions counted per lesson section of the Pop Quiz; option letters used to detect MCQ"),
    "PQ2": ("partly", "Presence of a Lesson Plan Location line; precision needs a reviewer"),
    "PQ3": ("partly", "Presence of feedback lines by code; with model checks, the model reads each message against the standard (a suggestion)"),
    "PQ4": ("yes", "Yellow run highlight or yellow shading on a correct option; a check-mark symbol does not count"),
    "PQ5": ("partly", "SLOs from the Document of Specifications vs the SLO tags in the questions; untagged questions need a reviewer"),
    "CE1": ("partly", "Question count, and the 70/30 split when the questions state their level; otherwise, with model checks, the model's Bloom's levels (a suggestion)"),
    "CE2": ("model", SUGGEST + " Needs --book-index."), "CE3": ("model", SUGGEST + " Needs --book-index."), "CE4": ("yes", "SLO tag present in every question block"),
    "CE5": ("partly", "Answers present (a key or under each question); completeness needs a reviewer"),
    "DB1": ("yes", "Items per lesson, MCQ-only and lower-order-only, read from the Data Bank tables"),
    "DB2": ("yes", "Presence of Subject, Chapter, Lesson and SLO fields on every item"),
    "DB3": ("partly", "Presence of feedback fields by code; with model checks, the model reads each message against the standard (a suggestion)"),
    "DB4": ("yes", "Yellow highlight on the correct answer"),
    "DB5": ("partly", "Near-duplicate wording against the Pop Quiz; reworded duplicates need a reviewer"),
    "WS1": ("partly", "Question count, and the 40/60 split when the questions state their level; otherwise, with model checks, the model's Bloom's levels (a suggestion)"),
    "WS2": ("no", "Formats are listed for the reviewer; the workbook does not say how many formats make a mix"),
    "WS3": ("yes", "SLO tag present in every question block"),
    "WS4": ("partly", "Questions per lesson from SLO / lesson tags; untagged questions need a reviewer"),
    "WS5": ("partly", "SLOs from the Document of Specifications vs the SLO tags in the Worksheet; untagged questions need a reviewer"),
    "WS6": ("partly", "Answer Key section found at the end; completeness needs a reviewer"),
    "ST1": ("yes", "Coverage map built from the Document of Specifications and the tags in the four assessment types"),
    "ST2": ("yes", "SLOs with no tagged question in any assessment, exactly as the workbook words it (explicit tags only)"),
    "DBS1": ("partly", "Sort order by chapter and lesson; subject tab needs a reviewer"),
    "WE1": ("partly", "Shape [Type]-[Identifier]-[Topic]-v[Version] checked; the meaning of each part is not"),
    "WE2": ("yes", "File name searched for 'Draft' / 'Final'"),
    "WE3": ("partly", "Version presence is checked; whether it matches the document's stage needs a reviewer"),
    "WE4": ("yes", "Page size, orientation, margins and effective line spacing"),
    "WE5": ("yes", "Slide size and line spacing"),
    "WE6": ("yes", "Effective font of every English text run"),
    "WE7": ("yes", "Effective complex-script font of every Urdu text run (Noto Nastaliq, per the team's decision)"),
    "WE8": ("yes", "Effective sizes and weights by paragraph role, using the Writing & Editing Guidelines' sizes for the package's grade"),
    "WE9": ("partly", "Only decidable when both language versions are in the package"),
    "WE10": ("yes", "All digits scanned for non-Western numerals"),
    "WE11": ("partly", "Paragraph direction vs script; mid-sentence switching is not checked"),
    "WE12": ("yes", "Decimal numbering of headings"),
    "WE13": ("yes", "Heading-like paragraphs checked for built-in Heading styles"),
    "WE14": ("yes", "Same rule as LP3, rolled up over all Lesson Plans"),
    "WE15": ("yes", "Table of Contents versus the page count stored in the file"),
    "WE16": ("yes", "Caption above each table, numbered"),
    "WE17": ("yes", "Caption below each figure, numbered, 11pt italic"),
    "WE18": ("yes", "Image pixels analysed for colour"),
    "WE19": ("partly", "Resolution checked; screenshots with browser chrome need a reviewer"),
    "WE20": ("yes", "Alt text on every figure"),
    "WE21": ("yes", "Latin-script citations checked for (Author, Year); Quran and Hadith references are not treated as citations"),
    "WE22": ("yes", "Reference list present when citing, alphabetical, hanging indent"),
    "WE24": ("no", HUMAN),
    "WE23": ("partly", "Run colours, shading and style colours; colours inherited from a theme are not resolved for slides"),
    "WE25": ("partly", "Footer (or slide footer placeholders) has page number, version, date and a name; the name is not compared with the real one"),
}
