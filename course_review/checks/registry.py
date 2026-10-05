"""Which checklist codes the tool decides, and why the others are left to a person.

The wording of every rule comes from the workbook; this only records the automation
status so the report can state it. A code missing from here is reported as
'not automated' rather than skipped.
"""

AI = "Needs reading comprehension of the content (planned: local-model assisted checks)"
HUMAN = "Needs a reviewer; not reliably decidable by code"
PHASE2 = "Needs SLO extraction and question tagging across documents (planned)"

# code -> (level, how)   level: 'yes' | 'partly' | 'no'
AUTOMATION = {
    "LP1": ("no", AI), "LP2": ("no", AI + "; also needs the textbook"),
    "LP3": ("yes", "Section headings located by profile vocabulary and checked for presence and order"),
    "LP4": ("yes", "Heading numbers read (typed or automatic) and checked for levels, order and skipped levels"),
    "LP5": ("yes", "Items under the SLO heading checked for bullet list formatting"),
    "LP6": ("no", AI), "LP7": ("no", AI), "LP8": ("no", HUMAN + " (needs the textbook)"), "LP9": ("no", AI),
    "LP10": ("yes", "Derived from WE1-WE4 and WE6-WE8 on this lesson's Lesson Plan file"),
    "FG1": ("no", AI), "FG2": ("no", AI), "FG3": ("no", AI), "FG4": ("no", AI + "; the five items are not defined in the workbook"),
    "FG5": ("no", HUMAN), "FG6": ("no", AI), "FG7": ("no", AI),
    "FG8": ("yes", "Derived from WE5-WE7 and WE23 on this lesson's Facilitator Guide file"),
    "PQ1": ("yes", "Questions counted per lesson section of the Pop Quiz; option letters used to detect MCQ"),
    "PQ2": ("partly", "Presence of a Lesson Plan Location line; precision needs a reviewer"),
    "PQ3": ("partly", "Presence of feedback lines; quality needs a reviewer"),
    "PQ4": ("yes", "Yellow run highlight or yellow shading on a correct option; a check-mark symbol does not count"),
    "PQ5": ("no", PHASE2),
    "CE1": ("partly", "Question count; the 70/30 split needs question levels"),
    "CE2": ("no", AI + "; needs the textbook"), "CE3": ("no", AI), "CE4": ("yes", "SLO tag present in every question block"),
    "CE5": ("no", "Whether an answer key is expected is not decidable by code"),
    "DB1": ("yes", "Items per lesson, MCQ-only and lower-order-only, read from the Data Bank tables"),
    "DB2": ("yes", "Presence of Subject, Chapter, Lesson and SLO fields on every item"),
    "DB3": ("partly", "Presence of feedback fields; quality needs a reviewer"),
    "DB4": ("yes", "Yellow highlight on the correct answer"),
    "DB5": ("partly", "Near-duplicate wording against the Pop Quiz; reworded duplicates need a reviewer"),
    "WS1": ("partly", "Question count; the 40/60 split needs question levels"),
    "WS2": ("no", "Formats are listed for the reviewer; the workbook does not say how many formats make a mix"),
    "WS3": ("yes", "SLO tag present in every question block"),
    "WS4": ("no", PHASE2), "WS5": ("no", PHASE2),
    "WS6": ("partly", "Answer Key section found at the end; completeness needs a reviewer"),
    "ST1": ("no", PHASE2), "ST2": ("no", PHASE2),
    "DBS1": ("partly", "Sort order by chapter and lesson; subject tab needs a reviewer"),
    "WE1": ("partly", "Shape [Type]-[Identifier]-[Topic]-v[Version] checked; the meaning of each part is not"),
    "WE2": ("yes", "File name searched for 'Draft' / 'Final'"),
    "WE3": ("partly", "Version presence is checked; whether it matches the document's stage needs a reviewer"),
    "WE4": ("yes", "Page size, orientation, margins and effective line spacing"),
    "WE5": ("yes", "Slide size and line spacing"),
    "WE6": ("yes", "Effective font of every English text run"),
    "WE7": ("yes", "Effective complex-script font of every Urdu text run"),
    "WE8": ("yes", "Effective sizes and weights by paragraph role (Title, Heading 1, body, table, caption, TOC)"),
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
    "WE21": ("no", AI), "WE22": ("no", AI), "WE24": ("no", HUMAN),
    "WE23": ("partly", "Run colours, shading and style colours; colours inherited from a theme are not resolved for slides"),
    "WE25": ("partly", "Footer has page number, version, date and a name; the name is not compared with the real one"),
}
