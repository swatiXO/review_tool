"""Plain wording for the notes written into the documents, for people who do not know the checklist codes.

Every note says, in everyday words, what happened or what to do. The checklist code stays at the end in
brackets, so a reviewer can still find the rule in the workbook.
"""
import re

LEGEND = ("Colours: Pink = fixed for you, nothing to do. Red = please fix. "
          "Turquoise = an AI check, please decide. Green = OK.")

# notes the checks put on a line -> what the reader does about it
NOTE_WORDS = [
    (r"heading has no decimal number", "Number this heading (1, 1.1, 1.1.1)."),
    (r"looks like a heading but uses the '.*?' style", "Use a Word heading style (Heading 1 or Heading 2) for this heading."),
    (r"body text in the '.*?' style", "This is normal text: use the Normal style, not a heading style."),
    (r"warm-up question beyond the third", "The warm-up should ask at most 3 questions. This one is extra."),
    (r"numbered list item; use a bullet", "Use a bullet here instead of a number."),
    (r"SLO item is .*?, should be a bullet", "Write this SLO as a bullet."),
    (r"numbered; Key Takeaways are written as bullets", "Write the Key Takeaways as bullets, not numbers."),
    (r"not a bullet; Key Takeaways are written as bullets", "Write the Key Takeaways as bullets."),
    (r"Urdu text set left-to-right", "Set this Urdu text right-to-left."),
    (r"colour used", "Remove the colour. Only black, white and grey are allowed."),
    (r"has no 'Table N:' caption", "Add a title above this table, like 'Table 1: ...'."),
    (r"has no 'Figure N:' caption|no 'Figure N:' caption below", "Add a caption under this picture, like 'Figure 1: ...'."),
    (r"no alt text", "Add alt text to this picture: right-click it, then Edit Alt Text."),
    (r"not a multiple-choice question", "Pop Quiz questions must be multiple choice."),
    (r"no 'Lesson Plan Location' given", "Say where in the Lesson Plan this question appears (e.g. 'After 4.2 Liquids')."),
    (r"no correct/incorrect feedback", "Add feedback for a correct answer and for a wrong answer."),
    (r"correct answer is not highlighted yellow", "Highlight the correct answer in yellow."),
    (r"question has no SLO tag", "Tag this question with the SLO it tests."),
    (r"higher-order question", "Too many higher-order questions for this exam's mix."),
    (r"lower-order question", "Too many lower-order questions for this exam's mix."),
    (r"copied word for word from the Lesson Plan", "This slide copies the Lesson Plan word for word. Shorten it for the slide."),
]

# whole-document results -> a to-do line (message as written by the check is used when there is no entry)
TODO_WORDS = {
    "WE1": lambda m: "Add the version to the file name, e.g. Lesson-Plan-Lesson-1-Chapter-1-v1.0.docx.",
    "WE3": lambda m: "Add a version number to the file name, e.g. ...-v1.0.docx.",
    "WE25": lambda m: "Add the version to the footer." if "version" in m.lower() else "Add a footer: name, version, date, page number.",
    "WE17": lambda m: "Add a caption under each picture, like 'Figure 1: ...'" + _count(m, " ({} pictures)") + ".",
    "WE18": lambda m: "Some pictures are in colour. Use black-and-white or greyscale pictures" + _count(m, " ({} found)") + ".",
    "WE20": lambda m: "Add alt text to the pictures (right-click a picture, then Edit Alt Text)" + _count(m, " ({} pictures)") + ".",
    "LPG3": lambda m: "Add the Book and SLO Coverage table at the end (columns: Book Heading, Lesson Plan Location, SLOs).",
    "LPG4": lambda m: "If the lesson uses hard words, add a 'Glossary Terms' list at the end of Concept Building.",
    "WE15": lambda m: "Add a table of contents if the file is longer than two pages.",
    "LP3": lambda m: "Some Lesson Plan sections are missing or out of order: " + m + ".",
    "WE8": lambda m: _sizes(m),
}
# results that only repeat others (LP10 sums up WE1-WE8, FG8 sums up WE5-WE23, WE14 is LP3)
REPEATS = {"LP10", "FG8", "WE14"}


def _sizes(message):
    """'Section heading (14 bold): 31% of text is wrong (12pt: 39)' -> 'Section headings should be 14 pt, bold.'"""
    parts = re.findall(r"([A-Za-z][A-Za-z ]*?) \((\d+)( bold)?\): \d+% of text is wrong", message or "")
    if not parts:
        return "Some text is not the size the guidelines ask."
    wants = [f"{name.strip().lower()} {size} pt" + (", bold" if bold else "") for name, size, bold in parts]
    return "Some text is not the right size. It should be: " + "; ".join(wants) + "."


def _count(message, fmt):
    m = re.search(r"\b(\d+)\b", message or "")
    return fmt.format(m.group(1)) if m else ""


def note_words(note):
    for rx, words in NOTE_WORDS:
        if re.search(rx, note or "", re.I):
            return words
    return (note or "").rstrip(".") + "."


def tag(code, name=None):
    return f" [{code}{' ' + name if name else ''}]" if code else ""


def fail_line(code, note, name=None):
    return "To fix: " + note_words(note) + tag(code, name)


def check_line(code, note, name=None):
    return "Please check: " + note_words(note) + tag(code, name)


def pass_line(code, note, name=None):
    return "OK: " + note_words(note) + tag(code, name)


def todo_line(code, message, name=None):
    words = TODO_WORDS.get(code)
    return "To fix: " + (words(message) if words else note_words(message)) + tag(code, name)


def ai_lines(code, verdict, reason, first, name=None):
    reason = (reason or "").strip()
    if not first:
        return ["Same as the AI note above" + (f": {reason}" if reason and verdict == "fail" else ".") + tag(code, name)]
    if verdict == "pass":
        return [f"AI check, looks fine: {reason}" + tag(code, name),
                "If you agree, delete this note. If not, write what is missing."]
    if verdict == "fail":
        return [f"AI check, may need fixing: {reason}" + tag(code, name),
                "If you agree, fix it or tell the writer. If not, delete this note."]
    return ["The AI could not decide this one. Please check it yourself." + tag(code, name)]


def fixed_line(words):
    return "Fixed for you: " + words


def source_note(note):
    """How the file was recognised, in plain words, for the summary at the top."""
    m = re.match(r"Recognised as (.+?) (?:from its content|by the model)", note or "")
    if m:
        return f"Treated as a {m.group(1).title()}, worked out from what is in the file."
    if (note or "").startswith("converted from"):
        return "Converted from an old Word or PowerPoint file (.doc / .ppt)."
    return note
