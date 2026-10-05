"""Write the three outputs: filled checklist workbook, JSON, and an HTML report."""
import html
import json
from collections import Counter
from datetime import datetime
from pathlib import Path

from . import workbook
from .checks.registry import AUTOMATION
from .models import FAIL, NA, PASS, REVIEW

STATUS_LABEL = {PASS: "Pass", FAIL: "Fail", NA: "N/A", REVIEW: "Needs review"}


def _summary_rows(rules):
    rows = []
    for code, rule in rules.items():
        level, how = AUTOMATION.get(code, ("no", "No implementation is bound to this code"))
        rows.append((code, {"yes": "Yes", "partly": "Partly", "no": "No"}[level], how, rule.text))
    return rows


def write_outputs(pkg, res, rules, layout, checklist_path, out_dir, zip_name):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    meta = {
        "Package": zip_name,
        "Reviewed at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Checklist workbook": f"{Path(checklist_path).name} (sha1 {workbook.file_sha1(checklist_path)})",
        "Rule overrides": "None. The workbook is followed as written.",
        "Documents classified": len(pkg.docs),
        "Files not classified": len(pkg.unclassified),
        "How to read the checklist sheets": "Pass/Fail/N/A were decided by code. A blank cell was not decided; the Notes column says why.",
    }
    subject_cells = {c: (s, n) for c, (s, n, _) in res.subject.items()}
    xlsx_path = out / "Course-Review-Checklist-filled.xlsx"
    workbook.fill_workbook(checklist_path, xlsx_path, res.grid, subject_cells, _summary_rows(rules), meta)

    data = {
        "meta": meta,
        "inventory": res.inventory,
        "subject_level": {c: {"status": s, "note": n, "documents": [f.to_dict() for f in fs]} for c, (s, n, fs) in res.subject.items()},
        "findings": [f.to_dict() for f in res.findings],
        "automation": {c: {"level": AUTOMATION.get(c, ("no", ""))[0], "how": AUTOMATION.get(c, ("no", ""))[1]} for c in rules},
    }
    (out / "review.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf8")
    (out / "report.html").write_text(_html(pkg, res, rules, layout, meta), encoding="utf8")
    return xlsx_path


CSS = """
:root{--bg:#f5f7f9;--fg:#14202b;--mut:#566472;--line:#d5dce3;--card:#fff;--pass:#146c43;--fail:#b42318;--rev:#8a5a00;--na:#566472}
@media (prefers-color-scheme:dark){:root{--bg:#0e1519;--fg:#e3eaf0;--mut:#93a2ae;--line:#27333b;--card:#151e24;--pass:#5fd3a0;--fail:#ff8a80;--rev:#e3a04a;--na:#93a2ae}}
body{background:var(--bg);color:var(--fg);font:14px/1.5 system-ui,Segoe UI,sans-serif;margin:0;padding:24px 16px}
main{max-width:1100px;margin:auto;display:flex;flex-direction:column;gap:28px}
h1{font-size:1.5rem;margin:0}h2{font-size:1.1rem;margin:0 0 8px}
table{border-collapse:collapse;width:100%;background:var(--card)}th,td{border:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}
th{font-size:.75rem;text-transform:uppercase;letter-spacing:.05em;color:var(--mut)}
.scroll{overflow-x:auto}.pass{color:var(--pass);font-weight:600}.fail{color:var(--fail);font-weight:600}.review{color:var(--rev);font-weight:600}.na{color:var(--na)}
.mut{color:var(--mut)}code{font-family:ui-monospace,Consolas,monospace;font-size:.85em}
details{margin:2px 0}summary{cursor:pointer}
.cell{min-width:2.4rem;text-align:center;font-weight:600}
"""


def _cls(status, partial=False):
    if status == PASS:
        return "pass" if not partial else "review"
    return {FAIL: "fail", REVIEW: "review", NA: "na"}.get(status, "")


def _html(pkg, res, rules, layout, meta):
    e = html.escape
    counts = Counter(f.status for f in res.findings)
    parts = [f"<!doctype html><meta charset=utf-8><title>Course review: {e(pkg.subject)}</title><style>{CSS}</style><main>"]
    parts.append(f"<header><h1>Course review: {e(pkg.subject)}</h1><p class=mut>{e(meta['Reviewed at'])} · checklist {e(meta['Checklist workbook'])} · "
                 f"{len(pkg.docs)} documents · rule overrides: none</p>"
                 f"<p><span class=fail>{counts[FAIL]} fails</span> · <span class=pass>{counts[PASS]} passes</span> · "
                 f"<span class=review>{counts[REVIEW]} need review</span> · {counts[NA]} not applicable. "
                 f"A pass marked partial covers only the part of the rule that code can decide.</p></header>")

    # subject level
    parts.append("<section><h2>Subject-level rules</h2><div class=scroll><table><tr><th>Code<th>Rule<th>Result<th>Detail</tr>")
    for code, (status, note, fs) in sorted(res.subject.items()):
        label = {"pass": "Pass", "fail": "Fail", "na": "N/A", None: "Not decided"}[status]
        cls = {"pass": "pass", "fail": "fail", "na": "na", None: "review"}[status]
        ev = ""
        if fs:
            lines = [f"<li><code>{e(f.doc or '')}</code>: {e(STATUS_LABEL[f.status])}{' (partial)' if f.partial else ''}. {e(f.message)}</li>" for f in fs[:40]]
            ev = f"<details><summary>per-document results ({len(fs)})</summary><ul>{''.join(lines)}</ul></details>"
        parts.append(f"<tr><td><b>{code}</b><td>{e(rules[code].text)}<td class={cls}>{label}<td>{e(note)}{ev}</tr>")
    parts.append("</table></div></section>")

    # per-lesson grids
    for sheet, rows in res.grid.items():
        spec = layout.matrix[sheet]
        parts.append(f"<section><h2>{e(sheet)}</h2><div class=scroll><table><tr><th>{'Chapter' if spec['scope']=='chapter' else 'Lesson'}")
        parts.extend(f"<th title=\"{e(rules[c].text)}\">{c}" for c in spec["codes"])
        parts.append("<th>Notes</tr>")
        for row in rows:
            tds = []
            for c in spec["codes"]:
                v = row["cells"].get(c)
                txt, cls = {"pass": ("Pass", "pass"), "fail": ("Fail", "fail"), "na": ("N/A", "na"), None: ("·", "mut")}[v]
                tds.append(f"<td class='cell {cls}'>{txt}")
            parts.append(f"<tr><td>{e(row['label'])}{''.join(tds)}<td>{e(row['note'])}</tr>")
        parts.append("</table></div></section>")

    # inventory
    inv = res.inventory
    parts.append("<section><h2>Package inventory</h2>")
    parts.append(f"<p class=mut>Documents by type: {', '.join(f'{k} {v}' for k, v in sorted(inv['counts'].items(), key=lambda x: str(x[0])))}</p>")
    if inv["missing"]:
        parts.append("<p><b>Expected but not found</b></p><ul>" + "".join(f"<li>{e(a)}: {e(b)}</li>" for a, b in inv["missing"]) + "</ul>")
    if inv["unclassified"]:
        parts.append("<p><b>Files that could not be placed</b></p><ul>" + "".join(f"<li><code>{e(p)}</code></li>" for p in inv["unclassified"]) + "</ul>")
    notes = [(d["path"], n) for d in inv["documents"] for n in d["notes"]]
    if notes:
        parts.append("<p><b>Notes</b></p><ul>" + "".join(f"<li><code>{e(p)}</code>: {e(n)}</li>" for p, n in notes[:60]) + "</ul>")
    if inv["parse_errors"]:
        parts.append("<p><b>Files that could not be read</b></p><ul>" + "".join(f"<li><code>{e(p)}</code>: {e(m)}</li>" for p, m in inv["parse_errors"].items()) + "</ul>")
    parts.append("</section>")

    # automation table
    parts.append("<section><h2>What the tool decides</h2><div class=scroll><table><tr><th>Code<th>By tool<th>How / why not<th>Rule</tr>")
    for code, rule in rules.items():
        level, how = AUTOMATION.get(code, ("no", "No implementation bound"))
        parts.append(f"<tr><td>{code}<td>{ {'yes':'Yes','partly':'Partly','no':'No'}[level] }<td>{e(how)}<td>{e(rule.text)}</tr>")
    parts.append("</table></div></section></main>")
    return "".join(parts)
