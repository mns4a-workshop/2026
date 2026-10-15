#!/usr/bin/env python3
"""
Generates the LaTeX file with one MNS4A badge (A6, 105x148 mm) per person
listed in the agenda spreadsheet.

Uses the columns "Name: (to be displayed on the badge)" and
"Institution: (to be displayed on the badge)". Long names and institutions
are split automatically into balanced lines. Institution is printed as
"- INPE -"; if it is empty, only the name is printed.

* Nothing is dropped: repeated people still get a badge, but they are flagged
  ("REPEATED - check" in the .tex and a list in the terminal) so you can check.
* If crachas_template.tex does not exist yet, it is created automatically
  (same layout as your main.tex: background image + name at 85 mm).
* It only writes the .tex file (it does not compile). Each run creates a new
  file: crachas.tex, then crachas_02.tex, crachas_03.tex...

Usage:
    python gerar_crachas.py mns4a_agenda.xlsx

Requirements: pip install openpyxl
"""
import argparse
import itertools
import math
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

# ----------------------------------------------------------------------
# Default LaTeX template. It is written to disk (crachas_template.tex) when
# no template file exists yet; edit that file afterwards to customise.
# ----------------------------------------------------------------------
DEFAULT_TEMPLATE = r'''\documentclass[10pt]{article}
\usepackage[paperwidth=105mm, paperheight=148mm, top=0mm, bottom=0mm, left=0mm, right=0mm]{geometry}
\usepackage{graphicx}
\usepackage{tikz}
\pagestyle{empty}
\setlength{\parindent}{0pt}

% Background image of the badge
\newcommand{\backgroundImage}{cracha_colorido_final.png}

% Vertical position of the name (distance from the top of the badge)
\newcommand{\nameposition}{85mm}

% One badge per page. #1 = name (lines separated by \\) and institution
\newcommand{\badge}[1]{%
  \begin{tikzpicture}[remember picture, overlay]
    \node[anchor=north west, inner sep=0pt] at (current page.north west)
      {\includegraphics[width=\paperwidth, height=\paperheight]{\backgroundImage}};
  \end{tikzpicture}%
  \vspace*{\nameposition}%
  \begin{center}
    \textbf{\fontsize{12}{18}\selectfont #1}
  \end{center}%
  \newpage}

'''

# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
LATEX_SPECIAL = {
    "\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
    "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}
UNICODE_TO_LATEX = {
    "–": "--", "—": "---", "‘": "`", "’": "'", "“": "``", "”": "''",
    "σ": r"$\sigma$", "φ": r"$\varphi$", "μ": r"$\mu$", " ": " ",
}


def tex_escape(text: str) -> str:
    out = []
    for ch in text:
        out.append(LATEX_SPECIAL.get(ch) or UNICODE_TO_LATEX.get(ch) or ch)
    return "".join(out)


def clean(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def fix_case(name: str) -> str:
    """'IAN GABRIEL' -> 'Ian Gabriel' (only when the whole name is upper case)."""
    if not name.isupper():
        return name
    small = {"da", "de", "do", "das", "dos", "e"}
    words = name.title().split()
    return " ".join(w.lower() if i and w.lower() in small else w for i, w in enumerate(words))


def key(text: str) -> str:
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z]", "", text.lower())


def balanced_lines(text: str, max_chars: int) -> list:
    """Splits text at spaces into the fewest lines of <= max_chars characters,
    as even as possible (so there are no orphan words)."""
    words = text.split()
    if len(text) <= max_chars or len(words) < 2:
        return [text]
    for n_lines in range(max(2, math.ceil(len(text) / max_chars)), len(words) + 1):
        best = None
        for cuts in itertools.combinations(range(1, len(words)), n_lines - 1):
            bounds = (0,) + cuts + (len(words),)
            lines = [" ".join(words[a:b]) for a, b in zip(bounds, bounds[1:])]
            longest = max(len(l) for l in lines)
            if best is None or longest < best[0]:
                best = (longest, lines)
        if best[0] <= max_chars or n_lines == len(words):
            return best[1]
    return [text]


# ----------------------------------------------------------------------
# Spreadsheet
# ----------------------------------------------------------------------
def find_columns(header):
    cols = {}
    for i, h in enumerate(header):
        norm = clean(h).lower()
        if "badge" not in norm:
            continue
        if norm.startswith("name") and "name" not in cols:
            cols["name"] = i
        elif norm.startswith("institution") and "institution" not in cols:
            cols["institution"] = i
    missing = {"name", "institution"} - set(cols)
    if missing:
        sys.exit(f"Columns not found: {sorted(missing)} (looking for headers containing "
                 f"'to be displayed on the badge').")
    return cols


def read_people(xlsx: Path, sheet: str):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    cols = find_columns(rows[0])

    people, skipped = [], []
    for n, row in enumerate(rows[1:], start=2):
        name = fix_case(clean(row[cols["name"]]))
        inst = clean(row[cols["institution"]])
        if inst.lower() in ("to be confirmed", "-", "n/a", "none"):
            inst = ""
        if not name:
            if inst:
                skipped.append((n, "(no name)", "institution filled but no badge name"))
            continue
        if name.lower() == "to be confirmed":
            skipped.append((n, name, "name still 'To be confirmed'"))
            continue
        people.append({"row": n, "name": name, "institution": inst})

    # Nothing is dropped: repeated people still get a badge, only flagged.
    count = {}
    for p in people:
        count[key(p["name"])] = count.get(key(p["name"]), 0) + 1
    repeated = []
    for p in people:
        if count[key(p["name"])] > 1:
            p["repeat"] = True
            repeated.append((p["row"], p["name"]))
    return people, skipped, repeated


# ----------------------------------------------------------------------
# LaTeX generation
# ----------------------------------------------------------------------
def badge_text(p, name_chars=24, inst_chars=32) -> str:
    name = r" \\ ".join(tex_escape(l) for l in balanced_lines(p["name"], name_chars))
    if not p["institution"]:
        return name
    inst_lines = [tex_escape(l) for l in balanced_lines(p["institution"], inst_chars)]
    inst_lines[0] = "- " + inst_lines[0]
    inst_lines[-1] = inst_lines[-1] + " -"
    return name + r" \\ \vspace{1em} " + r" \\ ".join(inst_lines)


def build_tex(template: Path, people) -> str:
    src = template.read_text(encoding="utf8")
    marker = r"\begin{document}"
    preamble = src.split(marker)[0]
    if r"\newcommand{\badge}" not in preamble:
        sys.exit("The template does not define \\badge before \\begin{document}.")
    body = []
    for p in people:
        flag = "  <<< REPEATED - check" if p.get("repeat") else ""
        body.append(f"% row {p['row']}{flag}\n\\badge{{{badge_text(p)}}}\n")
    return preamble.rstrip() + "\n\n" + marker + "\n\n" + "\n".join(body) + "\n\\end{document}\n"


def unique_output(path: Path) -> Path:
    """crachas.tex -> crachas.tex, then crachas_02.tex, _03... (never overwrites)."""
    n = 1
    while True:
        cand = path if n == 1 else path.with_name(f"{path.stem}_{n:02d}{path.suffix}")
        if not cand.exists():
            return cand
        n += 1


# ----------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx", type=Path, help="agenda spreadsheet")
    ap.add_argument("--sheet", default="Programa", help="sheet name (default: Programa)")
    ap.add_argument("--template", type=Path, default=Path("crachas_template.tex"))
    ap.add_argument("--output", type=Path, default=Path("crachas.tex"))
    args = ap.parse_args()

    if not args.xlsx.exists():
        sys.exit(f"Spreadsheet '{args.xlsx}' not found.")
    if not args.template.exists():
        alt = Path(__file__).resolve().parent / args.template.name
        if alt.exists():
            args.template = alt
        else:
            args.template.write_text(DEFAULT_TEMPLATE, encoding="utf8")
            print(f"Template not found - created {args.template}")

    people, skipped, repeated = read_people(args.xlsx, args.sheet)
    args.output = unique_output(args.output)
    args.output.write_text(build_tex(args.template, people), encoding="utf8")
    print(f"{len(people)} badges written to {args.output}")

    if repeated:
        print(f"\n{len(repeated)} repeated rows (badge generated for each; marked "
              f"'REPEATED - check' in the .tex):")
        for n, name in repeated:
            print(f"  row {n}: {name}")
    if skipped:
        print(f"\n{len(skipped)} rows skipped:")
        for n, name, why in skipped:
            print(f"  row {n}: {name} - {why}")


if __name__ == "__main__":
    main()
