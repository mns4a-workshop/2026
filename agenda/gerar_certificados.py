#!/usr/bin/env python3
"""
Generates one MNS4A certificate per person listed in the agenda spreadsheet.

* Presenters (rows with a title and a session): presentation certificate,
  built from the columns "Full Name", "Title" and "Session" (the "Thematic
  Session" column; the "SESSION 01 · MODELING · " prefix is removed). The
  "Type" column picks the text: Oral Presentation, Poster Presentation or
  Invited.
* Attendees (rows with a name but neither title nor session): participation
  certificate with the name only. People who also present receive both
  certificates. Use --no-attendees to leave attendees out.

It reuses the preamble of certificado_mns4a_texto.tex (fonts, colours,
\\oral / \\poster / \\convidado / \\ouvinte macros, signatories, background
image) and replaces the body of the document with one call per person.

If certificado_mns4a_texto.tex does not exist yet, it is created automatically.
It only writes the .tex file (it does not compile). Each run creates a new
file: certificados_apresentacao.tex, then certificados_apresentacao_02.tex...

Usage:
    python gerar_certificados.py mns4a_agenda.xlsx

Requirements: pip install openpyxl
"""
import argparse
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

# ----------------------------------------------------------------------
# Default LaTeX template. It is written to disk (certificado_mns4a_texto.tex)
# when no template file exists yet; edit that file afterwards to customise.
# ----------------------------------------------------------------------
DEFAULT_TEMPLATE = r'''% =====================================================================
% MNS4A certificate - all text placed over the background image
%
% Put the background image (logos, MNS4A header, frame; WITHOUT any of
% the text below) in the same folder as fundo.jpg (or change \fundo).
% Compile with: pdflatex certificado_mns4a_texto.tex
%
% Usage (one page per certificate):
%   \certificado{Name}{Paragraph text}
% Use \\ inside the paragraph for the line break.
% The fixed texts (title, "This is to certify that", signatories) are
% defined in the configuration block below.
% =====================================================================
\documentclass[a4paper,landscape]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[a4paper,landscape,margin=0cm]{geometry}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tikz}
\usepackage{helvet}
\usepackage{mathpazo}   % serif font for the name
\renewcommand{\familydefault}{\sfdefault}
\pagestyle{empty}
\setlength{\parindent}{0pt}

\definecolor{navy}{HTML}{0B2A5B}
\definecolor{bluetext}{HTML}{1F4E8C}

% ---------- Configuration ----------
\newcommand{\fundo}{fundo.jpg}   % background image

% Fixed texts
\newcommand{\titulo}{CERTIFICATE OF RECOGNITION}
\newcommand{\certifythat}{This is to certify that}

% Signatories
\newcommand{\assinanteA}{Dr. Esfhan Alam Kherani}
\newcommand{\cargoA}{Chair of NMS4A}
\newcommand{\assinanteB}{Dr. L\'aysa Cristina Araujo Resende Chagas}
\newcommand{\cargoB}{President of SBGEA}

% Vertical positions (distance from the top of the page) - adjust if needed
\newcommand{\ytitulo}{8.0cm}
\newcommand{\ycertify}{9.6cm}
\newcommand{\ynome}{11.6cm}
\newcommand{\ycorpo}{14.2cm}
\newcommand{\yassin}{18.0cm}

% ---------- One certificate ----------
% #1 name | #2 paragraph text
\newcommand{\certificado}[2]{%
  \thispagestyle{empty}%
  \begin{tikzpicture}[remember picture,overlay,shift={(current page.north west)}]
    \node[anchor=north west,inner sep=0] at (0,0)
      {\includegraphics[width=\paperwidth,height=\paperheight]{\fundo}};
    % Title
    \node[font=\LARGE\color{bluetext}]
      at (0.5\paperwidth,-\ytitulo) {\titulo};
    % "This is to certify that"
    \node[font=\large\color{navy}]
      at (0.5\paperwidth,-\ycertify) {\certifythat};
    % Name
    \node[font=\fontsize{28}{32}\selectfont\bfseries\rmfamily\color{navy}]
      at (0.5\paperwidth,-\ynome) {#1};
    % Body
    \node[font=\Large\color{navy},text width=0.95\paperwidth,align=center]
      at (0.5\paperwidth,-\ycorpo) {#2};
    % Signatories
    \node[align=center,font=\normalsize\color{navy}]
      at (0.328\paperwidth,-\yassin) {{\bfseries\assinanteA}\\\cargoA};
    \node[align=center,font=\normalsize\color{navy}]
      at (0.715\paperwidth,-\yassin) {{\bfseries\assinanteB}\\\cargoB};
  \end{tikzpicture}%
  \clearpage}

% ---------- Default texts for presentation certificates ----------
% #1 presenter name | #2 title of the work | #3 session name
\newcommand{\localdata}{of the MNS4A Workshop,
  held in S\~ao Jos\'e dos Campos (SP), Brazil, from 19 to 21 October 2026.}

% Oral presentation
\newcommand{\oral}[3]{%
  \certificado{#1}{has presented the oral talk entitled {\bfseries ``#2''}
    in the Session ``#3'' \localdata}}

% Poster presentation
\newcommand{\poster}[3]{%
  \certificado{#1}{has presented the poster entitled {\bfseries ``#2''}
    in the Session ``#3'' \localdata}}

% Invited talk
\newcommand{\convidado}[3]{%
  \certificado{#1}{has presented the invited talk entitled {\bfseries ``#2''}
    in the Session ``#3'' \localdata}}

% ---------- Default text for attendee (listener) certificates ----------
% #1 attendee name
\newcommand{\ouvinte}[1]{{%
  \renewcommand{\titulo}{CERTIFICATE OF PARTICIPATION}%
  \certificado{#1}{has attended the MNS4A Workshop,
    held in S\~ao Jos\'e dos Campos (SP), Brazil, from 19 to 21 October 2026.}}}

\begin{document}

% Examples (one of each type)
\oral{Presenter Name}{Title of the Presented Work}{Session Name}
\poster{Presenter Name}{Title of the Presented Work}{Session Name}
\convidado{Presenter Name}{Title of the Presented Work}{Session Name}
\ouvinte{Attendee Name}

\end{document}
'''


# ----------------------------------------------------------------------
# LaTeX helpers
# ----------------------------------------------------------------------
LATEX_SPECIAL = {
    "\\": r"\textbackslash{}",
    "&": r"\&",
    "%": r"\%",
    "$": r"\$",
    "#": r"\#",
    "_": r"\_",
    "{": r"\{",
    "}": r"\}",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}

# Characters that the utf8 input encoding cannot typeset in text mode
UNICODE_TO_LATEX = {
    "–": "--",
    "—": "---",
    "‘": "`",
    "’": "'",
    "“": "``",
    "”": "''",
    "σ": r"$\sigma$",
    "φ": r"$\varphi$",
    "μ": r"$\mu$",
    "°": r"$^\circ$",
    "×": r"$\times$",
    "\u00a0": " ",
}


def tex_escape(text: str) -> str:
    out = []
    for ch in text:
        if ch in LATEX_SPECIAL:
            out.append(LATEX_SPECIAL[ch])
        elif ch in UNICODE_TO_LATEX:
            out.append(UNICODE_TO_LATEX[ch])
        else:
            out.append(ch)
    return "".join(out)


def clean(value) -> str:
    """Normalise a spreadsheet cell to a single-line string."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def clean_session(value: str) -> str:
    """'SESSION 01 · MODELING · Modeling of ...' -> 'Modeling of ...'"""
    parts = [p.strip() for p in value.split("·")]
    return parts[-1] if len(parts) >= 3 else value


def fix_case(name: str) -> str:
    """'IAN GABRIEL OLIVEIRA DA SILVA' -> 'Ian Gabriel Oliveira da Silva'."""
    if not name.isupper():
        return name
    small = {"da", "de", "do", "das", "dos", "e"}
    words = name.title().split()
    return " ".join(w.lower() if i and w.lower() in small else w for i, w in enumerate(words))


def key(text: str) -> str:
    """Accent/case-insensitive key used to detect the same person twice."""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z]", "", text.lower())


def slug(text: str) -> str:
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_") or "certificado"


# ----------------------------------------------------------------------
# Spreadsheet
# ----------------------------------------------------------------------
# "Type" column value (substring, lower case) -> LaTeX macro of the template
TYPE_MACROS = {"oral": "oral", "poster": "poster", "invited": "convidado"}


def find_columns(header):
    """Maps logical names to column indexes, tolerant to ':' and spaces."""
    norm = [clean(h).lower().rstrip(":").strip() for h in header]
    cols = {}
    for i, name in enumerate(norm):
        if name == "full name" and "name" not in cols:
            cols["name"] = i
        elif name == "type" and "type" not in cols:
            cols["type"] = i
        elif name == "title" and "title" not in cols:
            cols["title"] = i
        elif (name.startswith("thematic session") or name == "session") and "session" not in cols:
            cols["session"] = i
    missing = {"name", "title", "session", "type"} - set(cols)
    if missing:
        sys.exit(f"Columns not found in the spreadsheet: {sorted(missing)}\n"
                 f"Headers found: {norm}")
    return cols


def read_presenters(xlsx: Path, sheet: str):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    cols = find_columns(rows[0])

    presenters, attendees, skipped = [], [], []
    for n, row in enumerate(rows[1:], start=2):
        name = fix_case(clean(row[cols["name"]]))
        title = clean(row[cols["title"]])
        session = clean_session(clean(row[cols["session"]]))
        if not name:
            continue
        kind = clean(row[cols["type"]]).lower()
        if not title and not session:
            attendees.append({"row": n, "name": name})
        elif not title or title.lower() == "to be confirmed" or not session:
            skipped.append((n, name, "title or session missing / still 'To be confirmed'"))
        else:
            macro = next((m for word, m in TYPE_MACROS.items() if word in kind), None)
            if macro is None:
                skipped.append((n, name, f"unknown presentation type '{kind}'"))
                continue
            presenters.append({"row": n, "name": name, "title": title,
                               "session": session, "macro": macro})

    # Nothing is dropped: people who present AND attend get both certificates,
    # and repeated rows still produce a certificate. Repeats are only flagged
    # (comment in the .tex + message in the terminal) so you can check them.
    repeated = []
    for group, kind in ((attendees, "attendee"), (presenters, "presenter")):
        count = {}
        for item in group:
            k = (key(item["name"]), key(item.get("title", "")))
            count[k] = count.get(k, 0) + 1
        for item in group:
            if count[(key(item["name"]), key(item.get("title", "")))] > 1:
                item["repeat"] = True
                repeated.append((item["row"], item["name"], kind))
    return presenters, attendees, skipped, repeated


# ----------------------------------------------------------------------
# LaTeX generation
# ----------------------------------------------------------------------
def build_tex(template: Path, presenters, attendees) -> str:
    src = template.read_text(encoding="utf8")
    marker = r"\begin{document}"
    if marker not in src:
        sys.exit("The template has no \\begin{document}.")
    preamble = src.split(marker)[0]
    needed = {p["macro"] for p in presenters} | ({"ouvinte"} if attendees else set())
    for macro in sorted(needed):
        if "\\newcommand{\\" + macro + "}" not in preamble:
            sys.exit(f"The template does not define \\{macro} before \\begin{{document}}.")

    body = []
    for p in presenters:
        body.append(
            f"% row {p['row']}{'  <<< REPEATED - check' if p.get('repeat') else ''}\n"
            f"\\{p['macro']}\n"
            f"  {{{tex_escape(p['name'])}}}\n"
            f"  {{{tex_escape(p['title'])}}}\n"
            f"  {{{tex_escape(p['session'])}}}\n"
        )
    for a in attendees:
        flag = "  <<< REPEATED - check" if a.get("repeat") else ""
        body.append(f"% row {a['row']}{flag}\n\\ouvinte{{{tex_escape(a['name'])}}}\n")
    return preamble + marker + "\n\n" + "\n".join(body) + "\n\\end{document}\n"


def unique_output(path: Path) -> Path:
    """certificados.tex -> certificados.tex, then certificados_02.tex, _03...
    Never overwrites a file generated by a previous run."""
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
    ap.add_argument("--template", type=Path, default=Path("certificado_mns4a_texto.tex"))
    ap.add_argument("--output", type=Path, default=Path("certificados_apresentacao.tex"))
    ap.add_argument("--no-attendees", action="store_true",
                    help="only presentation certificates (skip attendees)")
    args = ap.parse_args()

    # template: current folder first, then the folder of this script;
    # if it does not exist anywhere, create it from the built-in default
    if not args.template.exists():
        alt = Path(__file__).resolve().parent / args.template.name
        if alt.exists():
            args.template = alt
        else:
            args.template.write_text(DEFAULT_TEMPLATE, encoding="utf8")
            print(f"Template not found - created {args.template}")
    if not args.xlsx.exists():
        sys.exit(f"Spreadsheet '{args.xlsx}' not found.")

    presenters, attendees, skipped, repeated = read_presenters(args.xlsx, args.sheet)
    if args.no_attendees:
        attendees = []
    args.output = unique_output(args.output)  # never overwrite a previous run
    args.output.write_text(build_tex(args.template, presenters, attendees), encoding="utf8")
    print(f"{len(presenters)} presentation + {len(attendees)} attendee certificates "
          f"written to {args.output}")

    if repeated and not args.no_attendees or any(k == "presenter" for _, _, k in repeated):
        shown = [r for r in repeated if not (args.no_attendees and r[2] == "attendee")]
        if shown:
            print(f"\n{len(shown)} repeated rows (certificate generated for each; marked "
                  f"'REPEATED - check' in the .tex):")
            for n, name, kind in shown:
                print(f"  row {n}: {name} ({kind})")

    if skipped:
        print(f"\n{len(skipped)} rows skipped:")
        for n, name, why in skipped:
            print(f"  row {n}: {name} - {why}")


if __name__ == "__main__":
    main()
