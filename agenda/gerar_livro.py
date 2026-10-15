#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_livro.py — Livro-resumo (Book of Abstracts) em LaTeX, UMA PÁGINA POR ENTRADA.

Usa exatamente a mesma fonte de dados do gerar_agenda.py: a planilha do programa (a exportação do
formulário serve como está) e a página index.html (que dá os dias, os temas e os horários da agenda).
Os dados são lidos e conferidos pelas mesmas funções do gerar_agenda.py; este arquivo precisa ficar na
MESMA PASTA dele.

Como usar
---------
    python gerar_livro.py planilha.xlsx                       # grava livro_resumo.tex
    python gerar_livro.py planilha.xlsx --pdf                 # grava o .tex e já compila o PDF (precisa de LaTeX)
    python gerar_livro.py planilha.xlsx --pagina index.html --saida livro/livro_resumo.tex
    python gerar_livro.py planilha.xlsx --titulo "MNS4A Workshop" --subtitulo "19–21 October 2026 · INPE"

Opções: --pagina (padrão index.html), --saida (padrão livro_resumo.tex), --aba, --titulo, --subtitulo,
        --fonte / --fonte-sans (fontes do XeLaTeX/LuaLaTeX), --estrito, --ordem-das-linhas (como no
        gerar_agenda.py), --pdf.

O que entra no livro, e em que ordem
------------------------------------
1. Capa, página "About this book" e sumário (com a hora ou o número do pôster de cada entrada).
2. Uma página por apresentação, na ordem do programa: dia 1, dia 2, dia 3 (convidadas e orais, pelo horário
   da agenda). Quem ainda não tem horário entra no fim do dia da sua sessão temática, marcado
   "Time to be announced" — um livro de resumos não pode perder trabalho só porque a agenda não fechou.
3. Uma página por pôster, agrupados por tema e numerados P01, P02... exatamente como no site.
Linhas com problema na planilha ficam de fora e são listadas no fim (como no gerar_agenda.py).

Cada página tem: tipo (Invited talk / Oral presentation / Poster Pxx), dia, data, horário e sessão;
título; autores com os números das instituições (o apresentador vem sublinhado); instituições; abstract.
O tamanho da letra do abstract é ajustado ao tamanho do texto para ele caber na página; se algum abstract
for grande demais, o script avisa (e, com --pdf, confere no PDF se cada entrada ficou em uma página só).

Compilação: XeLaTeX ou LuaLaTeX (recomendado; aceita qualquer caractere Unicode que a fonte tenha) ou
pdfLaTeX (aceita letras gregas e símbolos científicos comuns, que o script converte). No Overleaf, basta
subir o .tex. Fontes: TeX Gyre Pagella e Heros (já vêm no TeX Live e no Overleaf). Rodar duas vezes para o
sumário (o --pdf já faz isso).

Requisitos: Python 3 e openpyxl  (pip install openpyxl)
"""

import argparse
import datetime
import re
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    import gerar_agenda as ga
except ImportError:                                       # pragma: no cover
    sys.exit("O gerar_livro.py precisa estar na mesma pasta do gerar_agenda.py.")

VERSAO = "2026-10-06"

TIPOS = {"invited": "Invited talk", "oral": "Oral presentation", "poster": "Poster"}

# ============================================================================
# Texto -> LaTeX
# ============================================================================
ESPECIAIS = {
    "\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#", "_": r"\_",
    "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}", "^": r"\textasciicircum{}",
}

# Letras gregas e símbolos científicos -> modo matemático (funcionam com qualquer motor e qualquer fonte)
_GREGO = {
    "α": "alpha", "β": "beta", "γ": "gamma", "δ": "delta", "ε": "varepsilon", "ϵ": "epsilon", "ζ": "zeta",
    "η": "eta", "θ": "theta", "ϑ": "vartheta", "ι": "iota", "κ": "kappa", "λ": "lambda", "μ": "mu", "ν": "nu",
    "ξ": "xi", "π": "pi", "ρ": "rho", "σ": "sigma", "ς": "varsigma", "τ": "tau", "υ": "upsilon", "φ": "varphi",
    "ϕ": "phi", "χ": "chi", "ψ": "psi", "ω": "omega",
    "Γ": "Gamma", "Δ": "Delta", "Θ": "Theta", "Λ": "Lambda", "Ξ": "Xi", "Π": "Pi", "Σ": "Sigma",
    "Φ": "Phi", "Ψ": "Psi", "Ω": "Omega",
}
MATEMATICA = {c: "\\" + n for c, n in _GREGO.items()}
MATEMATICA.update({
    "±": r"\pm", "×": r"\times", "÷": r"\div", "≥": r"\geq", "≤": r"\leq", "≠": r"\neq", "≈": r"\approx",
    "∼": r"\sim", "≃": r"\simeq", "∝": r"\propto", "∞": r"\infty", "→": r"\rightarrow", "←": r"\leftarrow",
    "↑": r"\uparrow", "↓": r"\downarrow", "↔": r"\leftrightarrow", "⇒": r"\Rightarrow", "∂": r"\partial",
    "∇": r"\nabla", "∑": r"\sum", "∫": r"\int", "√": r"\surd", "∥": r"\parallel", "⊥": r"\perp",
    "−": "-", "⋅": r"\cdot", "′": r"^{\prime}", "″": r"^{\prime\prime}", "∆": r"\Delta",
    "≳": r"\gtrsim", "≲": r"\lesssim", "⊙": r"\odot", "≪": r"\ll", "≫": r"\gg", "∈": r"\in", "∘": r"\circ", "≡": r"\equiv", "∼": r"\sim",
    "⁰": "^{0}", "¹": "^{1}", "²": "^{2}", "³": "^{3}", "⁴": "^{4}", "⁵": "^{5}", "⁶": "^{6}", "⁷": "^{7}",
    "⁸": "^{8}", "⁹": "^{9}", "⁺": "^{+}", "⁻": "^{-}",
    "₀": "_{0}", "₁": "_{1}", "₂": "_{2}", "₃": "_{3}", "₄": "_{4}", "₅": "_{5}", "₆": "_{6}", "₇": "_{7}",
    "₈": "_{8}", "₉": "_{9}", "₊": "_{+}", "₋": "_{-}",
})
# espaços e invisíveis
ESPACOS = {"\u00a0": "~", "\u202f": r"\,", "\u2009": r"\,", "\u2002": " ", "\u2003": " ", "\u2007": "~",
           "\u200b": "", "\u200c": "", "\u200d": "", "\ufeff": "", "\u00ad": r"\-", "\u2011": "-", "\u2010": "-",
           "\u2212": None}                                  # None = trata como matemática (acima)

# o que o script considera "seguro" para a fonte: latim (até U+024F), pontuação geral, moeda, setas/símbolos mapeados
def _seguro(c):
    o = ord(c)
    return (o < 0x250 or 0x2000 <= o <= 0x206F or 0x20A0 <= o <= 0x20CF or c in MATEMATICA or c in ESPACOS
            or c in SUBSTITUICOES
            or o in (0x2122, 0x2126, 0x212B, 0x2030))


# Equações coladas do Word chegam como 〖Es〗_c ou 〖23.2〗^° S (U+3016/U+3017 + _ ou ^ para índice/expoente).
# O script tira os colchetes e vira índice/expoente de verdade.
_SUB0, _SUB1, _SUP0, _SUP1 = "\ue000", "\ue001", "\ue002", "\ue003"
_TOKEN = r"(?:\{[^}]*\}|[A-Za-z0-9\u00b0+\-\u2212])"
SUBSTITUICOES = {"\u02bb": "\u2018", "\u02bc": "\u2019", "\u3016": "", "\u3017": ""}   # okina, apóstrofo modificador, colchetes do Word


def _equacao_word(s):
    if "\u3016" not in s:
        return s

    def indices(t):
        return re.sub(r"([_^])(" + _TOKEN + ")",
                      lambda m: (_SUB0 if m.group(1) == "_" else _SUP0) + m.group(2).strip("{}")
                      + (_SUB1 if m.group(1) == "_" else _SUP1), t)

    return re.sub(r"\u3016(.*?)\u3017((?:[_^]" + _TOKEN + ")*)",
                  lambda m: indices(m.group(1)) + indices(m.group(2)), s)


def _aspas(s):
    """Aspas retas -> tipográficas (abre depois de espaço/início/abre-parêntese; fecha no resto)."""
    saida, prev = [], " "
    for c in s:
        if c == '"':
            c = "\u201c" if (prev.isspace() or prev in "([{\u2014\u2013-/") else "\u201d"
        elif c == "'":
            c = "\u2018" if (prev.isspace() or prev in "([{\u201c\u2014\u2013-/") else "\u2019"
        saida.append(c)
        prev = c
    return "".join(saida)


_URL_RE = re.compile(r"https?://[^\s<>]+[^\s<>.,;:)\]}'\"”’]")


def tex(s, mover=False):
    """
    Escapa um texto para LaTeX. mover=True é para argumentos "móveis" (sumário, marcadores do PDF): comandos
    frágeis ganham \\protect/\\texorpdfstring e URLs não viram \\url.
    """
    s = _aspas(_equacao_word(str(s)))
    partes, pos = [], 0
    for m in (_URL_RE.finditer(s) if not mover else []):
        partes.append(("t", s[pos:m.start()]))
        partes.append(("u", m.group(0)))
        pos = m.end()
    partes.append(("t", s[pos:]))
    saida = []
    for tipo, p in partes:
        if tipo == "u":
            saida.append(r"\url{" + p.replace("\\", "") + "}")
            continue
        for c in p:
            if c in ESPECIAIS:
                saida.append(ESPECIAIS[c])
            elif c in SUBSTITUICOES:
                saida.append(SUBSTITUICOES[c])
            elif c in (_SUB0, _SUP0):
                saida.append(r"\textsubscript{" if c == _SUB0 else r"\textsuperscript{")
            elif c in (_SUB1, _SUP1):
                saida.append("}")
            elif c in MATEMATICA:
                cmd = r"\ensuremath{" + MATEMATICA[c] + "}"
                saida.append(r"\texorpdfstring{" + cmd + "}{" + c + "}" if mover else cmd)
            elif c in ESPACOS and ESPACOS[c] is not None:
                saida.append(ESPACOS[c])
            elif not c.isspace() and not _seguro(c):
                # fora do que o script sabe: o XeLaTeX/LuaLaTeX imprime o caractere se a fonte tiver; o pdfLaTeX
                # mostra [U+XXXX] em vez de parar com erro
                saida.append(("\\texorpdfstring{\\unichar{%X}}{%s}" % (ord(c), c)) if mover
                             else "\\unichar{%X}" % ord(c))
            else:
                saida.append(c)
    return "".join(saida)


def caracteres_suspeitos(textos):
    """Caracteres fora do que o script sabe converter (podem faltar na fonte ou quebrar o pdfLaTeX)."""
    achados = {}
    for t in textos:
        for c in str(t):
            if not _seguro(c) and not c.isspace():
                achados[c] = achados.get(c, 0) + 1
    return achados


# ============================================================================
# Dados: mesma leitura e mesma validação do gerar_agenda.py
# ============================================================================
def carregar(args):
    pagina = Path(args.pagina)
    if not pagina.exists():
        sys.exit(f"Não encontrei a página: {pagina} (use --pagina; ela dá os dias, temas e horários)")
    linhas = pagina.read_bytes().decode("utf-8").replace("\r\n", "\n").split("\n")
    abas = [a.strip() for a in args.aba.split(",") if a.strip()] if args.aba else None
    registros, ignoradas = ga.ler_planilha(args.planilha, abas)
    grade, meta = ga.ler_grade(linhas)
    if not grade:
        raise ga.ErroPlanilha("Não encontrei os horários da agenda na página (<!-- DAY n --> e slots).")
    blocos_p, _secao = ga.ler_posters(linhas, meta)
    if not blocos_p:                                       # página sem seção de pôsteres: um bloco por tema da agenda
        blocos_p = [{"tema": i["tema"], "sessao": d, "ocupados": 0, "reservados": 0}
                    for d, i in sorted(meta.items()) if i.get("tema")]
    erros, avisos, colocados, sem_horario, excluidos, colocados_p = ga.validar(
        registros, grade, args.ordem_das_linhas, args.estrito, blocos_p, meta)
    return {"texto": linhas, "registros": registros, "ignoradas": ignoradas, "grade": grade, "meta": meta,
            "blocos_p": blocos_p, "erros": erros, "avisos": avisos, "colocados": colocados,
            "sem_horario": sem_horario, "excluidos": excluidos, "colocados_p": colocados_p}


def ler_cabecalho_da_pagina(linhas, meta):
    """(título, subtítulo) padrão, tirados da própria página: <title>, datas dos dias e o local do botão do topo."""
    t = "\n".join(linhas)
    m = re.search(r"<title>([^<]+)</title>", t)
    titulo = ga.html.unescape(m.group(1).strip()) if m else "Workshop"
    ano = re.search(r"\d{1,2}\s*[\u2013-]\s*\d{1,2}\s+[A-Za-z]+\s+(20\d\d)", t)
    periodo = ga.periodo_do_evento(meta)
    datas = (periodo + (" " + ano.group(1) if ano else "")).strip()
    m = re.search(r'class="btn-header-sub">([^<]+)<', t)
    local = ga.html.unescape(m.group(1).strip()) if m else ""
    sub = " \u00b7 ".join(x for x in (datas, local) if x)
    return titulo, sub


def montar_entradas(d):
    """
    Lista de grupos na ordem do livro:  [{'titulo': str, 'entradas': [ {...} ]}]
    cada entrada: {'r', 'tipo', 'rotulo' (hora ou Pxx), 'kicker', 'info', 'sessao'}
    """
    meta, grade = d["meta"], d["grade"]
    grupos, sem_destino = [], []
    pendentes = list(d["sem_horario"])
    for dia in sorted(meta):
        tema, data = meta[dia].get("tema", ""), meta[dia].get("data", "")
        entradas = []
        for (dd, ini), r in sorted(d["colocados"].items()):
            if dd != dia:
                continue
            g = grade[(dd, ini)]
            entradas.append({"r": r, "tipo": r["tipo"], "rotulo": ini, "hora": g["tempo"].replace("-", "\u2013"),
                             "sessao": tema})
        for r in [x for x in pendentes if (x["dia"] or x["dia_sessao"]) == dia]:
            entradas.append({"r": r, "tipo": r["tipo"], "rotulo": "", "hora": "", "sessao": tema})
            pendentes.remove(r)
        for e in entradas:
            e["dia"], e["data"] = dia, data
        if entradas:
            grupos.append({"titulo": f"Day {dia} \u00b7 {data} \u2014 {tema}".replace(" \u00b7  \u2014", " \u2014"),
                           "entradas": entradas, "tipo": "dia"})
    for r in pendentes:                                    # sem horário e sem sessão que bata com um dia da agenda
        sem_destino.append({"r": r, "tipo": r["tipo"], "rotulo": "", "hora": "", "sessao": "", "dia": None,
                            "data": ""})
    if sem_destino:
        grupos.append({"titulo": "Session to be announced", "entradas": sem_destino, "tipo": "dia"})
    periodo = ga.periodo_do_evento(meta)
    for b in d["blocos_p"]:
        if not b.get("cartoes"):
            continue
        entradas = [{"r": r, "tipo": "poster", "rotulo": f"P{n:02d}", "hora": "", "sessao": b["tema"],
                     "dia": None, "data": ""} for n, r in b["cartoes"]]
        grupos.append({"titulo": f"Poster session \u2014 {b['tema']}", "entradas": entradas, "tipo": "poster",
                       "periodo": periodo})
    return grupos


# ============================================================================
# Tamanho da letra: o abstract precisa caber na página
# ============================================================================
# Estimativa de altura (mm) numa página A4 com texto de 166 x 245 mm. É uma aproximação conservadora; o
# conferidor do --pdf diz se alguma entrada passou de uma página.
# Cada nível: (macro e (pt, entrelinha) do abstract), (macro, pt, entrelinha) dos autores e das instituições.
# Tamanhos da classe article 11pt: normalsize 10.95/13.6, small 10/12, footnotesize 9/11, scriptsize 8/9.5.
NIVEIS = [
    (("\\normalsize", 10.95, 13.6), ("\\normalsize", 10.95, 13.6), ("\\small", 10.0, 12.0)),
    (("\\small", 10.0, 12.0), ("\\normalsize", 10.95, 13.6), ("\\footnotesize", 9.0, 11.0)),
    (("\\footnotesize", 9.0, 11.0), ("\\small", 10.0, 12.0), ("\\scriptsize", 8.0, 9.5)),
    (("\\scriptsize", 8.0, 9.5), ("\\footnotesize", 9.0, 11.0), ("\\scriptsize", 8.0, 9.5)),
]
PT = 0.3528                                                # mm por ponto


def _linhas(n_chars, pt, largura_mm=166.0, fator=0.44):
    cpl = max(10, int(largura_mm / (pt * PT * fator)))
    return max(1, -(-n_chars // cpl))


def escolher_nivel(r):
    """(índice do nível, estourou): o menor nível de redução em que a entrada cabe numa página."""
    titulo = _linhas(len(r["titulo"]), 17.3, fator=0.60) * 17.3 * 1.2 * PT
    chars = [len(p) for p in r["abstract"]]
    for k, (ab, au, af) in enumerate(NIVEIS):
        autores = _linhas(sum(len(a["n"]) + 4 for a in r["autores"]) + (40 if r["apresentador"] else 0),
                          au[1]) * au[2] * PT
        afil = sum(_linhas(len(a["t"]) + 4, af[1]) for a in r["afiliacoes"]) * af[2] * PT
        fixo = 14 + titulo + autores + afil + 30            # cabeçalho, filetes, rótulo "Abstract", folgas
        h = sum(_linhas(c, ab[1]) for c in chars) * ab[2] * PT + max(0, len(chars) - 1) * ab[1] * PT * 0.6
        ultimo = k == len(NIVEIS) - 1
        if fixo + h <= 245.0 * (1.08 if ultimo else 0.93):    # no último nível, só avisa se passar bem do limite
            return k, False
    return len(NIVEIS) - 1, True


# ============================================================================
# LaTeX
# ============================================================================
PREAMBULO = r"""% !TEX program = xelatex
% Gerado por gerar_livro.py __VERSAO__ — não edite à mão (rode o gerador de novo).
% Compile com XeLaTeX ou LuaLaTeX (recomendado); pdfLaTeX também funciona. Rode 2 vezes para o sumário.
\documentclass[11pt,a4paper]{article}
\usepackage{iftex}
\ifPDFTeX
  \usepackage[utf8]{inputenc}
  \usepackage[T1]{fontenc}
  \usepackage{textcomp}
  \usepackage{mathpazo}
  \usepackage[scaled=.92]{helvet}
\else
  \usepackage{fontspec}
  \setmainfont{__FONTE__}[Ligatures=TeX]
  \setsansfont{__FONTE_SANS__}[Ligatures=TeX]
\fi
\ifPDFTeX
  \DeclareRobustCommand{\unichar}[1]{\textbf{[U+#1]}}
\else
  \DeclareRobustCommand{\unichar}[1]{\char"#1\relax}
\fi
\usepackage[english]{babel}
\usepackage[a4paper,top=27mm,bottom=24mm,left=22mm,right=22mm,headheight=14pt,headsep=8mm,footskip=11mm]{geometry}
\usepackage{microtype}
\usepackage{amssymb}
\usepackage{xcolor}
\usepackage{fancyhdr}
\usepackage{tocloft}
\usepackage{url}
\usepackage[colorlinks=true,linkcolor=navy,urlcolor=accent,citecolor=navy,plainpages=false,pdfusetitle,
            unicode=true,pdfencoding=auto]{hyperref}

\definecolor{navy}{HTML}{14305A}
\definecolor{accent}{HTML}{1C7F78}
\definecolor{muted}{HTML}{5C6B7F}
\definecolor{rule}{HTML}{B8C4D3}
\urlstyle{same}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}
\emergencystretch=2em
\hypersetup{pdftitle={__TITULO__ — Book of Abstracts},pdfsubject={Book of Abstracts}}

% ---- cabeçalho / rodapé
\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\sffamily\footnotesize\color{muted}__TITULO__ \textperiodcentered\ Book of Abstracts}
\fancyhead[R]{\sffamily\footnotesize\color{muted}\nouppercase{\rightmark}}
\fancyfoot[C]{\sffamily\footnotesize\color{muted}\thepage}
\renewcommand{\headrulewidth}{0.4pt}
\renewcommand{\headrule}{\hbox to\headwidth{\color{rule}\leaders\hrule height \headrulewidth\hfill}}
\renewcommand{\footrulewidth}{0pt}
\newcommand{\setsession}[1]{\markright{#1}}

% ---- sumário
\renewcommand{\contentsname}{Contents}
\renewcommand{\cfttoctitlefont}{\sffamily\bfseries\Huge\color{navy}}
\setlength{\cftbeforetoctitleskip}{0pt}
\setlength{\cftaftertoctitleskip}{14pt}
\renewcommand{\cftsecfont}{\sffamily\bfseries\color{navy}}
\renewcommand{\cftsecpagefont}{\sffamily\bfseries\color{navy}}
\setlength{\cftsecindent}{0pt}
\setlength{\cftbeforesecskip}{12pt}
\renewcommand{\cftsecleader}{\cftdotfill{\cftdotsep}}
\setlength{\cftsubsecindent}{0pt}
\setlength{\cftsubsecnumwidth}{3.6em}
\setlength{\cftbeforesubsecskip}{3pt}
\renewcommand{\cftsubsecfont}{\small}
\renewcommand{\cftsubsecpagefont}{\small}
\renewcommand{\cftsubsecaftersnum}{\hspace{0pt}}

% ---- página de uma entrada
\newcommand{\EntryKicker}[3]{% tipo, linha de informação, sessão
  \noindent{\sffamily\bfseries\small\color{accent}\MakeUppercase{#1}}\hfill{\sffamily\small\color{muted}#2}\par
  \vspace{3pt}{\color{navy}\hrule height 1.1pt}\vspace{4pt}
  {\sffamily\small\color{muted}#3}\par}
\newcommand{\EntryTitle}[1]{%
  \vspace{16pt}{\raggedright\sffamily\bfseries\LARGE\color{navy}\hyphenpenalty=10000\exhyphenpenalty=10000 #1\par}}
\newcommand{\EntryAuthors}[2][\normalsize]{%
  \vspace{12pt}{\raggedright #1 #2\par}}
\newcommand{\Presenter}[1]{\underline{#1}}
\newcommand{\EntryAffils}[1][\small]{\vspace{8pt}\begingroup #1\color{muted}\raggedright}
\newcommand{\EndEntryAffils}{\par\endgroup}
\newcommand{\Affil}[2]{\hangindent=1.1em\hangafter=1\noindent\ifx\relax#1\relax\else\textsuperscript{#1}\,\fi #2\par}
\newcommand{\AbstractHeading}{%
  \vspace{16pt}{\sffamily\bfseries\small\color{accent}ABSTRACT}\par\vspace{2pt}{\color{rule}\hrule height 0.5pt}\vspace{8pt}}
\newenvironment{EntryAbstract}[1][\normalsize]
  {\begingroup #1 \setlength{\parskip}{0.55em}\sloppy\noindent\ignorespaces}
  {\par\endgroup}
\newcommand{\EntryPage}{\clearpage}

\begin{document}
\sloppy
"""

CAPA = r"""
% ---------------------------------------------------------------- capa
\pagenumbering{roman}
\thispagestyle{empty}
\begin{titlepage}
\vspace*{3.2cm}
{\color{accent}\rule{3.2cm}{3pt}}\par\vspace{1.0cm}
{\raggedright\sffamily\bfseries\fontsize{40}{44}\selectfont\color{navy} __TITULO_TEX__\par}
\vspace{0.6cm}
{\raggedright\sffamily\fontsize{24}{28}\selectfont\color{accent} Book of Abstracts\par}
\vfill
__SUBTITULO_BLOCO__
\vspace{1.2cm}
\end{titlepage}

% ---------------------------------------------------------------- sobre
\clearpage
\thispagestyle{plain}
{\sffamily\bfseries\Huge\color{navy} About this book\par}
\vspace{14pt}
\noindent This book gathers the abstracts of the __N_INV_ORAL__invited talks, oral presentations and
posters of __TITULO_TEX__. The talks follow the programme order, day by day; the posters are grouped by
scientific theme and numbered as on the workshop website. Each contribution has its own page.

\medskip\noindent In the author lists, the numbers refer to the institutions listed below the authors, and the
presenting author is \Presenter{underlined}. Abstracts are reproduced as submitted by the authors.

\clearpage
\tableofcontents
\clearpage
\pagenumbering{arabic}
"""

FIM = "\n\\end{document}\n"


def _autores_tex(r):
    partes = []
    ap = r["apresentador"]
    achou = False
    for a in r["autores"]:
        nome = tex(a["n"])
        if ap and ga.mesma_pessoa(ap, a["n"]):
            nome = r"\Presenter{" + nome + "}"
            achou = True
        marcas = ",".join(a["m"])
        partes.append(nome + (r"\textsuperscript{" + marcas + "}" if marcas else ""))
    texto = ", ".join(partes)
    if ap and not achou:
        texto += r"\par\smallskip\noindent{\small\color{muted}Presenting author: \Presenter{" + tex(ap) + "}}"
    return texto


def render_entrada(e, n_pagina_id):
    r = e["r"]
    tipo = TIPOS[e["tipo"]]
    if e["tipo"] == "poster":
        kicker = f"{tipo} {e['rotulo']}"
        info = "On display every day"
        sessao = f"Poster session \u00b7 Theme: {e['sessao']}"
    else:
        kicker = tipo
        if e["dia"]:
            partes = [f"Day {e['dia']}"] + ([e["data"]] if e["data"] else [])
            partes.append(e["hora"] if e["hora"] else "Time to be announced")
            info = " \u00b7 ".join(partes)
        else:
            info = "Time to be announced"
        sessao = f"Session: {e['sessao']}" if e["sessao"] else "Session to be announced"
    nivel, grande = escolher_nivel(r)
    macro_ab, macro_au, macro_af = (NIVEIS[nivel][0][0], NIVEIS[nivel][1][0], NIVEIS[nivel][2][0])
    primeiro = r["autores"][0]["n"] if r["autores"] else ""
    et_al = r" \emph{et al.}" if len(r["autores"]) > 1 else ""
    rot_toc = tex(e["rotulo"], mover=True)
    toc = (r"\phantomsection\addcontentsline{toc}{subsection}{\protect\numberline{" + rot_toc + "}"
           + tex(r["titulo"], mover=True) + (" \u2014 " + tex(primeiro, mover=True) + et_al if primeiro else "") + "}")
    L = [
        "",
        f"% ---- entrada {n_pagina_id}: {r['titulo'][:70].replace(chr(10), ' ')}",
        r"\EntryPage",
        r"\setsession{" + tex(e["sessao"] or "", mover=True) + "}",
        toc,
        r"\label{e-" + str(n_pagina_id) + "}",
        r"\EntryKicker{" + tex(kicker) + "}{" + tex(info) + "}{" + tex(sessao) + "}",
        r"\EntryTitle{" + tex(r["titulo"]) + "}",
        r"\EntryAuthors[" + macro_au + "]{" + _autores_tex(r) + "}",
    ]
    if r["afiliacoes"]:
        L.append(r"\EntryAffils[" + macro_af + "]")
        for a in r["afiliacoes"]:
            L.append(r"\Affil{" + (tex(a["m"]) if a["m"] else r"\relax") + "}{" + tex(a["t"]) + "}")
        L.append(r"\EndEntryAffils")
    L.append(r"\AbstractHeading")
    L.append(r"\begin{EntryAbstract}[" + macro_ab + "]")
    if r["abstract"]:
        L.append("\n\n".join(tex(p) for p in r["abstract"]))
    else:
        L.append(r"\emph{Abstract not available.}")
    L.append(r"\end{EntryAbstract}")
    return "\n".join(L), grande


def montar_tex(d, titulo, subtitulo, fonte="TeX Gyre Pagella", fonte_sans="TeX Gyre Heros"):
    grupos = montar_entradas(d)
    corpo, grandes, k = [], [], 0
    for g in grupos:
        titulo_grupo = tex(g["titulo"], mover=True)
        primeiro = True
        for e in g["entradas"]:
            k += 1
            bloco, grande = render_entrada(e, k)
            e["id"] = k
            if primeiro:                                     # entrada do sumário para o grupo (sem página própria)
                bloco = bloco.replace(r"\phantomsection\addcontentsline{toc}{subsection}",
                                      r"\phantomsection\addcontentsline{toc}{section}{" + titulo_grupo + "}\n"
                                      r"\addcontentsline{toc}{subsection}", 1)
                primeiro = False
            if grande:
                grandes.append(e)
            corpo.append(bloco)
    n_poster = sum(1 for g in grupos for e in g["entradas"] if e["tipo"] == "poster")
    pre = (PREAMBULO.replace("__VERSAO__", VERSAO).replace("__TITULO__", tex(titulo))
           .replace("__FONTE_SANS__", fonte_sans).replace("__FONTE__", fonte))
    sub = (r"{\raggedright\sffamily\large\color{muted} " + tex(subtitulo) + r"\par}") if subtitulo else ""
    capa = (CAPA.replace("__TITULO_TEX__", tex(titulo)).replace("__SUBTITULO_BLOCO__", sub)
            .replace("__N_INV_ORAL__", ""))
    texto = pre + capa + "\n".join(corpo) + FIM
    return texto, grupos, grandes, n_poster


# ============================================================================
# Compilação (opcional) e conferência de "uma página por entrada"
# ============================================================================
def compilar(tex_path, n_entradas):
    d = tex_path.parent
    motor = next((m for m in ("xelatex", "lualatex", "pdflatex") if shutil.which(m)), None)
    if not motor:
        return None, "Não achei xelatex/lualatex/pdflatex no PATH; compile o .tex no Overleaf ou instale o TeX Live/MiKTeX."
    cmd = [motor, "-interaction=nonstopmode", "-halt-on-error", tex_path.name]
    for _ in range(2):
        p = subprocess.run(cmd, cwd=str(d), capture_output=True, text=True, errors="replace")
        if p.returncode != 0:
            fim = "\n".join(p.stdout.splitlines()[-25:])
            return None, f"{motor} falhou. Final do log:\n{fim}"
    pdf = tex_path.with_suffix(".pdf")
    aux = tex_path.with_suffix(".aux").read_text(encoding="utf-8", errors="replace")
    paginas = {int(m.group(1)): int(m.group(2))
               for m in re.finditer(r"\\newlabel\{e-(\d+)\}\{\{[^}]*\}\{(\d+)\}", aux)}
    log = tex_path.with_suffix(".log").read_text(encoding="utf-8", errors="replace")
    m = re.search(r"Output written on .*\((\d+) pages?", log)
    total = int(m.group(1)) if m else None
    longas = []
    ordenadas = sorted(paginas)
    for i, k in enumerate(ordenadas):
        prox = paginas.get(ordenadas[i + 1]) if i + 1 < len(ordenadas) else None
        if prox is not None and prox - paginas[k] != 1:
            longas.append(k)
    return {"pdf": pdf, "motor": motor, "paginas": total, "longas": longas, "pag_entrada": paginas}, None


def main():
    ap = argparse.ArgumentParser(description="Livro-resumo em LaTeX (uma página por entrada) a partir da planilha do programa.")
    ap.add_argument("planilha", help="arquivo .xlsx (o mesmo do gerar_agenda.py)")
    ap.add_argument("--pagina", default="index.html", help="página do site, para dias/temas/horários (padrão: index.html)")
    ap.add_argument("--saida", default="livro_resumo.tex", help="arquivo .tex de saída (padrão: livro_resumo.tex)")
    ap.add_argument("--aba", default=None, help="aba a ler; para várias, separe por vírgula")
    ap.add_argument("--titulo", default=None, help="título da capa (padrão: o <title> da página)")
    ap.add_argument("--subtitulo", default=None, help="linha da capa com datas e local (padrão: tirada da página)")
    ap.add_argument("--fonte", default="TeX Gyre Pagella",
                    help="fonte do texto (XeLaTeX/LuaLaTeX); troque por uma de cobertura maior, ex.: \"DejaVu Serif\" "
                         "ou \"Noto Serif\", se houver nomes em cirílico, chinês etc. (padrão: TeX Gyre Pagella)")
    ap.add_argument("--fonte-sans", default="TeX Gyre Heros", help="fonte dos títulos (padrão: TeX Gyre Heros)")
    ap.add_argument("--ordem-das-linhas", action="store_true",
                    help="quem não tem horário/ordem entra na ordem da planilha, nos horários livres do dia")
    ap.add_argument("--estrito", action="store_true", help="se houver problema na planilha, não escreve nada")
    ap.add_argument("--pdf", action="store_true", help="compila o PDF (precisa de LaTeX) e confere uma página por entrada")
    args = ap.parse_args()

    for fluxo in (sys.stdout, sys.stderr):
        try:
            fluxo.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    print(f"gerar_livro.py \u00b7 versão {VERSAO}")

    try:
        d = carregar(args)
    except ga.ErroPlanilha as e:
        sys.exit(f"\nERRO: {e}\n")
    if d["erros"]:
        print(f"\nA planilha tem {len(d['erros'])} problema(s). Nada foi escrito:\n")
        for e in d["erros"]:
            print("  \u2717", e)
        sys.exit(1)

    if not any(g_["entradas"] for g_ in montar_entradas(d)):
        sys.exit("\nA planilha não tem nenhuma apresentação válida para o livro; nada foi escrito.\n")
    tit_pag, sub_pag = ler_cabecalho_da_pagina(d["texto"], d["meta"])
    titulo = args.titulo or tit_pag
    subtitulo = args.subtitulo if args.subtitulo is not None else sub_pag
    texto, grupos, grandes, n_poster = montar_tex(d, titulo, subtitulo, args.fonte, args.fonte_sans)
    total = sum(len(g["entradas"]) for g in grupos)

    saida = Path(args.saida)
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(texto, encoding="utf-8", newline="\n")

    print(f"\nPlanilha: {args.planilha}  ({len(d['registros'])} apresentações lidas"
          + (f"; {d['ignoradas']} inscrições sem apresentação foram ignoradas" if d["ignoradas"] else "") + ")")
    print(f"Livro:    {total} entradas \u2192 {total} páginas ({total - n_poster} apresentações do programa + {n_poster} pôsteres)")
    for g in grupos:
        print(f"          {g['titulo']}: {len(g['entradas'])}")
    sem_h = [e for g in grupos for e in g["entradas"] if e["tipo"] != "poster" and not e["hora"]]
    if sem_h:
        print(f"\nSem horário ainda: {len(sem_h)} apresentação(ões) entraram no fim do dia da sua sessão, "
              "como \"Time to be announced\".")
    sem_abs = [e for g in grupos for e in g["entradas"] if not e["r"]["abstract"]]
    if sem_abs:
        print(f"\nSem abstract ({len(sem_abs)}): a página sai com \"Abstract not available.\":")
        for e in sem_abs:
            print("  \u00b7", ga.rotulo(e["r"]), "\u2014", e["r"]["titulo"][:70])
    if grandes:
        print(f"\nABSTRACT MUITO LONGO ({len(grandes)}): mesmo com a letra menor pode passar de uma página:")
        for e in grandes:
            print("  \u00b7", ga.rotulo(e["r"]), "\u2014", e["r"]["titulo"][:70])
    todos = [t for g in grupos for e in g["entradas"] for t in
             [e["r"]["titulo"], e["r"]["apresentador"], *e["r"]["abstract"], *[a["n"] for a in e["r"]["autores"]],
              *[a["t"] for a in e["r"]["afiliacoes"]]]]
    susp = caracteres_suspeitos(todos)
    if susp:
        lista = ", ".join(f"{c} (U+{ord(c):04X}, {n}x)" for c, n in sorted(susp.items(), key=lambda x: -x[1])[:12])
        print(f"\nCaracteres que podem faltar na fonte (use XeLaTeX/LuaLaTeX e confira no PDF): {lista}")
    if d["excluidos"]:
        print(f"\nFICARAM DE FORA do livro ({len(d['excluidos'])}) \u2014 corrija na planilha e rode de novo:")
        for e in d["excluidos"]:
            print("  \u2717", e)
    if d["avisos"]:
        print(f"\nConfira na planilha ({len(d['avisos'])}):")
        for a in d["avisos"]:
            print("  !", a)
    print(f"\nArquivo escrito em: {saida}")

    if args.pdf:
        res, erro = compilar(saida, total)
        if erro:
            print("\n" + erro)
            sys.exit(2)
        print(f"PDF:      {res['pdf']}  ({res['paginas']} páginas, compilado com {res['motor']})")
        if res["longas"]:
            print("\nATENÇÃO: estas entradas ocuparam mais de uma página (reduza o abstract na planilha):")
            ent = {e["id"]: e for g in grupos for e in g["entradas"]}
            for k in res["longas"]:
                print("  \u00b7", ga.rotulo(ent[k]["r"]), "\u2014", ent[k]["r"]["titulo"][:70])
        else:
            print("Conferido: cada entrada ocupa exatamente uma página.")


if __name__ == "__main__":
    main()
