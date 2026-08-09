# -*- coding: utf-8 -*-
"""Conversor de markdown simples para blocos estruturados.

Usado para renderizar a narrativa da IA nos relatórios (Word/PDF/PPTX/XLSX),
da mesma forma que o front-end faz com mdToHtml().
"""

from __future__ import annotations

import re


def parse_blocks(md: str):
    """Converte markdown em uma lista de blocos.

    Cada bloco é um dict com:
      - tipo: 'h1'..'h4' | 'p' | 'ul' | 'ol' | 'table' | 'blockquote' | 'hr'
      - Para 'table': 'linhas' (lista de listas de str)
      - Para 'ul'/'ol': 'itens' (lista de str)
      - Para demais: 'texto' (str com marcações %B%..%B% e %I%..%I%)
    """
    if not md:
        return []
    blocos = []
    lista_atual = None
    tabela_atual = None
    in_code = False
    code_buf = []

    linhas = md.split("\n")
    i = 0
    while i < len(linhas):
        ln = linhas[i].rstrip()
        original = linhas[i]
        stripped = ln.strip()

        # fenced code block
        if stripped.startswith("```"):
            if not in_code:
                in_code = True
                code_buf = []
                i += 1
                continue
            else:
                in_code = False
                if code_buf:
                    blocos.append({"tipo": "p", "texto": _marcar("\n".join(code_buf))})
                i += 1
                continue
        if in_code:
            code_buf.append(original)
            i += 1
            continue

        if not stripped:
            lista_atual = None
            tabela_atual = None
            i += 1
            continue

        # título
        m = re.match(r"^(#{1,4})\s+(.+)$", stripped)
        if m:
            lista_atual = None
            tabela_atual = None
            nivel = len(m.group(1))
            blocos.append({"tipo": f"h{nivel + 1}", "texto": _marcar(m.group(2))})
            i += 1
            continue

        # tabela (linha contendo |)
        if "|" in ln:
            cells = [c.strip() for c in ln.split("|")]
            if cells and cells[0] == "":
                cells = cells[1:]
            if cells and cells[-1] == "":
                cells = cells[:-1]
            # separador de cabeçalho (---)
            if all(re.match(r"^:?-{2,}:?$", c) for c in cells):
                i += 1
                continue
            if tabela_atual is None:
                tabela_atual = []
                blocos.append({"tipo": "table", "linhas": tabela_atual})
            tabela_atual.append([_marcar(c) for c in cells])
            i += 1
            continue
        else:
            tabela_atual = None

        # lista
        m = re.match(r"^\s*([-*+]|\d+\.)\s+(.+)$", ln)
        if m:
            tipo = "ul" if m.group(1) in ("-", "*", "+") else "ol"
            if lista_atual is None or lista_atual["tipo"] != tipo:
                lista_atual = {"tipo": tipo, "itens": []}
                blocos.append(lista_atual)
            lista_atual["itens"].append(_marcar(m.group(2)))
            i += 1
            continue
        lista_atual = None

        # blockquote
        m = re.match(r"^>\s?(.*)$", stripped)
        if m:
            blocos.append({"tipo": "blockquote", "texto": _marcar(m.group(1))})
            i += 1
            continue

        # hr
        if re.match(r"^-{3,}$", stripped):
            blocos.append({"tipo": "hr"})
            i += 1
            continue

        # parágrafo
        blocos.append({"tipo": "p", "texto": _marcar(ln)})
        i += 1

    return blocos


def _marcar(texto: str) -> str:
    """Troca ** e * por marcadores %B% %I% (para docx/pptx)."""
    s = str(texto)
    # negrito primeiro para não colidir com itálico simples
    s = re.sub(r"\*\*(.+?)\*\*", r"%B%\1%B%", s, flags=re.S)
    s = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"%I%\1%I%", s)
    s = s.replace("`", "")
    return s


def runs_do_texto(texto_marcado: str):
    """Divide texto marcado (%B%...%B%, %I%...%I%) em runs (texto, bold, italic)."""
    if not texto_marcado:
        return [("", False, False)]
    runs = []
    buf = []
    bold = False
    italic = False
    i = 0
    s = texto_marcado
    while i < len(s):
        if s.startswith("%B%", i):
            if buf:
                runs.append(("".join(buf), bold, italic))
                buf = []
            bold = not bold
            i += 3
            continue
        if s.startswith("%I%", i):
            if buf:
                runs.append(("".join(buf), bold, italic))
                buf = []
            italic = not italic
            i += 3
            continue
        buf.append(s[i])
        i += 1
    if buf:
        runs.append(("".join(buf), bold, italic))
    return runs


def runs_para_reportlab(texto_marcado: str) -> str:
    """Converte texto marcado em mini-markup para reportlab (<b>/<i>)."""
    if not texto_marcado:
        return ""
    from xml.sax.saxutils import escape
    s = escape(texto_marcado)
    s = re.sub(r"%B%(.+?)%B%", r"<b>\1</b>", s, flags=re.S)
    s = re.sub(r"%I%(.+?)%I%", r"<i>\1</i>", s, flags=re.S)
    return s
