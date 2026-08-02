# -*- coding: utf-8 -*-
"""Geração de relatórios em Word (.docx) e PDF.

- Word: python-docx, com os gráficos embutidos.
- PDF: reportlab (platypus), com os gráficos embutidos.
"""
from __future__ import annotations

import io
import base64

from docx import Document
from docx.shared import Pt, Inches, RGBColor

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.enums import TA_CENTER

AZUL = RGBColor(0x1F, 0x3B, 0x73)
VERDE = RGBColor(0x1E, 0x8E, 0x3E)
VERMELHO = RGBColor(0xC0, 0x39, 0x2B)
CINZA = RGBColor(0x55, 0x55, 0x55)
AMARELO = RGBColor(0xF3, 0x9C, 0x12)


def _b64_para_bytes(b64: str) -> bytes:
    return base64.b64decode(b64.split(",", 1)[1] if "," in b64 else b64)


# ==========================================================================
# WORD
# ==========================================================================

def gerar_word(resultado: dict, graf: dict) -> io.BytesIO:
    doc = Document()
    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(11)

    doc.add_heading("Raio-X do Caixa — CFO de Bolso", level=0)

    p = doc.add_paragraph()
    r = p.add_run("Números calculados em regime de caixa, direto dos seus dados.")
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = CINZA

    doc.add_heading("O essencial (em 1 olhada)", level=1)
    t = doc.add_table(rows=4, cols=2)
    t.style = "Light Grid Accent 1"
    linhas = [
        ("Entrou", f"R$ {resultado.get('total_entradas', 0):,.2f}"),
        ("Saiu", f"R$ {resultado.get('total_saidas', 0):,.2f}"),
        ("Sobrou / Faltou", f"R$ {resultado.get('saldo', 0):,.2f}"),
        ("Margem do período", f"{resultado.get('margem')}%" if resultado.get("margem") is not None else "não calculável"),
    ]
    for i, (k, v) in enumerate(linhas):
        t.cell(i, 0).text = k
        t.cell(i, 1).text = v
    _estilizar_tabela_docx(t)

    if graf.get("comparativo"):
        doc.add_picture(io.BytesIO(_b64_para_bytes(graf["comparativo"])), width=Inches(5.5))
        doc.add_paragraph()

    if resultado.get("gastos_categorias"):
        doc.add_heading("Pra onde foi o dinheiro", level=1)
        total_s = resultado.get("total_saidas", 1) or 1
        linhas_cat = [("Categoria", "Valor", "%")]
        for cat, val in resultado["gastos_categorias"].items():
            linhas_cat.append((cat, f"R$ {val:,.2f}", f"{val / total_s * 100:.1f}%"))
        t = doc.add_table(rows=len(linhas_cat), cols=3)
        t.style = "Light Grid Accent 1"
        for i, linha in enumerate(linhas_cat):
            for j, v in enumerate(linha):
                t.cell(i, j).text = v
        _estilizar_tabela_docx(t)
        if graf.get("pizza_gastos"):
            doc.add_picture(io.BytesIO(_b64_para_bytes(graf["pizza_gastos"])), width=Inches(5))
            doc.add_paragraph()

    if graf.get("fluxo_mensal"):
        doc.add_heading("Fluxo de caixa por mês", level=1)
        doc.add_picture(io.BytesIO(_b64_para_bytes(graf["fluxo_mensal"])), width=Inches(5.5))
        doc.add_paragraph()

    if resultado.get("alertas"):
        doc.add_heading("Alertas", level=1)
        for a in resultado["alertas"]:
            cor = VERMELHO if a["nivel"] == "alto" else AMARELO
            p = doc.add_paragraph()
            r = p.add_run(f"[{a['nivel'].upper()}] {a['titulo']} — ")
            r.bold = True
            r.font.color.rgb = cor
            p.add_run(a["detalhe"])

    if resultado.get("narrativa_ia"):
        doc.add_heading("Análise da IA (DeepSeek)", level=1)
        doc.add_paragraph(resultado["narrativa_ia"])
    elif resultado.get("erro_ia"):
        doc.add_heading("IA indisponível", level=1)
        p = doc.add_paragraph()
        r = p.add_run("Não foi possível consultar a IA: ")
        r.bold = True
        p.add_run(resultado["erro_ia"])

    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run("Lembrete: é gestão de caixa gerencial. Para imposto/nota fiscal, fale com seu contador.")
    r.italic = True
    r.font.size = Pt(8)
    r.font.color.rgb = CINZA

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def _estilizar_tabela_docx(t):
    for i, row in enumerate(t.rows):
        for cell in row.cells:
            for pr in cell.paragraphs:
                for r in pr.runs:
                    r.font.size = Pt(10)
                    if i == 0:
                        r.bold = True


AMARELO_RL = HexColor("#F39C12")


# ==========================================================================
# PDF
# ==========================================================================

def gerar_pdf(resultado: dict, graf: dict) -> io.BytesIO:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=15 * mm, bottomMargin=15 * mm)

    estilos = getSampleStyleSheet()
    titulo = ParagraphStyle("Titulo", parent=estilos["Title"], fontSize=20,
                            textColor=HexColor("#1F3B73"), spaceAfter=4)
    h2 = ParagraphStyle("H2", parent=estilos["Heading2"], fontSize=14,
                        textColor=HexColor("#1F3B73"), spaceBefore=12, spaceAfter=4)
    corpo = ParagraphStyle("Corpo", parent=estilos["BodyText"], fontSize=10, leading=14)
    nota = ParagraphStyle("Nota", parent=estilos["BodyText"], fontSize=8,
                          textColor=HexColor("#555555"), italic=True)

    elementos = []
    elementos.append(Paragraph("Raio-X do Caixa — CFO de Bolso", titulo))
    elementos.append(Paragraph("Números calculados em regime de caixa, direto dos seus dados.", nota))
    elementos.append(Spacer(1, 8))

    elementos.append(Paragraph("O essencial (em 1 olhada)", h2))
    linhas = [
        ["Entrou", f"R$ {resultado.get('total_entradas', 0):,.2f}"],
        ["Saiu", f"R$ {resultado.get('total_saidas', 0):,.2f}"],
        ["Sobrou / Faltou", f"R$ {resultado.get('saldo', 0):,.2f}"],
        ["Margem do período", f"{resultado.get('margem')}%" if resultado.get("margem") is not None else "não calculável"],
    ]
    tbl = Table(linhas, colWidths=[80 * mm, 90 * mm])
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#BBBBBB")),
        ("BACKGROUND", (0, 0), (0, -1), HexColor("#E8EDF4")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elementos.append(tbl)
    elementos.append(Spacer(1, 8))

    if graf.get("comparativo"):
        img = Image(io.BytesIO(_b64_para_bytes(graf["comparativo"])), width=150 * mm, height=90 * mm)
        img.hAlign = "CENTER"
        elementos.append(img)

    if resultado.get("gastos_categorias"):
        elementos.append(Paragraph("Pra onde foi o dinheiro", h2))
        total_s = resultado.get("total_saidas", 1) or 1
        linhas_cat = [["Categoria", "Valor", "%"]]
        for cat, val in resultado["gastos_categorias"].items():
            linhas_cat.append([cat, f"R$ {val:,.2f}", f"{val / total_s * 100:.1f}%"])
        tbl = Table(linhas_cat, colWidths=[90 * mm, 45 * mm, 35 * mm])
        tbl.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("GRID", (0, 0), (-1, -1), 0.4, HexColor("#CCCCCC")),
            ("BACKGROUND", (0, 0), (-1, 0), HexColor("#1F3B73")),
            ("TEXTCOLOR", (0, 0), (-1, 0), HexColor("#FFFFFF")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        elementos.append(tbl)
        elementos.append(Spacer(1, 8))
        if graf.get("pizza_gastos"):
            img = Image(io.BytesIO(_b64_para_bytes(graf["pizza_gastos"])), width=120 * mm, height=90 * mm)
            img.hAlign = "CENTER"
            elementos.append(img)

    if graf.get("fluxo_mensal"):
        elementos.append(Paragraph("Fluxo de caixa por mês", h2))
        img = Image(io.BytesIO(_b64_para_bytes(graf["fluxo_mensal"])), width=150 * mm, height=90 * mm)
        img.hAlign = "CENTER"
        elementos.append(img)

    if resultado.get("alertas"):
        elementos.append(Paragraph("Alertas", h2))
        for a in resultado["alertas"]:
            cor = HexColor("#C0392B") if a["nivel"] == "alto" else AMARELO_RL
            elementos.append(Paragraph(f"<font color='{cor.hexval()}'><b>[{a['nivel'].upper()}] {a['titulo']}</b></font> — {a['detalhe']}", corpo))

    if resultado.get("narrativa_ia"):
        elementos.append(Paragraph("Análise da IA (DeepSeek)", h2))
        for bloco in resultado["narrativa_ia"].split("\n\n"):
            elementos.append(Paragraph(bloco.replace("\n", "<br/>"), corpo))
            elementos.append(Spacer(1, 4))

    elementos.append(Spacer(1, 12))
    elementos.append(Paragraph("Lembrete: é gestão de caixa gerencial. Para imposto/nota fiscal, fale com seu contador.", nota))

    doc.build(elementos)
    buf.seek(0)
    return buf
