# -*- coding: utf-8 -*-
"""Geração de gráficos em PNG (BytesIO) para o relatório e a página."""
from __future__ import annotations

import io
import base64

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.unicode_minus"] = False

AZUL = "#1F3B73"
VERDE = "#1E8E3E"
VERMELHO = "#C0392B"
AMARELO = "#F39C12"
CINZA = "#555555"


def _png_b64(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return "data:image/png;base64," + base64.b64encode(buf.read()).decode("ascii")


def grafico_comparativo(resultado: dict) -> str:
    """Entradas x Saídas x Saldo (agregado)."""
    fig, ax = plt.subplots(figsize=(7, 4.5))
    labels = ["Entradas", "Saídas", "Saldo"]
    vals = [resultado["total_entradas"], resultado["total_saidas"], resultado["saldo"]]
    cores = [VERDE, VERMELHO, AZUL]
    b = ax.bar(labels, [v / 1000 for v in vals], color=cores)
    for bar, v in zip(b, vals):
        ax.annotate(f"R$ {v:,.0f}", (bar.get_x() + bar.get_width() / 2, bar.get_height()),
                    ha="center", va="bottom", fontsize=9)
    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_ylabel("R$ mil")
    ax.set_title("Entrou, saiu e sobrou (período)")
    return _png_b64(fig)


def grafico_pizza_gastos(resultado: dict) -> str:
    """Pizza das categorias de saída."""
    cats = resultado["gastos_categorias"]
    if not cats:
        return None
    cores = ["#C0392B", "#E67E22", "#F1C40F", "#2ECC71", "#3498DB", "#9B59B6",
             "#E74C3C", "#1ABC9C", "#95A5A6", "#34495E"]
    labels = list(cats.keys())
    vals = list(cats.values())
    fig, ax = plt.subplots(figsize=(7, 5))
    wedges, _, autotexts = ax.pie(vals, labels=None, autopct=lambda p: f"{p:.1f}%",
                                  startangle=90, colors=cores[:len(vals)],
                                  textprops={"fontsize": 9})
    ax.legend(wedges, [f"{l} — R$ {v:,.0f}" for l, v in zip(labels, vals)],
              loc="center left", bbox_to_anchor=(1, 0.5), fontsize=9)
    ax.set_title("Pra onde foi o dinheiro (saídas)")
    return _png_b64(fig)


def grafico_fluxo_mensal(resultado: dict) -> str:
    """Barras entradas/saídas + linha de saldo por mês."""
    por_mes = resultado["por_mes"]
    if not por_mes:
        return None
    meses = list(por_mes.keys())
    ent = [por_mes[m]["entradas"] for m in meses]
    sai = [por_mes[m]["saidas"] for m in meses]
    sal = [por_mes[m]["saldo"] for m in meses]
    fig, ax = plt.subplots(figsize=(8, 4.5))
    x = range(len(meses))
    bw = 0.28
    ax.bar([i - bw for i in x], [v / 1000 for v in ent], bw, label="Entradas", color=VERDE)
    ax.bar([i for i in x], [v / 1000 for v in sai], bw, label="Saídas", color=VERMELHO)
    ax2 = ax.twinx()
    ax2.plot(x, [v / 1000 for v in sal], "o-", color=AZUL, linewidth=2.5, label="Saldo")
    for i, v in enumerate(sal):
        ax2.annotate(f"R$ {v:,.0f}", (i, v / 1000), textcoords="offset points",
                     xytext=(0, 8), ha="center", fontsize=8)
    ax.set_xticks(list(x))
    ax.set_xticklabels(meses)
    ax.set_ylabel("R$ mil (entradas/saídas)")
    ax2.set_ylabel("R$ mil (saldo)")
    ax.axhline(0, color="black", linewidth=0.8)
    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=9)
    ax.set_title("Fluxo de caixa por mês")
    return _png_b64(fig)


def gerar_todos(resultado: dict) -> dict:
    """Gera todos os gráficos e devolve dict de base64."""
    return {
        "comparativo": grafico_comparativo(resultado),
        "pizza_gastos": grafico_pizza_gastos(resultado),
        "fluxo_mensal": grafico_fluxo_mensal(resultado),
    }
