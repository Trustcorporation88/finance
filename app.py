# -*- coding: utf-8 -*-
"""CFO de Bolso - app web.

Roda com:  python app.py
Publicação (Railway/Vercel/VPS): veja README.md
"""
from __future__ import annotations

import os
import io
import re
import traceback

from flask import Flask, request, render_template, jsonify, send_file
from werkzeug.utils import secure_filename

import analise
import graficos
import deepseek_client
import relatorio

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB

# Permite ser embutido (iframe) por origens confiáveis (ex.: excel.trustcorp.com.br)
# IMPORTANTE: NÃO usar X-Frame-Options (valor ALLOWALL é inválido e bloqueia no Chrome).
# Usar apenas Content-Security-Policy frame-ancestors, que é a forma correta e moderna.
@app.after_request
def _permitir_iframe(resp):
    resp.headers.pop("X-Frame-Options", None)
    resp.headers["Content-Security-Policy"] = (
        "frame-ancestors 'self' https://excel.trustcorp.com.br"
    )
    return resp

ALLOWED = {"csv", "xlsx", "xls", "txt"}


def _arquivo_permitido(nome):
    return "." in nome and nome.rsplit(".", 1)[1].lower() in ALLOWED


def _processar_dados(arquivo_bytes=None, nome_arquivo="", texto=None, intencao=""):
    """Lê os dados, calcula o raio-x e consulta a IA."""
    if texto:
        df = analise.ler_texto(texto)
    else:
        df = analise.ler_dataframe(arquivo_bytes, nome_arquivo)

    df = analise.categorizar(df)
    resultado = analise.calcular(df)

    resumo = analise.resumo_para_ia(resultado)
    if intencao and intencao.strip():
        resumo += (
            "\n\nOBJETIVO DO USUÁRIO (pedido feito antes de enviar os dados): "
            f"{intencao.strip()}\nResponda a análise priorizando esse pedido, mas mantendo o raio-x completo."
        )
    try:
        narrativa = deepseek_client.analisar_numeros(resumo)
    except Exception as e:
        narrativa = None
        erro_ia = str(e)
    else:
        erro_ia = None

    resultado["narrativa_ia"] = narrativa
    resultado["erro_ia"] = erro_ia
    resultado["resumo"] = resumo
    resultado["graf"] = graficos.gerar_todos(resultado)
    return df, resultado


@app.route("/")
def index():
    return render_template("index.html", has_key=deepseek_client.tem_chave())


@app.route("/analisar", methods=["POST"])
def analisar():
    try:
        texto = request.form.get("texto", "")
        intencao = request.form.get("intencao", "")
        arquivo = request.files.get("arquivo")
        arquivo_bytes = None
        nome_arquivo = ""
        if arquivo and arquivo.filename:
            if not _arquivo_permitido(arquivo.filename):
                return jsonify({"erro": f"Formato não suportado. Use: CSV, XLSX ou TXT."}), 400
            arquivo_bytes = arquivo.read()
            nome_arquivo = secure_filename(arquivo.filename)

        if not texto and not arquivo_bytes:
            return jsonify({"erro": "Envie um arquivo ou cole os dados em texto."}), 400

        df, resultado = _processar_dados(arquivo_bytes, nome_arquivo, texto, intencao)
        # remove gráficos pesados do JSON (são montados no HTML)
        graf = resultado.pop("graf", {})
        return jsonify({"resultado": resultado, "graficos": graf})
    except ValueError as e:
        return jsonify({"erro": str(e)}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": f"Erro ao analisar: {e}"}), 500


@app.route("/relatorio/word", methods=["POST"])
def relatorio_word():
    data = request.get_json(force=True)
    resultado = data.get("resultado", {})
    graf = data.get("graficos", {})
    bloqueio = _checar_bloqueio(resultado)
    if bloqueio:
        return jsonify({"erro": bloqueio}), 422
    buf = relatorio.gerar_word(resultado, graf)
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     as_attachment=True, download_name="raio-x-do-caixa.docx")


def _checar_bloqueio(resultado):
    """Bloqueia download se houver divergência de nível alto que o usuário não confirmou."""
    divergencias = resultado.get("divergencias", [])
    criticas = [d for d in divergencias if d.get("nivel") == "alto"]
    if not criticas:
        return None
    if resultado.get("divergencias_confirmadas"):
        return None
    titulos = "; ".join(d["titulo"] for d in criticas[:3])
    return (f"Há divergência(s) crítica(s) não confirmada(s): {titulos}. "
            f"Revise os dados e confirme a revisão para liberar o download.")


@app.route("/relatorio/pdf", methods=["POST"])
def relatorio_pdf():
    data = request.get_json(force=True)
    resultado = data.get("resultado", {})
    graf = data.get("graficos", {})
    bloqueio = _checar_bloqueio(resultado)
    if bloqueio:
        return jsonify({"erro": bloqueio}), 422
    buf = relatorio.gerar_pdf(resultado, graf)
    return send_file(buf, mimetype="application/pdf", as_attachment=True, download_name="raio-x-do-caixa.pdf")


@app.route("/relatorio/pptx", methods=["POST"])
def relatorio_pptx():
    data = request.get_json(force=True)
    resultado = data.get("resultado", {})
    graf = data.get("graficos", {})
    bloqueio = _checar_bloqueio(resultado)
    if bloqueio:
        return jsonify({"erro": bloqueio}), 422
    buf = relatorio.gerar_pptx(resultado, graf)
    return send_file(buf,
                     mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                     as_attachment=True, download_name="raio-x-do-caixa.pptx")


@app.route("/perguntar", methods=["POST"])
def perguntar():
    data = request.get_json(force=True)
    pergunta = (data.get("pergunta") or "").strip()
    resumo = data.get("resumo") or ""
    if not pergunta:
        return jsonify({"erro": "Digite uma pergunta."}), 400
    try:
        resposta = deepseek_client.perguntar(resumo, pergunta)
        return jsonify({"resposta": resposta})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


if __name__ == "__main__":
    porta = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=porta, debug=False)
