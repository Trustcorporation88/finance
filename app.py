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
    """Lê os dados, calcula o raio-x e consulta a IA.

    Se o arquivo for uma planilha complexa (múltiplas abas, modelo, DRE...),
    usa a leitura avançada e devolve uma análise estrutural + resposta da IA.
    """
    if texto:
        df = analise.ler_texto(texto)
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

    # arquivo
    try:
        df = analise.ler_dataframe(arquivo_bytes, nome_arquivo)
    except ValueError:
        # planilha complexa -> leitura avançada
        return _processar_planilha_completa(arquivo_bytes, nome_arquivo, intencao)

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


def _processar_planilha_completa(arquivo_bytes, nome_arquivo, intencao=""):
    """Lê a planilha inteira e devolve uma análise estrutural com resposta da IA."""
    info = analise.ler_planilha_completa(arquivo_bytes, nome_arquivo)
    estrutura = info["resumo_texto"]

    pergunta = intencao.strip() if intencao and intencao.strip() else (
        "Faça um resumo executivo desta planilha: o que ela contém, os principais "
        "números por aba, pontos de atenção e recomendações práticas."
    )
    try:
        narrativa = deepseek_client.perguntar_planilha(estrutura, pergunta)
        erro_ia = None
    except Exception as e:
        narrativa = None
        erro_ia = str(e)

    resultado = {
        "tipo_analise": "planilha_completa",
        "total_entradas": 0,
        "total_saidas": 0,
        "saldo": 0,
        "margem": None,
        "gastos_categorias": {},
        "entradas_categorias": {},
        "por_mes": None,
        "n_entradas": 0,
        "n_saidas": 0,
        "n_transacoes": info["n_abas"],
        "alertas": [{
            "nivel": "medio",
            "titulo": "Planilha com múltiplas abas",
            "detalhe": f"Esta planilha tem {info['n_abas']} abas e foi analisada como um modelo completo (não um extrato simples). As abas são: {', '.join(a['nome'] for a in info['abas'][:8])}.",
        }],
        "divergencias": [],
        "conferencia": {
            "lidos": info["n_abas"],
            "entradas": 0,
            "saidas": 0,
            "sem_data": 0,
            "sem_descricao": 0,
            "observacao": "Planilha completa com múltiplas abas — análise estrutural via IA.",
        },
        "narrativa_ia": narrativa,
        "erro_ia": erro_ia,
        "resumo": estrutura,
        "graf": {},
        "planilha_info": {
            "n_abas": info["n_abas"],
            "abas": info["abas"],  # dicts completos (nome, cabecalhos, amostra)
            "estrutura": estrutura,
        },
    }
    return None, resultado


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
        resultado.setdefault("nome_arquivo", nome_arquivo or "texto-colado.txt")
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
    estrutura = data.get("estrutura") or ""
    historico = data.get("historico") or []
    if not pergunta:
        return jsonify({"erro": "Digite uma pergunta."}), 400
    try:
        if estrutura:
            resposta = deepseek_client.perguntar_planilha_historico(historico, estrutura, pergunta)
        elif historico:
            resposta = deepseek_client.perguntar_historico(historico, pergunta)
        else:
            resposta = deepseek_client.perguntar(resumo, pergunta)
        return jsonify({"resposta": resposta})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/analisar/aba", methods=["POST"])
def analisar_aba():
    """Analisa uma aba específica de uma planilha complexa."""
    try:
        arquivo = request.files.get("arquivo")
        aba = (request.form.get("aba") or "").strip()
        intencao = request.form.get("intencao", "")
        if not arquivo or not arquivo.filename:
            return jsonify({"erro": "Envie o arquivo."}), 400
        if not _arquivo_permitido(arquivo.filename):
            return jsonify({"erro": "Formato não suportado."}), 400
        arquivo_bytes = arquivo.read()
        nome_arquivo = secure_filename(arquivo.filename)

        info = analise.ler_aba_especifica(arquivo_bytes, nome_arquivo, aba)
        estrutura = info["resumo_texto"]
        pergunta = intencao.strip() if intencao.strip() else (
            f"Analise a aba '{aba}' em detalhe: o que ela contém, principais números, pontos de atenção e recomendações."
        )
        try:
            narrativa = deepseek_client.perguntar_planilha(estrutura, pergunta)
            erro_ia = None
        except Exception as e:
            narrativa = None
            erro_ia = str(e)

        resultado = {
            "tipo_analise": "aba_especifica",
            "nome_arquivo": nome_arquivo,
            "aba_analisada": aba,
            "alertas": [],
            "divergencias": [],
            "conferencia": {"observacao": f"Aba '{aba}' analisada em detalhe."},
            "narrativa_ia": narrativa,
            "erro_ia": erro_ia,
            "resumo": estrutura,
            "graf": {},
            "planilha_info": {"n_abas": 1, "abas": [{"nome": aba}], "estrutura": estrutura},
        }
        return jsonify({"resultado": resultado, "graficos": {}})
    except ValueError as e:
        return jsonify({"erro": str(e)}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": f"Erro ao analisar aba: {e}"}), 500


@app.route("/comparar", methods=["POST"])
def comparar():
    """Compara duas planilhas e aponta diferenças."""
    try:
        arq1 = request.files.get("arquivo1")
        arq2 = request.files.get("arquivo2")
        if not arq1 or not arq1.filename or not arq2 or not arq2.filename:
            return jsonify({"erro": "Envie as duas planilhas."}), 400
        n1 = secure_filename(arq1.filename)
        n2 = secure_filename(arq2.filename)
        info = analise.comparar_planilhas(arq1.read(), n1, arq2.read(), n2)
        try:
            narrativa = deepseek_client.comparar_planilhas(info["resumo_texto"])
            erro_ia = None
        except Exception as e:
            narrativa = None
            erro_ia = str(e)
        resultado = {
            "tipo_analise": "comparacao",
            "nome_arquivo": f"{n1} vs {n2}",
            "alertas": [],
            "divergencias": [],
            "conferencia": {"observacao": "Comparação de duas planilhas."},
            "narrativa_ia": narrativa,
            "erro_ia": erro_ia,
            "resumo": info["resumo_texto"],
            "graf": {},
            "comparacao": {"planilha1": info["planilha1"], "planilha2": info["planilha2"]},
        }
        return jsonify({"resultado": resultado, "graficos": {}})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": f"Erro ao comparar: {e}"}), 500


@app.route("/previsao", methods=["POST"])
def previsao():
    """Gera previsão dos próximos 3 meses com base na análise atual."""
    data = request.get_json(force=True)
    resultado = data.get("resultado", {})
    try:
        contexto = analise.preparar_previsao(resultado)
        resposta = deepseek_client.prever_proximos_meses(contexto)
        return jsonify({"resposta": resposta, "contexto": contexto})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/relatorio/planilha/<formato>", methods=["POST"])
def relatorio_planilha(formato):
    """Downloads para planilhas completas (multi-abas). Formatos: pptx, word, pdf, xlsx."""
    data = request.get_json(force=True)
    resultado = data.get("resultado", {})
    info = data.get("planilha_info") or resultado.get("planilha_info", {})
    nome_base = str(resultado.get("nome_arquivo", "planilha")).rsplit(".", 1)[0][:50] or "planilha"
    try:
        if formato == "pptx":
            buf = relatorio.gerar_pptx_planilha(resultado, info)
            return send_file(buf,
                             mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                             as_attachment=True, download_name=f"{nome_base}-apresentacao.pptx")
        if formato == "word":
            buf = relatorio.gerar_word_planilha(resultado, info)
            return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                             as_attachment=True, download_name=f"{nome_base}-relatorio.docx")
        if formato == "pdf":
            buf = relatorio.gerar_pdf_planilha(resultado, info)
            return send_file(buf, mimetype="application/pdf", as_attachment=True, download_name=f"{nome_base}-relatorio.pdf")
        if formato == "xlsx":
            buf = relatorio.gerar_xlsx_processado(resultado, info)
            return send_file(buf,
                             mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             as_attachment=True, download_name=f"{nome_base}-processada.xlsx")
        return jsonify({"erro": "Formato não suportado."}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": f"Erro ao gerar: {e}"}), 500


# ---------- logs de uso (simples, em memória) ----------
import collections
import time as _time

USO = collections.Counter()          # contagem por tipo de ação
USO_IP = collections.Counter()       # contagem por IP
USO_DETALHES = []                     # últimas ações


def _registrar_uso(acao, ip=""):
    USO[acao] += 1
    if ip:
        USO_IP[ip] += 1
    USO_DETALHES.append({"t": _time.time(), "acao": acao, "ip": ip})
    if len(USO_DETALHES) > 200:
        USO_DETALHES.pop(0)


@app.after_request
def _log_uso_geral(resp):
    try:
        _registrar_uso(f"HTTP {resp.status_code} {request.method} {request.path}", request.remote_addr or "")
    except Exception:
        pass
    return resp


@app.route("/api/uso", methods=["GET"])
def api_uso():
    """Painel simples de uso (contagem de acessos/ações)."""
    total = sum(USO.values())
    top_rotas = USO.most_common(15)
    top_ips = USO_IP.most_common(10)
    return jsonify({
        "total_requisicoes": total,
        "top_rotas": [{"rota": r, "count": c} for r, c in top_rotas],
        "top_ips": [{"ip": ip, "count": c} for ip, c in top_ips],
        "recentes": USO_DETALHES[-20:],
    })


if __name__ == "__main__":
    porta = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=porta, debug=False)
