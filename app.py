# -*- coding: utf-8 -*-
"""CFO de Bolso - app web.

Roda com:  python app.py
Publicação (Railway/Vercel/VPS): veja README.md
"""
from __future__ import annotations

import os
import io
import re
import json
import time
import time as _time
import hmac
import hashlib
import collections
import traceback

from flask import Flask, request, render_template, jsonify, send_file, session
from werkzeug.utils import secure_filename

import analise
import graficos
import deepseek_client
import relatorio
import supabase_client

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB
# Sessão: usa chave do ambiente (segura) ou gera uma estável em memória
app.secret_key = os.environ.get("APP_SECRET_KEY", "cfo-bolso-dev-secret-key-troque-em-producao")

# Auth simples: ADMIN_USER + ADMIN_PASSWORD via ambiente (opcional)
ADMIN_USER = os.environ.get("ADMIN_USER", "").strip()
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "").strip()
AUTH_ATIVO = bool(ADMIN_USER and ADMIN_PASSWORD)

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

# formatos aceitos para upload; XLSM (com macros) é bloqueado por segurança
ALLOWED = {"csv", "xlsx", "xls", "txt", "ofx", "qif"}
BLOCKED_EXTS = {"xlsm"}


def _arquivo_permitido(nome):
    ext = nome.rsplit(".", 1)[-1].lower() if "." in nome else ""
    if ext in BLOCKED_EXTS:
        return False
    return ext in ALLOWED


def _usuario_atual():
    """Devolve o nome do usuário logado (ou 'anonimo' se auth desativado)."""
    if AUTH_ATIVO:
        return session.get("usuario", "")
    return "anonimo"


def _exigir_auth():
    """Retorna erro se auth ativo e usuário não logado."""
    if AUTH_ATIVO and not _usuario_atual():
        return jsonify({"erro": "Faça login para continuar.", "precisa_login": True}), 401
    return None


# ---------- histórico por usuário (Supabase se ativo; senão memória) ----------
HISTORICOS = collections.defaultdict(list)   # fallback em memória


def _salvar_historico(usuario, item):
    # tenta Supabase primeiro
    if supabase_client.ativo():
        ok = supabase_client.salvar_analise(usuario, item.get("nome", "análise"),
                                            item.get("tipo", "extrato"), item.get("resultado", {}))
        if ok:
            return
    # fallback em memória
    HISTORICOS[usuario].insert(0, item)
    if len(HISTORICOS[usuario]) > 30:
        HISTORICOS[usuario] = HISTORICOS[usuario][:30]


def _listar_historico(usuario):
    if supabase_client.ativo():
        itens = supabase_client.listar_analises(usuario)
        return [{
            "id": i.get("id"),
            "nome": i.get("nome", "análise"),
            "tipo": i.get("tipo", "extrato"),
            "data": (i.get("criado_em") or "")[:16].replace("T", " "),
        } for i in itens]
    return HISTORICOS.get(usuario, [])


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
    ext = (nome_arquivo or "").lower().rsplit(".", 1)[-1]
    try:
        if ext == "ofx":
            df = analise.ler_ofx(arquivo_bytes.decode("utf-8", errors="replace"))
        elif ext == "qif":
            df = analise.ler_qif(arquivo_bytes.decode("utf-8", errors="replace"))
        else:
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
    # Detecção de pedido de estudo completo / análise geral profunda
    alvo = (intencao or "").lower()
    palavras_estudo = ["estudo", "completo", "análise geral", "analise geral", "geral", "completa",
                       "detalhado", "raiz", "entenda", "entender", "fundo", "profunda"]
    eh_estudo = any(p in alvo for p in palavras_estudo)
    try:
        if eh_estudo:
            narrativa = deepseek_client.estudo_completo(estrutura)
        else:
            narrativa = deepseek_client.perguntar_planilha(estrutura, pergunta)
        erro_ia = None
    except Exception as e:
        narrativa = None
        erro_ia = str(e)

    resultado = {
        "tipo_analise": "planilha_completa",
        "modo_estudo": eh_estudo,
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

        bloqueio = _exigir_auth()
        if bloqueio:
            return bloqueio
        df, resultado = _processar_dados(arquivo_bytes, nome_arquivo, texto, intencao)
        resultado.setdefault("nome_arquivo", nome_arquivo or "texto-colado.txt")
        # salva no histórico do usuário
        _salvar_historico(_usuario_atual(), {
            "id": int(time.time() * 1000),
            "nome": resultado.get("nome_arquivo", "análise"),
            "tipo": resultado.get("tipo_analise", "extrato"),
            "data": time.strftime("%d/%m/%Y %H:%M"),
            "resultado": resultado,
        })
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
            "planilha_info": {
                "n_abas": 1,
                "abas": info.get("abas") or [{"nome": aba}],
                "estrutura": estrutura,
            },
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




@app.route("/orcado-realizado", methods=["POST"])
def orcado_realizado():
    """Compara orçado x realizado (duas análises) e a IA aponta desvios."""
    data = request.get_json(force=True)
    orcado = data.get("orcado", {})
    realizado = data.get("realizado", {})
    if not orcado or not realizado:
        return jsonify({"erro": "Envie orçado e realizado."}), 400
    try:
        contexto = analise.preparar_orcado_realizado(orcado, realizado)
        resposta = deepseek_client.analisar_numeros(
            contexto + "\n\nResponda como auditor: aponte os maiores desvios (em R$ e %), o que estourou, o que ficou abaixo e recomendações."
        )
        return jsonify({"resposta": resposta, "contexto": contexto})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/alerta-caixa", methods=["POST"])
def alerta_caixa():
    """Analisa se o caixa projetado corre risco de ficar negativo."""
    data = request.get_json(force=True)
    resultado = data.get("resultado", {})
    saldo = resultado.get("saldo", 0)
    total_sai = resultado.get("total_saidas", 0)
    por_mes = resultado.get("por_mes") or {}
    alertas = []
    for mes, v in por_mes.items():
        if v.get("saldo", 0) < 0:
            alertas.append(f"{mes}: saldo {v['saldo']:,.2f} NEGATIVO")
    nivel = "verde"
    mensagem = f"Saldo do período: R$ {saldo:,.2f}. "
    if saldo < 0:
        nivel = "vermelho"
        mensagem += "O caixa está NEGATIVO."
    elif total_sai > 0 and saldo < total_sai * 0.15:
        nivel = "amarelo"
        mensagem += "O caixa está apertado (menos de 15% das saídas de reserva)."
    else:
        mensagem += "O caixa está saudável."
    if alertas:
        nivel = "vermelho"
        mensagem += " Meses com saldo negativo: " + "; ".join(alertas)
    return jsonify({"nivel": nivel, "mensagem": mensagem, "meses_negativos": alertas})

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


# ---------- autenticação ----------
@app.route("/auth/status", methods=["GET"])
def auth_status():
    return jsonify({
        "ativo": AUTH_ATIVO,
        "supabase": supabase_client.ativo(),
        "logado": bool(_usuario_atual()),
    })


@app.route("/auth/login", methods=["POST"])
def auth_login():
    data = request.get_json(force=True)
    user = (data.get("usuario") or "").strip()
    senha = (data.get("senha") or "").strip()
    # Supabase Auth tem prioridade se ativo
    if supabase_client.ativo():
        r = supabase_client.sign_in(user, senha)
        if r.get("ok"):
            session["usuario"] = user
            session["supabase_token"] = r.get("access_token", "")
            return jsonify({"ok": True, "logado": True, "usuario": user, "provedor": "supabase"})
        return jsonify({"ok": False, "erro": r.get("erro", "E-mail ou senha inválidos.")}), 401
    if not AUTH_ATIVO:
        return jsonify({"ok": True, "logado": True})
    # comparação segura (constant-time)
    ok_user = hmac.compare_digest(user, ADMIN_USER)
    ok_senha = hmac.compare_digest(senha, ADMIN_PASSWORD)
    if ok_user and ok_senha:
        session["usuario"] = user
        return jsonify({"ok": True, "logado": True, "usuario": user})
    return jsonify({"ok": False, "erro": "Usuário ou senha inválidos."}), 401


@app.route("/auth/signup", methods=["POST"])
def auth_signup():
    data = request.get_json(force=True)
    email = (data.get("email") or "").strip()
    senha = (data.get("senha") or "").strip()
    if not supabase_client.ativo():
        return jsonify({"ok": False, "erro": "Cadastro desativado (Supabase não configurado)."}), 400
    if not email or len(senha) < 6:
        return jsonify({"ok": False, "erro": "Informe e-mail e senha com pelo menos 6 caracteres."}), 400
    r = supabase_client.sign_up(email, senha)
    if not r.get("ok"):
        return jsonify({"ok": False, "erro": r.get("erro", "Falha no cadastro.")}), 400
    session["usuario"] = email
    return jsonify({"ok": True, "logado": True, "usuario": email})


@app.route("/auth/magic", methods=["POST"])
def auth_magic():
    data = request.get_json(force=True)
    email = (data.get("email") or "").strip()
    if not supabase_client.ativo():
        return jsonify({"ok": False, "erro": "Acesso por link desativado (Supabase não configurado)."}), 400
    r = supabase_client.sign_in_magic(email)
    if r.get("ok"):
        return jsonify({"ok": True, "mensagem": r.get("mensagem", "Link enviado.")})
    return jsonify({"ok": False, "erro": r.get("erro", "Falha ao enviar link.")}), 400


@app.route("/auth/logout", methods=["POST"])
def auth_logout():
    session.pop("usuario", None)
    session.pop("supabase_token", None)
    if supabase_client.ativo():
        try:
            supabase_client._get_client().auth.sign_out()
        except Exception:
            pass
    return jsonify({"ok": True})


@app.route("/historico", methods=["GET"])
def historico():
    bloqueio = _exigir_auth()
    if bloqueio:
        return bloqueio
    itens = _listar_historico(_usuario_atual())
    if supabase_client.ativo():
        # itens já vêm no formato resumido
        return jsonify({"itens": itens})
    resumo = [{
        "id": i["id"], "nome": i["nome"], "tipo": i["tipo"], "data": i["data"],
    } for i in itens]
    return jsonify({"itens": resumo})


@app.route("/historico/<int:hid>", methods=["GET"])
def historico_item(hid):
    bloqueio = _exigir_auth()
    if bloqueio:
        return bloqueio
    usuario = _usuario_atual()
    if supabase_client.ativo():
        item = supabase_client.buscar_analise(usuario, hid)
        if item:
            return jsonify({"resultado": item.get("resultado", {})})
        return jsonify({"erro": "Análise não encontrada."}), 404
    for i in HISTORICOS.get(usuario, []):
        if i["id"] == hid:
            return jsonify({"resultado": i["resultado"]})
    return jsonify({"erro": "Análise não encontrada."}), 404


@app.route("/historico/<int:hid>", methods=["DELETE"])
def historico_delete(hid):
    bloqueio = _exigir_auth()
    if bloqueio:
        return bloqueio
    usuario = _usuario_atual()
    if supabase_client.ativo():
        supabase_client.deletar_analise(usuario, hid)
        return jsonify({"ok": True})
    HISTORICOS[usuario] = [i for i in HISTORICOS.get(usuario, []) if i["id"] != hid]
    return jsonify({"ok": True})


# ---------- logs de uso (Supabase se ativo; senão em memória) ----------
USO = collections.Counter()          # contagem por tipo de ação
USO_IP = collections.Counter()       # contagem por IP
USO_DETALHES = []                     # últimas ações


def _registrar_uso(acao, ip=""):
    # Supabase primeiro
    if supabase_client.ativo():
        try:
            supabase_client.registrar_log(acao, ip)
        except Exception:
            pass
        return
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
    if supabase_client.ativo():
        dados = supabase_client.resumo_logs()
        return jsonify({
            "supabase": True,
            "top_rotas": [{"rota": r, "count": c} for r, c in dados.get("top_acoes", [])],
            "top_ips": [{"ip": ip, "count": c} for ip, c in dados.get("top_ips", [])],
            "recentes": dados.get("recentes", []),
        })
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
    # inicializa Supabase (cria tabelas e bucket) se configurado
    try:
        if supabase_client.ativo():
            supabase_client.garantir_bucket()
            supabase_client.inicializar_schema()
    except Exception:
        pass
    porta = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=porta, debug=False)
