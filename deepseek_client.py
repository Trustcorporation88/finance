# -*- coding: utf-8 -*-
"""Cliente DeepSeek API - transforma os números calculados em texto CFO.

Usa a API da DeepSeek (compatível com OpenAI). A chave vem de
DEEPSEEK_API_KEY no ambiente / Railway / arquivo .env.
"""
from __future__ import annotations

import os
import json

import requests

API_URL = "https://api.deepseek.com/chat/completions"
MODELO_PADRAO = "deepseek-chat"

SYSTEM_PROMPT = """Você é um CFO fracional / analista financeiro sênior que atende pequenos negócios (dono de PME).

Sua missão: pegar os números que o usuário vai te dar (já calculados, em regime de caixa) e devolver um RAIO-X claro, sem jargão contábil. O dono quer responder 3 perguntas:
(1) quanto entrou e saiu? (2) sobrou ou faltou, e por quê? (3) o que eu faço com isso?

REGRAS DE OURO (inegociáveis):
1. NUNCA invente número. Use SOMENTE os números fornecidos. Se um dado faltar, sinalize — não chute.
2. Todo número vem com 'e daí?' — o que significa e o que fazer.
3. Linguagem de dono, não de contador: "sobrou R$ 4 mil" em vez de "resultado líquido positivo".
4. Use os alertas automáticos fornecidos e acrescente padrões que você notar (concentração, gasto crescente, custo fixo alto).
5. Entregue SEMPRE nesta estrutura:

## Raio-X do Caixa — [período]

### O essencial (em 1 olhada)
Entrou: R$ X · Saiu: R$ Y · Sobrou/Faltou: R$ Z
Margem do período: __%

### Pra onde foi o dinheiro
[top categorias de gasto com %]

### 3 alertas
🔴/🟡 [o que merece atenção, com o número que comprova]

### Recomendações (o que eu faria)
[1 a 3 ações práticas e priorizadas, cada uma com impacto em R$ quando possível]

Responda em português do Brasil. Seja direto e prático. Não faça preâmbulos."""


def tem_chave() -> bool:
    """Verifica se há uma chave configurada (sem chamar a API)."""
    return bool(os.environ.get("DEEPSEEK_API_KEY", "").strip())


def analisar_numeros(resumo_texto: str, api_key: str = None, modelo: str = None) -> str:
    """Envia os números calculados para a DeepSeek e devolve o texto do raio-x.

    Retorna string com o markdown do relatório. Se não houver chave, devolve
    um texto explicando que a análise por IA está desativada.
    """
    api_key = api_key or os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return None

    modelo = modelo or os.environ.get("DEEPSEEK_MODEL", MODELO_PADRAO)

    payload = {
        "model": modelo,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": resumo_texto},
        ],
        "temperature": 0.4,
        "max_tokens": 1500,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    try:
        resp = requests.post(API_URL, headers=headers, json=payload, timeout=90)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as e:
        corpo = e.response.text[:400] if e.response is not None else str(e)
        raise RuntimeError(f"Erro na API DeepSeek (HTTP {e.response.status_code if e.response else '?'}): {corpo}")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha de conexão com a DeepSeek: {e}")


def resumo_texto_seguro(resumo: str) -> str:
    """Usado quando a IA está desativada (sem chave)."""
    return resumo


SYSTEM_PROMPT_CHAT = """Você é um CFO fracional / analista financeiro sênior que atende pequenos negócios (dono de PME), em linguagem simples e prática.

O usuário acabou de analisar os dados financeiros dele. Você receberá:
1. UM RESUMO DOS NÚMEROS CALCULADOS (entradas, saídas, saldo, categorias, fluxo mensal, alertas).
2. A PERGUNTA do usuário.

REGRAS:
- Responda SOMENTE com base nos números fornecidos. NUNCA invente valor, projeção ou categoria.
- Se faltar dado para responder, diga o que falta e sugira o que enviar.
- Linguagem de dono, não de contador. Números sempre com 'e daí?' (o que significa e o que fazer).
- Se o usuário pedir projeção ou plano de corte, faça com os números que existem e deixe claro o que é estimativa.
- Responda em português do Brasil, de forma direta e prática."""


def perguntar(resumo_texto: str, pergunta: str, api_key: str = None, modelo: str = None) -> str:
    """Envia a pergunta do usuário junto com o contexto da análise para a DeepSeek."""
    api_key = api_key or os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return "IA desativada: configure a chave DEEPSEEK_API_KEY para usar o chat."

    modelo = modelo or os.environ.get("DEEPSEEK_MODEL", MODELO_PADRAO)

    payload = {
        "model": modelo,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT_CHAT},
            {"role": "user", "content": f"RESUMO DOS NÚMEROS:\n{resumo_texto}\n\nPERGUNTA DO USUÁRIO:\n{pergunta}"},
        ],
        "temperature": 0.4,
        "max_tokens": 1500,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    try:
        resp = requests.post(API_URL, headers=headers, json=payload, timeout=90)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as e:
        corpo = e.response.text[:400] if e.response is not None else str(e)
        raise RuntimeError(f"Erro na API DeepSeek (HTTP {e.response.status_code if e.response else '?'}): {corpo}")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha de conexão com a DeepSeek: {e}")


SYSTEM_PROMPT_PLANILHA = """Você é um analista financeiro e de planilhas sênior que ajuda donos de PME.

O usuário enviou uma PLANILHA COMPLETA (pode ter várias abas: orçamentos, DRE, fluxo de caixa, cenários, clientes, pessoal, etc.). Você receberá:
1. A ESTRUTURA da planilha (abas, cabeçalhos, amostra de dados de cada aba).
2. A PERGUNTA do usuário sobre essa planilha.

REGRAS:
- Responda COM BASE na estrutura e nos dados fornecidos da planilha. Não invente números que não estão visíveis.
- Se a pergunta pedir um número que está em uma aba, procure na amostra/estrutura e cite de qual aba veio.
- Se não houver dados suficientes para responder com precisão, diga o que está vendo e o que faltaria para responder melhor.
- Linguagem de dono de empresa, simples e prática, em português do Brasil.
- Se a planilha tem cenários, comparações ou métricas (EBITDA, margem, caixa), explique o que significam em linguagem simples.
- Seja útil: além de responder, aponte se há algo relevante (risco, inconsistência, destaque)."""


def perguntar_planilha(estrutura_texto: str, pergunta: str, api_key: str = None, modelo: str = None) -> str:
    """Responde perguntas sobre uma planilha completa (múltiplas abas)."""
    api_key = api_key or os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return "IA desativada: configure a chave DEEPSEEK_API_KEY para usar o chat."

    modelo = modelo or os.environ.get("DEEPSEEK_MODEL", MODELO_PADRAO)

    payload = {
        "model": modelo,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT_PLANILHA},
            {"role": "user", "content": "ESTRUTURA DA PLANILHA:\n" + estrutura_texto + "\n\nPERGUNTA DO USUÁRIO:\n" + pergunta},
        ],
        "temperature": 0.4,
        "max_tokens": 2000,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    try:
        resp = requests.post(API_URL, headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as e:
        corpo = e.response.text[:400] if e.response is not None else str(e)
        raise RuntimeError(f"Erro na API DeepSeek (HTTP {e.response.status_code if e.response else '?'}): {corpo}")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha de conexão com a DeepSeek: {e}")



def perguntar_historico(historico, pergunta, api_key=None, modelo=None):
    """Chat multi-turno: recebe histórico de mensagens e responde mantendo contexto.

    historico: lista de dicts {"role": "user"/"assistant", "content": "..."}
    """
    api_key = api_key or os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return "IA desativada: configure a chave DEEPSEEK_API_KEY para usar o chat."

    modelo = modelo or os.environ.get("DEEPSEEK_MODEL", MODELO_PADRAO)

    messages = [{"role": "system", "content": SYSTEM_PROMPT_CHAT}]
    for msg in historico[-10:]:
        if msg.get("role") in ("user", "assistant") and msg.get("content"):
            messages.append({"role": msg["role"], "content": str(msg["content"])[:2000]})
    messages.append({"role": "user", "content": pergunta})

    payload = {
        "model": modelo,
        "messages": messages,
        "temperature": 0.4,
        "max_tokens": 2000,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    try:
        resp = requests.post(API_URL, headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as e:
        corpo = e.response.text[:400] if e.response is not None else str(e)
        raise RuntimeError(f"Erro na API DeepSeek (HTTP {e.response.status_code if e.response else '?'}): {corpo}")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha de conexão com a DeepSeek: {e}")


def perguntar_planilha_historico(historico, estrutura_texto, pergunta, api_key=None, modelo=None):
    """Chat multi-turno para planilhas completas (mantém contexto da planilha)."""
    api_key = api_key or os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return "IA desativada: configure a chave DEEPSEEK_API_KEY para usar o chat."

    modelo = modelo or os.environ.get("DEEPSEEK_MODEL", MODELO_PADRAO)

    messages = [{"role": "system", "content": SYSTEM_PROMPT_PLANILHA}]
    messages.append({"role": "system", "content": "ESTRUTURA DA PLANILHA (referência, use sempre que preciso):\n" + estrutura_texto})
    for msg in historico[-8:]:
        if msg.get("role") in ("user", "assistant") and msg.get("content"):
            messages.append({"role": msg["role"], "content": str(msg["content"])[:2000]})
    messages.append({"role": "user", "content": pergunta})

    payload = {
        "model": modelo,
        "messages": messages,
        "temperature": 0.4,
        "max_tokens": 2000,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    try:
        resp = requests.post(API_URL, headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as e:
        corpo = e.response.text[:400] if e.response is not None else str(e)
        raise RuntimeError(f"Erro na API DeepSeek (HTTP {e.response.status_code if e.response else '?'}): {corpo}")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha de conexão com a DeepSeek: {e}")



SYSTEM_PROMPT_COMPARAR = """Você é um auditor de planilhas sênior. O usuário enviou DUAS planilhas (ou duas versões) e quer saber as diferenças.

Você receberá a estrutura/resumo das duas. Compare e aponte:
1. Abas que existem em uma e não na outra.
2. Diferenças de valores nos pontos equivalentes (orçado vs realizado, versão A vs B).
3. O que mudou de mais relevante e o que isso significa.
4. Riscos ou divergências que merecem atenção.

Responda em português do Brasil, linguagem de dono de empresa. Use tabelas quando ajudar. Não invente valores que não estejam nos dados fornecidos."""


SYSTEM_PROMPT_PREVISAO = """Você é um CFO analítico. O usuário quer uma PREVISÃO dos próximos 3 meses do caixa.

Você receberá os números reais (entradas, saídas, saldo, categorias, fluxo mensal). Projete os próximos 3 meses:
1. Estimativa de entradas, saídas e saldo por mês.
2. Cenário pessimista, realista e otimista.
3. Pontos de atenção (mês que aperta, risco de queimar caixa).
4. Recomendações práticas.

Deixe CLARO que é uma projeção/estimativa baseada nos dados atuais, não garantia. Responda em português do Brasil."""


def comparar_planilhas(estrutura_texto: str, api_key=None, modelo=None):
    """IA compara duas planilhas e aponta diferenças."""
    api_key = api_key or os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return "IA desativada: configure a chave DEEPSEEK_API_KEY para usar o chat."
    modelo = modelo or os.environ.get("DEEPSEEK_MODEL", MODELO_PADRAO)
    payload = {
        "model": modelo,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT_COMPARAR},
            {"role": "user", "content": estrutura_texto},
        ],
        "temperature": 0.3,
        "max_tokens": 2000,
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    try:
        resp = requests.post(API_URL, headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha de conexão com a DeepSeek: {e}")


def prever_proximos_meses(contexto: str, api_key=None, modelo=None):
    """IA projeta os próximos 3 meses com base nos números."""
    api_key = api_key or os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return "IA desativada: configure a chave DEEPSEEK_API_KEY para usar o chat."
    modelo = modelo or os.environ.get("DEEPSEEK_MODEL", MODELO_PADRAO)
    payload = {
        "model": modelo,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT_PREVISAO},
            {"role": "user", "content": contexto},
        ],
        "temperature": 0.4,
        "max_tokens": 2000,
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    try:
        resp = requests.post(API_URL, headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha de conexão com a DeepSeek: {e}")
