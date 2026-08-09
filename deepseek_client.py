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



SYSTEM_PROMPT_ESTUDO = """Você é um analista financeiro sênior (CFO) conduzindo um ESTUDO COMPLETO. Seu trabalho é igual ao de um consultor contratado para dissecar a planilha e entregar um raio-X com TODOS os números relevantes — nada de resumo superficial.

Você receberá a ESTRUTURA da planilha (abas, cabeçalhos, amostras de cada aba). Entregue:

## 1. Resumo Executivo (3 cenários se houver)
Monte UMA TABELA comparativa com colunas: Indicador | Cenário 1 | Cenário 2 | Cenário 3
Exemplo:
| Indicador | Cenário 1 | Cenário 2 | Cenário 3 |
|-----------|-----------|-----------|-----------|
| Vendas totais (ago-dez/26) | R$ X | R$ Y | R$ Z |
| Gastos totais (ago-dez/26) | R$ X | R$ Y | R$ Z |
| EBITDA | R$ X | R$ Y | R$ Z |
| Lucro líquido | R$ X | R$ Y | R$ Z |
| Geração de caixa | R$ X | R$ Y | R$ Z |
| Caixa final | R$ X | R$ Y | R$ Z |
Números exatos, sem arredondar.

## 2. Análise de Cada Cenário (um por um, com profundidade)
Para cada cenário, monte UMA TABELA com colunas: Mês | Receita | Custo | EBITDA | Saldo Caixa.
Exemplo:
| Mês/Ano | Receita | Custo | EBITDA | Saldo Caixa |
|---------|---------|-------|--------|-------------|
| Ago/26  | R$ X    | R$ Y  | R$ Z   | R$ W        |
Depois:
- Qual mês o caixa vira negativo
- Margem do período e tendência

## 3. DRE × Fluxo de Caixa (por que fecha diferente?)
- Receita contábil (DRE) vs. receita caixa (CF)
- Custos contábeis vs. pagos
- Resultado DRE vs. resultado Caixa
- Explicação simples da diferença

## 4. De onde vem a receita? (Concentração)
- Liste clientes/produtos/linhas com valor e %
- Aponte se há concentração perigosa (>40% em 1 cliente)
- Cite nome do cliente e valor (ex.: "Roche R$ 46.605")

## 5. Para onde vai o dinheiro? (Estrutura de custos)
- Categorias de gasto com valor e %
- Fixo vs. Variável
- O que pesa mais e por quê

## 6. Pessoal / Headcount
- Quantas pessoas, custo total, % da receita
- Se houver pró-labore / sócios, aponte

## 7. Fluxo de Caixa Projetado
- Tabela com colunas: Mês/Ano | Saldo Inicial | Entradas | Saídas | Saldo Final
- Destaque o pior mês e o melhor mês
- Quanto de capital de giro precisa

## 8. Indicadores-Chave
- Margem bruta, margem líquida, EBITDA %
- Ponto de equilíbrio (se calcular)
- Cobertura de caixa (meses)

## 9. Alertas e Riscos (3 a 5)
- Cada alerta com o NÚMERO que comprova (ex.: "Folha consome 72% da receita")
- 🔴 para crítico, 🟡 para atenção

## 10. Recomendações (3 a 5)
- Ações práticas priorizadas
- Impacto estimado em R$ quando possível

## 11. Sugestão de Apresentação
- 5-6 slides para mostrar ao gestor/diretor

REGRAS DE OURO:
- NUNCA faça resumo genérico. Extraia TODO número disponível da estrutura.
- TODO número precisa de RÓTULO e DATA: não escreva "R$ 54.610" solto — escreva "Receita Ago/26: R$ 54.610" ou "Custo fixo mensal (12x): R$ 118.622". O leitor precisa saber o que é e de quando é cada valor.
- Tabelas devem ter cabeçalho com nomes das colunas (ex.: | Mês | Receita | Custo | EBITDA |).
- Cite SEMPRE de qual aba veio o número (ex.: "aba CashFlow1").
- Se um número não estiver visível na amostra, escreva: "(não visível na amostra da aba X)" — NÃO invente.
- Use tabelas markdown para comparar cenários e meses.
- Formato: markdown (## para seções, | para tabelas, - para listas).
- Linguagem de dono de empresa, português do Brasil.
- Seja cirúrgico nos números. Não diga "aproximadamente" — use os valores exatos da planilha.
- max_tokens é alto de propósito: use todo o espaço para detalhar.

No final: "Quer que eu aprofunde algum ponto específico?" """


def estudo_completo(estrutura_texto: str, api_key: str = None, modelo: str = None) -> str:
    """Gera um estudo completo da planilha (análise geral profunda)."""
    api_key = api_key or os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        return "IA desativada: configure a chave DEEPSEEK_API_KEY para usar o chat."
    modelo = modelo or os.environ.get("DEEPSEEK_MODEL", MODELO_PADRAO)
    payload = {
        "model": modelo,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT_ESTUDO},
            {"role": "user", "content": "ESTRUTURA COMPLETA DA PLANILHA:\n" + estrutura_texto},
        ],
        "temperature": 0.3,
        "max_tokens": 8000,
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    try:
        resp = requests.post(API_URL, headers=headers, json=payload, timeout=180)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as e:
        corpo = e.response.text[:400] if e.response is not None else str(e)
        raise RuntimeError(f"Erro na API DeepSeek (HTTP {e.response.status_code if e.response else '?'}): {corpo}")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha de conexão com a DeepSeek: {e}")
