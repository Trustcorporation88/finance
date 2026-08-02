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
