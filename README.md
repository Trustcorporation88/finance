# 🚀 CFO de Bolso — FIRST Ambiental

Sobe a sua planilha (CSV/Excel) ou cola os números e recebe o **Raio-X do caixa**: entrou, saiu, sobrou, margem, pra onde foi o dinheiro, alertas automáticos e análise com a IA da **DeepSeek** (usando a sua API key). Download do relatório em **Word** e **PDF**. Identidade visual FIRST Ambiental (verde claro + logo).

## Testar localmente (2 minutos)

```bash
cd cfo-de-bolso-app

# 1. instalar dependências
pip install -r requirements.txt

# 2. criar o arquivo de chave
copy .env.example .env        # Windows
# cp .env.example .env        # Mac/Linux
#   -> edite o .env e coloque sua chave DeepSeek em DEEPSEEK_API_KEY

# 3. rodar
python app.py
```

Abra **http://localhost:5000** — pronto.

> Sem a chave no `.env` o app funciona igual, só mostra os cálculos automáticos e avisa que a IA está desativada.

## Como funciona

1. Você envia CSV/XLSX ou cola texto (formato `data;descrição;valor` ou `descrição;valor` por linha).
2. O backend **calcula tudo localmente** (pandas): entradas, saídas, saldo, margem, categorias de gasto, fluxo por mês e alertas. **Nenhum número é inventado** — o cálculo é 100% determinístico no seu servidor.
3. Os números prontos são enviados para a **DeepSeek** (`deepseek-chat`), que escreve o raio-x em linguagem de dono de PME: o que sobrou, por quê e o que fazer.
4. Gera gráficos (matplotlib) e permite baixar o relatório em **Word (.docx)** e **PDF**.

## Publicar (Railway Pro — recomendado)

O projeto inclui `Dockerfile` + `railway.json` (config pronta):

1. Suba a pasta `cfo-de-bolso-app` para um repositório Git (GitHub).
2. No Railway: **New Project → Deploy from GitHub repo** → selecione o repositório.
3. Em **Settings → Variables** adicione:
   - `DEEPSEEK_API_KEY=sk-...` (sua chave)
   - `DEEPSEEK_MODEL=deepseek-chat` (opcional, padrão)
4. Deploy. O Railway injeta a variável `PORT` automaticamente.
5. Pronto — URL pública tipo `https://cfo-de-bolso.up.railway.app`.

> 🔒 A chave fica **só no servidor** (variável de ambiente). Nunca é exposta no navegador.

## Publicar (Vercel)

Para Vercel (serverless), a abordagem recomendada é subir este app num backend (Railway/Render/VPS) e apontar o Vercel para ele, OU usar o Vercel só como frontend. Flask + matplotlib + geração de PDF não rodam bem em serverless sem ajustes. Se quiser mesmo no Vercel, o caminho mais simples:

```json
// vercel.json (experimental — requer Python runtime)
{ "builds": [{ "src": "app.py", "use": "@vercel/python" }] }
```

e coloque `DEEPSEEK_API_KEY` em **Vercel → Project → Settings → Environment Variables**. Teste com cuidado: o pacote completo pode estourar o limite de tamanho do serverless.

## Variáveis de ambiente

| Variável | Obrigatória | Descrição |
|---|---|---|
| `DEEPSEEK_API_KEY` | sim (para IA) | Sua chave da DeepSeek |
| `DEEPSEEK_MODEL` | não | Padrão `deepseek-chat` |
| `PORT` | não | Definida automaticamente em Railway |

## Estrutura

```
cfo-de-bolso-app/
├── app.py                # servidor Flask + rotas
├── analise.py            # parsing, categorização e cálculos
├── graficos.py           # geração de gráficos (matplotlib)
├── deepseek_client.py    # integração com a API DeepSeek
├── relatorio.py          # geração de Word e PDF
├── templates/index.html  # página web
├── static/logo-first.png # logo FIRST Ambiental
├── requirements.txt
├── Dockerfile            # build para Railway
├── railway.json          # config de deploy do Railway
└── .env.example
```

## Dica de formato dos dados

- **Colunas** reconhecidas automaticamente: `valor`/`vlr`, `data`, `descrição`, `tipo` (entrada/saída).
- **Sinal**: positivo = entrada (dinheiro que entrou), negativo = saída. Se sua planilha tiver coluna de tipo, use-a.
- **Texto colado**: uma linha por lançamento, `data;descrição;valor` ou `descrição;valor`.

> ⚠️ Gestão de caixa gerencial, não apuração fiscal. Para imposto/nota, consulte seu contador.
