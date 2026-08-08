# -*- coding: utf-8 -*-
"""Integração com Supabase (Auth + Postgres + Storage).

Ativo apenas quando SUPABASE_URL e SUPABASE_ANON_KEY estão configurados.
Sem isso, o app cai no modo local (sessão + memória) e funciona normalmente.
"""
from __future__ import annotations

import os
import io
import json
import time

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip()
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "").strip()
SUPABASE_SERVICE_ROLE = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()

ATIVO = bool(SUPABASE_URL and SUPABASE_ANON_KEY)

_client = None
_admin = None

# nomes de tabelas/bucket (configuráveis)
TABELA_ANALISES = os.environ.get("SUPABASE_TABELA_ANALISES", "analises")
TABELA_LOGS = os.environ.get("SUPABASE_TABELA_LOGS", "logs_uso")
BUCKET_RELATORIOS = os.environ.get("SUPABASE_BUCKET", "relatorios")


def _get_client(service_role=False):
    global _client, _admin
    if not ATIVO:
        return None
    from supabase import create_client
    if service_role and SUPABASE_SERVICE_ROLE:
        if _admin is None:
            _admin = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE)
        return _admin
    if _client is None:
        _client = create_client(SUPABASE_URL, SUPABASE_ANON_KEY)
    return _client


def ativo() -> bool:
    return ATIVO


def inicializar_schema():
    """Cria tabelas se não existirem (via SQL, usando service role)."""
    if not ATIVO or not SUPABASE_SERVICE_ROLE:
        return False
    try:
        client = _get_client(service_role=True)
        sql = f"""
        create table if not exists {TABELA_ANALISES} (
          id bigint generated always as identity primary key,
          usuario text not null,
          nome text,
          tipo text,
          resultado jsonb,
          criado_em timestamptz default now()
        );
        create table if not exists {TABELA_LOGS} (
          id bigint generated always as identity primary key,
          acao text,
          ip text,
          criado_em timestamptz default now()
        );
        """
        client.rpc("exec_sql", {"query": sql}).execute()
        return True
    except Exception:
        return False


# --------------------------------------------------------------------------
# AUTH
# --------------------------------------------------------------------------

def sign_up(email: str, senha: str) -> dict:
    """Cria um novo usuário. Retorna {'ok': bool, 'erro'?}."""
    client = _get_client()
    if not client:
        return {"ok": False, "erro": "Supabase não configurado."}
    try:
        resp = client.auth.sign_up({"email": email, "password": senha})
        user = resp.user
        if user and user.identities and len(user.identities) > 0:
            return {"ok": True, "usuario": email}
        # e-mail já existe ou precisa confirmação
        return {"ok": True, "usuario": email, "confirmacao": True}
    except Exception as e:
        return {"ok": False, "erro": str(e)}


def sign_in(email: str, senha: str) -> dict:
    client = _get_client()
    if not client:
        return {"ok": False, "erro": "Supabase não configurado."}
    try:
        resp = client.auth.sign_in_with_password({"email": email, "password": senha})
        return {"ok": True, "usuario": email, "access_token": resp.session.access_token,
                "refresh_token": resp.session.refresh_token}
    except Exception as e:
        return {"ok": False, "erro": "E-mail ou senha inválidos."}


def sign_in_magic(email: str) -> dict:
    client = _get_client()
    if not client:
        return {"ok": False, "erro": "Supabase não configurado."}
    try:
        client.auth.sign_in_with_otp({"email": email})
        return {"ok": True, "mensagem": "Link de acesso enviado para o e-mail."}
    except Exception as e:
        return {"ok": False, "erro": str(e)}


def get_user(access_token: str):
    client = _get_client()
    if not client:
        return None
    try:
        resp = client.auth.get_user(access_token)
        return resp.user
    except Exception:
        return None


def criar_usuario_admin(email: str, senha: str):
    """Cria usuário via service role (admin)."""
    if not SUPABASE_SERVICE_ROLE:
        return {"ok": False, "erro": "Service role não configurado."}
    try:
        client = _get_client(service_role=True)
        client.auth.admin.create_user({"email": email, "password": senha, "email_confirm": True})
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "erro": str(e)}


# --------------------------------------------------------------------------
# POSTGRES: histórico de análises
# --------------------------------------------------------------------------

def salvar_analise(usuario: str, nome: str, tipo: str, resultado: dict) -> bool:
    client = _get_client()
    if not client:
        return False
    try:
        client.table(TABELA_ANALISES).insert({
            "usuario": usuario,
            "nome": nome,
            "tipo": tipo,
            "resultado": resultado,
        }).execute()
        return True
    except Exception:
        return False


def listar_analises(usuario: str, limite=30) -> list:
    client = _get_client()
    if not client:
        return []
    try:
        resp = client.table(TABELA_ANALISES).select("id,nome,tipo,criado_em").eq(
            "usuario", usuario).order("criado_em", desc=True).limit(limite).execute()
        return resp.data or []
    except Exception:
        return []


def buscar_analise(usuario: str, analise_id: int):
    client = _get_client()
    if not client:
        return None
    try:
        resp = client.table(TABELA_ANALISES).select("*").eq("usuario", usuario).eq(
            "id", analise_id).execute()
        if resp.data:
            return resp.data[0]
        return None
    except Exception:
        return None


def deletar_analise(usuario: str, analise_id: int) -> bool:
    client = _get_client()
    if not client:
        return False
    try:
        client.table(TABELA_ANALISES).delete().eq("usuario", usuario).eq("id", analise_id).execute()
        return True
    except Exception:
        return False


# --------------------------------------------------------------------------
# POSTGRES: logs de uso
# --------------------------------------------------------------------------

def registrar_log(acao: str, ip: str = "") -> bool:
    client = _get_client()
    if not client:
        return False
    try:
        client.table(TABELA_LOGS).insert({"acao": acao, "ip": ip}).execute()
        return True
    except Exception:
        return False


def resumo_logs(limit=20) -> dict:
    client = _get_client()
    if not client:
        return {}
    try:
        resp = client.table(TABELA_LOGS).select("acao,ip").order("criado_em", desc=True).limit(limit).execute()
        acoes = {}
        ips = {}
        for row in (resp.data or []):
            acoes[row.get("acao", "?")] = acoes.get(row.get("acao", "?"), 0) + 1
            ip = row.get("ip", "")
            if ip:
                ips[ip] = ips.get(ip, 0) + 1
        return {
            "recentes": resp.data or [],
            "top_acoes": sorted(acoes.items(), key=lambda x: -x[1])[:10],
            "top_ips": sorted(ips.items(), key=lambda x: -x[1])[:10],
        }
    except Exception:
        return {}


# --------------------------------------------------------------------------
# STORAGE: relatórios gerados
# --------------------------------------------------------------------------

def salvar_relatorio(usuario: str, nome: str, conteudo: bytes, tipo: str = "application/octet-stream") -> str:
    """Salva um relatório no Storage e devolve URL pública (ou None)."""
    client = _get_client()
    if not client:
        return ""
    try:
        caminho = f"{usuario}/{int(time.time())}-{nome}"
        client.storage.from_(BUCKET_RELATORIOS).upload(caminho, conteudo,
                                                       {"content-type": tipo})
        return f"{SUPABASE_URL}/storage/v1/object/public/{BUCKET_RELATORIOS}/{caminho}"
    except Exception:
        return ""


def garantir_bucket():
    """Cria o bucket se não existir (service role)."""
    if not ATIVO or not SUPABASE_SERVICE_ROLE:
        return
    try:
        client = _get_client(service_role=True)
        client.storage.create_bucket(BUCKET_RELATORIOS, {"public": True})
    except Exception:
        pass
