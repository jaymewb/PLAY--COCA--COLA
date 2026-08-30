from datetime import datetime
from io import BytesIO
import os
import re
import sqlite3
import urllib.parse
import json
import time

from flask import Flask, request, render_template_string, jsonify, session, redirect, url_for
import pandas as pd
import requests


# ============================================================
# CONFIGURAÇÃO
# ============================================================

app = Flask(__name__)
app.secret_key = "1020_chave_secreta_cads"

SHEET_URL = (
    "https://1drv.ms/x/c/b96adcc2e8fff38f/"
    "IQArU8O1IbaiQIK2QvhMQw4gAfT_UMwNBlGCdRsJ23RyRHg?e=BRjezC"
)

DB_NAME = "rh_escala.db"

# Cache da planilha para evitar baixar/ler o Excel do OneDrive a cada abertura/filtro.
PLANILHA_CACHE_TTL = 120
_planilha_cache = {"timestamp": 0, "df": None, "colunas": None, "sync_timestamp": 0}


# ============================================================
# BANCO DE DADOS
# ============================================================

def conectar():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_banco():

    conn = conectar()
    cursor = conn.cursor()

    # --------------------------------------------------------
    # CADASTRO PRINCIPAL
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cadastro_colaboradores (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            nome TEXT,
            telefone TEXT,
            data_admissao TEXT,
            status TEXT,
            unidade TEXT,

            dados_completos TEXT,

            atualizado_em TEXT,

            UNIQUE(nome, telefone)
        )
    """)

    # --------------------------------------------------------
    # HISTÓRICO DE ENVIO
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historico_envios (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            hora TEXT,
            nome TEXT,
            sucesso BOOLEAN,
            detalhe TEXT,
            mensagem TEXT
        )
    """)

    conn.commit()
    conn.close()


inicializar_banco()
with conectar() as _conn:
    _conn.execute("CREATE INDEX IF NOT EXISTS idx_cadastro_nome ON cadastro_colaboradores(nome COLLATE NOCASE)")
    _conn.execute("CREATE INDEX IF NOT EXISTS idx_cadastro_telefone ON cadastro_colaboradores(telefone)")
    _conn.execute("CREATE INDEX IF NOT EXISTS idx_cadastro_status ON cadastro_colaboradores(status)")
    _conn.execute("CREATE INDEX IF NOT EXISTS idx_cadastro_unidade ON cadastro_colaboradores(unidade)")
    _conn.execute("CREATE INDEX IF NOT EXISTS idx_cadastro_admissao ON cadastro_colaboradores(data_admissao)")
    _conn.commit()


# ============================================================
# NORMALIZAÇÃO DE TEXTO
# ============================================================

def normalizar_texto(valor):

    if valor is None:
        return ""

    texto = str(valor).strip()

    texto = (
        texto
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
    )

    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def normalizar_coluna(nome):

    texto = normalizar_texto(nome).lower()

    substituicoes = {
        "á": "a", "à": "a", "ã": "a", "â": "a", "ä": "a",
        "é": "e", "è": "e", "ê": "e", "ë": "e",
        "í": "i", "ì": "i", "î": "i", "ï": "i",
        "ó": "o", "ò": "o", "õ": "o", "ô": "o", "ö": "o",
        "ú": "u", "ù": "u", "û": "u", "ü": "u", "ç": "c",
    }

    for origem, destino in substituicoes.items():
        texto = texto.replace(origem, destino)

    texto = re.sub(r"[^a-z0-9]+", " ", texto)

    return texto.strip()


# ============================================================
# LOCALIZAR COLUNAS DA PLANILHA
# ============================================================

def localizar_coluna(df, possibilidades):

    mapa = {}

    for coluna in df.columns:
        chave = normalizar_coluna(coluna)
        mapa[chave] = coluna

    for possibilidade in possibilidades:
        chave = normalizar_coluna(possibilidade)
        if chave in mapa:
            return mapa[chave]

    for chave, coluna in mapa.items():
        for possibilidade in possibilidades:
            p = normalizar_coluna(possibilidade)
            if p in chave or chave in p:
                return coluna

    return None


def descobrir_colunas(df):

    colunas = {}
    colunas["nome"] = localizar_coluna(df, ["Nome", "Nome do Colaborador", "Nome Colaborador", "NOME"])
    colunas["telefone"] = localizar_coluna(df, ["TELEFONE", "Telefone", "Celular", "WhatsApp", "Telefone Celular"])
    colunas["data_admissao"] = localizar_coluna(df, ["Data admissão", "Data de admissão", "Data Admissão", "Admissão", "Data adm"])
    colunas["status"] = localizar_coluna(df, ["Status", "STATUS"])
    colunas["unidade"] = localizar_coluna(df, ["Unidade", "UNIDADE", "Posto", "POSTO"])

    return colunas


# ============================================================
# TELEFONE & DATA
# ============================================================

def limpar_telefone(valor):

    if pd.isna(valor):
        return ""

    telefone = str(valor)
    if telefone.endswith(".0"):
        telefone = telefone[:-2]

    telefone = re.sub(r"\D", "", telefone)
    if telefone.startswith("55") and len(telefone) >= 12:
        telefone = telefone[2:]

    return telefone


def formatar_telefone(telefone):

    telefone = limpar_telefone(telefone)
    if len(telefone) == 11:
        return f"({telefone[:2]}) {telefone[2:7]}-{telefone[7:]}"
    if len(telefone) == 10:
        return f"({telefone[:2]}) {telefone[2:6]}-{telefone[6:]}"
    return telefone


def formatar_data(valor):

    if valor is None or pd.isna(valor):
        return ""

    try:
        if isinstance(valor, datetime):
            return valor.strftime("%d/%m/%Y")
        data = pd.to_datetime(valor, dayfirst=True)
        return data.strftime("%d/%m/%Y")
    except Exception:
        return normalizar_texto(valor)


def converter_data_filtro(valor):

    if not valor:
        return None
    try:
        return datetime.strptime(valor, "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return valor


# ============================================================
# BAIXAR E LER PLANILHA
# ============================================================

def baixar_planilha():

    urls = [SHEET_URL, SHEET_URL + "&download=1"]
    ultimo_erro = None

    for url in urls:
        try:
            resposta = requests.get(
                url,
                timeout=30,
                allow_redirects=True,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
            )
            resposta.raise_for_status()
            conteudo = resposta.content
            inicio = conteudo[:20]

            if inicio.startswith(b"PK") or "excel" in resposta.headers.get("Content-Type", "").lower():
                return conteudo

            texto = conteudo.decode("utf-8", errors="ignore")
            possiveis_urls = re.findall(r'https?://[^"\']+', texto)

            for possivel in possiveis_urls:
                if "download" in possivel.lower() or "onedrive" in possivel.lower():
                    try:
                        r2 = requests.get(possivel, timeout=30, allow_redirects=True)
                        if r2.ok and r2.content[:2] == b"PK":
                            return r2.content
                    except Exception:
                        pass
        except Exception as erro:
            ultimo_erro = erro

    raise RuntimeError(f"Não foi possível baixar a planilha do OneDrive: {ultimo_erro}")


def carregar_planilha(forcar=False):
    agora = time.time()
    if (not forcar and _planilha_cache["df"] is not None
            and agora - _planilha_cache["timestamp"] < PLANILHA_CACHE_TTL):
        return _planilha_cache["df"]

    arquivo = baixar_planilha()
    excel = pd.ExcelFile(BytesIO(arquivo))
    aba = excel.sheet_names[0]
    df = pd.read_excel(BytesIO(arquivo), sheet_name=aba, dtype=object)
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df.columns = [normalizar_texto(c) for c in df.columns]
    _planilha_cache.update({"timestamp": agora, "df": df})
    return df


def sincronizar_cadastro(df):

    colunas = descobrir_colunas(df)
    nome_col, telefone_col, admissao_col, status_col, unidade_col = (
        colunas["nome"], colunas["telefone"], colunas["data_admissao"], colunas["status"], colunas["unidade"]
    )

    if not nome_col or not telefone_col or not admissao_col or not status_col:
        raise RuntimeError("Colunas essenciais não foram encontradas na planilha.")

    conn = conectar()
    cursor = conn.cursor()
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    quantidade = 0

    for _, linha in df.iterrows():
        nome = normalizar_texto(linha.get(nome_col, ""))
        if not nome:
            continue

        telefone = limpar_telefone(linha.get(telefone_col, ""))
        data_admissao = formatar_data(linha.get(admissao_col, ""))
        status = normalizar_texto(linha.get(status_col, ""))
        unidade = normalizar_texto(linha.get(unidade_col, "")) if unidade_col else ""

        dados_completos = {}
        for coluna in df.columns:
            valor = linha.get(coluna, "")
            if pd.isna(valor):
                valor = ""
            if isinstance(valor, datetime):
                valor = valor.strftime("%d/%m/%Y")
            dados_completos[str(coluna)] = str(valor)

        dados_json = json.dumps(dados_completos, ensure_ascii=False)

        cursor.execute("SELECT id FROM cadastro_colaboradores WHERE nome = ? AND telefone = ?", (nome, telefone))
        existente = cursor.fetchone()

        if existente:
            cursor.execute("""
                UPDATE cadastro_colaboradores
                SET data_admissao = ?, status = ?, unidade = ?, dados_completos = ?, atualizado_em = ?
                WHERE id = ?
            """, (data_admissao, status, unidade, dados_json, agora, existente["id"]))
        else:
            cursor.execute("""
                INSERT INTO cadastro_colaboradores (nome, telefone, data_admissao, status, unidade, dados_completos, atualizado_em)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (nome, telefone, data_admissao, status, unidade, dados_json, agora))

        quantidade += 1

    conn.commit()
    conn.close()
    return quantidade, colunas


def carregar_cadastro():
    conn = conectar()
    registros = conn.execute("SELECT * FROM cadastro_colaboradores ORDER BY nome COLLATE NOCASE").fetchall()
    conn.close()
    return registros


def aplicar_filtros(registros, busca, filtro_posto, filtro_status, filtro_pendencia, data_selecionada):
    resultado = []
    postos = {f.lower() for f in filtro_posto}
    status_set = {f.lower() for f in filtro_status}
    pend_set = {f.lower() for f in filtro_pendencia}
    busca_lower = busca.lower()
    data_filtro = converter_data_filtro(data_selecionada) if data_selecionada else None
    for registro in registros:
        nome = registro["nome"] or ""
        status = registro["status"] or ""
        unidade = registro["unidade"] or ""
        data_admissao = registro["data_admissao"] or ""

        nome_lower = nome.lower()
        status_lower = status.lower()
        unidade_lower = unidade.lower()
        if busca_lower and busca_lower not in nome_lower:
            continue
        if postos and unidade_lower not in postos:
            continue
        if status_set and status_lower not in status_set:
            continue
        if pend_set and status_lower not in pend_set:
            continue
        if data_filtro and data_admissao != data_filtro:
            continue

        resultado.append(registro)
    return resultado


def montar_card(registro):
    nome = registro["nome"] or ""
    telefone = limpar_telefone(registro["telefone"] or "")
    status = registro["status"] or ""
    primeiro_nome = nome.split()[0] if nome.split() else ""
    hora = datetime.now().hour

    saudacao = "bom dia" if hora < 12 else ("boa tarde" if hora < 18 else "boa noite")
    mensagem = f"Prezado(a) {primeiro_nome}, {saudacao}! Identificamos a pendência: {status}."

    link = f"https://api.whatsapp.com/send?phone=55{telefone}&text=" + urllib.parse.quote(mensagem) if (telefone and len(telefone) >= 10) else "#"

    return {
        "id": registro["id"],
        "nome": nome,
        "telefone_bruto": telefone,
        "telefone_formatado": formatar_telefone(telefone),
        "data_admissao": registro["data_admissao"],
        "status": status,
        "pendencia": status,
        "mensagem": mensagem,
        "posto": registro["unidade"] or "Sem Unidade",
        "link": link
    }


def buscar_historico():
    conn = conectar()
    dados = conn.execute("SELECT id, hora, nome, sucesso, detalhe, mensagem FROM historico_envios ORDER BY id DESC").fetchall()
    conn.close()
    return dados


def registrar_log(nome, sucesso, detalhe, mensagem):
    conn = conectar()
    conn.execute("""
        INSERT INTO historico_envios (hora, nome, sucesso, detalhe, mensagem)
        VALUES (?, ?, ?, ?, ?)
    """, (datetime.now().strftime("%d/%m/%Y %H:%M:%S"), nome, 1 if sucesso else 0, detalhe, mensagem))
    conn.commit()
    conn.close()


# ============================================================
# ROTAS (LOGIN E SISTEMA)
# ============================================================

@app.route("/", methods=["GET"])
def index():
    # A página usa o cadastro local. Nenhum download do OneDrive
    # acontece durante filtros, busca ou navegação do painel.
    erro = None
    sincronizados = 0

    busca = request.args.get("busca", "").strip()
    filtro_posto = request.args.getlist("filtro_posto")
    filtro_status = request.args.getlist("filtro_status")
    filtro_pendencia = request.args.getlist("filtro_pendencia")
    data_selecionada = request.args.get("data_selecionada", "").strip()

    registros = carregar_cadastro()
    registros_filtrados = aplicar_filtros(registros, busca, filtro_posto, filtro_status, filtro_pendencia, data_selecionada)
    dados = [montar_card(r) for r in registros_filtrados]

    unidades_disponiveis = sorted({r["unidade"] for r in registros if r["unidade"]}, key=lambda x: x.lower())
    status_disponiveis = sorted({r["status"] for r in registros if r["status"]}, key=lambda x: x.lower())
    pendencias_disponiveis = status_disponiveis.copy()
    historico = buscar_historico()

    return render_template_string(
        HTML_PAINEL,
        dados=dados,
        historico=historico,
        busca=busca,
        filtro_posto=filtro_posto,
        filtro_status=filtro_status,
        filtro_pendencia=filtro_pendencia,
        data_selecionada=data_selecionada,
        unidades_disponiveis=unidades_disponiveis,
        status_disponiveis=status_disponiveis,
        pendencias_disponiveis=pendencias_disponiveis,
        erro=erro,
        sincronizados=sincronizados
    )


@app.route("/registrar-envio", methods=["POST"])
def registrar_envio():
    dados = request.get_json(silent=True) or {}
    registrar_log(dados.get("nome", ""), dados.get("sucesso", True), dados.get("detalhe", f"Contato: {dados.get('telefone', '')}"), dados.get("mensagem", ""))
    return jsonify({"ok": True})


@app.route("/limpar-historico", methods=["POST"])
def limpar_historico():
    conn = conectar()
    conn.execute("DELETE FROM historico_envios")
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


# ============================================================
# HTML LOGIN E PAINEL
# ============================================================

HTML_PAINEL = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Painel RH - CADS / WhatsApp</title>
<style>
* { box-sizing: border-box; }
body { font-family: "Segoe UI", Arial, sans-serif; background: #0f172a; color: white; margin: 0; height: 100vh; overflow: hidden; }
.topo-fixo { padding: 12px 18px; background: #0f172a; border-bottom: 1px solid #334155; box-shadow: 0 4px 12px rgba(0,0,0,.35); position: relative; z-index: 1000; }
.header-painel { display: flex; justify-content: space-between; align-items: center; background: #1e293b; padding: 12px 18px; border-radius: 8px; margin-bottom: 10px; }
.header-painel h2 { margin: 0; font-size: 19px; }
.relogio-24h { color: #38bdf8; background: #0f172a; padding: 6px 12px; border: 1px solid #334155; border-radius: 6px; font-weight: bold; }
.painel-controles { display: flex; gap: 10px; flex-wrap: wrap; align-items: stretch; }
.search-form { flex: 1; min-width: 500px; display: flex; gap: 10px; flex-wrap: wrap; align-items: center; background: #1e293b; border: 1px solid #334155; padding: 10px; border-radius: 8px; }
.search-box-item { flex: 1; min-width: 150px; position: relative; }
.search-box-item label { display: block; font-size: 10px; color: #94a3b8; margin-bottom: 3px; }
input[type="text"] { width: 100%; padding: 8px 10px; border-radius: 7px; border: 1px solid #475569; background: #0f172a; color: white; outline: none; }
input[type="text"]:focus { border-color: #38bdf8; }
.multiselect-btn { padding: 12px 14px; min-height: 44px; border: 1px solid #475569; border-radius: 8px; background: #0f172a; color: white; cursor: pointer; display: flex; justify-content: space-between; align-items: center; font-size: 13px; font-weight: 600; user-select: none; transition: border-color .12s, background .12s; }
.multiselect-btn:hover { border-color: #38bdf8; background: #172033; }
.filter-clear { width: 100%; padding: 9px 10px; margin-bottom: 5px; border: 1px solid #ef4444; border-radius: 6px; background: #7f1d1d; color: white; cursor: pointer; font-size: 11px; font-weight: bold; }
.multiselect-content { display: none; position: absolute; top: 100%; left: 0; right: 0; background: #1e293b; border: 1px solid #334155; border-radius: 7px; padding: 7px; max-height: 220px; overflow-y: auto; z-index: 5000; }
.multiselect-content.show { display: block; }
.dropdown-item { display: flex; gap: 7px; align-items: center; padding: 6px; font-size: 12px; color: #cbd5e1; cursor: pointer; border-radius: 5px; }
.dropdown-item:hover { background: #334155; }
.calendario-box { background: #1e293b; border: 1px solid #334155; border-radius: 8px; width: 215px; padding: 8px; text-align: center; }
.cal-header { display: flex; justify-content: space-between; align-items: center; color: #38bdf8; font-weight: bold; font-size: 12px; margin-bottom: 5px; }
.cal-header button { background: #334155; color: white; border: none; padding: 3px 7px; border-radius: 4px; cursor: pointer; }
.cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; font-size: 10px; }
.cal-day-name { color: #94a3b8; font-weight: bold; padding: 3px; }
.cal-day { background: #0f172a; color: #cbd5e1; padding: 5px 0; border-radius: 3px; cursor: pointer; }
.cal-day:hover { background: #334155; color: white; }
.cal-day.selected { background: #38bdf8; color: #0f172a; font-weight: bold; }
.cal-day.today { border: 1px solid #22c55e; }
.painel-bot { width: 220px; background: #1e293b; border: 1px solid #334155; border-radius: 8px; padding: 10px; display: flex; flex-direction: column; gap: 6px; }
.btn-acao { border: none; padding: 7px; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 11px; }
.btn-iniciar { background: #22c55e; color: #0f172a; }
.btn-pausar { background: #eab308; color: #0f172a; }
.btn-home { background: #3b82f6; color: white; text-decoration: none; text-align: center; }
.cronometro-box { background: #0f172a; border: 1px solid #334155; padding: 5px; border-radius: 5px; text-align: center; color: #38bdf8; font-size: 11px; font-weight: bold; }
.barra-progresso-container { height: 6px; background: #334155; border-radius: 5px; overflow: hidden; margin-top: 4px; }
.barra-progresso-fill { height: 100%; width: 0%; background: #22c55e; transition: width .3s; }
.main-layout { display: flex; height: calc(100vh - 150px); overflow: hidden; }
.conteudo-scroll { flex: 1; overflow-y: auto; padding: 18px; }
.grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 12px; }
.card { background: #1e293b; border: 1px solid #334155; border-radius: 9px; padding: 13px; box-shadow: 0 4px 10px rgba(0,0,0,.25); transition: .3s; }
.card-ativo { border-color: #38bdf8; box-shadow: 0 0 15px rgba(56,189,248,.45); }
.posto { display: inline-block; border: 1px solid #facc15; color: #facc15; padding: 3px 7px; border-radius: 4px; font-size: 10px; font-weight: bold; margin-bottom: 7px; }
.nome { font-size: 15px; font-weight: bold; margin-bottom: 6px; }
.info { color: #cbd5e1; font-size: 12px; margin: 4px 0; }
.pendencia { margin-top: 8px; padding: 8px; background: #0f172a; border-radius: 6px; border: 1px solid #475569; }
.pendencia strong { color: #facc15; }
.preview-msg { margin-top: 7px; padding: 8px; background: #0f172a; border: 1px dashed #475569; border-radius: 6px; color: #38bdf8; font-size: 11px; word-break: break-word; }
.btn { display: block; margin-top: 9px; padding: 8px; background: #22c55e; color: #0f172a; text-align: center; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: bold; }
.btn:hover { background: #16a34a; }
.relatorio-lateral { width: 340px; background: #1e293b; border-left: 1px solid #334155; padding: 13px; overflow-y: auto; }
.relatorio-acoes { display: flex; gap: 5px; margin-bottom: 8px; }
.btn-limpar-log { background: #ef4444; color: white; border: none; border-radius: 5px; padding: 5px 8px; cursor: pointer; font-size: 10px; }
.log-item { background: #0f172a; padding: 8px; border-radius: 6px; margin-bottom: 6px; border-left: 4px solid #334155; font-size: 11px; }
.log-sucesso { border-left-color: #22c55e; }
.log-erro { border-left-color: #ef4444; }
.log-msg-preview { color: #94a3b8; margin-top: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.75); z-index: 9999; align-items: center; justify-content: center; }
.modal-conteudo { width: 600px; max-width: 90%; background: #1e293b; border: 1px solid #334155; border-radius: 10px; padding: 20px; }
.modal-texto { background: #0f172a; padding: 14px; border-radius: 6px; margin: 10px 0; max-height: 350px; overflow-y: auto; white-space: pre-wrap; word-break: break-word; }
.btn-fechar-modal { background: #3b82f6; color: white; border: none; padding: 7px 14px; border-radius: 5px; cursor: pointer; }
@media(max-width: 600px) {
    body {
        height: auto;
        min-height: 100vh;
        overflow-x: hidden;
        overflow-y: auto;
    }

    .topo-fixo {
        position: sticky;
        top: 0;
        padding: 8px;
    }

    .header-painel {
        padding: 9px 10px;
        margin-bottom: 7px;
        gap: 6px;
    }

    .header-painel h2 {
        font-size: 14px;
    }

    .relogio-24h {
        padding: 4px 7px;
        font-size: 10px;
    }

    .header-painel a {
        padding: 5px 7px !important;
        font-size: 10px !important;
    }

    .painel-controles {
        display: flex;
        flex-direction: column;
        gap: 7px;
    }

    .search-form {
        min-width: 0;
        width: 100%;
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 6px;
        padding: 7px;
    }

    .search-box-item {
        min-width: 0;
    }

    .search-box-item:first-child {
        grid-column: 1 / -1;
    }

    .search-box-item label {
        font-size: 9px;
    }

    input[type="text"] {
        padding: 7px 8px;
        font-size: 12px;
    }

    .multiselect-btn {
        min-height: 38px;
        padding: 8px;
        font-size: 11px;
    }

    .multiselect-content {
        max-height: 180px;
    }

    .dropdown-item {
        padding: 8px 6px;
        font-size: 11px;
    }

    .calendario-box,
    .painel-bot {
        width: 100%;
    }

    .calendario-box {
        padding: 6px;
    }

    .main-layout {
        flex-direction: column;
        height: auto;
        overflow: visible;
    }

    .conteudo-scroll {
        padding: 8px;
        overflow: visible;
    }

    .grid {
        grid-template-columns: 1fr;
        gap: 8px;
    }

    .card {
        padding: 10px;
        border-radius: 7px;
    }

    .nome {
        font-size: 14px;
    }

    .info,
    .preview-msg {
        font-size: 11px;
    }

    .relatorio-lateral {
        width: 100%;
        border-left: none;
        border-top: 1px solid #334155;
        height: 260px;
        padding: 9px;
    }
}
</style>
</head>
<body>

<div class="topo-fixo">
    <div class="header-painel">
        <h2>📋 Painel RH - CADS / WhatsApp</h2>
        <div style="display: flex; gap: 10px; align-items: center;">
            <div id="relogio" class="relogio-24h">--:--:--</div>
        </div>
    </div>

    {% if erro %}
    <div style="background:#7f1d1d; padding:8px; border-radius:6px; margin-bottom:8px; font-size:12px;">
        ⚠️ Erro ao atualizar planilha: {{ erro }}
    </div>
    {% endif %}

    <div class="painel-controles">
        <form class="search-form" method="get" id="searchForm" action="/">
            <input type="hidden" name="data_selecionada" id="dataSelecionadaInput" value="{{ data_selecionada }}">

            <div class="search-box-item">
                <label>Pesquisar Colaborador</label>
                <input type="text" id="buscaInput" name="busca" placeholder="Digite o nome..." value="{{ busca }}">
            </div>

            <div class="search-box-item">
                <label>Unidade</label>
                <div class="multiselect-btn" onclick="toggleDropdown('dropdownUnidade')">
                    <span id="labelUnidade">{% if filtro_posto %}{{ filtro_posto|length }} selecionada(s){% else %}Todas{% endif %}</span>
                    <span>▼</span>
                </div>
                <div class="multiselect-content" id="dropdownUnidade">
                    <button type="button" class="filter-clear" onclick="limparFiltro('unidade', event)">✕ Limpar este filtro</button>
                    <label class="dropdown-item" style="font-weight:bold; border-bottom:1px solid #334155;">
                        <input type="checkbox" onchange="toggleTodos(this, 'chk-unidade')"> Marcar Todos
                    </label>
                    {% for u in unidades_disponiveis %}
                    <label class="dropdown-item">
                        <input type="checkbox" name="filtro_posto" value="{{ u }}" class="chk-unidade" {% if u in filtro_posto %}checked{% endif %} onchange="aplicarFiltrosRapido()">
                        {{ u }}
                    </label>
                    {% endfor %}
                </div>
            </div>

            <div class="search-box-item">
                <label>Status</label>
                <div class="multiselect-btn" onclick="toggleDropdown('dropdownStatus')">
                    <span>{% if filtro_status %}{{ filtro_status|length }} selecionado(s){% else %}Todos{% endif %}</span>
                    <span>▼</span>
                </div>
                <div class="multiselect-content" id="dropdownStatus">
                    <button type="button" class="filter-clear" onclick="limparFiltro('status', event)">✕ Limpar este filtro</button>
                    <label class="dropdown-item" style="font-weight:bold; border-bottom:1px solid #334155;">
                        <input type="checkbox" onchange="toggleTodos(this, 'chk-status')"> Marcar Todos
                    </label>
                    {% for s in status_disponiveis %}
                    <label class="dropdown-item">
                        <input type="checkbox" name="filtro_status" value="{{ s }}" class="chk-status" {% if s in filtro_status %}checked{% endif %} onchange="aplicarFiltrosRapido()">
                        {{ s }}
                    </label>
                    {% endfor %}
                </div>
            </div>

            <div class="search-box-item">
                <label>Pendência</label>
                <div class="multiselect-btn" onclick="toggleDropdown('dropdownPendencia')">
                    <span>{% if filtro_pendencia %}{{ filtro_pendencia|length }} selecionada(s){% else %}Todas{% endif %}</span>
                    <span>▼</span>
                </div>
                <div class="multiselect-content" id="dropdownPendencia">
                    <button type="button" class="filter-clear" onclick="limparFiltro('pendencia', event)">✕ Limpar este filtro</button>
                    <label class="dropdown-item" style="font-weight:bold; border-bottom:1px solid #334155;">
                        <input type="checkbox" onchange="toggleTodos(this, 'chk-pendencia')"> Marcar Todos
                    </label>
                    {% for pend in pendencias_disponiveis %}
                    <label class="dropdown-item">
                        <input type="checkbox" name="filtro_pendencia" value="{{ pend }}" class="chk-pendencia" {% if pend in filtro_pendencia %}checked{% endif %} onchange="aplicarFiltrosRapido()">
                        {{ pend }}
                    </label>
                    {% endfor %}
                </div>
            </div>
        </form>

        <div class="calendario-box">
            <div class="cal-header">
                <button type="button" onclick="mudarMes(-1)">◀</button>
                <span id="mesAnoTitulo">Mês Ano</span>
                <button type="button" onclick="mudarMes(1)">▶</button>
            </div>
            <div class="cal-grid" id="calendarioGrid"></div>
            {% if data_selecionada %}
            <button type="button" onclick="limparData()" style="background:none; border:none; color:#f43f5e; cursor:pointer; font-size:10px; margin-top:4px;">❌ Limpar {{ data_selecionada }}</button>
            {% endif %}
        </div>

        <div class="painel-bot">
            <span style="font-size:11px; color:#38bdf8; font-weight:bold;">🤖 Disparo em Massa</span>
            <button type="button" class="btn-acao btn-iniciar" id="btnIniciar" onclick="alternarBot()">▶ Iniciar Massa</button>
            <button type="button" class="btn-acao btn-pausar" onclick="pausarDisparos()">⏸ Pausar</button>
            <div class="cronometro-box">
                <span id="timerTexto">00:00</span>
                <div class="barra-progresso-container">
                    <div class="barra-progresso-fill" id="barraProgresso"></div>
                </div>
            </div>
            <a href="/" class="btn-acao btn-home">🏠 Página Inicial</a>
            <span id="statusBot" style="font-size:10px; color:#94a3b8;">Pronto</span>
        </div>
    </div>
</div>

<div class="main-layout">
    <div class="conteudo-scroll">
        <div class="grid">
            {% if dados %}
                {% for p in dados %}
                {% set card_id = loop.index %}
                <div class="card" id="card-{{ card_id }}"
                     data-nome="{{ p.nome|lower }}"
                     data-telefone="{{ p.telefone_bruto }}"
                     data-posto="{{ p.posto|lower }}"
                     data-status="{{ p.status|lower }}"
                     data-data="{{ p.data_admissao }}">
                    <div class="posto">🏢 {{ p.posto }}</div>
                    <div class="nome">👤 {{ p.nome }}</div>
                    <div class="info">📱 <strong>Telefone:</strong> {{ p.telefone_formatado }}</div>
                    <div class="info">📅 <strong>Data admissão:</strong> {{ p.data_admissao }}</div>
                    {% if p.status %}
                    <div class="pendencia">⚠️ <strong>Pendência:</strong> {{ p.status }}</div>
                    {% endif %}
                    <div class="preview-msg" id="preview-{{ card_id }}">💬 <strong>Mensagem:</strong> {{ p.mensagem }}</div>
                    <a class="btn" href="{{ p.link }}" target="_blank" id="btn-{{ card_id }}" data-linklimpo="{{ p.link }}" onclick="registrarEnvioManual('{{ p.nome|e }}', '{{ p.telefone_bruto }}', {{ card_id }})">💬 Enviar Mensagem</a>
                </div>
                {% endfor %}
            {% else %}
                <div style="grid-column:1/-1; text-align:center; padding:40px; color:#94a3b8;">
                    <h3>📭 Nenhum registro encontrado</h3>
                    <p>Não há colaboradores para os filtros selecionados.</p>
                </div>
            {% endif %}
        </div>
    </div>

    <div class="relatorio-lateral">
        <h3 style="color:#38bdf8; font-size:15px; margin-top:0;">📊 Histórico de Envios</h3>
        <div class="relatorio-acoes">
            <input type="text" id="buscaHistorico" placeholder="Pesquisar..." oninput="filtrarHistorico()">
            <button class="btn-limpar-log" onclick="limparHistorico()">Limpar</button>
        </div>
        <div id="containerLogs">
            {% if historico %}
                {% for h in historico %}
                <div class="log-item {% if h['sucesso'] %}log-sucesso{% else %}log-erro{% endif %}" data-logtext="{{ h['nome']|lower }} {{ h['detalhe']|lower }} {{ h['mensagem']|lower }}">
                    <strong>{{ h['hora'] }}</strong> - {{ h['nome'] }}
                    <div style="color:#cbd5e1; margin-top:3px;">{{ h['detalhe'] }}</div>
                    <div class="log-msg-preview">💬 {{ h['mensagem'] }}</div>
                    <button style="margin-top:5px; background:#334155; color:#38bdf8; border:none; border-radius:4px; padding:4px 7px; cursor:pointer; font-size:10px;" onclick="abrirModalMensagem('{{ h['nome']|e }}', '{{ h['mensagem']|e }}')">🔍 Maximizar</button>
                </div>
                {% endfor %}
            {% else %}
                <p style="color:#64748b; font-size:11px;">Nenhum registro.</p>
            {% endif %}
        </div>
    </div>
</div>

<div class="modal-overlay" id="modalMensagem">
    <div class="modal-conteudo">
        <h3 style="color:#38bdf8; margin-top:0;">💬 Mensagem Enviada <span id="modalNomeColaborador"></span></h3>
        <div class="modal-texto" id="modalTextoConteudo"></div>
        <button class="btn-fechar-modal" onclick="fecharModalMensagem()">Fechar</button>
    </div>
</div>

<script>
// Não recarrega a página automaticamente: isso eliminava filtros e deixava o painel lento.
// A sessão continua válida normalmente.

function atualizarRelogio() {
    const agora = new Date();
    document.getElementById("relogio").innerText = agora.toLocaleTimeString("pt-BR");
}
setInterval(atualizarRelogio, 1000);
atualizarRelogio();

// FILTROS SUPER RÁPIDOS — tudo acontece no navegador, sem recarregar a página.
let timeoutBusca = null;
const inputBusca = document.getElementById("buscaInput");

if (inputBusca) {
    inputBusca.addEventListener("input", function () {
        clearTimeout(timeoutBusca);
        timeoutBusca = setTimeout(aplicarFiltrosRapido, 40);
    });
}

function valoresMarcados(classe) {
    return Array.from(document.querySelectorAll("." + classe + ":checked"))
        .map(cb => cb.value.toLowerCase());
}

function aplicarFiltrosRapido() {
    const busca = (document.getElementById("buscaInput")?.value || "").trim().toLowerCase();
    const postos = valoresMarcados("chk-unidade");
    const status = valoresMarcados("chk-status");
    const pendencias = valoresMarcados("chk-pendencia");
    const data = document.getElementById("dataSelecionadaInput")?.value || "";

    const cards = document.querySelectorAll(".card");
    let visiveis = 0;

    cards.forEach(card => {
        const nome = card.dataset.nome || "";
        const posto = card.dataset.posto || "";
        const cardStatus = card.dataset.status || "";
        const cardData = card.dataset.data || "";

        const okBusca = !busca || nome.includes(busca);
        const okPosto = !postos.length || postos.includes(posto);
        const okStatus = !status.length || status.includes(cardStatus);
        const okPendencia = !pendencias.length || pendencias.includes(cardStatus);

        // O calendário usa data no formato YYYY-MM-DD, enquanto o card está DD/MM/YYYY.
        let okData = true;
        if (data) {
            const partes = data.split("-");
            const dataBR = partes.length === 3 ? `${partes[2]}/${partes[1]}/${partes[0]}` : data;
            okData = cardData === dataBR;
        }

        const mostrar = okBusca && okPosto && okStatus && okPendencia && okData;
        card.style.display = mostrar ? "" : "none";
        if (mostrar) visiveis++;
    });

    atualizarLabelsFiltros(postos, status, pendencias);

    const url = new URL(window.location.href);
    url.search = "";
    if (busca) url.searchParams.set("busca", busca);
    if (data) url.searchParams.set("data_selecionada", data);
    postos.forEach(v => url.searchParams.append("filtro_posto", v));
    status.forEach(v => url.searchParams.append("filtro_status", v));
    pendencias.forEach(v => url.searchParams.append("filtro_pendencia", v));
    history.replaceState(null, "", url);

    // Feedback instantâneo sem reconstruir a tela.
    let vazio = document.getElementById("resultadoVazioRapido");
    if (!visiveis && cards.length) {
        if (!vazio) {
            vazio = document.createElement("div");
            vazio.id = "resultadoVazioRapido";
            vazio.style.cssText = "grid-column:1/-1;text-align:center;padding:40px;color:#94a3b8;";
            vazio.innerHTML = "<h3>📭 Nenhum registro encontrado</h3><p>Não há colaboradores para os filtros selecionados.</p>";
            document.querySelector(".grid").appendChild(vazio);
        }
    } else if (vazio) {
        vazio.remove();
    }
}

function atualizarLabelsFiltros(postos, status, pendencias) {
    const lu = document.getElementById("labelUnidade");
    if (lu) lu.textContent = postos.length ? postos.length + " selecionada(s)" : "Todas";

    const ls = document.querySelector("#dropdownStatus")?.parentElement?.querySelector(".multiselect-btn span");
    if (ls) ls.textContent = status.length ? status.length + " selecionado(s)" : "Todos";

    const lp = document.querySelector("#dropdownPendencia")?.parentElement?.querySelector(".multiselect-btn span");
    if (lp) lp.textContent = pendencias.length ? pendencias.length + " selecionada(s)" : "Todas";
}

function limparFiltro(tipo, event) {
    if (event) event.stopPropagation();
    const mapa = {
        unidade: ".chk-unidade",
        status: ".chk-status",
        pendencia: ".chk-pendencia"
    };
    document.querySelectorAll(mapa[tipo]).forEach(cb => cb.checked = false);

    const dropdown = {
        unidade: "dropdownUnidade",
        status: "dropdownStatus",
        pendencia: "dropdownPendencia"
    }[tipo];
    const el = document.getElementById(dropdown);
    if (el) el.classList.remove("show");
    aplicarFiltrosRapido();
}

function toggleDropdown(id) {
    document.querySelectorAll(".multiselect-content").forEach(function (elemento) {
        if (elemento.id !== id) elemento.classList.remove("show");
    });
    document.getElementById(id).classList.toggle("show");
}

window.onclick = function(event) {
    if (!event.target.closest(".multiselect-btn") && !event.target.closest(".multiselect-content")) {
        document.querySelectorAll(".multiselect-content").forEach(function(el) {
            el.classList.remove("show");
        });
    }
};

function toggleTodos(master, classe) {
    document.querySelectorAll("." + classe).forEach(function(cb) {
        cb.checked = master.checked;
    });
    aplicarFiltrosRapido();
}

function submeterFormulario() {
    aplicarFiltrosRapido();
}

let calendarioData = new Date();
const dataAtualSelecionada = "{{ data_selecionada }}";

function renderizarCalendario() {
    const ano = calendarioData.getFullYear();
    const mes = calendarioData.getMonth();
    const nomesMeses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    document.getElementById("mesAnoTitulo").innerText = nomesMeses[mes] + " " + ano;

    const grid = document.getElementById("calendarioGrid");
    grid.innerHTML = "";

    const diasSemana = ["D", "S", "T", "Q", "Q", "S", "S"];
    diasSemana.forEach(function(dia) {
        const el = document.createElement("div");
        el.className = "cal-day-name";
        el.innerText = dia;
        grid.appendChild(el);
    });

    const primeiroDia = new Date(ano, mes, 1).getDay();
    const ultimoDia = new Date(ano, mes + 1, 0).getDate();

    for (let i = 0; i < primeiroDia; i++) {
        grid.appendChild(document.createElement("div"));
    }

    const hoje = new Date();
    for (let dia = 1; dia <= ultimoDia; dia++) {
        const el = document.createElement("div");
        el.className = "cal-day";
        el.innerText = dia;

        if (dia === hoje.getDate() && mes === hoje.getMonth() && ano === hoje.getFullYear()) {
            el.classList.add("today");
        }

        const dataISO = ano + "-" + String(mes + 1).padStart(2, "0") + "-" + String(dia).padStart(2, "0");
        if (dataAtualSelecionada === dataISO) {
            el.classList.add("selected");
        }

        el.onclick = function() {
            document.getElementById("dataSelecionadaInput").value = dataISO;
            aplicarFiltrosRapido();
        };
        grid.appendChild(el);
    }
}

function mudarMes(quantidade) {
    calendarioData.setMonth(calendarioData.getMonth() + quantidade);
    renderizarCalendario();
}

function limparData() {
    document.getElementById("dataSelecionadaInput").value = "";
    aplicarFiltrosRapido();
}
renderizarCalendario();

// DISPARO RÁPIDO AUTOMATIZADO CORRIGIDO
let rodandoBot = false;
let indiceAtual = 0;
let intervaloBot = null;
let segundosDecorridos = 0;
let timerIntervalo = null;

function alternarBot() {
    const cards = document.querySelectorAll(".card");
    if (!cards.length) {
        alert("Nenhum contato na tela para disparar!");
        return;
    }

    if (!rodandoBot) {
        rodandoBot = true;
        indiceAtual = 0;
        document.getElementById("btnIniciar").innerText = "⏹ Parar Massa";
        document.getElementById("btnIniciar").style.background = "#ef4444";
        document.getElementById("statusBot").innerText = "Enviando em Massa...";

        segundosDecorridos = 0;
        timerIntervalo = setInterval(function() {
            segundosDecorridos++;
            const minutos = String(Math.floor(segundosDecorridos / 60)).padStart(2, "0");
            const segundos = String(segundosDecorridos % 60).padStart(2, "0");
            document.getElementById("timerTexto").innerText = minutos + ":" + segundos;
        }, 1000);

        processarProximoDisparo(cards);

        intervaloBot = setInterval(function() {
            if (!rodandoBot) return;
            processarProximoDisparo(cards);
        }, 9000);
    } else {
        pausarDisparos();
    }
}

function processarProximoDisparo(cards) {
    if (!rodandoBot) return;

    cards.forEach(function(card) {
        card.classList.remove("card-ativo");
    });

    if (indiceAtual >= cards.length) {
        pausarDisparos();
        document.getElementById("statusBot").innerText = "Concluído!";
        return;
    }

    const card = cards[indiceAtual];
    card.classList.add("card-ativo");
    card.scrollIntoView({ behavior: "smooth", block: "nearest" });

    const nome = card.getAttribute("data-nome");
    const telefone = card.getAttribute("data-telefone");
    const botao = card.querySelector(".btn");
    const url = botao.getAttribute("data-linklimpo");

    if (telefone && telefone.length >= 10 && url && url !== "#") {
        const janela = window.open(url, "_blank");
        registrarLogNoBanco(nome, telefone, true, card);
        setTimeout(function() {
            if (janela) janela.close();
        }, 7000);
    } else {
        registrarLogNoBanco(nome, telefone, false, card);
    }

    indiceAtual++;
    const progresso = (indiceAtual / cards.length) * 100;
    document.getElementById("barraProgresso").style.width = progresso + "%";
}

function pausarDisparos() {
    rodandoBot = false;
    if (intervaloBot) { clearInterval(intervaloBot); intervaloBot = null; }
    if (timerIntervalo) { clearInterval(timerIntervalo); timerIntervalo = null; }
    document.getElementById("btnIniciar").innerText = "▶ Iniciar Massa";
    document.getElementById("btnIniciar").style.background = "#22c55e";
    document.getElementById("statusBot").innerText = "Pausado";
}

function registrarLogNoBanco(nome, telefone, sucesso, card) {
    const preview = card.querySelector(".preview-msg");
    let mensagem = preview ? preview.innerText.replace("💬 Mensagem:", "").trim() : "";

    fetch("/registrar-envio", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ nome: nome, telefone: telefone, sucesso: sucesso, detalhe: sucesso ? "Contato enviado" : "Contato inválido", mensagem: mensagem })
    });
}

function registrarEnvioManual(nome, telefone, cardId) {
    const card = document.getElementById("card-" + cardId);
    if (!card) return;
    registrarLogNoBanco(nome, telefone, !!(telefone && telefone.length >= 10), card);
}

function filtrarHistorico() {
    const busca = document.getElementById("buscaHistorico").value.toLowerCase();
    document.querySelectorAll(".log-item").forEach(function(item) {
        const texto = item.getAttribute("data-logtext") || "";
        item.style.display = texto.includes(busca) ? "" : "none";
    });
}

function limparHistorico() {
    if (!confirm("Deseja realmente limpar todo o histórico?")) return;
    fetch("/limpar-historico", { method: "POST" }).then(function() { location.reload(); });
}

function abrirModalMensagem(nome, mensagem) {
    document.getElementById("modalNomeColaborador").innerText = "- " + nome;
    document.getElementById("modalTextoConteudo").innerText = mensagem;
    document.getElementById("modalMensagem").style.display = "flex";
}

function fecharModalMensagem() {
    document.getElementById("modalMensagem").style.display = "none";
}
</script>
</body>
</html>
"""


# ============================================================
# EXECUÇÃO
# ============================================================



if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=True)
