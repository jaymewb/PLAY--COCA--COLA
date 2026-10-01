import base64
from datetime import datetime
from io import BytesIO
import os
import re
import sqlite3
import urllib.parse
import json

from flask import Flask, request, render_template_string, jsonify, session, redirect, url_for
import pandas as pd
import requests

# ============================================================
# CONFIGURAÇÃO
# ============================================================

app = Flask(__name__)
app.secret_key = "1020_chave_secreta_cads"

SHEET_URL = "https://1drv.ms/x/c/b96adcc2e8fff38f/IQC1usMRQOkVSI0pS5n6PpnEAUoTEYza4vN2utVrIm0_UYU?e=Opi2eN&nav=MTVfezA5Mzc2MUE0LTIxOUMtNEQ2MC04QkM4LURCNUJDNzk2RDQwM30"

ABA_PLANILHA = "TESTE PRÉ ADMISSÃO"
DB_NAME = "rh_escala.db"


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

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cadastro_colaboradores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            telefone TEXT,
            data_admissao TEXT,
            status TEXT,
            unidade TEXT,
            cargo TEXT,
            origem TEXT DEFAULT "principal",
            dados_completos TEXT,
            atualizado_em TEXT,
            UNIQUE(nome, telefone)
        )
    """)

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


def garantir_colunas_banco():
    conn = conectar()
    c = conn.cursor()
    cols = [r[1] for r in c.execute("PRAGMA table_info(cadastro_colaboradores)").fetchall()]
    if "cargo" not in cols: c.execute("ALTER TABLE cadastro_colaboradores ADD COLUMN cargo TEXT")
    if "origem" not in cols: c.execute("ALTER TABLE cadastro_colaboradores ADD COLUMN origem TEXT DEFAULT 'principal'")
    if "foto" not in cols: c.execute("ALTER TABLE cadastro_colaboradores ADD COLUMN foto TEXT")
    conn.commit()
    conn.close()


garantir_colunas_banco()


# ============================================================
# NORMALIZAÇÃO DE TEXTO
# ============================================================

def normalizar_texto(valor):
    if valor is None or pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.lower() in ("nan", "nat", "none"):
        return ""
    texto = texto.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", texto).strip()


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
    str_val = str(valor).strip()
    if str_val.lower() in ("nat", "nan", "none", ""):
        return ""
    try:
        if isinstance(valor, datetime):
            return valor.strftime("%d/%m/%Y")
        data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if pd.isna(data):
            return str_val
        return data.strftime("%d/%m/%Y")
    except Exception:
        return str_val


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
    urls = [SHEET_URL, SHEET_URL + ("&" if "?" in SHEET_URL else "?") + "download=1"]
    ultimo_erro = None

    for url in urls:
        try:
            resposta = requests.get(url, timeout=30, allow_redirects=True,
                                    headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
            resposta.raise_for_status()
            conteudo = resposta.content
            if conteudo[:2] == b"PK":
                return conteudo
            if "excel" in resposta.headers.get("Content-Type", "").lower() or "spreadsheet" in resposta.headers.get(
                    "Content-Type", "").lower():
                return conteudo
            texto = conteudo.decode("utf-8", errors="ignore")
            possiveis_urls = re.findall(r'https?://[^"\']+', texto)
            for possivel in possiveis_urls:
                possivel = possivel.replace("\\u0026", "&").replace("&amp;", "&")
                if any(x in possivel.lower() for x in ("download", "onedrive", "1drv.ms", "sharepoint")):
                    try:
                        r2 = requests.get(possivel, timeout=30, allow_redirects=True,
                                          headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
                        if r2.ok and r2.content[:2] == b"PK":
                            return r2.content
                    except Exception:
                        pass
        except Exception as erro:
            ultimo_erro = erro

    raise RuntimeError(f"Não foi possível baixar a planilha do OneDrive: {ultimo_erro}")


def extrair_fotos_embutidas(arquivo, nome_aba):
    fotos = {}
    try:
        from openpyxl import load_workbook as owb
        wb = owb(filename=BytesIO(arquivo), read_only=False, data_only=False)
        if nome_aba not in wb.sheetnames:
            return fotos
        ws = wb[nome_aba]
        for img in getattr(ws, "_images", []):
            try:
                linha = img.anchor._from.row
                dados = img._data()
                if not dados:
                    continue
                extensao = str(getattr(img, "format", "") or "").lower()
                mime = "image/jpeg"
                if extensao in ("png", "gif", "webp"):
                    mime = f"image/{extensao}"
                fotos[linha - 1] = f"data:{mime};base64," + base64.b64encode(dados).decode("ascii")
            except Exception:
                continue
        wb.close()
    except Exception:
        pass
    return fotos


def carregar_planilha():
    arquivo = baixar_planilha()
    excel = pd.ExcelFile(BytesIO(arquivo))
    aba_encontrada = next((a for a in excel.sheet_names
                           if normalizar_texto(a).upper() == normalizar_texto(ABA_PLANILHA).upper()), None)
    if not aba_encontrada:
        raise RuntimeError(
            f"A aba '{ABA_PLANILHA}' não foi encontrada. Abas disponíveis: {', '.join(excel.sheet_names)}")

    df = pd.read_excel(BytesIO(arquivo), sheet_name=aba_encontrada, dtype=object, header=0)
    df = df.dropna(axis=0, how="all")
    df.columns = [normalizar_texto(c) for c in df.columns]

    fotos = extrair_fotos_embutidas(arquivo, aba_encontrada)
    df["__foto__"] = [fotos.get(i, "") for i in range(len(df))]
    return df


def coluna_por_letra(df, letra):
    numero = 0
    for caractere in letra.upper():
        numero = numero * 26 + (ord(caractere) - ord('A') + 1)
    indice = numero - 1
    if indice < 0 or indice >= len(df.columns):
        raise RuntimeError(
            f"A coluna {letra} não existe na aba '{ABA_PLANILHA}'. A planilha possui {len(df.columns)} colunas.")
    return df.columns[indice]


def descobrir_colunas_candidatos_por_posicao(df):
    return {
        "nome": coluna_por_letra(df, "E"),
        "telefone": coluna_por_letra(df, "G"),
        "data_admissao": coluna_por_letra(df, "K"),
        "unidade": coluna_por_letra(df, "M"),
        "observacao_admissao": coluna_por_letra(df, "AH"),
        "foto": "__foto__",
    }


def sincronizar_cadastro(df):
    if df is None or df.empty:
        return 0, {}
    col = descobrir_colunas_candidatos_por_posicao(df)
    return _sincronizar_dataframe(df, col, "principal", col["observacao_admissao"])


def _sincronizar_dataframe(df, col, origem, status_col):
    conn = conectar()
    cur = conn.cursor()
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    qtd = 0
    vistos = set()

    try:
        cur.execute("DELETE FROM cadastro_colaboradores")

        for _, row in df.iterrows():
            nome = normalizar_texto(row.get(col["nome"], ""))
            if not nome:
                continue

            telefone = limpar_telefone(row.get(col["telefone"], ""))
            chave = (normalizar_coluna(nome), telefone)

            if chave in vistos:
                continue
            vistos.add(chave)

            adm = formatar_data(row.get(col["data_admissao"], ""))

            status_val = row.get(status_col, "")
            status = normalizar_texto(status_val)
            if not status or status.lower() in ("nan", "nat", "none"):
                status = "Nenhuma"

            unidade_val = row.get(col.get("unidade"), "") if col.get("unidade") else ""
            unidade = normalizar_texto(unidade_val)
            if not unidade or unidade.lower() in ("nan", "nat", "none"):
                unidade = "Unidade não informada"

            cargo = normalizar_texto(row.get(col.get("cargo"), "")) if col.get("cargo") else ""
            foto = normalizar_texto(row.get(col.get("foto"), "")) if col.get("foto") else ""

            dados = {}
            for k in df.columns:
                if k == "__foto__":
                    continue
                v = row.get(k, "")
                if pd.isna(v) or str(v).lower() in ("nan", "nat", "none"):
                    v = ""
                v = v.strftime("%d/%m/%Y") if isinstance(v, datetime) else v
                dados[str(k)] = str(v)

            js = json.dumps(dados, ensure_ascii=False)
            cur.execute(
                """INSERT INTO cadastro_colaboradores
                (nome, telefone, data_admissao, status, unidade, cargo, origem, dados_completos, atualizado_em, foto)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (nome, telefone, adm, status, unidade, cargo, origem, js, agora, foto)
            )
            qtd += 1

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    return qtd, col


def carregar_cadastro():
    conn = conectar()
    registros = conn.execute("SELECT * FROM cadastro_colaboradores ORDER BY nome COLLATE NOCASE").fetchall()
    conn.close()
    return registros


def aplicar_filtros(registros, busca, filtro_posto, filtro_status, filtro_pendencia, data_selecionada):
    resultado = []
    for registro in registros:
        nome = registro["nome"] or ""
        status = registro["status"] or ""
        unidade = registro["unidade"] or ""
        data_admissao = registro["data_admissao"] or ""

        if busca and busca.lower() not in nome.lower():
            continue
        if filtro_posto and unidade.lower() not in [f.lower() for f in filtro_posto]:
            continue
        if filtro_status and status.lower() not in [f.lower() for f in filtro_status]:
            continue
        if filtro_pendencia and status.lower() not in [f.lower() for f in filtro_pendencia]:
            continue
        if data_selecionada:
            data_filtro = converter_data_filtro(data_selecionada)
            if data_admissao != data_filtro:
                continue

        resultado.append(registro)
    return resultado


def montar_card(registro):
    nome = registro["nome"] or ""
    telefone = limpar_telefone(registro["telefone"] or "")
    status = registro["status"] or "Nenhuma"
    cargo = registro["cargo"] if "cargo" in registro.keys() else ""
    primeiro_nome = nome.split()[0] if nome.split() else ""
    hora = datetime.now().hour

    saudacao = "bom dia" if hora < 12 else ("boa tarde" if hora < 18 else "boa noite")

    if status.lower() == "nenhuma":
        mensagem = f"Prezado(a) {primeiro_nome}, {saudacao}! Entramos em contato referente ao seu processo de admissão."
    else:
        mensagem = f"Prezado(a) {primeiro_nome}, {saudacao}! Identificamos a pendência: {status}."

    link = f"https://api.whatsapp.com/send?phone=55{telefone}&text=" + urllib.parse.quote(mensagem) if (
                telefone and len(telefone) >= 10) else "#"

    return {
        "id": registro["id"],
        "nome": nome,
        "telefone_bruto": telefone,
        "telefone_formatado": formatar_telefone(telefone),
        "data_admissao": registro["data_admissao"],
        "cargo": cargo,
        "status": status,
        "pendencia": status,
        "mensagem": mensagem,
        "posto": registro["unidade"] or "Unidade não informada",
        "link": link
    }


def buscar_historico():
    conn = conectar()
    dados = conn.execute(
        "SELECT id, hora, nome, sucesso, detalhe, mensagem FROM historico_envios ORDER BY id DESC").fetchall()
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

@app.route("/login", methods=["GET", "POST"])
def login():
    erro_login = None
    if request.method == "POST":
        senha = request.form.get("senha", "")
        if senha == "1020":
            session["autenticado"] = True
            return redirect(url_for("index"))
        else:
            erro_login = "Senha incorreta! Use 1020."
    return render_template_string(HTML_LOGIN, erro_login=erro_login)


@app.route("/logout")
def logout():
    session.pop("autenticado", None)
    return redirect(url_for("login"))


@app.route("/sincronizar", methods=["GET"])
def sincronizar():
    if not session.get("autenticado"):
        return redirect(url_for("login"))
    try:
        df = carregar_planilha()
        sincronizados, _ = sincronizar_cadastro(df)
        return redirect(url_for("index", sincronizados=sincronizados))
    except Exception as e:
        return redirect(url_for("index", erro=str(e)))


@app.route("/", methods=["GET"])
def index():
    if not session.get("autenticado"):
        return redirect(url_for("login"))

    erro = request.args.get("erro")
    sincronizados = request.args.get("sincronizados", 0)

    busca = request.args.get("busca", "").strip()
    filtro_posto = request.args.getlist("filtro_posto")
    filtro_status = request.args.getlist("filtro_status")
    filtro_pendencia = request.args.getlist("filtro_pendencia")
    data_selecionada = request.args.get("data_selecionada", "").strip()

    registros = carregar_cadastro()
    registros_filtrados = aplicar_filtros(registros, busca, filtro_posto, filtro_status, filtro_pendencia,
                                          data_selecionada)
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
    if not session.get("autenticado"):
        return jsonify({"ok": False}), 403
    dados = request.get_json(silent=True) or {}
    registrar_log(dados.get("nome", ""), dados.get("sucesso", True),
                  dados.get("detalhe", f"Contato: {dados.get('telefone', '')}"), dados.get("mensagem", ""))
    return jsonify({"ok": True})


@app.route("/limpar-historico", methods=["POST"])
def limpar_historico():
    if not session.get("autenticado"):
        return jsonify({"ok": False}), 403
    conn = conectar()
    conn.execute("DELETE FROM historico_envios")
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


# ============================================================
# HTML LOGIN E PAINEL
# ============================================================

HTML_LOGIN = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Login - Painel RH</title>
<style>
body { font-family: "Segoe UI", Arial, sans-serif; background: #0f172a; color: white; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
.login-box { background: #1e293b; padding: 30px; border-radius: 10px; border: 1px solid #334155; width: 320px; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.5); }
input { width: 100%; padding: 10px; margin: 15px 0; border-radius: 6px; border: 1px solid #475569; background: #0f172a; color: white; outline: none; box-sizing: border-box; }
button { width: 100%; padding: 10px; background: #38bdf8; color: #0f172a; font-weight: bold; border: none; border-radius: 6px; cursor: pointer; }
button:hover { background: #0ea5e9; }
.erro { color: #f43f5e; font-size: 12px; margin-bottom: 10px; }
</style>
</head>
<body>
<div class="login-box">
    <h2>🔒 Restrito</h2>
    <p style="font-size: 13px; color: #94a3b8;">Digite a senha de acesso (1020)</p>
    {% if erro_login %}<div class="erro">{{ erro_login }}</div>{% endif %}
    <form method="post">
        <input type="password" name="senha" placeholder="Senha..." required autofocus>
        <button type="submit">Entrar</button>
    </form>
</div>
</body>
</html>
"""

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
.multiselect-btn { padding: 8px 10px; border: 1px solid #475569; border-radius: 7px; background: #0f172a; color: white; cursor: pointer; display: flex; justify-content: space-between; align-items: center; font-size: 12px; }
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
@media(max-width: 1000px) {
    .main-layout { flex-direction: column; height: auto; overflow-y: auto; }
    .relatorio-lateral { width: 100%; border-left: none; border-top: 1px solid #334155; height: 300px; }
    .search-form { min-width: 100%; }
}
</style>
</head>
<body>

<div class="topo-fixo">
    <div class="header-painel">
        <div style="display: flex; gap: 15px; align-items: center;">
            <h2>📋 Painel RH - CADS / WhatsApp</h2>
            <a href="/sincronizar" style="background:#38bdf8; color:#0f172a; padding:6px 12px; border-radius:6px; text-decoration:none; font-size:11px; font-weight:bold;">🔄 Sincronizar Planilha</a>
        </div>
        <div style="display: flex; gap: 10px; align-items: center;">
            <div id="relogio" class="relogio-24h">--:--:--</div>
            <a href="/logout" style="background:#ef4444; color:white; padding:6px 10px; border-radius:6px; text-decoration:none; font-size:11px; font-weight:bold;">Sair</a>
        </div>
    </div>

    {% if erro %}
    <div style="background:#7f1d1d; padding:8px; border-radius:6px; margin-bottom:8px; font-size:12px;">
        ⚠️️ Erro ao atualizar planilha: {{ erro }}
    </div>
    {% endif %}

    {% if sincronizados %}
    <div style="background:#065f46; padding:8px; border-radius:6px; margin-bottom:8px; font-size:12px;">
        ✅ Planilha sincronizada com sucesso! {{ sincronizados }} registos atualizados.
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
                    <label class="dropdown-item" style="font-weight:bold; border-bottom:1px solid #334155;">
                        <input type="checkbox" onchange="toggleTodos(this, 'chk-unidade')"> Marcar Todos
                    </label>
                    {% for u in unidades_disponiveis %}
                    <label class="dropdown-item">
                        <input type="checkbox" name="filtro_posto" value="{{ u }}" class="chk-unidade" {% if u in filtro_posto %}checked{% endif %} onchange="submeterFormulario()">
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
                    <label class="dropdown-item" style="font-weight:bold; border-bottom:1px solid #334155;">
                        <input type="checkbox" onchange="toggleTodos(this, 'chk-status')"> Marcar Todos
                    </label>
                    {% for s in status_disponiveis %}
                    <label class="dropdown-item">
                        <input type="checkbox" name="filtro_status" value="{{ s }}" class="chk-status" {% if s in filtro_status %}checked{% endif %} onchange="submeterFormulario()">
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
                    <label class="dropdown-item" style="font-weight:bold; border-bottom:1px solid #334155;">
                        <input type="checkbox" onchange="toggleTodos(this, 'chk-pendencia')"> Marcar Todos
                    </label>
                    {% for pend in pendencias_disponiveis %}
                    <label class="dropdown-item">
                        <input type="checkbox" name="filtro_pendencia" value="{{ pend }}" class="chk-pendencia" {% if pend in filtro_pendencia %}checked{% endif %} onchange="submeterFormulario()">
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
                <div class="card" id="card-{{ card_id }}" data-nome="{{ p.nome }}" data-telefone="{{ p.telefone_bruto }}">
                    <div class="posto">🏢 {{ p.posto }}</div>
                    <div class="nome">👤 {{ p.nome }}</div>
                    <div class="info">📱 <strong>Telefone:</strong> {{ p.telefone_formatado }}</div>
                    <div class="info">📅 <strong>Data admissão:</strong> {{ p.data_admissao }}</div>
                    {% if p.cargo %}<div class="info">💼 <strong>Cargo:</strong> {{ p.cargo }}</div>{% endif %}
                    <div class="pendencia">⚠️ <strong>Pendência:</strong> {{ p.status }}</div>
                    <div class="preview-msg" id="preview-{{ card_id }}">💬 <strong>Mensagem:</strong> {{ p.mensagem }}</div>
                    <a class="btn" href="{{ p.link }}" target="_blank" id="btn-{{ card_id }}" data-linklimpo="{{ p.link }}" onclick="registrarEnvioManual('{{ p.nome|e }}', '{{ p.telefone_bruto }}', {{ card_id }})">💬 Enviar Mensagem</a>
                </div>
                {% endfor %}
            {% else %}
                <div style="grid-column:1/-1; text-align:center; padding:40px; color:#94a3b8;">
                    <h3>📭 Nenhum registo encontrado</h3>
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
                <p style="color:#64748b; font-size:11px;">Nenhum registo.</p>
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
function atualizarRelogio() {
    const agora = new Date();
    document.getElementById("relogio").innerText = agora.toLocaleTimeString("pt-BR");
}
setInterval(atualizarRelogio, 1000);
atualizarRelogio();

let timeoutBusca = null;
const inputBusca = document.getElementById("buscaInput");
if (inputBusca) {
    inputBusca.addEventListener("input", function () {
        clearTimeout(timeoutBusca);
        timeoutBusca = setTimeout(function () {
            document.getElementById("searchForm").submit();
        }, 300);
    });
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
    document.getElementById("searchForm").submit();
}

function submeterFormulario() {
    document.getElementById("searchForm").submit();
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
            document.getElementById("searchForm").submit();
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
    document.getElementById("searchForm").submit();
}
renderizarCalendario();

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
    print("=" * 70)
    print("AVI - PAINEL RH / CADS / WHATSAPP (Corrigido e Otimizado)")
    print("=" * 70)
    print()
    print("Banco:", DB_NAME)
    print("Planilha:", SHEET_URL)
    print()
    print("Acesse no navegador:")
    print("http://127.0.0.1:5000")
    print()
    print("=" * 70)

    app.run(host="0.0.0.0", port=5000, debug=True)