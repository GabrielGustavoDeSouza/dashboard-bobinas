"""
Dashboard de Controle de Matéria-Prima — Bobinas BSW | Grupo Delga

- Visitantes veem os dados automaticamente (sem upload)
- Admin atualiza os dados via upload protegido por senha (Secrets: ADMIN_PASSWORD)
- Dados persistem no GitHub (não somem quando o app dorme)

Fontes de dados, em ordem de prioridade:
  1. GitHub   (Secrets: GITHUB_TOKEN)
  2. SharePoint (Secrets: TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHAREPOINT_*)
  3. Arquivo local data/dados_atuais.xlsx (útil para testar no computador)
"""
import base64
import hmac
import html as html_lib
import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import streamlit.components.v1 as components

# ============================================================
# CONFIGURAÇÃO
# ============================================================
st.set_page_config(
    page_title="Grupo Delga | Dashboard Bobinas BSW",
    page_icon="🔵",
    layout="wide",
    initial_sidebar_state="collapsed",
)

GITHUB_REPO = "GabrielGustavoDeSouza/dashboard-bobinas"
GITHUB_DATA_PATH = "data/dados_atuais.xlsx"
GITHUB_BRANCH = "main"
LOCAL_DATA_PATH = Path(__file__).parent / "data" / "dados_atuais.xlsx"
ASSETS_DIR = Path(__file__).parent / "assets"

AZUL = "#1400FF"
UNIDADE_COLORS = {"Ferraz": "#1E88E5", "Diadema": "#43A047", "Jarinu": "#FB8C00", "Sul": "#8E24AA"}
ORDEM_PLANTAS = ["FERRAZ", "DIADEMA", "JARINU", "SUL"]

COR = {
    "azul": "#4D6BFF", "azul_claro": "#4DA3FF", "verde": "#16A34A", "laranja": "#F59E0B",
    "vermelho": "#E53E3E", "coral": "#F97066", "roxo": "#7C3AED", "cinza": "#94A3B8",
    "texto": "#1F2937", "texto2": "#64748B", "borda": "#E2E6F0",
}
CHART_COLORS = ["#4D6BFF", "#F59E0B", "#16A34A", "#7C3AED", "#F97066", "#06B6D4",
                "#EAB308", "#22C55E", "#A78BFA", "#FB7185", "#0EA5E9", "#84CC16"]

PLOTLY_LAYOUT = dict(
    paper_bgcolor="#FFFFFF",
    plot_bgcolor="#FFFFFF",
    font=dict(color="#475569", family="Inter, Arial", size=12),
    margin=dict(l=20, r=20, t=50, b=20),
    legend=dict(bgcolor="rgba(255,255,255,0)", font=dict(color="#64748B")),
)
GRID = dict(gridcolor="#EEF1F6", zerolinecolor="#EEF1F6")

# Etapas da aba "A.Propostas": (coluna_excel, label, conta_no_pct)
STAGE_DEFS = [
    ("FORMALIZADO COM COMPRAS", "Formaliz. c/ Compras", True),
    ("DATA ENVIO P/ USINA", "Envio à Usina", True),
    ("PRAZO DA USINA PARA RETORNO", "Prazo Retorno Usina", True),
    ("DATA DE DEVOLUÇÃO DA CONSULTA", "Devolução Consulta", True),
    ("ENVIO DE PLANO DE CORTE PARA COMPRAS", "Plano Corte → Compras", True),
    ("ENVIO DO PLANO DE CORTE PARA USINAS", "Plano Corte → Usina", True),
    ("PRAZO PARA CADASTRO DO NOVO PN (LIBERAR PROGRAMAÇÃO)", "Prazo Cadastro PN", True),
    ("DATA DE CADASTRO DO NOVO PN", "Cadastro PN", False),  # informativa, não conta no %
    ("PRAZO DE RECEBIMENTO NAS NOVAS ESPECIFICAÇÕES (APROXIMADO)", "Recebimento", True),
]
DONE = ("done", "na")

# ============================================================
# ESTILO (tema claro fixado em .streamlit/config.toml)
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, .stApp { font-family: 'Inter', sans-serif; }
.stApp { background:#F4F6FB; }
header[data-testid="stHeader"] { background:transparent; }
.block-container { padding-top:1.6rem; padding-bottom:3rem; max-width:1500px; }
h1,h2,h3,h4 { color:#1F2937 !important; font-family:'Inter',sans-serif; }

/* Cabeçalho */
.app-head { display:flex; align-items:center; justify-content:space-between; gap:16px; flex-wrap:wrap;
            padding-bottom:14px; margin-bottom:18px; border-bottom:1px solid #E2E6F0; }
.app-head-l { display:flex; align-items:center; gap:14px; }
.app-logo { background:linear-gradient(135deg,#1400FF,#0A00AA); color:#fff; font-weight:800; font-size:18px;
            border-radius:10px; padding:10px 12px; }
.app-title { font-size:22px; font-weight:800; color:#1F2937; line-height:1.1; }
.app-sub { font-size:11.5px; color:#64748B; letter-spacing:.5px; text-transform:uppercase; margin-top:3px; }
.app-upd { font-size:12px; color:#64748B; background:#fff; border:1px solid #E2E6F0; border-radius:20px; padding:5px 12px; }

/* Seções */
.sec { margin:26px 0 12px 0; }
.sec-t { font-size:16px; font-weight:700; color:#1F2937; }
.sec-s { font-size:12.5px; color:#64748B; margin-top:2px; }

/* Big numbers */
.kpi-row { display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:12px; }
.kpi { background:#fff; border:1px solid #E2E6F0; border-radius:12px; padding:14px 16px;
       box-shadow:0 1px 3px rgba(15,23,42,.04); border-top:3px solid var(--c,#4D6BFF); }
.kpi-l { font-size:11px; color:#64748B; font-weight:600; text-transform:uppercase; letter-spacing:.4px; }
.kpi-v { font-size:26px; font-weight:800; color:#1F2937; margin-top:4px; line-height:1.1; }
.kpi-s { font-size:11.5px; color:#94A3B8; margin-top:3px; }

/* Abas */
.stTabs [data-baseweb="tab"] p { font-weight:600; font-size:14px; }

/* Cartão de gráfico */
div[data-testid="stPlotlyChart"] { background:#fff; border:1px solid #E2E6F0; border-radius:12px; padding:4px; }

/* Tabela clara */
.lt-wrap { border:1px solid #E2E6F0; border-radius:12px; overflow:hidden; background:#fff; }
table.lt { width:100%; border-collapse:collapse; }
table.lt th { background:#F8FAFD; color:#64748B; text-align:left; font-size:11px; text-transform:uppercase;
              letter-spacing:.4px; font-weight:700; padding:9px 12px; border-bottom:1px solid #E2E6F0; }
table.lt td { color:#1F2937; font-size:13px; padding:9px 12px; border-bottom:1px solid #F1F3F8; }
table.lt td.n, table.lt th.n { text-align:right; font-variant-numeric:tabular-nums; }
table.lt tr:last-child td { border-bottom:none; }

/* ===== Linha do tempo ===== */
details.blk { background:#fff; border:1px solid #E2E6F0; border-radius:12px; margin-bottom:12px; overflow:hidden; }
details.blk > summary { list-style:none; cursor:pointer; display:flex; align-items:center; gap:12px;
                        padding:12px 16px; background:#F8FAFD; }
details.blk > summary::-webkit-details-marker { display:none; }
details.blk > summary::before { content:"▸"; color:#94A3B8; font-size:12px; transition:transform .15s; }
details.blk[open] > summary::before { transform:rotate(90deg); }
details.blk[open] > summary { border-bottom:1px solid #E2E6F0; }
.blk-dot { width:12px; height:12px; border-radius:3px; }
.blk-name { font-weight:700; color:#1F2937; font-size:14.5px; }
.blk-sub { color:#64748B; font-size:12px; }
.blk-stats { margin-left:auto; color:#64748B; font-size:12px; }
.blk-stats b { color:#1F2937; font-size:14px; }
.blk-body { overflow-x:auto; }
.blk-inner { min-width:1050px; }

.tl-row { display:flex; align-items:center; gap:16px; padding:12px 16px; border-bottom:1px solid #F1F3F8; }
.tl-row:last-child { border-bottom:none; }
.tl-row:hover { background:#FAFBFE; }
.tl-row.inv { background:#FFF7F7; }
.tl-head { background:#FCFDFE; padding-top:10px; padding-bottom:10px; position:sticky; top:0; }
.tl-info { width:250px; flex-shrink:0; min-width:0; }
.tl-top { display:flex; align-items:center; gap:8px; }
.tl-code { font-weight:700; color:#1F2937; font-size:13.5px; white-space:nowrap; }
.tl-badge { font-size:10px; font-weight:700; padding:2px 8px; border-radius:20px; white-space:nowrap; }
.tl-desc { color:#64748B; font-size:11.5px; margin-top:3px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.tl-meta { color:#94A3B8; font-size:10.5px; margin-top:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.tl-hl { font-size:10px; color:#94A3B8; font-weight:700; text-transform:uppercase; letter-spacing:.5px; }
.tl-track { flex:1; display:flex; min-width:0; }
.tl-cell { flex:1; position:relative; display:flex; flex-direction:column; align-items:center; min-width:0; }
.tl-cell .ln { position:absolute; top:9px; height:3px; width:50%; background:#E2E6F0; }
.tl-cell .ln.l { left:0; } .tl-cell .ln.r { right:0; }
.tl-cell .ln.on { background:#1400FF; }
.tl-cell .ln.bad { background:#FECACA; }
.tl-dot { width:20px; height:20px; border-radius:50%; border:3px solid #D1D7E3; background:#fff; box-sizing:border-box; z-index:1; }
.tl-dot.done, .tl-dot.na { background:#1400FF; border-color:#1400FF; }
.tl-dot.planned { border-color:#4DA3FF; }
.tl-dot.pending { border-color:#F59E0B; }
.tl-dot.info { border-style:dashed; opacity:.75; }
.tl-dot.bad { background:#FCA5A5; border-color:#EF4444; }
.tl-dot.tip { cursor:help; box-shadow:0 0 0 4px rgba(20,0,255,.15); }
.tl-val { font-size:10px; color:#64748B; margin-top:5px; text-align:center; line-height:1.2; padding:0 2px;
          max-width:100%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tl-lbl { font-size:10px; color:#334155; font-weight:600; text-align:center; line-height:1.25; margin-top:4px; padding:0 2px; }
.tl-lbl.info { color:#94A3B8; font-style:italic; font-weight:500; }
.tl-pill { font-size:8.5px; font-weight:700; text-transform:uppercase; letter-spacing:.4px; padding:1px 6px; border-radius:10px;
           background:#EEF1FB; color:#1400FF; border:1px solid #CBD8FB; max-width:100%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tl-pill.eng { background:#FFF7ED; color:#C2610C; border-color:#FED7AA; }
.tl-pill.log { background:#F0FDF4; color:#166534; border-color:#BBF7D0; }
.tl-pill.sol { background:#FAF5FF; color:#7C3AED; border-color:#DDD6FE; }
.tl-pill.all, .tl-pill.none { background:#F8FAFC; color:#94A3B8; border-color:#E2E6F0; }
.tl-pct { width:62px; flex-shrink:0; text-align:right; font-size:18px; font-weight:800; color:#1F2937; }
.tl-pct.inv { color:#E53E3E; font-size:12px; }
.legend { font-size:11.5px; color:#64748B; margin:0 0 10px 2px; }
.legend span { margin-right:14px; white-space:nowrap; }
.legend i { display:inline-block; width:10px; height:10px; border-radius:50%; border:2px solid; vertical-align:-1px; margin-right:4px; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# UTILITÁRIOS
# ============================================================
def get_secret(key):
    try:
        return st.secrets[key]
    except Exception:
        return None


def esc(v):
    return html_lib.escape(str(v))


def fmt_num(v, dec=0):
    """Formata número no padrão brasileiro: 12.345,6"""
    try:
        s = f"{float(v):,.{dec}f}"
    except (TypeError, ValueError):
        return "—"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def get_unidade_color(nome):
    for key, color in UNIDADE_COLORS.items():
        if key.lower() in str(nome).lower():
            return color
    return COR["azul"]


def parse_numero_brasileiro(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    txt = str(valor).strip().replace("R$", "").replace(" ", "").replace("\xa0", "")
    if not txt or txt.lower() in ("nan", "none", "-"):
        return 0.0
    if "," in txt and "." in txt:
        if txt.rfind(",") > txt.rfind("."):
            txt = txt.replace(".", "").replace(",", ".")
        else:
            txt = txt.replace(",", "")
    elif "," in txt:
        txt = txt.replace(".", "").replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return 0.0


def normalizar_texto(valor, default="Desconhecida"):
    if pd.isna(valor):
        return default
    txt = " ".join(str(valor).replace("\n", " / ").split())
    return txt.upper() if txt else default


def find_col(df, contains=None, exact=None):
    for c in df.columns:
        u = str(c).strip().upper()
        if exact is not None and u == exact:
            return c
        if contains is not None and contains in u:
            return c
    return None


def ordenar_plantas(plantas):
    return sorted(plantas, key=lambda p: (ORDEM_PLANTAS.index(p) if p in ORDEM_PLANTAS else 99, p))


# ============================================================
# COMPONENTES VISUAIS
# ============================================================
def section(title, subtitle=None):
    sub = f'<div class="sec-s">{esc(subtitle)}</div>' if subtitle else ""
    st.markdown(f'<div class="sec"><div class="sec-t">{esc(title)}</div>{sub}</div>', unsafe_allow_html=True)


def kpi_row(items):
    """items: lista de (rótulo, valor, subtítulo|None, cor|None)."""
    cells = []
    for label, value, sub, color in items:
        sub_html = f'<div class="kpi-s">{esc(sub)}</div>' if sub else ""
        cells.append(
            f'<div class="kpi" style="--c:{color or COR["azul"]}">'
            f'<div class="kpi-l">{esc(label)}</div><div class="kpi-v">{esc(value)}</div>{sub_html}</div>'
        )
    st.markdown(f'<div class="kpi-row">{"".join(cells)}</div>', unsafe_allow_html=True)


def light_table(df_table, numeric_cols=()):
    head = "".join(f'<th class="{"n" if c in numeric_cols else ""}">{esc(c)}</th>' for c in df_table.columns)
    rows = []
    for _, r in df_table.iterrows():
        tds = "".join(f'<td class="{"n" if c in numeric_cols else ""}">{esc(r[c])}</td>' for c in df_table.columns)
        rows.append(f"<tr>{tds}</tr>")
    st.markdown(
        f'<div class="lt-wrap"><table class="lt"><thead><tr>{head}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>',
        unsafe_allow_html=True,
    )


def render_chart(fig, empty_msg=None):
    if fig is not None:
        st.plotly_chart(fig, width="stretch", theme=None, config={"displayModeBar": False})
    elif empty_msg:
        st.info(empty_msg)


def _layout(fig, title, height=380, **extra):
    layout = {**PLOTLY_LAYOUT, **extra}
    fig.update_layout(**layout, height=height,
                      title=dict(text=title, font=dict(size=14, color=COR["texto"]), x=0.02))
    return fig


# ============================================================
# CARGA DE DADOS
# ============================================================
def _gh_headers(token, raw=False):
    return {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.raw" if raw else "application/vnd.github.v3+json",
    }


@st.cache_data(ttl=120, show_spinner=False)
def fetch_from_github():
    """Retorna (bytes_do_excel, data_da_última_atualização) ou (None, None)."""
    token = get_secret("GITHUB_TOKEN")
    if not token:
        return None, None
    base = f"https://api.github.com/repos/{GITHUB_REPO}"
    r = requests.get(f"{base}/contents/{GITHUB_DATA_PATH}", headers=_gh_headers(token, raw=True),
                     params={"ref": GITHUB_BRANCH}, timeout=30)
    if r.status_code != 200:
        return None, None
    updated = None
    try:
        c = requests.get(f"{base}/commits", headers=_gh_headers(token),
                         params={"path": GITHUB_DATA_PATH, "sha": GITHUB_BRANCH, "per_page": 1}, timeout=15)
        if c.ok and c.json():
            updated = pd.Timestamp(c.json()[0]["commit"]["committer"]["date"]).tz_convert("America/Sao_Paulo")
    except Exception:
        pass
    return r.content, updated


@st.cache_data(ttl=300, show_spinner=False)
def fetch_from_sharepoint():
    if not get_secret("TENANT_ID"):
        return None, None
    tok = requests.post(
        f"https://login.microsoftonline.com/{st.secrets['TENANT_ID']}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": st.secrets["CLIENT_ID"],
              "client_secret": st.secrets["CLIENT_SECRET"], "scope": "https://graph.microsoft.com/.default"},
        timeout=30,
    )
    tok.raise_for_status()
    headers = {"Authorization": f"Bearer {tok.json()['access_token']}"}
    site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{st.secrets['SHAREPOINT_DOMAIN']}:/sites/{st.secrets['SHAREPOINT_SITE_PATH']}",
        headers=headers, timeout=30)
    site.raise_for_status()
    f = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site.json()['id']}/drive/root:/{st.secrets['SHAREPOINT_FILE_PATH']}:/content",
        headers=headers, timeout=60)
    f.raise_for_status()
    return f.content, None


def fetch_from_local():
    if LOCAL_DATA_PATH.exists():
        mtime = pd.Timestamp(datetime.fromtimestamp(LOCAL_DATA_PATH.stat().st_mtime))
        return LOCAL_DATA_PATH.read_bytes(), mtime
    return None, None


def save_data_to_github(file_bytes, filename):
    token = get_secret("GITHUB_TOKEN")
    if not token:
        return False, "Token do GitHub não configurado nos Secrets."
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_DATA_PATH}"
    r = requests.get(url, headers=_gh_headers(token), params={"ref": GITHUB_BRANCH}, timeout=30)
    sha = r.json().get("sha") if r.status_code == 200 else None
    payload = {
        "message": f"Atualização de dados: {filename} ({datetime.now().strftime('%d/%m/%Y %H:%M')})",
        "content": base64.b64encode(file_bytes).decode("utf-8"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    r = requests.put(url, headers=_gh_headers(token), json=payload, timeout=60)
    if r.status_code in (200, 201):
        fetch_from_github.clear()
        return True, "Dados atualizados com sucesso!"
    return False, f"Erro ao salvar: {r.status_code} - {r.json().get('message', '')}"


def find_header_row(buf, sheet_name, markers, max_scan=15):
    buf.seek(0)
    raw = pd.read_excel(buf, sheet_name=sheet_name, header=None, nrows=max_scan)
    for i in range(len(raw)):
        vals = [str(v).replace("\n", " ").strip().upper() for v in raw.iloc[i] if pd.notna(v)]
        if any(m.upper() in v for m in markers for v in vals):
            return i
    return 0


def read_controle(buf):
    buf.seek(0)
    df = pd.read_excel(buf, sheet_name="Controle", header=0)
    if sum(str(c).startswith("Unnamed") for c in df.columns) > len(df.columns) * 0.5:
        buf.seek(0)
        raw = pd.read_excel(buf, sheet_name="Controle", header=None)
        for i in range(min(5, len(raw))):
            vals = [str(v).replace("\n", " ").strip() for v in raw.iloc[i] if pd.notna(v)]
            if any(k in v for v in vals for k in ("Código", "Bobina", "NECESSIDADE", "Tipo")):
                buf.seek(0)
                df = pd.read_excel(buf, sheet_name="Controle", header=i)
                break
    return df


def read_propostas(buf):
    """Lê a aba 'A.Propostas'. Retorna (df, dept_map) — dept_map vem da linha acima do cabeçalho."""
    try:
        buf.seek(0)
        sheet = next((s for s in pd.ExcelFile(buf).sheet_names
                      if s.strip().upper().replace(" ", "") in ("A.PROPOSTAS", "APROPOSTAS")), None)
        if sheet is None:
            return None, {}
        hdr = find_header_row(buf, sheet, ["CÓDIGO DELGA", "CODIGO DELGA"])
        buf.seek(0)
        df = pd.read_excel(buf, sheet_name=sheet, header=hdr, keep_default_na=False, na_values=[""])
        dept_map = {}
        if hdr > 0:
            buf.seek(0)
            raw = pd.read_excel(buf, sheet_name=sheet, header=None)
            dept_row, head_row = raw.iloc[hdr - 1], raw.iloc[hdr]
            for i, col in enumerate(head_row):
                dept = dept_row.iloc[i] if i < len(dept_row) else None
                if pd.notna(col) and pd.notna(dept) and str(dept).strip():
                    dept_map[str(col).replace("\n", " ").strip()] = str(dept).strip()
        return df, dept_map
    except Exception:
        return None, {}


@st.cache_data(show_spinner=False)
def parse_workbook(content: bytes):
    buf = io.BytesIO(content)
    controle = read_controle(buf)
    buf.seek(0)
    formulas = pd.read_excel(buf, sheet_name="Formulas")
    propostas, dept_map = read_propostas(buf)
    return controle, formulas, propostas, dept_map


def load_data():
    for source, nome in ((fetch_from_github, "GitHub"), (fetch_from_sharepoint, "SharePoint"),
                         (fetch_from_local, "arquivo local")):
        try:
            content, updated = source()
        except Exception:
            continue
        if content:
            return parse_workbook(content), updated, nome
    return None, None, None


# ============================================================
# PROCESSAMENTO — aba Controle / Formulas
# ============================================================
def process_data(df_raw):
    df = df_raw.copy()
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    col_codigo = next((c for c in df.columns if "Código" in c and "Bobina" in c), None)
    if col_codigo:
        df = df[df[col_codigo].notna() & (df[col_codigo] != "")]

    def first(pred):
        return next((c for c in df.columns if pred(c) and "MÉDIA" not in c.upper()), None)

    col_names = {
        "jan": first(lambda c: "Janeiro" in c),
        "fev": first(lambda c: "Fevereiro" in c),
        "mar": first(lambda c: "Março" in c or "Marco" in c),
        "abr": first(lambda c: "Abril" in c),
        "mai": first(lambda c: "Maio" in c),
        "media": next((c for c in df.columns if "MÉDIA" in c.upper() and "FEV" in c.upper()), None),
    }
    col_names = {k: v for k, v in col_names.items() if v}
    for col in col_names.values():
        df[col] = df[col].apply(parse_numero_brasileiro)
    return df, col_names


def _to_float(v, default=0.0):
    try:
        return float(v) if pd.notna(v) else default
    except (TypeError, ValueError):
        return None


def parse_formulas(df_f):
    unidades, usinas = [], []
    for i in range(min(10, len(df_f))):
        row = df_f.iloc[i]
        nome = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if nome.lower() == "total":
            break
        if not nome or nome.lower() in ("nan", "usinas"):
            continue
        bob = _to_float(row.iloc[1])
        if bob is None:
            continue
        pct = _to_float(row.iloc[4]) or 0
        unidades.append({
            "unidade": nome, "bobinas": int(bob),
            "peso_total": _to_float(row.iloc[2]) or 0, "peso_analisado": _to_float(row.iloc[3]) or 0,
            "pct": pct * 100 if pct <= 1 else pct,
        })

    start = next((i + 1 for i in range(len(df_f))
                  if pd.notna(df_f.iloc[i, 0]) and str(df_f.iloc[i, 0]).strip().lower() == "usinas"), None)
    if start:
        for i in range(start, len(df_f)):
            row = df_f.iloc[i]
            nome = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if nome.lower() == "total":
                break
            if not nome or nome.lower() == "nan" or nome == "0":
                continue
            bob = _to_float(row.iloc[1])
            if bob is None:
                continue
            rep = _to_float(row.iloc[3]) or 0
            usinas.append({"usina": nome, "bobinas": int(bob), "peso": _to_float(row.iloc[2]) or 0,
                           "pct_representacao": rep * 100 if rep <= 1 else rep})
    return pd.DataFrame(unidades), pd.DataFrame(usinas)


# ============================================================
# PROCESSAMENTO — aba A.Propostas
# ============================================================
def classify_stage_value(v):
    """done (data passada) | planned (data futura) | na | pending (texto) | empty"""
    if pd.isna(v) or (isinstance(v, str) and not v.strip()):
        return "empty", ""
    if isinstance(v, (pd.Timestamp, datetime)):
        d = pd.Timestamp(v)
        return ("done" if d.date() <= datetime.now().date() else "planned"), d.strftime("%d/%m/%y")
    txt = str(v).strip()
    up = txt.upper()
    if up in ("N/A", "NA", "NÃO SE APLICA", "NAO SE APLICA", "-"):
        return "na", "N/A"
    if "PENDENTE" in up:
        return "pending", "Pendente"
    return "pending", txt[:40]


def classify_formalizacao(valor, data_envio_usina):
    """Se já foi enviado à usina (ou envio N/A), a formalização é considerada concluída."""
    status, texto = classify_stage_value(valor)
    if status == "done":
        return status, texto
    if classify_stage_value(data_envio_usina)[0] in DONE:
        return "done", texto
    return status, texto


@st.cache_data(show_spinner=False)
def process_propostas(df_raw, dept_map):
    if df_raw is None:
        return None
    df = df_raw.copy()
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    col_codigo = find_col(df, exact="CÓDIGO DELGA")
    if col_codigo is None:
        return None
    df = df[df[col_codigo].notna() & (df[col_codigo].astype(str).str.strip() != "")].copy()
    if df.empty:
        return None

    # Remove Código Delga duplicado, mantendo a linha mais avançada (SIM > N/A > NÃO)
    col_passado = find_col(df, contains="PASSADO PARA USINA")
    if col_passado:
        rank = {"SIM": 2, "N/A": 1}
        df["_rank"] = df[col_passado].apply(lambda x: rank.get(str(x).strip().upper(), 0))
        df = df.sort_values("_rank").drop_duplicates(subset=[col_codigo], keep="last").drop(columns=["_rank"])
    else:
        df = df.drop_duplicates(subset=[col_codigo])
    df = df.rename(columns={col_codigo: "CÓDIGO DELGA"})

    col_desc = find_col(df, contains="DESCRIÇÃO")
    col_planta = find_col(df, contains="PLANTA DELGA")
    col_fonte = find_col(df, exact="FONTE")
    col_reducao = find_col(df, contains="REDUÇÃO POTENCIAL")
    col_consumo = find_col(df, contains="MÉDIA CONSUMO")
    col_envio = find_col(df, exact="DATA ENVIO P/ USINA")
    col_viab = find_col(df, contains="VIABILIDADE")
    col_area = find_col(df, contains="PENDENTE")
    cols_projeto = [c for c in (find_col(df, contains="QUAL PROJETO"), find_col(df, contains="NOME DO PROJETO")) if c]

    def projeto(row):
        for c in cols_projeto:
            if pd.notna(row.get(c)):
                v = str(row[c]).split("\n")[0].strip()  # 2ª linha costuma ser anotação
                if v and v.lower() not in ("nan", "none"):
                    return v
        return ""

    df["_DESCRICAO"] = df[col_desc].astype(str).str.strip() if col_desc else ""
    df["_PLANTA"] = df[col_planta].apply(normalizar_texto) if col_planta else "DESCONHECIDA"
    df["_FONTE"] = df[col_fonte].apply(normalizar_texto) if col_fonte else ""
    df["_PASSADO"] = df[col_passado].apply(lambda x: normalizar_texto(x, "")) if col_passado else ""
    df["_REDUCAO"] = df[col_reducao].apply(parse_numero_brasileiro) if col_reducao else 0.0
    df["_CONSUMO"] = df[col_consumo].apply(parse_numero_brasileiro) if col_consumo else 0.0
    df["_VIABILIDADE"] = df[col_viab].apply(lambda x: normalizar_texto(x, "")) if col_viab else ""
    df["_AREA_PENDENTE"] = df[col_area].apply(lambda x: normalizar_texto(x, "")) if col_area else ""
    df["_PROJETO"] = df.apply(projeto, axis=1)

    stages_def = [s for s in STAGE_DEFS if s[0] in df.columns]
    n_pct = sum(1 for s in stages_def if s[2]) or 1

    out = {k: [] for k in ("_STAGES", "_PCT", "_BADGE", "_BADGE_COR", "_ETAPA", "_INVIAVEL")}
    for _, row in df.iterrows():
        stages, completed = [], 0
        for col, label, counts in stages_def:
            if col == "FORMALIZADO COM COMPRAS":
                status, texto = classify_formalizacao(row.get(col), row.get(col_envio) if col_envio else None)
            else:
                status, texto = classify_stage_value(row.get(col))
            if counts and status in DONE:
                completed += 1
            stages.append({"label": label, "status": status, "texto": texto, "counts": counts,
                           "dept": dept_map.get(col, "")})

        viab, passado = row["_VIABILIDADE"], row["_PASSADO"]
        inviavel = "INVIÁV" in viab or "INVIAV" in viab
        pct = round(completed / n_pct * 100)
        pendentes = [s["label"] for s in stages if s["counts"] and s["status"] not in DONE]
        etapa = pendentes[0] if pendentes else "Concluído"

        if inviavel:
            pct, badge, cor, etapa = 0, "Inviável", COR["vermelho"], "Inviável"
        elif "NÃO" in passado or passado == "NAO":
            pct, badge, cor, etapa = 0, "Não enviado à usina", COR["coral"], "Não enviado à usina"
        elif "KAIZEN" in passado:
            badge, cor = "Kaizen Plano de Corte", COR["roxo"]
        elif not pendentes:
            pct, badge, cor = 100, "Concluído", COR["verde"]
        elif completed == 0:
            badge, cor = "Aguardando início", COR["texto2"]
        else:
            badge, cor = "Em andamento", COR["azul"]

        for k, v in zip(out, (stages, pct, badge, cor, etapa, inviavel)):
            out[k].append(v)

    for k, v in out.items():
        df[k] = v
    df["_ENVIADA"] = df["_PASSADO"].astype(str).str.contains("SIM", na=False)
    return df


# ============================================================
# LINHA DO TEMPO (HTML)
# ============================================================
def fluxo_svg(active: int) -> str:
    """Fallback desenhado em SVG caso as imagens de assets/ não existam."""
    quadrants = [
        ("1. ÁREA TÉCNICA", "#1D4ED8", "#DBEAFE", ["Análise Técnica", "Elaboração da Proposta", "Encaminhamento p/ Compras"]),
        ("2. ÁREA COMERCIAL", "#B91C1C", "#FEE2E2", ["Consulta à Usina", "Negociação Comercial", "Acordo Delga x Usina"]),
        ("3. LOGÍSTICA / PCP", "#6D28D9", "#EDE9FE", ["Confirmação PCP", "Planejamento Entrega", "Planejamento Produção"]),
        ("4. MANUFATURA", "#15803D", "#DCFCE7", ["Liberação p/ Fabricação", "Produção", "FIM"]),
    ]
    W, QW, QH, X0, Y0, GAP = 680, 158, 170, 12, 40, 8
    p = [f'<svg viewBox="0 0 {W} 230" xmlns="http://www.w3.org/2000/svg" style="display:block;width:660px">',
         f'<rect width="{W}" height="230" rx="10" fill="#F8FAFC"/>',
         f'<text x="{W // 2}" y="24" text-anchor="middle" font-family="Arial" font-weight="800" font-size="13" fill="#1F2937">FLUXO PROCESSO BSW</text>']
    for i, (nome, hc, bc, items) in enumerate(quadrants):
        x, act = X0 + i * (QW + GAP), (i + 1 == active)
        op = "1" if act else "0.4"
        p.append(f'<g opacity="{op}"><rect x="{x}" y="{Y0}" width="{QW}" height="{QH}" rx="8" fill="{bc}" '
                 f'stroke="{"#16A34A" if act else hc}" stroke-width="{2.5 if act else 1}"/>'
                 f'<rect x="{x}" y="{Y0}" width="{QW}" height="28" rx="8" fill="{hc}"/>'
                 f'<text x="{x + QW // 2}" y="{Y0 + 18}" text-anchor="middle" font-family="Arial" font-weight="700" font-size="10" fill="#fff">{nome}</text>')
        for j, it in enumerate(items):
            y = Y0 + 42 + j * 40
            p.append(f'<rect x="{x + 10}" y="{y}" width="{QW - 20}" height="26" rx="5" fill="#fff" stroke="{hc}" stroke-opacity=".4"/>'
                     f'<text x="{x + QW // 2}" y="{y + 17}" text-anchor="middle" font-family="Arial" font-size="9" fill="#1F2937">{it}</text>')
        p.append("</g>")
        if act:
            p.append(f'<rect x="{x + 20}" y="{Y0 + QH + 6}" width="{QW - 40}" height="18" rx="9" fill="#16A34A"/>'
                     f'<text x="{x + QW // 2}" y="{Y0 + QH + 19}" text-anchor="middle" font-family="Arial" font-weight="700" font-size="9" fill="#fff">ETAPA ATUAL</text>')
    p.append("</svg>")
    return "".join(p)


@st.cache_data(show_spinner=False)
def fluxo_tooltip_html():
    """Divs ocultas com a imagem de cada quadrante do fluxo (lidas pelo JS do tooltip)."""
    parts = []
    for i in range(1, 5):
        img = ASSETS_DIR / f"fluxo_{i}.jpg"
        if img.exists():
            b64 = base64.b64encode(img.read_bytes()).decode()
            inner = f'<img src="data:image/jpeg;base64,{b64}" style="width:760px;max-width:88vw;border-radius:8px;display:block">'
        else:
            inner = fluxo_svg(i)
        parts.append(f'<div id="bsw-fluxo-{i}" style="display:none">{inner}</div>')
    return "".join(parts)


def flow_stage(pct):
    return 0 if pct <= 0 else 1 if pct <= 25 else 2 if pct <= 50 else 3 if pct <= 75 else 4


def _pill_class(dept):
    d = dept.upper()
    if not d:
        return "none"
    for key, cls in (("ENGENH", "eng"), ("LOG", "log"), ("SOLIC", "sol"), ("TODOS", "all")):
        if key in d:
            return cls
    return ""


def _track_html(stages, header=False, inviavel=False, fstage=0):
    n = len(stages)
    filled = 0
    for s in stages:
        if s["status"] not in DONE:
            break
        filled += 1
    last_done = max((i for i, s in enumerate(stages) if s["status"] in DONE), default=-1)
    cells = []
    for i, s in enumerate(stages):
        if header:
            lbl_cls = "tl-lbl" + ("" if s["counts"] else " info")
            cells.append(
                f'<div class="tl-cell"><span class="tl-pill {_pill_class(s["dept"])}">{esc(s["dept"]) or "&nbsp;"}</span>'
                f'<div class="{lbl_cls}">{esc(s["label"])}</div></div>'
            )
            continue
        bad = " bad" if inviavel else ""
        lines = ""
        if i > 0:
            lines += f'<span class="ln l{" on" if i < filled and not inviavel else ""}{bad}"></span>'
        if i < n - 1:
            lines += f'<span class="ln r{" on" if i + 1 < filled and not inviavel else ""}{bad}"></span>'
        if inviavel:
            dot_cls, attr = "tl-dot bad", ""
        else:
            dot_cls = f'tl-dot {s["status"]}' + ("" if s["counts"] else " info")
            attr = ""
            if i == last_done and fstage > 0:
                dot_cls += " tip"
                attr = f' data-fstage="{fstage}"'
        val = esc(s["texto"] or "—")
        cells.append(f'<div class="tl-cell">{lines}<div class="{dot_cls}"{attr}></div>'
                     f'<div class="tl-val" title="{val}">{val}</div></div>')
    return "".join(cells)


def _row_html(r):
    inv = bool(r["_INVIAVEL"])
    meta = " · ".join(x for x in (r["_PROJETO"], r["_FONTE"]) if x) or "Projeto não informado"
    cor = r["_BADGE_COR"]
    pct_html = ('<div class="tl-pct inv">inviável</div>' if inv
                else f'<div class="tl-pct">{r["_PCT"]}%</div>')
    return (
        f'<div class="tl-row{" inv" if inv else ""}">'
        f'<div class="tl-info"><div class="tl-top"><span class="tl-code">{esc(r["CÓDIGO DELGA"])}</span>'
        f'<span class="tl-badge" style="background:{cor}1A;color:{cor};border:1px solid {cor}40">{esc(r["_BADGE"])}</span></div>'
        f'<div class="tl-desc" title="{esc(r["_DESCRICAO"])}">{esc(r["_DESCRICAO"])}</div>'
        f'<div class="tl-meta">{esc(meta)}</div></div>'
        f'<div class="tl-track">{_track_html(r["_STAGES"], inviavel=inv, fstage=0 if inv else flow_stage(r["_PCT"]))}</div>'
        f'{pct_html}</div>'
    )


def timeline_block_html(planta, df_g, aberto):
    ativos = df_g[~df_g["_INVIAVEL"]]
    media = ativos["_PCT"].mean() if len(ativos) else 0
    concl = int((df_g["_PCT"] == 100).sum())
    header = (f'<div class="tl-row tl-head"><div class="tl-info"><span class="tl-hl">Proposta</span></div>'
              f'<div class="tl-track">{_track_html(df_g.iloc[0]["_STAGES"], header=True)}</div>'
              f'<div class="tl-pct"><span class="tl-hl">%</span></div></div>')
    rows = "".join(_row_html(r) for _, r in df_g.iterrows())
    return (
        f'<details class="blk"{" open" if aberto else ""}><summary>'
        f'<span class="blk-dot" style="background:{get_unidade_color(planta)}"></span>'
        f'<span class="blk-name">{esc(planta)}</span><span class="blk-sub">{len(df_g)} propostas</span>'
        f'<span class="blk-stats"><b>{media:.0f}%</b> progresso médio &nbsp;·&nbsp; <b>{concl}</b> concluídas</span>'
        f'</summary><div class="blk-body"><div class="blk-inner">{header}{rows}</div></div></details>'
    )


TOOLTIP_JS = """
<script>
(function(){
  var par = window.parent.document, tries = 0;
  function bind(){
    tries++;
    var dots = par.querySelectorAll(".tl-dot.tip");
    if(!dots.length){ if(tries < 15) setTimeout(bind, 350); return; }
    var tip = par.getElementById("bsw-flow-tip");
    if(!tip){
      tip = par.createElement("div"); tip.id = "bsw-flow-tip";
      tip.style.cssText = "display:none;position:fixed;z-index:99999;pointer-events:none;background:#fff;"
        + "border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,.22);padding:6px;border:1px solid #E2E6F0;";
      par.body.appendChild(tip);
    }
    function pos(e){
      var w = tip.offsetWidth || 780, h = tip.offsetHeight || 520, M = 14, W = window.parent.innerWidth, H = window.parent.innerHeight;
      var x = e.clientX + M, y = e.clientY - h / 2;
      if(x + w > W - M) x = e.clientX - w - M;
      y = Math.max(M, Math.min(y, H - h - M));
      tip.style.left = x + "px"; tip.style.top = y + "px";
    }
    dots.forEach(function(d){
      if(d.dataset.bound) return; d.dataset.bound = "1";
      var src = par.getElementById("bsw-fluxo-" + d.dataset.fstage);
      if(!src) return;
      d.addEventListener("mouseenter", function(e){ tip.innerHTML = src.innerHTML; tip.style.display = "block"; pos(e); });
      d.addEventListener("mousemove", pos);
      d.addEventListener("mouseleave", function(){ tip.style.display = "none"; });
    });
  }
  setTimeout(bind, 400);
  new MutationObserver(function(){ tries = 0; setTimeout(bind, 400); })
    .observe(par.body, {childList: true, subtree: true});
})();
</script>
"""


# ============================================================
# GRÁFICOS
# ============================================================
def create_area_chart(df, col_names):
    meses, keys = ["Jan", "Fev", "Mar", "Abr", "Mai"], ["jan", "fev", "mar", "abr", "mai"]
    valores = [round(df[col_names[k]].sum(), 1) if k in col_names else 0 for k in keys]
    fig = go.Figure(go.Scatter(
        x=meses, y=valores, fill="tozeroy", fillcolor="rgba(77,107,255,0.12)",
        line=dict(color=COR["azul"], width=3), mode="lines+markers", marker=dict(size=8, color=COR["azul"]),
        hovertemplate="%{x}/2026 <b>%{y:,.0f} ton</b><extra></extra>",
    ))
    return _layout(fig, "Necessidade mensal (ton)", yaxis=dict(**GRID), xaxis=dict(**GRID))


def _pie(labels, values, colors, title):
    fig = go.Figure(go.Pie(
        labels=[str(x) for x in labels], values=list(values), hole=0.55, marker=dict(colors=colors),
        textinfo="percent", textfont=dict(size=12, color="#fff"), sort=False,
        hovertemplate="%{label} <b>%{value:,.1f} ton</b> (%{percent})<extra></extra>",
    ))
    return _layout(fig, title, legend=dict(orientation="h", y=-0.05, font=dict(color="#64748B")))


def create_tipo_pie_chart(df, col_media):
    col = find_col(df, exact="TIPO")
    if not col:
        return None
    d = df[df[col].notna() & (df[col].astype(str).str.strip() != "")].copy()
    grupos = {"Z": "BZ", "Q": "BQ", "F": "BF"}
    d["g"] = d[col].astype(str).str.strip().str.upper().str[-1].map(grupos)
    dist = d.dropna(subset=["g"]).groupby("g")[col_media].sum().sort_values(ascending=False)
    return _pie(dist.index, dist.values, CHART_COLORS, "Por tipo de bobina") if len(dist) else None


def create_unidade_pie_chart(df, col_media):
    col = next((c for c in df.columns if "Unidade" in c and "Delga" in c), None)
    if not col:
        return None
    d = df[df[col].notna() & (df[col].astype(str).str.strip() != "")]
    dist = d.groupby(col)[col_media].sum().sort_values(ascending=False)
    return _pie(dist.index, dist.values, [get_unidade_color(n) for n in dist.index], "Por unidade Delga") if len(dist) else None


def create_thickness_chart(df, col_media):
    col = next((c for c in df.columns if "Esp" in c and "mm" in c), None)
    if not col:
        return None
    d = df.copy()
    d["esp"] = pd.to_numeric(d[col], errors="coerce")
    d = d[d["esp"].notna()]
    labels = ["0-1", "1-2", "2-4", "4-6", "6-8", "8-10", "10-15", "15-20", "20+"]
    d["faixa"] = pd.cut(d["esp"], bins=[0, 1, 2, 4, 6, 8, 10, 15, 20, 50], labels=labels)
    dist = d.groupby("faixa", observed=True)[col_media].sum()
    dist = dist[dist > 0]
    if not len(dist):
        return None
    fig = go.Figure(go.Bar(x=[str(x) for x in dist.index], y=dist.values, marker_color=COR["azul"],
                           hovertemplate="%{x} mm <b>%{y:,.1f} ton</b><extra></extra>"))
    return _layout(fig, "Por faixa de espessura (mm)", yaxis=dict(title="ton", **GRID), xaxis=dict(**GRID))


def _hbar(labels, values, colors, title, fmt="%{x:,.1f} ton", height=None, text=None):
    fig = go.Figure(go.Bar(
        x=list(values), y=[str(x) for x in labels], orientation="h", marker_color=colors,
        text=text, textposition="outside", cliponaxis=False, textfont=dict(color=COR["texto"], size=12),
        hovertemplate=f"%{{y}} <b>{fmt}</b><extra></extra>",
    ))
    return _layout(fig, title, height=height or max(320, len(labels) * 34 + 90),
                   yaxis=dict(automargin=True, **GRID), xaxis=dict(**GRID))


def create_usinas_chart(df_usinas, top_n=15):
    if df_usinas.empty:
        return None
    d = df_usinas.nlargest(top_n, "peso").sort_values("peso")
    return _hbar(d["usina"], d["peso"], COR["azul"], f"Top {top_n} usinas por peso (ton)")


def create_progress_chart(df_unidades):
    if df_unidades.empty:
        return None
    cores = [get_unidade_color(u) for u in df_unidades["unidade"]]
    fig = go.Figure([
        go.Bar(name="Peso total", x=df_unidades["unidade"], y=df_unidades["peso_total"],
               marker=dict(color=cores, opacity=0.35), hovertemplate="%{x} total: <b>%{y:,.1f} ton</b><extra></extra>"),
        go.Bar(name="Peso analisado", x=df_unidades["unidade"], y=df_unidades["peso_analisado"],
               marker=dict(color=cores), hovertemplate="%{x} analisado: <b>%{y:,.1f} ton</b><extra></extra>"),
    ])
    return _layout(fig, "Peso total vs analisado por unidade (ton)", barmode="group",
                   yaxis=dict(**GRID), xaxis=dict(**GRID), legend=dict(orientation="h", y=1.1, x=1, xanchor="right"))


def create_group_bar(df, col_media, group_col, title, top_n=10, by_unit_color=False):
    d = df[df[group_col].notna() & (df[group_col].astype(str).str.strip() != "")]
    dist = d.groupby(group_col)[col_media].sum().sort_values().tail(top_n)
    if not len(dist):
        return None
    cores = [get_unidade_color(x) for x in dist.index] if by_unit_color else COR["azul"]
    return _hbar(dist.index, dist.values, cores, title)


def create_etapa_chart(df_p):
    """Quantas propostas estão paradas em cada etapa do processo."""
    ordem = ["Não enviado à usina"] + [s[1] for s in STAGE_DEFS if s[2]] + ["Concluído", "Inviável"]
    cont = df_p["_ETAPA"].value_counts()
    labels = [e for e in ordem if cont.get(e, 0) > 0]
    if not labels:
        return None
    cor_map = {"Concluído": COR["verde"], "Inviável": COR["vermelho"], "Não enviado à usina": COR["coral"]}
    labels = labels[::-1]  # primeira etapa no topo
    vals = [int(cont[e]) for e in labels]
    return _hbar(labels, vals, [cor_map.get(e, COR["azul"]) for e in labels],
                 "Onde estão as propostas (etapa atual)", fmt="%{x} propostas", text=vals)


# ============================================================
# ABAS
# ============================================================
def tab_visao_geral(df, col_names, col_media, df_usinas):
    c1, c2 = st.columns([2, 1])
    with c1:
        render_chart(create_area_chart(df, col_names))
    with c2:
        render_chart(create_tipo_pie_chart(df, col_media), "Coluna 'Tipo' não encontrada.")
    c3, c4 = st.columns(2)
    with c3:
        render_chart(create_thickness_chart(df, col_media), "Coluna de espessura não encontrada.")
    with c4:
        render_chart(create_unidade_pie_chart(df, col_media), "Coluna 'Unidade Delga' não encontrada.")
    render_chart(create_usinas_chart(df_usinas), "Dados de usinas não encontrados na aba Formulas.")


def tab_analises(df, col_media, df_unidades):
    if not df_unidades.empty:
        section("Progresso da análise por unidade")
        render_chart(create_progress_chart(df_unidades))
        t = pd.DataFrame({
            "Unidade": df_unidades["unidade"],
            "Bobinas": df_unidades["bobinas"].map(fmt_num),
            "Peso total (ton)": df_unidades["peso_total"].map(lambda v: fmt_num(v, 1)),
            "Peso analisado (ton)": df_unidades["peso_analisado"].map(lambda v: fmt_num(v, 1)),
            "% concluído": df_unidades["pct"].map(lambda v: f"{fmt_num(v, 1)}%"),
        })
        light_table(t, numeric_cols=t.columns[1:])

    section("Necessidade por unidade e beneficiador")
    c1, c2 = st.columns(2)
    with c1:
        col_u = next((c for c in df.columns if "Unidade" in c and "Delga" in c), None)
        render_chart(create_group_bar(df, col_media, col_u, "Por unidade Delga (ton)", by_unit_color=True) if col_u else None,
                     "Coluna 'Unidade Delga' não encontrada.")
    with c2:
        col_b = next((c for c in df.columns if "Beneficiador" in c), None)
        render_chart(create_group_bar(df, col_media, col_b, "Top 10 beneficiadores (ton)") if col_b else None,
                     "Coluna 'Beneficiador' não encontrada.")

    col_abc = find_col(df, exact="ABC")
    if col_abc:
        d = df[df[col_abc].notna() & (df[col_abc].astype(str).str.strip() != "")]
        if len(d):
            section("Classificação ABC")
            agg = d.groupby(col_abc)[col_media].agg(["sum", "count"]).sort_values("sum", ascending=False).reset_index()
            t = pd.DataFrame({"Classe": agg[col_abc], "Necessidade (ton)": agg["sum"].map(lambda v: fmt_num(v, 1)),
                              "Qtd bobinas": agg["count"].map(fmt_num)})
            light_table(t, numeric_cols=("Necessidade (ton)", "Qtd bobinas"))


def tab_acompanhamento(df_p):
    if df_p is None or df_p.empty:
        st.info('Aba "A.Propostas" não encontrada no Excel. Envie um arquivo que contenha essa aba para ver o acompanhamento.')
        return

    # ---------- 1. Big numbers gerais ----------
    ativos = df_p[~df_p["_INVIAVEL"]]
    total = len(df_p)
    enviadas = int(df_p["_ENVIADA"].sum())
    concl = int((df_p["_PCT"] == 100).sum())
    inviaveis = int(df_p["_INVIAVEL"].sum())
    andamento = total - concl - inviaveis
    section("Visão geral das propostas", "Todas as unidades")
    kpi_row([
        ("Propostas", fmt_num(total), None, COR["azul"]),
        ("Enviadas à usina", fmt_num(enviadas), f"{enviadas / total * 100:.0f}% do total", COR["azul_claro"]),
        ("Em aberto", fmt_num(andamento), "não concluídas nem inviáveis", COR["laranja"]),
        ("Concluídas", fmt_num(concl), f"{concl / total * 100:.0f}% do total", COR["verde"]),
        ("Inviáveis", fmt_num(inviaveis), None, COR["vermelho"]),
        ("Progresso médio", f"{ativos['_PCT'].mean():.0f}%" if len(ativos) else "—", "sem inviáveis", COR["roxo"]),
    ])

    # ---------- 2. Filtro único: unidade ----------
    plantas = ordenar_plantas(df_p["_PLANTA"].unique().tolist())
    section("Unidade")
    opcoes = ["Todas"] + plantas
    if hasattr(st, "segmented_control"):
        sel = st.segmented_control("Unidade", opcoes, default="Todas", label_visibility="collapsed",
                                   key="acomp_unidade") or "Todas"
    else:
        sel = st.radio("Unidade", opcoes, horizontal=True, label_visibility="collapsed", key="acomp_unidade")
    df_u = df_p if sel == "Todas" else df_p[df_p["_PLANTA"] == sel]
    ativos_u = df_u[~df_u["_INVIAVEL"]]

    # ---------- 3. Volume e situação da unidade ----------
    section(f"Volume e situação — {sel.title() if sel != 'Todas' else 'todas as unidades'}")
    kpi_row([
        ("Propostas", fmt_num(len(df_u)), f"{int(df_u['_ENVIADA'].sum())} enviadas à usina", get_unidade_color(sel)),
        ("Volume (média consumo)", fmt_num(df_u["_CONSUMO"].sum()), None, COR["azul_claro"]),
        ("Redução potencial", fmt_num(df_u["_REDUCAO"].sum()), None, COR["verde"]),
        ("Concluídas", fmt_num(int((df_u["_PCT"] == 100).sum())), None, COR["verde"]),
        ("Progresso médio", f"{ativos_u['_PCT'].mean():.0f}%" if len(ativos_u) else "—", "sem inviáveis", COR["roxo"]),
    ])

    st.write("")
    c1, c2 = st.columns([3, 2])
    with c1:
        render_chart(create_etapa_chart(df_u))
    with c2:
        if sel == "Todas":
            st.markdown('<div class="sec-t" style="margin:4px 0 8px">Resumo por unidade</div>', unsafe_allow_html=True)
            g = df_p.groupby("_PLANTA")
            resumo = pd.DataFrame({
                "Unidade": g.size().index,
                "Propostas": g.size().values,
                "Volume": g["_CONSUMO"].sum().values,
                "Redução": g["_REDUCAO"].sum().values,
                "Progresso": [df_p[(df_p["_PLANTA"] == p) & ~df_p["_INVIAVEL"]]["_PCT"].mean() for p in g.size().index],
            })
            resumo = resumo.set_index("Unidade").loc[ordenar_plantas(resumo["Unidade"].tolist())].reset_index()
            resumo["Volume"] = resumo["Volume"].map(fmt_num)
            resumo["Redução"] = resumo["Redução"].map(fmt_num)
            resumo["Progresso"] = resumo["Progresso"].map(lambda v: f"{v:.0f}%" if pd.notna(v) else "—")
            light_table(resumo, numeric_cols=("Propostas", "Volume", "Redução", "Progresso"))
        area = df_u.loc[df_u["_AREA_PENDENTE"] != "", "_AREA_PENDENTE"].value_counts()
        if len(area):
            st.markdown('<div class="sec-t" style="margin:16px 0 8px">Pendências por área</div>', unsafe_allow_html=True)
            light_table(pd.DataFrame({"Área": area.index, "Propostas": area.values}), numeric_cols=("Propostas",))

    # ---------- 4. Lista de propostas ----------
    section("Propostas", "Ordenadas da menos para a mais avançada. Clique na unidade para expandir. "
                         "Passe o mouse na bolinha destacada para ver a etapa no fluxo BSW.")
    st.markdown(
        '<div class="legend">'
        '<span><i style="background:#1400FF;border-color:#1400FF"></i>concluída / N/A</span>'
        '<span><i style="border-color:#F59E0B"></i>pendente</span>'
        '<span><i style="border-color:#4DA3FF"></i>prevista (data futura)</span>'
        '<span><i style="border-color:#D1D7E3"></i>não iniciada</span>'
        '<span><i style="border-color:#94A3B8;border-style:dashed"></i>informativa (não conta no %)</span>'
        '</div>', unsafe_allow_html=True)

    blocos = []
    plantas_u = ordenar_plantas(df_u["_PLANTA"].unique().tolist())
    for planta in plantas_u:
        g = df_u[df_u["_PLANTA"] == planta].sort_values(["_INVIAVEL", "_PCT"])
        blocos.append(timeline_block_html(planta, g, aberto=(sel != "Todas" or len(plantas_u) == 1)))
    st.markdown("".join(blocos) + fluxo_tooltip_html(), unsafe_allow_html=True)
    if hasattr(st, "iframe"):          # Streamlit >= 1.5x
        st.iframe(TOOLTIP_JS, height=1)
    else:                              # versões antigas
        components.html(TOOLTIP_JS, height=0)


# ============================================================
# APP
# ============================================================
def sidebar():
    with st.sidebar:
        try:
            st.image("logo_delga.png", width="stretch")
        except Exception:
            st.markdown('<div style="font-size:24px;font-weight:900;color:#1400FF">GRUPO DELGA</div>', unsafe_allow_html=True)
        st.caption("Controle de Matéria-Prima · Bobinas BSW")
        st.divider()

        st.markdown("**🔐 Atualizar dados**")
        admin_pwd = get_secret("ADMIN_PASSWORD")
        if not admin_pwd:
            st.info("Defina ADMIN_PASSWORD nos Secrets do app para habilitar a atualização.")
            return
        senha = st.text_input("Senha", type="password", key="admin_pwd", label_visibility="collapsed",
                              placeholder="Senha do administrador")
        if not senha:
            return
        if not hmac.compare_digest(senha, str(admin_pwd)):
            st.error("Senha incorreta.")
            return
        arq = st.file_uploader("Excel atualizado", type=["xlsx", "xls"], key="admin_upload")
        if arq and st.button("📤 Salvar e publicar", type="primary", width="stretch"):
            with st.spinner("Salvando..."):
                ok, msg = save_data_to_github(arq.getvalue(), arq.name)
            (st.success if ok else st.error)(msg)
        st.caption("Rotina: dados atualizados toda segunda-feira.")


def main():
    sidebar()
    dados, updated, fonte = load_data()

    upd_txt = (f"Dados atualizados em {updated.strftime('%d/%m/%Y %H:%M')}" if updated is not None
               else f"Fonte: {fonte}" if fonte else "Sem dados")
    st.markdown(
        '<div class="app-head"><div class="app-head-l"><div class="app-logo">BSW</div><div>'
        '<div class="app-title">Controle de Matéria-Prima</div>'
        '<div class="app-sub">Bobinas BSW · Jan a Mai/2026 · Grupo Delga</div></div></div>'
        f'<div class="app-upd">{esc(upd_txt)}</div></div>',
        unsafe_allow_html=True,
    )

    if dados is None:
        st.info("Nenhum dado disponível ainda. Administrador: abra o menu lateral (›) e envie o Excel.")
        st.stop()

    df_raw, df_formulas, df_prop_raw, dept_map = dados
    df, col_names = process_data(df_raw)
    df_unidades, df_usinas = parse_formulas(df_formulas)
    df_prop = process_propostas(df_prop_raw, dept_map)

    col_media = col_names.get("media")
    if not col_media:
        st.error("Coluna de necessidade média não encontrada no arquivo.")
        st.stop()

    # KPIs gerais de matéria-prima
    if not df_unidades.empty:
        peso = df_unidades["peso_total"].sum()
        anal = df_unidades["peso_analisado"].sum()
        kpi_row([
            ("Bobinas", fmt_num(df_unidades["bobinas"].sum()), None, COR["azul"]),
            ("Peso médio total (MP)", f"{fmt_num(peso)} ton", None, COR["azul_claro"]),
            ("Peso médio analisado (MP)", f"{fmt_num(anal)} ton", None, COR["verde"]),
            ("% analisado", f"{fmt_num(anal / peso * 100 if peso else 0, 1)}%", None, COR["roxo"]),
        ])
    st.write("")

    t1, t2, t3 = st.tabs(["📊 Visão Geral", "🔍 Análises", "🛠️ Acompanhamento"])
    with t1:
        tab_visao_geral(df, col_names, col_media, df_usinas)
    with t2:
        tab_analises(df, col_media, df_unidades)
    with t3:
        tab_acompanhamento(df_prop)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error("O dashboard encontrou um erro e foi protegido para não ficar em tela branca.")
        st.exception(e)
