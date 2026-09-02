import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import warnings, os, pickle, datetime

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="Dashboard Executivo — Grupo Delga 2026",
    page_icon="https://grupodelga.com.br/wp-content/uploads/2024/11/logo-fa-e-clientes-grupo-whatsapp-9-300x300.png",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── PALETA DELGA ──────────────────────────────────────────────────────────────
NAVY   = "#1C2B4A"
RED    = "#C8202E"
SILVER = "#8A9BB0"
LIGHT  = "#F4F6F9"
WHITE  = "#FFFFFF"
GREEN  = "#1A7A3A"
AMBER  = "#E8A838"
TEAL   = "#20C997"

# ── TIPOS VÁLIDOS DE PROJETO ───────────────────────────────────────────────────
# Entram no DRE: BSW, Kaizen, Kaizen - Ganho Recorrente, Redução de Custo, Você Resolve
# NÃO entram no DRE: Kaizen - Custo Evitado, Kaizen - Capital de Giro
DRE_TIPOS = {
    'BSW', 'Kaizen', 'Kaizen - Ganho Recorrente',
    'Redução de custo', 'Redução de Custo', 'Redução de Custo ', 'Redução de Custos',
    'Você Resolve', 'Você resolve',
    'Estratégia Comercial', 'kaizen'
}
NAO_DRE_TIPOS = {
    'Kaizen - Custo Evitado', 'Kaizen - Capital de Giro',
    'Meta Executiva', 'Meta Executiva '
}
VALID_TIPOS = DRE_TIPOS | NAO_DRE_TIPOS

# Agrupamento para o gráfico de pilares (exibe todos os subtipos de Kaizen)
PILARES_EXIBE = {
    'Kaizen': 'Kaizen',
    'kaizen': 'Kaizen',
    'Kaizen - Ganho Recorrente': 'Kaizen - Ganho Recorrente',
    'Kaizen - Custo Evitado':    'Kaizen - Custo Evitado',
    'Kaizen - Capital de Giro':  'Kaizen - Capital de Giro',
    'Redução de custo':          'Redução de Custo',
    'Redução de Custo':          'Redução de Custo',
    'Redução de Custo ':         'Redução de Custo',
    'Você Resolve':              'Você Resolve',
    'Você resolve':              'Você Resolve',
    'BSW':                       'BSW',
    'Meta Executiva':            'Meta Executiva',
    'Estratégia Comercial':      'Estratégia Comercial',
}

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{{font-family:'Inter',sans-serif;}}
.block-container{{padding-top:0!important;padding-bottom:2rem;max-width:1440px;}}

/* ── HEADER ── */
.dh{{background:linear-gradient(135deg,{NAVY} 0%,#243B55 100%);
     padding:18px 32px;border-radius:0 0 12px 12px;
     display:flex;align-items:center;gap:18px;margin-bottom:20px;
     box-shadow:0 2px 12px rgba(28,43,74,.18);}}
.dh img{{height:44px;border-radius:6px;}}
.dh-t h1{{color:white;font-size:20px;font-weight:700;margin:0;letter-spacing:-.2px;}}
.dh-t p{{color:rgba(255,255,255,.55);font-size:11px;margin:2px 0 0;}}
.dh-b{{margin-left:auto;background:rgba(255,255,255,.12);
       color:rgba(255,255,255,.85);font-size:10px;
       font-weight:500;padding:5px 14px;border-radius:8px;white-space:nowrap;
       letter-spacing:.4px;border:1px solid rgba(255,255,255,.18);}}
.dh-b span.lbl{{font-size:9px;opacity:.7;display:block;letter-spacing:.6px;text-transform:uppercase;margin-bottom:1px;}}

/* ── KPI CARDS ── */
.kpi-wrap{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:20px;}}
.kpi-6{{grid-template-columns:repeat(6,1fr);}}
.kpi-7{{grid-template-columns:repeat(7,1fr);}}
.kpi-card{{background:white;border-radius:12px;padding:18px 20px;
           border-left:4px solid {NAVY};
           box-shadow:0 1px 4px rgba(28,43,74,.06),0 4px 16px rgba(28,43,74,.04);
           transition:box-shadow .2s;}}
.kpi-card:hover{{box-shadow:0 2px 8px rgba(28,43,74,.1),0 8px 24px rgba(28,43,74,.06);}}
.kpi-card.cr{{border-left-color:{RED};}}
.kpi-card.cg{{border-left-color:{GREEN};}}
.kpi-card.ca{{border-left-color:{AMBER};}}
.kpi-card.cs{{border-left-color:{SILVER};}}
.kpi-card.ct{{border-left-color:{TEAL};}}
.kpi-l{{font-size:9px;font-weight:600;color:{SILVER};text-transform:uppercase;
        letter-spacing:.9px;margin-bottom:6px;}}
.kpi-v{{font-size:24px;font-weight:700;color:{NAVY};line-height:1.1;margin-bottom:3px;}}
.kpi-s{{font-size:11px;color:#555;margin-bottom:2px;}}
.kpi-d{{font-size:10px;color:{SILVER};}}

/* ── SECTION CARD ── */
.sc{{background:white;border-radius:12px;padding:20px 22px;
     box-shadow:0 1px 4px rgba(28,43,74,.06),0 4px 16px rgba(28,43,74,.04);
     margin-bottom:16px;}}
.st{{font-size:11px;font-weight:700;color:{NAVY};text-transform:uppercase;
     letter-spacing:.7px;border-bottom:2px solid {RED};
     padding-bottom:7px;margin-bottom:14px;display:inline-block;}}

/* ── NOTA ── */
.nota{{background:#FFFBF0;border-left:3px solid {AMBER};border-radius:6px;
       padding:11px 16px;font-size:11px;color:#444;line-height:1.7;margin:14px 0;}}

/* ── TABELA ── */
.dt{{width:100%;border-collapse:collapse;font-size:12px;}}
.dt thead tr{{background:{NAVY};}}
.dt thead th{{color:white;padding:10px 12px;text-align:left;font-weight:600;
              font-size:11px;white-space:nowrap;}}
.dt thead th:first-child{{border-radius:6px 0 0 0;}}
.dt thead th:last-child{{border-radius:0 6px 0 0;}}
.dt tbody tr:nth-child(even){{background:#FAFBFC;}}
.dt tbody tr:hover{{background:#F0F4FA;transition:background .1s;}}
.dt tbody td{{padding:8px 12px;border-bottom:1px solid #EEF0F3;vertical-align:middle;}}
.dt tbody tr.tr-tot td{{background:{LIGHT};font-weight:700;
                         border-top:2px solid {NAVY};border-bottom:none;}}

/* ── MACRO TABLE ── */
.mct{{width:100%;border-collapse:collapse;font-size:12px;}}
.mct td{{padding:10px 12px;border-bottom:1px solid #EEF0F3;vertical-align:middle;}}
.mct tr:hover{{background:#F7F9FC;}}
.mch th{{color:white;padding:10px 12px;font-weight:600;font-size:11px;
         text-align:left;white-space:nowrap;}}
.mc-tot td{{background:{LIGHT};font-weight:700;border-top:2px solid {NAVY};}}

/* ── PROGRESS ── */
.pb{{display:flex;align-items:center;gap:8px;}}
.pb-bg{{height:7px;background:#E2E8F0;border-radius:4px;overflow:hidden;display:inline-block;}}
.pb-f{{height:100%;border-radius:4px;}}

/* ── BADGES ── */
.bdg{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:600;}}
.bg{{background:#E6F4EC;color:{GREEN};}}
.ba{{background:#FFF3E0;color:{AMBER};}}
.br{{background:#FDECEA;color:{RED};}}
.bk{{background:#F2F3F5;color:#555;}}
.bn{{background:#E8EDF5;color:{NAVY};}}

/* ── IMPEDIMENTO ── */
.imp{{background:#FFF8E1;border-left:3px solid {AMBER};border-radius:4px;
      padding:3px 7px;font-size:10px;color:#555;margin-top:3px;line-height:1.5;}}

/* ── TOGGLE LABELS — sem quebra de linha ── */
[data-testid="stToggle"] label {{white-space:nowrap!important;font-size:12px!important;font-weight:500!important;}}
[data-testid="stToggle"] {{align-items:center!important;}}

/* ── SECTION TOGGLE — botão − / + minimalista (sem círculo) ── */
[data-testid="stColumn"]:last-child button[kind="secondary"]{{
  font-size:18px!important;font-weight:200!important;
  color:{SILVER}!important;
  background:transparent!important;
  border:none!important;
  border-bottom:1.5px solid #DDE2EA!important;
  border-radius:0!important;
  width:24px!important;height:24px!important;
  padding:0!important;min-width:unset!important;
  line-height:1!important;
  display:flex!important;align-items:center!important;justify-content:center!important;
  margin-top:4px;transition:color .15s,border-color .15s;
}}
[data-testid="stColumn"]:last-child button[kind="secondary"]:hover{{
  color:{NAVY}!important;border-bottom-color:{NAVY}!important;
}}

/* ── LOGIN ── */
.lw{{max-width:360px;margin:80px auto;padding:40px;background:white;
     border-radius:14px;box-shadow:0 8px 32px rgba(28,43,74,.14);text-align:center;}}

#MainMenu{{visibility:hidden;}}footer{{visibility:hidden;}}
.stDeployButton{{display:none;}}header[data-testid="stHeader"]{{display:none;}}


div[data-testid="stExpander"]>div:first-child{{
  background:{LIGHT}!important;border:1px solid #E2E8F0!important;
  border-radius:8px!important;padding:4px 10px!important;}}
</style>
""", unsafe_allow_html=True)

# ── SENHA ─────────────────────────────────────────────────────────────────────
SENHA = "Delga01"
DADOS_PATH = "/tmp/delga_dados.pkl"

def check_password():
    if st.session_state.get("auth"): return True
    st.markdown("""<div class="lw">
      <div style="font-size:44px;margin-bottom:12px;">📊</div>
      <div style="font-size:20px;font-weight:700;color:#1C2B4A;margin-bottom:4px;">Grupo Delga</div>
      <div style="font-size:12px;color:#8A9BB0;margin-bottom:22px;">Dashboard Executivo 2026 — Acesso Restrito</div>
    </div>""", unsafe_allow_html=True)
    _, col, _ = st.columns([1,2,1])
    with col:
        pw = st.text_input("Senha", type="password", placeholder="Senha de acesso",
                           label_visibility="collapsed")
        if st.button("Entrar →", use_container_width=True, type="primary"):
            if pw == SENHA:
                st.session_state["auth"] = True; st.rerun()
            else:
                st.error("Senha incorreta.")
    return False

if not check_password(): st.stop()

# ── FORMATAÇÃO ────────────────────────────────────────────────────────────────
def fmt_mi(v):
    if pd.isna(v) or v is None: return "R$ 0"
    v = float(v)
    if abs(v) >= 1_000_000: return f"R$ {v/1_000_000:.2f} Mi"
    if abs(v) >= 1_000:     return "R$ " + f"{v:,.0f}".replace(",","X").replace(".",",").replace("X",".")
    return f"R$ {v:.0f}"

def fmt_brl(v):
    if pd.isna(v) or v is None or float(v)==0: return "—"
    return "R$ " + f"{float(v):,.0f}".replace(",","X").replace(".",",").replace("X",".")

def fmt_date(v):
    if pd.isna(v) or v is None or str(v).strip() in ("nan",""): return "—"
    s = str(v).strip()
    if s.upper() in ("N/A","NA","S/A"): return s
    try:
        if isinstance(v,(datetime.datetime,datetime.date)): return v.strftime("%m/%Y")
        if " " in s and ":" in s: return pd.to_datetime(s).strftime("%m/%Y")
        return s[:7]
    except: return s

def safe(v, d=0.0):
    try: return float(v) if pd.notna(v) else d
    except: return d

def pbar_html(pct, w=72):
    pc = min(float(pct)*100, 100)
    c = GREEN if pct>=.30 else (AMBER if pct>=.15 else RED)
    return (f'<div class="pb"><div class="pb-bg" style="width:{w}px;">'
            f'<div class="pb-f" style="width:{pc:.0f}%;background:{c};"></div></div>'
            f'<span style="font-size:11px;font-weight:600;">{pc:.1f}%</span></div>')

def bdg_status(pct):
    return ('<span class="bdg bg">DESTAQUE ✓</span>' if pct>=.30
            else '<span class="bdg ba">EM EXECUÇÃO</span>')

def bdg_custos(v):
    v=str(v).strip()
    if v=="OK":                       return '<span class="bdg bg">✓ OK</span>'
    if v in("Não Ok","NOK","Não OK"): return '<span class="bdg br">✗ NOK</span>'
    if v in("","nan"):                return '<span class="bdg bk">Pendente</span>'
    return f'<span class="bdg bk">{v}</span>'

def bdg_st(v):
    v=str(v)
    if "Concluído" in v: return '<span class="bdg bg">✓ Concluído</span>'
    if "Execução"  in v: return '<span class="bdg ba">⏳ Execução</span>'
    if "Não"       in v: return '<span class="bdg bk">Não iniciado</span>'
    return f'<span class="bdg bk">{v[:20]}</span>'

def bdg_tipo(v):
    v=str(v).strip()
    if "BSW"         in v: return f'<span class="bdg bn">BSW</span>'
    if "Capital"     in v: return f'<span class="bdg" style="background:#EDE7F6;color:#512DA8;">Cap. Giro</span>'
    if "Evitado"     in v: return f'<span class="bdg" style="background:#E3F2FD;color:#0D47A1;">C. Evitado</span>'
    if "Recorrente"  in v: return f'<span class="bdg ba">Kaizen GR</span>'
    if "Kaizen"      in v or "kaizen" in v: return f'<span class="bdg ba">Kaizen</span>'
    if "Redução"     in v: return f'<span class="bdg bg">Red. Custo</span>'
    if "Você"        in v or "Voce" in v: return f'<span class="bdg bk">Você Resolve</span>'
    if "Meta"        in v: return f'<span class="bdg bn">Meta Exec.</span>'
    if "Estratégia"  in v: return f'<span class="bdg bn">Est. Comercial</span>'
    return f'<span class="bdg bk">{v[:15]}</span>'

def is_dre(tipo):
    """True se o tipo de projeto entra no DRE."""
    return str(tipo).strip() in DRE_TIPOS

# ── PERSISTÊNCIA ──────────────────────────────────────────────────────────────
def save_bytes(b):
    with open(DADOS_PATH,"wb") as f: pickle.dump(b,f)
def load_bytes():
    if os.path.exists(DADOS_PATH):
        with open(DADOS_PATH,"rb") as f: return pickle.load(f)
    return None

# ── CARGA DO EXCEL ────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(fb):
    import io
    def xls(sh): return pd.read_excel(io.BytesIO(fb), sheet_name=sh, header=None)
    d = {}
    d["u5"]  = xls("5 Unidades  +")
    d["par"] = xls("Pareto")
    for s in ["Diadema","Jarinu","Ferraz","São Leopoldo","Anchieta",
              "Compras ","Vendas","Corporativo"]:
        d[s] = xls(s)
    return d

# ── EXTRAÇÃO — KPIs GLOBAIS ───────────────────────────────────────────────────
def extract_kpis(d):
    df = d["u5"]
    return dict(
        meta       =safe(df.iloc[6,3]),
        portfolio  =safe(df.iloc[6,4]),   # E — Retorno Previsto (Anual)
        ret_val_ano=safe(df.iloc[6,5]),   # F — Retorno Validado (Anual)  [NOVO big number]
        prev2026   =safe(df.iloc[6,6]),
        validado   =safe(df.iloc[6,7]),
        real       =safe(df.iloc[6,9]),
        extra_dre  =safe(df.iloc[6,10]),
        pct_ating  =safe(df.iloc[6,11]),
        inic       =int(safe(df.iloc[6,13])),
    )

def extract_plantas(d):
    df = d["u5"]
    # col4=Diadema,5=Ferraz,6=SãoLeopoldo,7=Jarinu,8=Anchieta
    cfg = [("Diadema",4,"Diadema"),("Ferraz",5,"Ferraz"),
           ("São Leopoldo",6,"São Leopoldo"),("Jarinu",7,"Jarinu"),("Anchieta",8,"Anchieta")]
    res=[]
    for nome,col,sh in cfg:
        res.append(dict(nome=nome,sheet=sh,
            meta    =safe(df.iloc[22,col]),
            prev    =safe(df.iloc[23,col]),
            prev2026=safe(df.iloc[24,col]),
            val     =safe(df.iloc[25,col]),
            real    =safe(df.iloc[26,col]),
            pct     =safe(df.iloc[27,col]),
            extra   =safe(df.iloc[28,col])))  # row28 = Extra DRE (nova linha v27)
    return res

def extract_areas(d):
    df = d["u5"]
    cfg = [("Corporativo",9,"Corporativo"),("Compras",10,"Compras "),("Vendas",11,"Vendas")]
    res=[]
    for nome,col,sh in cfg:
        res.append(dict(nome=nome,sheet=sh,
            meta    =safe(df.iloc[22,col]),
            prev    =safe(df.iloc[23,col]),
            prev2026=safe(df.iloc[24,col]),
            val     =safe(df.iloc[25,col]),
            real    =safe(df.iloc[26,col]),
            pct     =safe(df.iloc[27,col]),
            extra   =safe(df.iloc[28,col])))  # row28 = Extra DRE (nova linha v27)
    return res

def extract_pilares_global(d):
    """Pilares do painel 5 Unidades (rows 12-16)."""
    df = d["u5"]
    res=[]
    for i in range(12,22):
        nome=df.iloc[i,3]
        if pd.notna(nome) and str(nome) not in ("TOTAL",""):
            try:
                res.append(dict(
                    nome=str(nome),qtd=int(safe(df.iloc[i,4])),
                    prev=safe(df.iloc[i,5]),val=safe(df.iloc[i,6]),
                    pct=safe(df.iloc[i,7])))
            except: pass
    return res

def extract_pilares_local(projetos):
    """
    Gera resumo de pilares a partir da lista de projetos de uma unidade,
    incluindo todos os subtipos de Kaizen. Retorna lista ordenada.
    """
    from collections import defaultdict
    qtd  = defaultdict(int)
    prev = defaultdict(float)
    real = defaultdict(float)
    for p in projetos:
        nome = PILARES_EXIBE.get(p["tipo"], p["tipo"])
        qtd[nome]  += 1
        prev[nome] += p["previsto"]
        real[nome] += p["real_ano"]
    ORDER = ["BSW","Kaizen","Kaizen - Ganho Recorrente",
             "Kaizen - Custo Evitado","Kaizen - Capital de Giro",
             "Redução de Custo","Você Resolve","Meta Executiva","Estratégia Comercial"]
    res=[]
    for k in ORDER:
        if k in qtd:
            entra_dre = k not in ("Kaizen - Custo Evitado","Kaizen - Capital de Giro","Meta Executiva","Meta Executiva ")
            res.append(dict(nome=k,qtd=qtd[k],prev=prev[k],real=real[k],dre=entra_dre))
    for k in sorted(qtd):
        if k not in ORDER:
            entra_dre = k not in ("Kaizen - Custo Evitado","Kaizen - Capital de Giro","Meta Executiva","Meta Executiva ")
            res.append(dict(nome=k,qtd=qtd[k],prev=prev[k],real=real[k],dre=entra_dre))
    return res

def extract_evolucao(d):
    df = d["u5"]
    meses=["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    # v42: col21=label, cols 22-33=dados mensais
    # row55=Previsto, row56=Real, row58=Acum.Prev, row59=Acum.Real, row60=Projeção Meta,
    # row61=Acumulado prev.Custos (NOVO), row62=Custos — mensal não-acumulado (NOVO)
    max_col = min(34, df.shape[1])
    def row(r): return [safe(df.iloc[r,c]) for c in range(22, max_col)]
    def pad(lst): return (lst + [0]*12)[:12]

    return dict(meses=meses,
                prev            =pad(row(55)),
                real            =pad(row(56)),
                acum_prev       =pad(row(58)),
                acum_real       =pad(row(59)),
                proj_meta       =pad(row(60)),
                acum_prev_custos=pad(row(61)),
                prev_custos     =pad(row(62)))

# ── EXTRAÇÃO DE PROJETOS — FUNÇÃO CENTRAL ─────────────────────────────────────
# Offsets (relativos à coluna "Total Ano") das 12 colunas mensais individuais
# de cada linha de projeto — válido para Plantas, Compras e Vendas (mesmo padrão:
# Jan,Fev,Mar,[Tot1Tri],Abr,Mai,Jun,[Tot2Tri],Jul,Ago,Set,[Tot3Tri],Out,Nov,Dez,[Tot4Tri],TotalAno)
MES_OFFSETS = [-16,-15,-14, -12,-11,-10, -8,-7,-6, -4,-3,-2]

def norm_sim_nao(v):
    """Normaliza célula 'Sim/Não' -> 'Sim', 'Não' ou '' (em branco/não preenchido)."""
    s = str(v).strip().upper() if pd.notna(v) else ""
    if s in ("SIM", "S", "YES", "Y"): return "Sim"
    if s in ("NÃO", "NAO", "N"):      return "Não"
    return ""

def extract_projetos(df, start_row, col_tipo=0, col_nome=2, col_resp=5,
                     col_termino=7, col_custos=12, col_saving=13,
                     col_status=14, col_onde=15, col_data_lib=16,
                     col_prev_real=18,  # col com 'Previsto'/'Real'
                     col_total_ano=36,  # col com Total Ano (linha Real = V.Real acumulado)
                     col_previsto=8,    # col com PREVISTO(R$) original do projeto
                     col_aguardando=None):  # col com 'Aguardando Custos ?' (Sim/Não)
    """
    Extrai projetos de uma aba usando a lógica:
      - Linha Previsto: col_tipo in VALID_TIPOS + col_nome preenchido + col_prev_real='Previsto'
      - Linha Real: row+1, col_prev_real='Real', col_total_ano = valor real acumulado
    Para antes de qualquer tipo não reconhecido (seção SPD, etc.)

    Também replica, por projeto, as mesmas fórmulas auxiliares (AU/AV) que a
    planilha usa nos KPIs oficiais:
      previsto_2026  = (Previsto/12) * Qtd.Meses com retorno   [= coluna AU]
      validado_anual = (Saving Validado/Qtd.Meses) * 12        [= coluna AV]
    A coluna "Quantidade de meses com retorno" fica sempre 2 posições à
    direita da coluna "Previsto (R$)" no layout original da planilha.
    """
    col_qtd_meses = col_previsto + 2
    res = []
    i = start_row
    max_row = min(start_row + 600, df.shape[0] - 1)
    while i <= max_row:
        tipo = str(df.iloc[i, col_tipo]).replace("\n"," ").replace("\r"," ").strip()
        nome = str(df.iloc[i, col_nome]).replace("\n"," ").replace("\r"," ").strip()
        c_pr = str(df.iloc[i, col_prev_real]).strip() if df.shape[1] > col_prev_real else ""

        if tipo in VALID_TIPOS and nome not in ("", "nan") and c_pr != "Real":
            # Normalmente a linha "Previsto" tem essa palavra escrita na coluna
            # de rótulo (c_pr == "Previsto"). Mas em pelo menos um projeto
            # (Diadema, linha do item 13 "Redução de Componente") essa célula
            # de rótulo veio corrompida/zerada na planilha (mostra 0 em vez do
            # texto "Previsto") — como a linha "Real" (row+1) NUNCA tem Tipo/
            # Nome preenchidos, basta garantir que não é literalmente "Real"
            # pra não perder projetos legítimos com esse tipo de falha de
            # digitação/preenchimento na planilha de origem.
            # V.Previsto = col8 (PREVISTO R$ original do projeto)
            tot_prev  = safe(df.iloc[i, col_previsto])
            qtd_meses = safe(df.iloc[i, col_qtd_meses]) if df.shape[1] > col_qtd_meses else 0.0
            val_saving_i = safe(df.iloc[i, col_saving])
            previsto_2026  = (tot_prev/12)*qtd_meses if qtd_meses else 0.0
            validado_anual = (val_saving_i/qtd_meses)*12 if qtd_meses else 0.0

            meses_previsto = [
                safe(df.iloc[i, col_total_ano+off]) if 0 <= col_total_ano+off < df.shape[1] else 0.0
                for off in MES_OFFSETS
            ]

            # Linha Real (row+1)
            tot_real = 0.0
            meses_real = [0.0]*12
            if i+1 <= max_row:
                c_pr_next = str(df.iloc[i+1, col_prev_real]).strip() if df.shape[1] > col_prev_real else ""
                if c_pr_next == "Real":
                    tot_real = safe(df.iloc[i+1, col_total_ano])
                    meses_real = [
                        safe(df.iloc[i+1, col_total_ano+off]) if 0 <= col_total_ano+off < df.shape[1] else 0.0
                        for off in MES_OFFSETS
                    ]

            # Onde está parado (col15) e Data de liberação (col16)
            onde_parado = ""
            data_lib    = ""
            if col_onde is not None and df.shape[1] > col_onde:
                v_onde = df.iloc[i, col_onde]
                v_str = str(v_onde).strip() if pd.notna(v_onde) else ""
                if v_str not in ("", "nan"):
                    onde_parado = v_str
            if col_data_lib is not None and df.shape[1] > col_data_lib:
                v_dl = df.iloc[i, col_data_lib]
                v_str2 = str(v_dl).strip() if pd.notna(v_dl) else ""
                if v_str2 not in ("", "nan"):
                    data_lib = fmt_date(v_dl)

            # "Aguardando Custos ?" (Sim/Não) — coluna nova entre o checklist e o
            # validador de Custos. Se a coluna não existir na planilha (versões
            # antigas), fica "" e simplesmente não entra nas contagens.
            val_aguardando = ""
            if col_aguardando is not None and df.shape[1] > col_aguardando:
                val_aguardando = norm_sim_nao(df.iloc[i, col_aguardando])

            res.append(dict(
                row0           = i,   # linha (0-based, pandas) da linha "Previsto" — usada p/ casar checkboxes
                tipo           = tipo,
                nome           = nome,
                resp           = str(df.iloc[i, col_resp]).strip() if pd.notna(df.iloc[i, col_resp]) else "—",
                termino        = fmt_date(df.iloc[i, col_termino]),
                previsto       = tot_prev,
                previsto_2026  = previsto_2026,
                validado_anual = validado_anual,
                real_ano       = tot_real,
                val_custos     = str(df.iloc[i, col_custos]).strip() if pd.notna(df.iloc[i, col_custos]) else "",
                val_saving     = val_saving_i,
                status         = str(df.iloc[i, col_status]).strip() if pd.notna(df.iloc[i, col_status]) else "",
                meses_previsto = meses_previsto,
                meses_real     = meses_real,
                onde_parado    = onde_parado,
                data_lib       = data_lib,
                entra_dre      = is_dre(tipo),
                val_aguardando = val_aguardando,   # "Sim" / "Não" / ""
            ))
            i += 2  # pula Previsto + Real
        elif tipo not in ("", "nan") and tipo not in VALID_TIPOS:
            # "SPD"/"Adesão" é o marcador real de fim da tabela de projetos
            # em todas as abas (confirmado). Qualquer OUTRO texto não
            # reconhecido aqui é, na prática, erro de digitação no "Tipo" do
            # projeto (ex.: "Redução de Custos" com "s" a mais) — se a gente
            # simplesmente parasse a extração ali, TODO o resto da planilha
            # abaixo dessa linha seria descartado silenciosamente. Por isso
            # só paramos de verdade em "SPD"; qualquer outro texto estranho
            # é pulado (não vira projeto) e registrado como aviso.
            if "SPD" in tipo.upper() or "ADESÃO" in tipo.upper() or "ADESAO" in tipo.upper():
                break
            msg = f'tipo de projeto não reconhecido, linha ignorada: "{tipo}" (verifique a coluna Tipo)'
            if msg not in LAYOUT_WARNINGS:
                LAYOUT_WARNINGS.append(msg)
            i += 1
        else:
            i += 1
    return res

# ── VALIDAÇÃO DE LAYOUT ────────────────────────────────────────────────────────
# A planilha já mudou de layout uma vez nesta versão (inserção da coluna
# "Aguardando Custos ?", que empurrou Custos/Saving/Status/etc. uma posição
# pra direita). Pra não voltar a ler dado errado silenciosamente se isso
# acontecer de novo, cada get_proj_* confere se o cabeçalho esperado realmente
# está na posição esperada e acumula um aviso aqui se não bater.
LAYOUT_WARNINGS = []

def _check_header(df, header_row0, col, expected_substr, sheet_label, field_label):
    ok = False
    if col is not None and header_row0 < df.shape[0] and col < df.shape[1]:
        v = df.iloc[header_row0, col]
        if pd.notna(v):
            txt = str(v).replace("\n", " ").strip().upper()
            ok = expected_substr.upper() in txt
    if not ok:
        msg = (f"{sheet_label}: coluna \"{field_label}\" não encontrada na posição esperada "
               f"(col {col}) — a planilha pode ter mudado de layout, confira os números dessa aba.")
        if msg not in LAYOUT_WARNINGS:
            LAYOUT_WARNINGS.append(msg)
    return ok

def get_proj_planta(d, sheet_key):
    df = d.get(sheet_key)
    if df is None: return []
    # Plantas v10 (col nova "Aguardando Custos ?" inserida em 08/2026):
    #   col0=tipo, col2=nome, col5=resp, col7=term,
    #   col12=Aguardando Custos?, col13=custos, col14=saving, col15=status,
    #   col16=onde_parado, col17=data_lib,
    #   col19=Previsto/Real, col37=Total Ano
    hr = 53  # linha de cabeçalho (Excel row 54, 0-based)
    _check_header(df, hr, 12, "Aguardando Custos", sheet_key, "Aguardando Custos ?")
    _check_header(df, hr, 13, "VALIDADOR", sheet_key, "Validador OK/NOK Custos")
    _check_header(df, hr, 14, "SAVING VALIDADO", sheet_key, "Saving Validado")
    return extract_projetos(df, start_row=54,
        col_tipo=0, col_nome=2, col_resp=5, col_termino=7,
        col_aguardando=12, col_custos=13, col_saving=14, col_status=15,
        col_onde=16, col_data_lib=17,
        col_prev_real=19, col_total_ano=37, col_previsto=8)

def get_proj_compras(d):
    df = d.get("Compras ")
    if df is None: return []
    # Compras v10 (mesma inserção de coluna que as plantas):
    #   col0=tipo,col3=nome,col5=resp,col7=term,
    #   col12=Aguardando Custos?,col13=custos,col14=saving,col15=status,
    #   col16=onde,col17=data_lib,col20=Prev/Real,col38=TotalAno
    hr = 29  # linha de cabeçalho (Excel row 30, 0-based)
    _check_header(df, hr, 12, "Aguardando Custos", "Compras", "Aguardando Custos ?")
    _check_header(df, hr, 13, "VALIDADOR", "Compras", "Validador OK/NOK Custos")
    _check_header(df, hr, 14, "SAVING VALIDADO", "Compras", "Saving Validado")
    return extract_projetos(df, start_row=30,
        col_tipo=0, col_nome=3, col_resp=5, col_termino=7,
        col_aguardando=12, col_custos=13, col_saving=14, col_status=15,
        col_onde=16, col_data_lib=17,
        col_prev_real=20, col_total_ano=38, col_previsto=8)

def get_proj_vendas(d):
    df = d.get("Vendas")
    if df is None: return []
    # Vendas v13 (mesma inserção de coluna, ponto de inserção 1 coluna antes
    # por causa do layout mais estreito dessa aba):
    # col0=tipo, col1=nome, col4=resp, col6=termino, col7=previsto,
    # col11=Aguardando Custos?, col12=custos, col13=saving, col14=status,
    # col15=onde, col16=data_lib, col18=Previsto/Real, col36=Total Ano
    hr = 32  # linha de cabeçalho (Excel row 33, 0-based)
    _check_header(df, hr, 11, "Aguardando Custos", "Vendas", "Aguardando Custos ?")
    _check_header(df, hr, 12, "VALIDADOR", "Vendas", "Validador OK/NOK Custos")
    _check_header(df, hr, 13, "SAVING VALIDADO", "Vendas", "Saving Validado")
    return extract_projetos(df, start_row=33,
        col_tipo=0, col_nome=1, col_resp=4, col_termino=6,
        col_aguardando=11, col_custos=12, col_saving=13, col_status=14,
        col_onde=15, col_data_lib=16,
        col_prev_real=18, col_total_ano=36, col_previsto=7)

def get_proj_corporativo(d):
    df = d.get("Corporativo")
    if df is None: return []
    # Corporativo: NÃO recebeu a coluna nova "Aguardando Custos ?" (fica
    # sem esse dado, val_aguardando="" pra esses projetos).
    # col0=tipo, col1=nome, col4=resp, col6=termino, col7=previsto,
    # col11=custos, col12=saving, col13=status, col14=onde, col15=data_lib,
    # col17=Previsto/Real, col35=Total Ano
    hr = 19  # linha de cabeçalho (Excel row 20, 0-based)
    _check_header(df, hr, 11, "VALIDADOR", "Corporativo", "Validador OK/NOK Custos")
    _check_header(df, hr, 12, "SAVING VALIDADO", "Corporativo", "Saving Validado")
    return extract_projetos(df, start_row=20,
        col_tipo=0, col_nome=1, col_resp=4, col_termino=6,
        col_custos=11, col_saving=12, col_status=13,
        col_onde=14, col_data_lib=15,
        col_prev_real=17, col_total_ano=35, col_previsto=7)

# ── CHECKLIST DE 3 ETAPAS (pré-requisito p/ enviar o projeto a Custos) ────────
# Cada unidade preenche, por projeto, 3 checkboxes de formulário na coluna do
# "Processo de entrega do projeto a custos":
#   1) A3 e Estrutura Desenvolvido
#   2) Memória de Cálculo desenvolvido
#   3) Formalizado com Dep de Custos
# Esses checkboxes NÃO são valores de célula — são objetos de desenho (legacy
# VML/Form Controls) e por isso não aparecem via pandas/openpyxl cell.value.
# O estado (marcado/desmarcado) é lido direto do XML interno do .xlsx.
CHECKLIST_SHEETS   = ["Diadema","Ferraz","São Leopoldo","Jarinu","Anchieta","Compras ","Vendas","Corporativo"]
CHECKLIST_CAPTIONS = {"a3": "a3", "mem": "memoria", "formaliz": "formalizado"}

@st.cache_data(show_spinner=False)
def parse_checklist_custos(fb):
    """
    Extrai, para cada aba de unidade, o estado dos 3 checkboxes de checklist
    por linha de projeto ("row0" = linha 0-based da linha "Previsto", mesma
    indexação usada em extract_projetos).

    Retorna: {sheet_name: {row0: {"a3":bool, "memoria":bool, "formalizado":bool}}}
    Em caso de qualquer erro de parsing (planilha sem esses controles, versão
    diferente do Excel, etc.) retorna {} silenciosamente — o dashboard segue
    funcionando normalmente, só sem o detalhamento do checklist.
    """
    import zipfile, re as _re, io as _io
    out = {}
    try:
        z = zipfile.ZipFile(_io.BytesIO(fb))
        wb_xml  = z.read("xl/workbook.xml").decode("utf-8", "ignore")
        wb_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
        rel_map = dict(_re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', wb_rels))

        sheet_path = {}
        for m in _re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb_xml):
            name, rid = m.group(1), m.group(2)
            target = rel_map.get(rid)
            if target:
                sheet_path[name] = "xl/" + target.lstrip("/")

        for sheet_name in CHECKLIST_SHEETS:
            sp = sheet_path.get(sheet_name)
            if not sp or sp not in z.namelist():
                continue
            sheet_file = sp.split("/")[-1]
            rels_path = f"xl/worksheets/_rels/{sheet_file}.rels"
            if rels_path not in z.namelist():
                continue
            rels_xml = z.read(rels_path).decode("utf-8", "ignore")
            vml_m = _re.search(r'Target="([^"]*vmlDrawing[^"]*)"', rels_xml)
            if not vml_m:
                continue
            vml_path = "xl/drawings/" + vml_m.group(1).split("/")[-1]
            if vml_path not in z.namelist():
                continue
            vml = z.read(vml_path).decode("utf-8", "ignore")

            shapes = _re.findall(r'<v:shape\b[^>]*type="#_x0000_t201".*?</v:shape>', vml, _re.S)
            per_row = {}
            for s in shapes:
                if 'ObjectType="Checkbox"' not in s:
                    continue
                tb = _re.search(r'<v:textbox.*?<div[^>]*>(.*?)</div>', s, _re.S)
                caption = _re.sub(r'<[^>]+>', ' ', tb.group(1)) if tb else ""
                caption = _re.sub(r'\s+', ' ', caption).strip().lower()
                am = _re.search(r'<x:Anchor>\s*([\d,\s]+)</x:Anchor>', s)
                if not am:
                    continue
                anchor_vals = [v.strip() for v in am.group(1).split(',')]
                if len(anchor_vals) < 3:
                    continue
                row0 = int(anchor_vals[2])   # linha inicial (0-based) do shape
                checked = '<x:Checked>1</x:Checked>' in s
                key = next((v for k, v in CHECKLIST_CAPTIONS.items() if k in caption), None)
                if key is None:
                    continue
                per_row.setdefault(row0, {})[key] = checked
            out[sheet_name] = per_row
    except Exception:
        return {}
    return out

def attach_checklist(p, sheet_checklist):
    """
    Anexa ao projeto p o estado do checklist de 3 etapas, usando p['row0'].
    Os 3 checkboxes de um projeto ficam ancorados na própria linha "Previsto"
    (row0) e/ou na linha "Real" seguinte (row0+1) — juntamos as duas.
    """
    data = {}
    i = p.get('row0')
    if i is not None and sheet_checklist:
        data.update(sheet_checklist.get(i, {}))
        data.update(sheet_checklist.get(i + 1, {}))
    p['chk_a3']          = bool(data.get('a3', False))
    p['chk_memoria']     = bool(data.get('memoria', False))
    p['chk_formalizado'] = bool(data.get('formalizado', False))
    p['chk_completo']    = p['chk_a3'] and p['chk_memoria'] and p['chk_formalizado']
    return p

# ── VISÃO GERAL / BSW — camada de agregação bottom-up ─────────────────────────
UNIDADE_SHEETS_PLANTAS = ["Diadema","Ferraz","São Leopoldo","Jarinu","Anchieta"]

@st.cache_data(show_spinner=False)
def get_todos_projetos(fb):
    """
    Lista unificada de TODOS os projetos (todas as unidades e áreas), cada um
    tagueado com sua 'unidade' de origem. É a base do modo de visão BSW, que
    filtra esta mesma lista por tipo=='BSW' e recalcula tudo a partir dela,
    usando as MESMAS fórmulas (inclusive as auxiliares AU/AV de Previsto 2026
    e Retorno Validado Anualizado) que a planilha usa nos KPIs oficiais.
    Validado projeto a projeto contra a tabela nativa "SAVING ESPECULADO POR
    PILAR" da aba 5 Unidades — bate exato (Previsto, Saving Validado e Real).
    """
    checklist = parse_checklist_custos(fb)
    todos = []
    for sh in UNIDADE_SHEETS_PLANTAS:
        for p in get_proj_planta(D, sh):
            p = dict(p); p['unidade'] = sh
            attach_checklist(p, checklist.get(sh, {}))
            todos.append(p)
    for p in get_proj_compras(D):
        p = dict(p); p['unidade'] = "Compras"
        attach_checklist(p, checklist.get("Compras ", {}))
        todos.append(p)
    for p in get_proj_vendas(D):
        p = dict(p); p['unidade'] = "Vendas"
        attach_checklist(p, checklist.get("Vendas", {}))
        todos.append(p)
    for p in get_proj_corporativo(D):
        p = dict(p); p['unidade'] = "Corporativo"
        attach_checklist(p, checklist.get("Corporativo", {}))
        todos.append(p)
    return todos

def compute_status_custos(projetos):
    """
    Agrega, por unidade, quantos projetos estão:
      - validado     : Custos = OK
      - nao_validado : Custos = Não Ok / NOK
      - em_branco    : Custos ainda sem validação (célula vazia)
    E, dentro de 'em_branco', usa o checklist de 3 etapas (A3/Estrutura,
    Memória de Cálculo, Formalizado com Custos) pra indicar o motivo:
      - checklist completo (3/3) e projeto entra no DRE -> falta_custos
        (unidade já fez a parte dela, falta só Custos aprovar)
      - checklist incompleto (ou não entra no DRE)       -> falta_unidade
        (unidade ainda precisa terminar e enviar pra Custos)
    """
    from collections import defaultdict
    res = defaultdict(lambda: dict(total=0, validado=0, nao_validado=0, em_branco=0,
                                    falta_custos=0, falta_unidade=0))
    for p in projetos:
        u = p.get('unidade', '—')
        r = res[u]
        r['total'] += 1
        custos = str(p.get('val_custos', '')).strip()
        if custos == "OK":
            r['validado'] += 1
        elif custos in ("Não Ok", "NOK", "Não OK"):
            r['nao_validado'] += 1
        else:
            r['em_branco'] += 1
            if p.get('chk_completo') and p.get('entra_dre'):
                r['falta_custos'] += 1
            else:
                r['falta_unidade'] += 1
    return dict(res)

def compute_kpis_bottom_up(projetos, meta):
    """
    Recalcula os big numbers a partir de uma lista de projetos já filtrada
    (ex.: só tipo=='BSW'). Meta NUNCA é recalculada aqui — é sempre a meta
    do grupo, inalterada pelo filtro Geral/BSW (única exceção pedida).
    """
    portfolio   = sum(p['previsto'] for p in projetos)
    ret_val_ano = sum(p['validado_anual'] for p in projetos)
    prev2026    = sum(p['previsto_2026'] for p in projetos)
    validado    = sum(p['val_saving'] for p in projetos)
    real        = sum(p['real_ano'] for p in projetos if p['entra_dre'])
    extra_dre   = sum(p['real_ano'] for p in projetos if not p['entra_dre'])
    pct_ating   = real/meta if meta > 0 else 0.0
    return dict(meta=meta, portfolio=portfolio, ret_val_ano=ret_val_ano,
                prev2026=prev2026, validado=validado, real=real,
                extra_dre=extra_dre, pct_ating=pct_ating, inic=len(projetos))

def compute_evolucao_bottom_up(projetos, ev_geral):
    """
    Recalcula as séries mensais do gráfico de Evolução somando as colunas
    mensais 'Previsto' e 'Real' de cada projeto filtrado. 'Projeção da Meta'
    é sempre a linha global (a meta não muda com o filtro).
    """
    prev = [0.0]*12
    real = [0.0]*12
    for p in projetos:
        mp = p.get('meses_previsto', [0.0]*12)
        mr = p.get('meses_real', [0.0]*12)
        for i in range(12):
            prev[i] += mp[i]
            real[i] += mr[i]
    acum_prev, acum_real, tp, tr = [], [], 0.0, 0.0
    for i in range(12):
        tp += prev[i]; tr += real[i]
        acum_prev.append(tp); acum_real.append(tr)

    # Previsto por Custos (2026) — mesma regra usada na linha nativa da
    # planilha: só projetos com OK de Custos e Saving Validado preenchido.
    prev_custos = [0.0]*12
    for p in projetos:
        if str(p.get('val_custos','')).strip() == "OK" and safe(p.get('val_saving',0)) > 0:
            mp = p.get('meses_previsto', [0.0]*12)
            for i in range(12):
                prev_custos[i] += mp[i]
    acum_prev_custos, tc = [], 0.0
    for v in prev_custos:
        tc += v
        acum_prev_custos.append(tc)

    return dict(meses=ev_geral["meses"], prev=prev, real=real,
                acum_prev=acum_prev, acum_real=acum_real,
                proj_meta=ev_geral["proj_meta"],
                prev_custos=prev_custos, acum_prev_custos=acum_prev_custos)

def compute_macro_bottom_up(projetos, lista_geral):
    """
    Recalcula as linhas da tabela macro (Plantas/Áreas) a partir da lista de
    projetos filtrada, mantendo a Meta de cada unidade inalterada (exceção).
    """
    res = []
    for it in lista_geral:
        proj_unidade = [p for p in projetos if p.get('unidade') == it['nome']]
        prev     = sum(p['previsto'] for p in proj_unidade)
        prev2026 = sum(p['previsto_2026'] for p in proj_unidade)
        val      = sum(p['val_saving'] for p in proj_unidade)
        real     = sum(p['real_ano'] for p in proj_unidade if p['entra_dre'])
        extra    = sum(p['real_ano'] for p in proj_unidade if not p['entra_dre'])
        pct      = real/it['meta'] if it['meta'] > 0 else 0.0
        res.append(dict(nome=it['nome'], sheet=it.get('sheet'), meta=it['meta'],
                         prev=prev, prev2026=prev2026, val=val, real=real,
                         extra=extra, pct=pct))
    return res

def extract_ranking(d):
    df = d["u5"]
    res=[]
    for i in range(53,137):
        uni=df.iloc[i,4]; nome=df.iloc[i,5]
        if not pd.notna(uni) or not pd.notna(nome): continue
        res.append(dict(
            pos=int(safe(df.iloc[i,3],i-52)),
            uni=str(uni),nome=str(nome),
            status=str(df.iloc[i,7]).strip() if pd.notna(df.iloc[i,7]) else "",
            custos=str(df.iloc[i,8]).strip() if pd.notna(df.iloc[i,8]) else "",
            prev26=safe(df.iloc[i,9]),prev_mo=safe(df.iloc[i,10]),real=safe(df.iloc[i,11]),
        ))
    return res

# ── GRÁFICOS ──────────────────────────────────────────────────────────────────
PAL = [NAVY,"#2C4F7C","#4A7AB5",SILVER,"#A8C8E8"]

def chart_funnel(kpis):
    stages = ["Meta do Grupo","Portfólio Previsto (Anualizado)","Previsto 2026","Validado Custos","Real DRE"]
    values = [kpis["meta"],kpis["portfolio"],kpis["prev2026"],kpis["validado"],kpis["real"]]
    pcts   = [f"{v/kpis['meta']*100:.1f}%" for v in values]
    colors = [NAVY,"#2C5F8A","#4A90D9",AMBER,GREEN]

    fig = go.Figure()
    # Barras horizontais simulando funil (comprimento proporcional)
    for idx,(stage,val,pct,color) in enumerate(zip(stages,values,pcts,colors)):
        w = val/kpis["meta"]
        fig.add_trace(go.Bar(
            x=[val], y=[stage],
            orientation="h",
            marker=dict(color=color, line=dict(width=0)),
            text=f"  <b>{fmt_mi(val)}</b>  <span style='opacity:.7'>({pct})</span>",
            textposition="outside",
            textfont=dict(size=12,color="#333"),
            hovertemplate=f"<b>{stage}</b><br>{fmt_mi(val)}<br>{pct} da meta<extra></extra>",
            showlegend=False,
            base=[(kpis["meta"]-val)/2],  # centraliza para efeito funil
        ))

    fig.update_layout(
        barmode="overlay",
        xaxis=dict(visible=False, range=[0, kpis["meta"]*1.3]),
        yaxis=dict(autorange="reversed",tickfont=dict(size=12,color="#444")),
        margin=dict(l=130,r=150,t=10,b=10),
        height=310,
        paper_bgcolor="white", plot_bgcolor="white",
        font=dict(family="Inter"),
    )
    return fig

def chart_gauge(pct):
    clr = GREEN if pct>=.30 else (AMBER if pct>=.15 else RED)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=pct*100,
        number=dict(suffix="%",font=dict(size=40,color=NAVY,family="Inter")),
        gauge=dict(
            axis=dict(range=[0,100],ticksuffix="%",tickfont=dict(size=10)),
            bar=dict(color=clr,thickness=0.28),
            bgcolor="white",borderwidth=0,
            steps=[dict(range=[0,30],color="#FFEBEE"),
                   dict(range=[30,70],color="#FFF3E0"),
                   dict(range=[70,100],color="#E8F5E9")],
            threshold=dict(line=dict(color=RED,width=3),thickness=.75,value=100),
        ),
        title=dict(text="<b>Atingimento da Meta</b>",font=dict(size=12,color=SILVER)),
    ))
    fig.update_layout(margin=dict(l=20,r=20,t=50,b=10),height=280,paper_bgcolor="white")
    return fig

def chart_evolucao(ev, series):
    """
    Gráfico de evolução com:
    - Linhas: Acumulado Previsto, Acumulado Real, Projeção da Meta
    - Barras mensais: Previsto Mensal (azul claro), Real Mensal (verde)
    - Hover ordenado do maior para o menor
    """
    BARRAS = {"Previsto Mensal", "Real Mensal"}

    # Configurações de cada série
    cfg = {
        "Acumulado Previsto":         dict(data=ev["acum_prev"],        color=NAVY,      dash="solid", type="line"),
        "Previsto por Custos (2026)": dict(data=ev["acum_prev_custos"], color="#6C3EB5", dash="dot",   type="line"),
        "Acumulado Real":             dict(data=ev["acum_real"],        color=GREEN,     dash="solid", type="line"),
        "Projeção da Meta":           dict(data=ev["proj_meta"],        color=RED,       dash="dash",  type="line"),
        "Previsto Mensal":            dict(data=ev["prev"],             color="#7EB3D8",               type="bar"),
        "Real Mensal":                dict(data=ev["real"],             color="#52A97C",               type="bar"),
    }

    fig = go.Figure()

    # Barras primeiro (ficam atrás das linhas)
    for s in series:
        if s not in cfg: continue
        c = cfg[s]
        if c["type"] == "bar":
            fig.add_trace(go.Bar(
                x=ev["meses"], y=c["data"], name=s,
                marker=dict(color=c["color"], opacity=0.75, line=dict(width=0)),
                hovertemplate=f"<b>{s}</b><br>%{{x}}: R$ %{{y:,.0f}}<extra></extra>",
            ))

    # Linhas por cima
    for s in series:
        if s not in cfg: continue
        c = cfg[s]
        if c["type"] == "line":
            fig.add_trace(go.Scatter(
                x=ev["meses"], y=c["data"], mode="lines+markers", name=s,
                line=dict(color=c["color"], width=2.5, dash=c["dash"]),
                marker=dict(size=6, color=c["color"]),
                hovertemplate=f"<b>{s}</b><br>%{{x}}: R$ %{{y:,.0f}}<extra></extra>",
            ))

    fig.update_layout(
        barmode="group",
        bargap=0.25,
        bargroupgap=0.05,
        xaxis=dict(showgrid=True, gridcolor="#F0F4F8"),
        yaxis=dict(tickformat=",.0f", showgrid=True, gridcolor="#F0F4F8", title="R$"),
        legend=dict(orientation="h", y=1.05, x=0.5, xanchor="center", font=dict(size=11)),
        margin=dict(l=80, r=20, t=50, b=40),
        height=420,
        paper_bgcolor="white", plot_bgcolor="white",
        # Hover ordenado do maior para o menor
        hovermode="x unified",
        hoverlabel=dict(bgcolor="white", font_size=12, font_family="Inter"),
        font=dict(family="Inter"),
    )
    # Ordenar hover do maior para o menor por valor
    fig.update_layout(hoversubplots="axis")
    return fig

def chart_donut(labels,values,colors):
    total = sum(values)
    if total == 0:
        total = 1  # evita ZeroDivisionError quando todos os valores são zero
    txt = [f"  {labels[i]}  {values[i]/total*100:.1f}%  {fmt_mi(values[i])}"
           for i in range(len(labels))]
    fig = go.Figure(go.Pie(labels=txt,values=values,hole=0.62,
        marker=dict(colors=colors),textinfo="none",
        hovertemplate="<b>%{label}</b><extra></extra>"))
    fig.update_layout(showlegend=True,
        legend=dict(orientation="v",y=0.5,x=0.55,yanchor="middle",font=dict(size=11)),
        margin=dict(l=10,r=10,t=10,b=30),height=280,
        paper_bgcolor="white",plot_bgcolor="white",
        annotations=[dict(text=f"<b>{fmt_mi(sum(values))}</b>",x=0.22,y=-0.08,
                          font_size=12,showarrow=False)],
        font=dict(family="Inter"))
    return fig

def chart_pilares(pilares_global, real_total):
    """Gráfico de pilares do painel global (5 Unidades)."""
    labels=[p["nome"] for p in pilares_global]
    previsto=[p["prev"] for p in pilares_global]
    validado=[p["val"]  for p in pilares_global]
    tv=sum(validado)
    real_est=[v/tv*real_total if tv>0 else 0 for v in validado]
    fig=go.Figure()
    fig.add_trace(go.Bar(name="Previsto",x=labels,y=previsto,marker_color="#B8D4E8"))
    fig.add_trace(go.Bar(name="Validado",x=labels,y=validado,marker_color=NAVY))
    fig.add_trace(go.Bar(name="Real DRE",x=labels,y=real_est, marker_color=GREEN))
    fig.update_layout(barmode="group",
        yaxis=dict(tickformat=",.0f",showgrid=True,gridcolor="#F0F4F8"),
        legend=dict(orientation="h",y=1.05,x=1,xanchor="right",font=dict(size=11)),
        margin=dict(l=60,r=20,t=40,b=60),height=300,
        paper_bgcolor="white",plot_bgcolor="white",
        bargap=0.28,font=dict(family="Inter"))
    return fig

# ── HTML HELPERS ──────────────────────────────────────────────────────────────
def th(*cols):
    ths = "".join(f"<th>{c}</th>" for c in cols)
    return f"<table class='dt'><thead><tr>{ths}</tr></thead><tbody>"

def render_proj_filtros(projetos, key_prefix=""):
    """
    Renderiza controles de filtro + ordenação no estilo Excel.
    Retorna a lista filtrada/ordenada de projetos.
    """
    if not projetos:
        return projetos

    TIPOS_DISP   = sorted({p["tipo"] for p in projetos if p["tipo"]})
    STATUS_DISP  = sorted({p["status"] for p in projetos if p["status"]})
    CUSTOS_DISP  = sorted({p["val_custos"] for p in projetos if p["val_custos"]})

    fc1, fc2, fc3, fc4, fc5 = st.columns([2,2,2,2,3])

    with fc1:
        f_tipo = st.multiselect("Tipo", TIPOS_DISP, default=[],
                                key=f"{key_prefix}_ftipo", placeholder="Todos")
    with fc2:
        f_status = st.multiselect("Status", STATUS_DISP, default=[],
                                  key=f"{key_prefix}_fstatus", placeholder="Todos")
    with fc3:
        f_custos = st.multiselect("Custos", CUSTOS_DISP, default=[],
                                  key=f"{key_prefix}_fcustos", placeholder="Todos")
    with fc4:
        sort_col = st.selectbox("Ordenar por",
            ["Nome (A→Z)", "Nome (Z→A)",
             "V.Previsto ↓", "V.Previsto ↑",
             "V.Validado ↓", "V.Validado ↑",
             "V.Real ↓",    "V.Real ↑",
             "Status (A→Z)","Tipo (A→Z)"],
            index=0, key=f"{key_prefix}_sort")
    with fc5:
        f_nome = st.text_input("🔍 Buscar projeto", value="",
                               key=f"{key_prefix}_fnome", placeholder="Filtrar por nome...")

    # Aplicar filtros
    res = projetos[:]
    if f_tipo:   res = [p for p in res if p["tipo"]       in f_tipo]
    if f_status: res = [p for p in res if p["status"]     in f_status]
    if f_custos: res = [p for p in res if p["val_custos"] in f_custos]
    if f_nome:   res = [p for p in res if f_nome.lower() in p["nome"].lower()]

    # Ordenar
    sort_map = {
        "Nome (A→Z)":    (lambda p: p["nome"].lower(),       False),
        "Nome (Z→A)":    (lambda p: p["nome"].lower(),       True),
        "V.Previsto ↓":  (lambda p: p["previsto"],           True),
        "V.Previsto ↑":  (lambda p: p["previsto"],           False),
        "V.Validado ↓":  (lambda p: p["val_saving"],         True),
        "V.Validado ↑":  (lambda p: p["val_saving"],         False),
        "V.Real ↓":      (lambda p: p["real_ano"],           True),
        "V.Real ↑":      (lambda p: p["real_ano"],           False),
        "Status (A→Z)":  (lambda p: p["status"].lower(),     False),
        "Tipo (A→Z)":    (lambda p: p["tipo"].lower(),       False),
    }
    key_fn, rev = sort_map.get(sort_col, (lambda p: p["nome"].lower(), False))
    res = sorted(res, key=key_fn, reverse=rev)

    return res


def projetos_por_pilar_html(projetos, key_prefix=""):
    """Exibe projetos agrupados por Tipo/Pilar com cabeçalho de totais."""
    if not projetos:
        return [], ""

    PILARES_ORDER = [
        "BSW","Kaizen","Kaizen - Ganho Recorrente","Kaizen - Custo Evitado",
        "Kaizen - Capital de Giro","Redução de Custo","Redução de custo",
        "Você Resolve","Você resolve","Meta Executiva","Meta Executiva ",
        "Estratégia Comercial",
    ]

    STATUS_DISP = sorted({p["status"] for p in projetos if p["status"]})
    fc1, fc2, fc3 = st.columns([3, 2, 4])
    with fc1:
        f_status = st.multiselect("Status:", STATUS_DISP, default=[],
                                  key=f"{key_prefix}_fstatus", placeholder="Todos")
    with fc2:
        sort_col = st.selectbox("Ordenar por",
            ["Nome (A→Z)","V.Previsto ↓","V.Previsto ↑","V.Real ↓","V.Real ↑"],
            index=0, key=f"{key_prefix}_sort")
    with fc3:
        f_nome = st.text_input("🔍 Buscar projeto", value="",
                               key=f"{key_prefix}_fnome", placeholder="Filtrar por nome...")

    res = projetos[:]
    if f_status: res = [p for p in res if p["status"] in f_status]
    if f_nome:   res = [p for p in res if f_nome.lower() in p["nome"].lower()]

    sort_map = {
        "Nome (A→Z)":   (lambda p: p["nome"].lower(), False),
        "V.Previsto ↓": (lambda p: p["previsto"],     True),
        "V.Previsto ↑": (lambda p: p["previsto"],     False),
        "V.Real ↓":     (lambda p: p["real_ano"],     True),
        "V.Real ↑":     (lambda p: p["real_ano"],     False),
    }
    key_fn, rev = sort_map.get(sort_col, (lambda p: p["nome"].lower(), False))
    res = sorted(res, key=key_fn, reverse=rev)

    from collections import OrderedDict
    grupos = OrderedDict()
    for tipo in PILARES_ORDER:
        grupos[tipo] = []
    for p in res:
        t = p["tipo"]
        if t not in grupos:
            grupos[t] = []
        grupos[t].append(p)

    # Renderiza cada pilar como bloco expandível via session_state
    # Retorna lista filtrada + flag vazia (HTML renderizado diretamente via st)
    for tipo, projs in grupos.items():
        if not projs:
            continue
        dre_flag = is_dre(tipo)
        dre_lbl  = "✓ DRE" if dre_flag else "↷ N/DRE"
        dre_clr  = "#7BDD9A" if dre_flag else "rgba(255,255,255,.45)"
        tot_prev = sum(p["previsto"]   for p in projs)
        tot_val  = sum(p["val_saving"] for p in projs)
        tot_real = sum(p["real_ano"]   for p in projs)
        n_p      = len(projs)

        # Cabeçalho do pilar (sempre visível)
        header_html = f"""<div style="background:{NAVY};border-radius:8px;
            padding:10px 16px;display:flex;align-items:center;gap:16px;margin-top:12px;">
          <div>
            <span style="color:white;font-size:12px;font-weight:700;">{tipo}</span>
            <span style="color:{dre_clr};font-size:9px;margin-left:8px;font-weight:600;">{dre_lbl}</span>
          </div>
          <div style="margin-left:auto;display:flex;gap:28px;align-items:center;">
            <div style="text-align:center;">
              <div style="color:rgba(255,255,255,.5);font-size:9px;text-transform:uppercase;letter-spacing:.5px;">Projetos</div>
              <div style="color:white;font-size:14px;font-weight:700;">{n_p}</div>
            </div>
            <div style="text-align:center;">
              <div style="color:rgba(255,255,255,.5);font-size:9px;text-transform:uppercase;letter-spacing:.5px;">Previsto</div>
              <div style="color:#C8D8EE;font-size:14px;font-weight:700;">{fmt_mi(tot_prev)}</div>
            </div>
            <div style="text-align:center;">
              <div style="color:rgba(255,255,255,.5);font-size:9px;text-transform:uppercase;letter-spacing:.5px;">Validado</div>
              <div style="color:#7BDD9A;font-size:14px;font-weight:700;">{fmt_mi(tot_val)}</div>
            </div>
            <div style="text-align:center;">
              <div style="color:rgba(255,255,255,.5);font-size:9px;text-transform:uppercase;letter-spacing:.5px;">Real Acum.</div>
              <div style="color:#7BDD9A;font-size:14px;font-weight:700;">{fmt_mi(tot_real)}</div>
            </div>
          </div>
        </div>"""

        # Renderiza cabeçalho do pilar (sempre visível)
        st.markdown(header_html, unsafe_allow_html=True)

        # Tabela de projetos — expander nativo com +/−
        with st.expander("", expanded=True):
            col_headers = "".join(
                f'<th style="padding:8px 12px;text-align:left;font-size:10px;font-weight:600;'
                f'color:{SILVER};text-transform:uppercase;letter-spacing:.4px;background:#F4F6F9;">{c}</th>'
                for c in ["Projeto","Responsável","Término","Previsto (R$)","Saving Validado",
                          "Real Acum.","Custos","Status","Onde Parado","Prev.Lib."]
            )
            rows = ""
            for p in projs:
                real_v = p["real_ano"]
                real_s = fmt_brl(real_v) if real_v and real_v != 0 else "—"
                rc     = GREEN if real_v > 0 else ("#DC3545" if real_v < 0 else "#999")
                concluido = "Concluído" in str(p.get("status",""))
                onde  = p.get("onde_parado","")
                dlib  = p.get("data_lib","")
                onde_html = f'<span style="font-size:10px;color:#555;">{onde}</span>' if (onde and not concluido) else '<span style="color:#ccc;font-size:10px;">—</span>'
                data_html = f'<span style="font-size:10px;color:{AMBER};font-weight:600;">{dlib}</span>' if (dlib and not concluido) else '<span style="color:#ccc;font-size:10px;">—</span>'

                # Detectar se término passou e projeto não está concluído
                termo_str = str(p.get("termino","")).strip()
                atrasado = False
                if not concluido and termo_str and termo_str != "—":
                    try:
                        termo_dt = datetime.datetime.strptime(termo_str[:7], "%m/%Y")
                        hoje = datetime.datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
                        if termo_dt < hoje:
                            atrasado = True
                    except:
                        pass

                row_style = "border-bottom:1px solid #EEF0F3;" + ("background:#FFF5F5;" if atrasado else "")
                txt_style = "color:#C8202E;" if atrasado else ""

                rows += f"""<tr style="{row_style}">
                  <td style="padding:8px 12px;font-size:11px;{txt_style}"><b>{p['nome']}</b></td>
                  <td style="padding:8px 12px;font-size:11px;white-space:nowrap;{txt_style}">{p['resp']}</td>
                  <td style="padding:8px 12px;font-size:11px;white-space:nowrap;{txt_style}">{p['termino']}</td>
                  <td style="padding:8px 12px;text-align:right;font-size:11px;{txt_style}">{fmt_brl(p['previsto'])}</td>
                  <td style="padding:8px 12px;text-align:right;font-size:11px;color:{TEAL};">{fmt_brl(p['val_saving'])}</td>
                  <td style="padding:8px 12px;text-align:right;font-size:11px;color:{rc};font-weight:600;">{real_s}</td>
                  <td style="padding:8px 12px;">{bdg_custos(p['val_custos'])}</td>
                  <td style="padding:8px 12px;white-space:nowrap;">{bdg_st(p['status'])}</td>
                  <td style="padding:8px 12px;">{onde_html}</td>
                  <td style="padding:8px 12px;">{data_html}</td>
                </tr>"""

            st.markdown(
                f'<table style="width:100%;border-collapse:collapse;font-size:12px;">'
                f'<thead><tr>{col_headers}</tr></thead>'
                f'<tbody>{rows}</tbody></table>',
                unsafe_allow_html=True
            )

    # Barra de TOTAL no fim — mesmo estilo dos pilares
    tot_all_prev = sum(p["previsto"]   for p in res) if res else 0
    tot_all_val  = sum(p["val_saving"] for p in res) if res else 0
    tot_all_real = sum(p["real_ano"]   for p in res) if res else 0
    n_all        = len(res)
    st.markdown(f"""<div style="background:{NAVY};border-radius:8px;
        padding:10px 16px;display:flex;align-items:center;gap:16px;margin-top:16px;
        border:2px solid rgba(255,255,255,.12);">
      <span style="color:white;font-size:12px;font-weight:700;letter-spacing:.3px;">TOTAL</span>
      <div style="margin-left:auto;display:flex;gap:28px;align-items:center;">
        <div style="text-align:center;">
          <div style="color:rgba(255,255,255,.5);font-size:9px;text-transform:uppercase;letter-spacing:.5px;">Projetos</div>
          <div style="color:white;font-size:14px;font-weight:700;">{n_all}</div>
        </div>
        <div style="text-align:center;">
          <div style="color:rgba(255,255,255,.5);font-size:9px;text-transform:uppercase;letter-spacing:.5px;">Previsto</div>
          <div style="color:#C8D8EE;font-size:14px;font-weight:700;">{fmt_mi(tot_all_prev)}</div>
        </div>
        <div style="text-align:center;">
          <div style="color:rgba(255,255,255,.5);font-size:9px;text-transform:uppercase;letter-spacing:.5px;">Validado</div>
          <div style="color:#7BDD9A;font-size:14px;font-weight:700;">{fmt_mi(tot_all_val)}</div>
        </div>
        <div style="text-align:center;">
          <div style="color:rgba(255,255,255,.5);font-size:9px;text-transform:uppercase;letter-spacing:.5px;">Real Acum.</div>
          <div style="color:#7BDD9A;font-size:14px;font-weight:700;">{fmt_mi(tot_all_real)}</div>
        </div>
      </div>
    </div>""", unsafe_allow_html=True)

    return res, ""

def proj_table_html(projetos):
    """Tabela de projetos com colunas Onde Parado e Data Liberação."""
    if not projetos:
        return "<p style='color:#999;font-size:12px;padding:6px 0;'>Nenhum projeto encontrado.</p>"
    rows = ""
    for p in projetos:
        real_v = p["real_ano"]
        real_s = fmt_brl(real_v) if real_v and real_v != 0 else "—"
        real_c = GREEN if real_v and real_v > 0 else "#999"

        dre_icon = (f'<span title="Entra no DRE" style="color:{GREEN};font-size:9px;">✓ DRE</span>'
                    if p["entra_dre"] else
                    f'<span title="Não entra no DRE" style="color:{SILVER};font-size:9px;">↷ N/DRE</span>')

        # Onde parado + Data lib — só mostra se não concluído
        concluido = "Concluído" in str(p.get("status",""))
        onde = p.get("onde_parado","")
        data_lib = p.get("data_lib","")
        if concluido:
            onde_html = '<span style="color:#ccc;font-size:10px;">—</span>'
            data_html = '<span style="color:#ccc;font-size:10px;">—</span>'
        else:
            onde_html = (f'<span style="font-size:10px;color:#555;">{onde}</span>' if onde
                         else '<span style="color:#ccc;font-size:10px;">—</span>')
            data_html = (f'<span style="font-size:10px;color:{AMBER};font-weight:600;">{data_lib}</span>' if data_lib
                         else '<span style="color:#ccc;font-size:10px;">—</span>')

        rows += f"""<tr>
          <td style="white-space:nowrap;">{bdg_tipo(p['tipo'])}<br>{dre_icon}</td>
          <td style="max-width:220px;font-size:11px;"><b>{p['nome']}</b></td>
          <td style="font-size:11px;white-space:nowrap;">{p['resp']}</td>
          <td style="font-size:11px;white-space:nowrap;">{p['termino']}</td>
          <td style="text-align:right;font-size:11px;">{fmt_brl(p['previsto'])}</td>
          <td style="text-align:right;font-size:11px;color:{TEAL};">{fmt_brl(p['val_saving'])}</td>
          <td style="text-align:right;font-size:11px;color:{real_c};font-weight:600;">{real_s}</td>
          <td>{bdg_custos(p['val_custos'])}</td>
          <td style="white-space:nowrap;">{bdg_st(p['status'])}</td>
          <td style="max-width:160px;">{onde_html}</td>
          <td style="white-space:nowrap;">{data_html}</td>
        </tr>"""
    return (th("Tipo","Projeto","Responsável","Término",
               "Previsto (R$)","Saving Validado","Real Acum.","Custos","Status",
               "Onde Parado","Previsão Lib.")
            + rows + "</tbody></table>")

def pilar_resumo_html(projetos):
    """Tabela local de pilares — nome completo conforme planilha, com indicador DRE."""
    pilares = extract_pilares_local(projetos)
    if not pilares: return ""
    rows = ""
    for p in pilares:
        dot_color = GREEN if p["dre"] else SILVER
        dre_txt   = "✓ DRE" if p["dre"] else "↷ N/DRE"
        dre_style = f"color:{GREEN};font-size:9px;" if p["dre"] else f"color:{SILVER};font-size:9px;"
        real_c = GREEN if p["real"] > 0 else SILVER
        rows += f"""<tr>
          <td style="font-size:11px;max-width:120px;">
            <b>{p['nome']}</b><br>
            <span style="{dre_style}">{dre_txt}</span>
          </td>
          <td style="text-align:center;font-size:11px;font-weight:700;color:{NAVY};">{p['qtd']}</td>
          <td style="text-align:right;font-size:11px;">{fmt_mi(p['prev'])}</td>
          <td style="text-align:right;font-size:11px;color:{real_c};font-weight:600;">{fmt_mi(p['real'])}</td>
        </tr>"""
    tot_qtd  = sum(p["qtd"]  for p in pilares)
    tot_prev = sum(p["prev"] for p in pilares)
    tot_real = sum(p["real"] for p in pilares)
    rows += f"""<tr class="tr-tot">
      <td style="font-size:11px;">TOTAL</td>
      <td style="text-align:center;font-size:11px;">{tot_qtd}</td>
      <td style="text-align:right;font-size:11px;">{fmt_mi(tot_prev)}</td>
      <td style="text-align:right;font-size:11px;color:{GREEN};">{fmt_mi(tot_real)}</td>
    </tr>"""
    # Linha de total
    tot_qtd  = sum(p["qtd"]  for p in pilares)
    tot_prev = sum(p["prev"] for p in pilares)
    tot_real = sum(p["real"] for p in pilares)
    real_c_tot = GREEN if tot_real > 0 else ("#DC3545" if tot_real < 0 else SILVER)
    rows += f"""<tr class="tr-tot">
      <td style="font-size:11px;">TOTAL</td>
      <td style="text-align:center;font-size:11px;">{tot_qtd}</td>
      <td style="text-align:right;font-size:11px;">{fmt_mi(tot_prev)}</td>
      <td style="text-align:right;font-size:11px;color:{real_c_tot};font-weight:700;">{fmt_mi(tot_real)}</td>
    </tr>"""
    return (th("Pilar","Qtd","Saving (R$)","Real Acum.") + rows + "</tbody></table>")

# Cabeçalho macro-tabela
# Larguras fixas por coluna — garante alinhamento header/rows/total
MC_WIDTHS = ["16%","8%","8%","9%","9%","9%","8%","7%","6%"]

def render_macro_table(items, show_expander_fn=None):
    """
    Renderiza tabela macro completa (header + rows + total) em HTML único.
    Garante alinhamento perfeito entre colunas.
    """
    col_names = [
        "Unidade / Área",
        "Meta 2026",
        "Retorno Previsto (Anualizado)",
        f'<span style="color:{AMBER}">Previsto 2026</span>',
        f'<span style="color:{TEAL}">Retorno Validado 2026</span>',
        f'<span style="color:{GREEN}">Retorno Real 2026</span>',
        f'<span style="color:#9B59B6">Extra DRE</span>',
        "% Meta","Status"
    ]
    # Header
    ths = "".join(
        f'<th style="background:{NAVY};color:white;padding:10px 12px;'
        f'font-size:11px;font-weight:600;width:{w};text-align:left;">{c}</th>'
        for c,w in zip(col_names, MC_WIDTHS)
    )
    html = f'<table style="width:100%;border-collapse:collapse;table-layout:fixed;font-size:12px;"><thead><tr>{ths}</tr></thead><tbody>'

    # Rows
    for it in items:
        html += f"""<tr style="border-bottom:1px solid #EEF0F3;">
          <td style="padding:10px 12px;font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{it['nome']}</td>
          <td style="padding:10px 12px;">{fmt_brl(it['meta'])}</td>
          <td style="padding:10px 12px;color:#F39C12;">{fmt_brl(it.get('prev',0))}</td>
          <td style="padding:10px 12px;color:{AMBER};">{fmt_brl(it.get('prev2026',0))}</td>
          <td style="padding:10px 12px;color:{TEAL};">{fmt_brl(it['val'])}</td>
          <td style="padding:10px 12px;color:{GREEN};font-weight:600;">{fmt_brl(it['real'])}</td>
          <td style="padding:10px 12px;color:#9B59B6;">{fmt_brl(it.get('extra',0))}</td>
          <td style="padding:10px 12px;">{pbar_html(it['pct'])}</td>
          <td style="padding:10px 12px;">{bdg_status(it['pct'])}</td>
        </tr>"""

    # Total
    tm=tp=tp26=tv=tr=te=0
    for it in items:
        tm+=it["meta"];tp+=it.get("prev",0);tp26+=it.get("prev2026",0)
        tv+=it["val"];tr+=it["real"];te+=it.get("extra",0)
    pt = tr/tm if tm>0 else 0
    html += f"""<tr style="background:{LIGHT};border-top:2px solid {NAVY};font-weight:700;">
      <td style="padding:10px 12px;">TOTAL</td>
      <td style="padding:10px 12px;">{fmt_brl(tm)}</td>
      <td style="padding:10px 12px;color:#F39C12;">{fmt_brl(tp)}</td>
      <td style="padding:10px 12px;color:{AMBER};">{fmt_brl(tp26)}</td>
      <td style="padding:10px 12px;color:{TEAL};">{fmt_brl(tv)}</td>
      <td style="padding:10px 12px;color:{GREEN};">{fmt_brl(tr)}</td>
      <td style="padding:10px 12px;color:#9B59B6;">{fmt_brl(te)}</td>
      <td style="padding:10px 12px;">{pbar_html(pt)}</td>
      <td style="padding:10px 12px;"></td>
    </tr>"""
    html += "</tbody></table>"
    return html

# Compat shims — mantidos para não quebrar código legado
def mc_header(): return ""
def mc_row(it): return ""
def mc_total(items): return ""

# ── HELPERS DE SEÇÃO MINIMIZÁVEL ─────────────────────────────────────────────
def section_open(key, title, default_open=True, accent_color=None):
    """Toggle simples — botão + / − discreto."""
    sk = f"sec_{key}"
    if sk not in st.session_state:
        st.session_state[sk] = default_open
    is_open = st.session_state[sk]
    icon = "−" if is_open else "+"
    ac = accent_color or RED
    col_t, col_b = st.columns([11, 1])
    with col_t:
        st.markdown(f'<span class="st" style="border-bottom-color:{ac};">{title}</span>',
                    unsafe_allow_html=True)
    with col_b:
        if st.button(icon, key=f"btn_{key}", help="Expandir / Minimizar"):
            st.session_state[sk] = not is_open
            st.rerun()
    return st.session_state[sk]

def paired_section_open(key, title_left, title_right, default_open=True, accent_color=None):
    """Toggle único para dois painéis lado a lado."""
    sk = f"sec_{key}"
    if sk not in st.session_state:
        st.session_state[sk] = default_open
    is_open = st.session_state[sk]
    icon = "−" if is_open else "+"
    ac = accent_color or RED
    c1, c2, c3 = st.columns([5, 5, 1])
    with c1:
        st.markdown(f'<span class="st" style="border-bottom-color:{ac};">{title_left}</span>',
                    unsafe_allow_html=True)
    with c2:
        st.markdown(f'<span class="st" style="border-bottom-color:{ac};">{title_right}</span>',
                    unsafe_allow_html=True)
    with c3:
        if st.button(icon, key=f"btn_{key}", help="Expandir / Minimizar"):
            st.session_state[sk] = not is_open
            st.rerun()
    return st.session_state[sk]

# ═══════════════════════════════════════════════════════════════════════════════
# INTERFACE PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

# HEADER
st.markdown(f"""<div class="dh">
  <img src="https://grupodelga.com.br/wp-content/uploads/2024/11/logo-fa-e-clientes-grupo-whatsapp-9-300x300.png"
       onerror="this.style.display='none'">
  <div class="dh-t">
    <h1>Dashboard Executivo — Grupo Delga 2026</h1>
    <p>Gestão Estratégica de Projetos e Redução de Custos</p>
  </div>
  <div class="dh-b"><span class="lbl">Atualizado em</span>{datetime.datetime.now().strftime("%d/%m/%Y")}</div>
</div>""", unsafe_allow_html=True)

# ADMIN UPLOAD
with st.expander("🔐 Administrador — Atualizar Planilha"):
    arquivo = st.file_uploader("Nova versão (.xlsx)", type=["xlsx"], key="up")
    if arquivo:
        b = arquivo.read(); save_bytes(b)
        st.cache_data.clear()
        # Limpa session_state de cache de dados para forçar recarregamento
        for k in list(st.session_state.keys()):
            if k.startswith("_cache_"): del st.session_state[k]
        st.success("✅ Planilha atualizada! Todos os usuários verão os novos dados.")

# CARGA
fb = load_bytes()
if fb is None:
    st.warning("⚠️ Nenhuma planilha carregada. Expanda o painel Administrador para fazer o upload.")
    st.stop()
try:
    D = load_data(fb)
except Exception as e:
    st.error(f"Erro ao processar planilha: {e}"); st.stop()

# EXTRAÇÃO
kpis    = extract_kpis(D)
plantas = extract_plantas(D)
areas   = extract_areas(D)
p_glob  = extract_pilares_global(D)
ev      = extract_evolucao(D)   # já inclui acum_prev_custos / prev_custos (linhas nativas da planilha)
ranking = extract_ranking(D)

# ── TOGGLE GERAL / BSW ─────────────────────────────────────────────────────────
st.markdown(f"""<style>
div[data-testid="stRadio"] > div{{gap:6px;background:{LIGHT};padding:4px;border-radius:10px;
  display:inline-flex;width:fit-content;}}
div[data-testid="stRadio"] label{{background:transparent;border-radius:8px;padding:6px 18px!important;
  margin:0!important;font-weight:600;font-size:13px;transition:all .15s;}}
div[data-testid="stRadio"] label:has(input:checked){{background:{NAVY};}}
div[data-testid="stRadio"] label:has(input:checked) p{{color:white!important;}}
div[data-testid="stRadio"] input{{display:none;}}
</style>""", unsafe_allow_html=True)

modo_visao = st.radio("Visão:", ["🏢 Geral", "🔵 BSW"], horizontal=True,
                       key="modo_visao", label_visibility="collapsed")
is_bsw = modo_visao.endswith("BSW")

todos_projetos = get_todos_projetos(fb)
projetos_bsw   = [p for p in todos_projetos if p["tipo"] == "BSW"]

if LAYOUT_WARNINGS:
    st.warning("⚠️ **A planilha pode ter mudado de layout — confira estes pontos:**\n\n" +
               "\n".join(f"- {w}" for w in LAYOUT_WARNINGS))

# ── STATUS DE VALIDAÇÃO POR CUSTOS (Validado / Não Validado / Em Branco) ──────
projetos_status_view = projetos_bsw if is_bsw else todos_projetos
status_custos_view   = compute_status_custos(projetos_status_view)
n_total_proj    = sum(v["total"]         for v in status_custos_view.values())
n_validado      = sum(v["validado"]      for v in status_custos_view.values())
n_nao_validado  = sum(v["nao_validado"]  for v in status_custos_view.values())
n_em_branco     = sum(v["em_branco"]     for v in status_custos_view.values())
n_falta_custos  = sum(v["falta_custos"]  for v in status_custos_view.values())
n_falta_unidade = sum(v["falta_unidade"] for v in status_custos_view.values())

# ── "AGUARDANDO CUSTOS ?" (coluna nova, preenchida Sim/Não por projeto) ───────
# Pedido do usuário: contar Sim/Não dessa coluna — a soma das duas é a
# quantidade real de projetos. Quando não fecha com n_total_proj, sobraram
# projetos sem essa célula preenchida (mostramos isso explicitamente).
n_aguard_sim   = sum(1 for p in projetos_status_view if p.get("val_aguardando") == "Sim")
n_aguard_nao   = sum(1 for p in projetos_status_view if p.get("val_aguardando") == "Não")
n_aguard_vazio = sum(1 for p in projetos_status_view if p.get("val_aguardando") not in ("Sim", "Não"))

# ── Nota de Total de Projetos: só 5 Unidades + Compras ────────────────────────
# Pedido do usuário: essa nota específica não deve contar Vendas nem
# Corporativo (Vendas/Corporativo têm layout e coluna "Aguardando Custos ?"
# diferentes/ausentes, o que confundia a leitura). Restrita a essas 6 abas.
UNIDADES_NOTA = UNIDADE_SHEETS_PLANTAS + ["Compras"]
projetos_nota    = [p for p in projetos_status_view if p.get("unidade") in UNIDADES_NOTA]
status_nota_view = {k: v for k, v in status_custos_view.items() if k in UNIDADES_NOTA}
n_total_proj_nota    = sum(v["total"]        for v in status_nota_view.values())
n_validado_nota      = sum(v["validado"]     for v in status_nota_view.values())
n_nao_validado_nota  = sum(v["nao_validado"] for v in status_nota_view.values())
n_aguard_sim_nota   = sum(1 for p in projetos_nota if p.get("val_aguardando") == "Sim")
n_aguard_nao_nota   = sum(1 for p in projetos_nota if p.get("val_aguardando") == "Não")
n_aguard_vazio_nota = sum(1 for p in projetos_nota if p.get("val_aguardando") not in ("Sim", "Não"))

# "Não formalizados com Custos": Total menos (OK + Não OK + Aguardando=Sim).
# São os projetos que a própria unidade ainda nem formalizou/enviou pra
# Custos — não estão aprovados, não estão reprovados, e nem sequer marcados
# como "Sim" (aguardando) na coluna "Aguardando Custos ?".
n_nao_formalizado_nota = n_total_proj_nota - (n_validado_nota + n_nao_validado_nota + n_aguard_sim_nota)

if is_bsw:
    st.markdown(f"""<div style="background:#EDE7F9;border-left:3px solid #6C3EB5;border-radius:6px;
        padding:8px 16px;font-size:11px;color:#444;margin-bottom:16px;">
      🔵 <b>Modo BSW ativo</b> — todos os valores, gráficos e tabelas abaixo mostram <b>apenas</b>
      projetos do pilar BSW ({len(projetos_bsw)} projetos), exceto a Meta Anual do Grupo (referência fixa).
      Os painéis <b>Ranking</b> e <b>GAP</b> não são filtrados por este modo.
    </div>""", unsafe_allow_html=True)
    kpis_view = compute_kpis_bottom_up(projetos_bsw, kpis["meta"])
    ev_view   = compute_evolucao_bottom_up(projetos_bsw, ev)
    plantas_view = compute_macro_bottom_up(projetos_bsw, plantas)
    areas_view   = compute_macro_bottom_up(projetos_bsw, areas)
else:
    kpis_view    = dict(meta=kpis["meta"], portfolio=kpis["portfolio"], ret_val_ano=kpis.get("ret_val_ano",0.0),
                         prev2026=kpis["prev2026"], validado=kpis["validado"], real=kpis["real"],
                         extra_dre=kpis.get("extra_dre",0.0), pct_ating=kpis["pct_ating"], inic=kpis.get("inic",0))
    ev_view      = ev
    plantas_view = plantas
    areas_view   = areas

meta=kpis_view["meta"]; portfolio=kpis_view["portfolio"]; ret_val_ano=kpis_view["ret_val_ano"]
prev2026=kpis_view["prev2026"]
validado=kpis_view["validado"]; real=kpis_view["real"]; extra_dre=kpis_view["extra_dre"]; pct_ating=kpis_view["pct_ating"]

# ── KPI CARDS ──────────────────────────────────────────────────────────────────
cob  = portfolio/meta*100 if meta>0 else 0
cova = ret_val_ano/portfolio*100 if portfolio>0 else 0
pp   = prev2026/portfolio*100 if portfolio>0 else 0
pv   = validado/prev2026*100 if prev2026>0 else 0
pct_iniciativas_validadas = n_validado/n_total_proj*100 if n_total_proj>0 else 0

def kpi(cls,lbl,vb,sub,det):
    return (f'<div class="kpi-card {cls}"><div class="kpi-l">{lbl}</div>'
            f'<div class="kpi-v">{vb}</div><div class="kpi-s">{sub}</div>'
            f'<div class="kpi-d">{det}</div></div>')

extra_dre_sub = "Ganho fora do DRE acumulado" if not is_bsw else "BSW é 100% DRE — não se aplica"
st.markdown(f"""<div class="kpi-wrap kpi-7">
  {kpi("","Meta Anual do Grupo (2026)",fmt_mi(meta),"","Objetivo 2026 — 100%")}
  {kpi("cs","Retorno Previsto (Anual)",fmt_mi(portfolio),"",f"{cob:.1f}% da meta coberta")}
  {kpi("ct","Retorno Validado (Anual)",fmt_mi(ret_val_ano),"",f"{cova:.1f}% do Retorno Previsto")}
  {kpi("ca","Previsto 2026",fmt_mi(prev2026),"",f"{pp:.1f}% do Retorno Previsto")}
  {kpi("","Validado por Custos (2026)",fmt_mi(validado),
       f"{pv:.1f}% do Previsto 2026",
       f"{pct_iniciativas_validadas:.1f}% das iniciativas aprovadas")}
  {kpi("cg","Retorno Real (DRE) (2026)",fmt_mi(real),"",f"{pct_ating*100:.1f}% de atingimento")}
  {kpi("cr","Extra DRE (Até o Momento)",fmt_mi(extra_dre),"",extra_dre_sub)}
</div>""", unsafe_allow_html=True)

_aguard_gap_nota = f' <span style="color:{RED};">({n_aguard_vazio_nota} projeto(s) sem essa célula preenchida)</span>' if n_aguard_vazio_nota else ""
st.markdown(f"""<div class="nota" style="display:flex;flex-direction:column;gap:4px;">
  <div style="display:flex;gap:28px;align-items:center;flex-wrap:wrap;">
    <span><b>Total de Projetos:</b> {n_total_proj_nota}</span>
    <span><b style="color:{GREEN};">Custos OK:</b> {n_validado_nota}</span>
    <span><b style="color:{RED};">Custos Não OK:</b> {n_nao_validado_nota}</span>
    <span><b style="color:{AMBER};">Aguardando Custos:</b> {n_aguard_sim_nota}</span>
    <span><b style="color:{NAVY};">Não Formalizados com Custos:</b> {n_nao_formalizado_nota}</span>
    <span style="color:{SILVER};font-size:10px;">(coluna "Aguardando Custos ?" — Sim: {n_aguard_sim_nota} · Não: {n_aguard_nao_nota}{_aguard_gap_nota})</span>
  </div>
  <div style="color:{SILVER};font-size:10px;">Considerando as 5 Unidades (Diadema, Ferraz, São Leopoldo, Jarinu, Anchieta) + Compras — não inclui Vendas nem Corporativo.</div>
</div>""", unsafe_allow_html=True)

st.markdown(f"""<div class="nota">
  <b>Metodologia:</b>&nbsp;
  <b style="color:{GREEN};">✓ DRE</b>: BSW · Kaizen · Kaizen GR · Redução de Custo · Você Resolve — impacto direto e mensurável no DRE.&nbsp;
  <b style="color:{SILVER};">↷ Não DRE</b>: Kaizen Custo Evitado · Kaizen Capital de Giro · Meta Executiva — geram valor operacional mas não reduzem GGF no DRE.
</div>""", unsafe_allow_html=True)

# ── EVOLUÇÃO ───────────────────────────────────────────────────────────────────
st.markdown('<div class="sc">', unsafe_allow_html=True)
is_ev = section_open("evolucao", "Evolução Mensal — Acumulado Previsto vs Real vs Meta")
if is_ev:
    series_all = ["Acumulado Previsto","Previsto por Custos (2026)","Acumulado Real",
                  "Projeção da Meta","Previsto Mensal","Real Mensal"]
    sel = st.multiselect("Séries:", series_all,
                         default=["Acumulado Previsto","Previsto por Custos (2026)","Acumulado Real",
                                  "Projeção da Meta","Previsto Mensal","Real Mensal"],
                         key="ev_sel")
    st.markdown(f'<p style="font-size:10px;color:{SILVER};margin:-6px 0 8px;">'
                f'<b>Previsto por Custos (2026)</b>: valor validado pelo departamento de Custos, '
                f'distribuído mês a mês diretamente na planilha (linha "Acumulado prev.Custos").</p>',
                unsafe_allow_html=True)
    if sel:
        st.plotly_chart(chart_evolucao(ev_view,sel), use_container_width=True, config={"displayModeBar":False})
st.markdown('</div>', unsafe_allow_html=True)

# ── FUNIL + GAUGE — botão único para o par ─────────────────────────────────────
is_fg = paired_section_open("funil_gauge",
                             "Funil de Conversão — Portfólio → DRE",
                             "Atingimento da Meta")
cfu, cga = st.columns([3, 2])
with cfu:
    st.markdown('<div class="sc" style="min-height:60px;">', unsafe_allow_html=True)
    if is_fg:
        st.markdown(f'<p style="font-size:11px;color:{SILVER};margin-bottom:8px;">Quanto do portfólio mapeado converte em resultado no DRE?</p>', unsafe_allow_html=True)
        st.plotly_chart(chart_funnel(kpis_view), use_container_width=True, config={"displayModeBar":False})
    st.markdown('</div>', unsafe_allow_html=True)
with cga:
    st.markdown('<div class="sc" style="min-height:60px;">', unsafe_allow_html=True)
    if is_fg:
        st.plotly_chart(chart_gauge(pct_ating), use_container_width=True, config={"displayModeBar":False})
        gap_val = meta - real
        st.markdown(f"""<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:4px;">
          <div style="background:{LIGHT};border-radius:8px;padding:12px;text-align:center;">
            <div style="font-size:9px;font-weight:600;color:{SILVER};text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px;">GAP para Meta</div>
            <div style="font-size:16px;font-weight:700;color:{RED};">{fmt_mi(gap_val)}</div>
          </div>
          <div style="background:{LIGHT};border-radius:8px;padding:12px;text-align:center;">
            <div style="font-size:9px;font-weight:600;color:{SILVER};text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px;">Validado / Meta</div>
            <div style="font-size:16px;font-weight:700;color:{NAVY};">{validado/meta*100:.1f}%</div>
          </div>
        </div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ── DONUTS — botão único para o par ────────────────────────────────────────────
is_dn = paired_section_open("donuts",
                             "Representatividade — Plantas",
                             "Representatividade — Áreas Funcionais")
cd1, cd2 = st.columns(2)
with cd1:
    st.markdown('<div class="sc" style="min-height:60px;">', unsafe_allow_html=True)
    if is_dn:
        st.plotly_chart(chart_donut([p["nome"] for p in plantas],[p["meta"] for p in plantas],PAL),
                        use_container_width=True, config={"displayModeBar":False})
    st.markdown('</div>', unsafe_allow_html=True)
with cd2:
    st.markdown('<div class="sc" style="min-height:60px;">', unsafe_allow_html=True)
    if is_dn:
        st.plotly_chart(chart_donut([a["nome"] for a in areas],[a["meta"] for a in areas],
                                    [NAVY,GREEN,"#20C997"]),
                        use_container_width=True, config={"displayModeBar":False})
    st.markdown('</div>', unsafe_allow_html=True)

# ── PILARES ────────────────────────────────────────────────────────────────────
st.markdown('<div class="sc">', unsafe_allow_html=True)
is_pil = section_open("pilares", "Distribuição por Tipo de Iniciativa — Grupo")

# Gráfico gerencial com toggles
def chart_pilares_gerencial(pilares_global, real_total, show_prev, show_val, show_real):
    labels   = [p["nome"] for p in pilares_global]
    previsto = [p["prev"] for p in pilares_global]
    validado_l = [p["val"] for p in pilares_global]
    # Usa 'real' direto se disponível (lido da tabela 5U), senão estima
    real_est = [p.get("real", 0) for p in pilares_global]
    if all(v == 0 for v in real_est):
        tv = sum(validado_l)
        real_est = [v/tv*real_total if tv>0 else 0 for v in validado_l]
    series = []
    if show_prev: series.append(dict(name="Previsto",  x=previsto,    color="#C8D8EE"))
    if show_val:  series.append(dict(name="Validado",  x=validado_l,  color=NAVY))
    if show_real: series.append(dict(name="Real DRE",  x=real_est,    color=GREEN))
    if not series: return None
    fig = go.Figure()
    for s in series:
        fig.add_trace(go.Bar(
            name=s["name"], y=labels, x=s["x"], orientation="h",
            marker=dict(color=s["color"], line=dict(width=0)),
            text=[fmt_mi(v) for v in s["x"]],
            textposition="outside", textfont=dict(size=10),
            hovertemplate=f"<b>%{{y}}</b><br>{s['name']}: R$ %{{x:,.0f}}<extra></extra>",
        ))
    fig.update_layout(
        barmode="group",
        xaxis=dict(tickformat=",.0f", showgrid=True, gridcolor="#F0F4F8",
                   tickprefix="R$ ", zeroline=False),
        yaxis=dict(autorange="reversed", tickfont=dict(size=12, color="#333"),
                   gridcolor="#F0F4F8"),
        legend=dict(orientation="h", y=1.06, x=0.5, xanchor="center",
                    font=dict(size=12), bgcolor="rgba(0,0,0,0)"),
        margin=dict(l=160, r=100, t=44, b=20),
        height=max(220, len(labels)*62),
        paper_bgcolor="white", plot_bgcolor="white",
        bargap=0.35, bargroupgap=0.06,
        font=dict(family="Inter"),
    )
    return fig

# build_pilares_grupo — agrupa projetos reais com subtipos Kaizen
@st.cache_data(show_spinner=False)
def build_pilares_grupo(fb_key):
    """
    Lê DIRETAMENTE da tabela 'Saving Especulado por Pilar' da aba 5 Unidades
    (rows 12-19, cols 3-7) — fonte única de verdade, mesmos valores do Excel.
    col3=Pilar, col4=Qtd, col5=Saving(Previsto), col6=Saving Validado, col7=Até o Momento(Real)
    """
    NAO_DRE_PIL = {"Kaizen - Custo Evitado","Kaizen - Capital de Giro",
                   "Meta Executiva","Meta Executiva "}
    df = D["u5"]
    res = []
    for ri in range(12, 21):  # rows 12-20 (20 = TOTAL, pular)
        nome = str(df.iloc[ri,3]).strip() if pd.notna(df.iloc[ri,3]) else ""
        if not nome or nome in ("TOTAL",""):
            continue
        qtd  = int(safe(df.iloc[ri,4]))
        prev = safe(df.iloc[ri,5])
        val  = safe(df.iloc[ri,6])
        real = safe(df.iloc[ri,7])
        dre  = nome not in NAO_DRE_PIL
        res.append(dict(nome=nome.strip(), qtd=qtd, prev=prev,
                        val=val, real=real, dre=dre))
    return res

if is_pil:
    cp1, cp2 = st.columns([5, 4])
    with cp1:
        _t1, _t2, _t3, _tsp = st.columns([2, 2, 2, 3])
        with _t1: show_prev = st.toggle("Previsto",  value=True,  key="tog_prev")
        with _t2: show_val  = st.toggle("Validado",  value=True,  key="tog_val")
        with _t3: show_real = st.toggle("Real DRE",  value=False, key="tog_real")
        p_grupo = build_pilares_grupo(hash(fb))
        if is_bsw:
            p_grupo = [p for p in p_grupo if p["nome"] == "BSW"]
        fig_pil = chart_pilares_gerencial(p_grupo, real, show_prev, show_val, show_real)
        if fig_pil:
            st.plotly_chart(fig_pil, use_container_width=True, config={"displayModeBar":False})
        else:
            st.info("Selecione ao menos uma série.")
    with cp2:
        st.markdown(f'<p class="st" style="border-bottom-color:{RED};">Resumo por Pilar — Grupo</p>',
                    unsafe_allow_html=True)
        rows_p = ""
        for p in p_grupo:
            dre_s = f"color:{GREEN};font-size:9px;font-weight:600;" if p["dre"] else f"color:{SILVER};font-size:9px;"
            dre_t = "✓ DRE" if p["dre"] else "↷ N/DRE"
            rows_p += f"""<tr>
              <td style="font-size:11px;font-weight:600;">{p['nome']}<br>
                <span style="{dre_s}">{dre_t}</span></td>
              <td style="text-align:center;font-size:11px;font-weight:700;">{p['qtd']}</td>
              <td style="text-align:right;font-size:11px;">{fmt_mi(p['prev'])}</td>
              <td style="text-align:right;font-size:11px;color:{TEAL};font-weight:600;">{fmt_mi(p['val'])}</td>
              <td style="text-align:right;font-size:11px;color:{GREEN};font-weight:600;">{fmt_mi(p['real'])}</td>
            </tr>"""
        tot_qtd_g=sum(p["qtd"] for p in p_grupo)
        tot_prev_g=sum(p["prev"] for p in p_grupo)
        tot_val_g=sum(p["val"] for p in p_grupo)
        tot_real_g=sum(p["real"] for p in p_grupo)
        rows_p += f"""<tr class="tr-tot">
          <td style="font-size:11px;">TOTAL</td>
          <td style="text-align:center;font-size:11px;">{tot_qtd_g}</td>
          <td style="text-align:right;font-size:11px;">{fmt_mi(tot_prev_g)}</td>
          <td style="text-align:right;font-size:11px;color:{TEAL};">{fmt_mi(tot_val_g)}</td>
          <td style="text-align:right;font-size:11px;color:{GREEN};">{fmt_mi(tot_real_g)}</td>
        </tr>"""
        st.markdown(th("Pilar","Qtd","Saving (R$)","Saving Validado","Até o Momento")+rows_p+"</tbody></table>",
                    unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PLANTAS INDUSTRIAIS
# ═══════════════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="sc">', unsafe_allow_html=True)
is_plantas = section_open("plantas", "Plantas Industriais — Performance Consolidada")
if is_plantas:
    st.markdown(render_macro_table(plantas_view), unsafe_allow_html=True)

for p in plantas_view:
    if not is_plantas:
        break
    with st.expander(f"＋  Ver projetos de {p['nome']}", expanded=False):
        proj = get_proj_planta(D, p["sheet"])
        n = len(proj)
        if is_bsw:
            proj = [x for x in proj if x["tipo"] == "BSW"]
        if proj:
            proj_v, pilar_html = projetos_por_pilar_html(proj, key_prefix=f"plt_{p['nome']}")
            st.markdown(f"<p style='font-size:11px;color:{SILVER};margin:4px 0 8px;'>"
                        f"<b>{len(proj_v)}</b> de {n} projetos</p>", unsafe_allow_html=True)
            st.markdown(pilar_html, unsafe_allow_html=True)
        else:
            st.markdown("<p style='color:#999;font-size:12px;'>Sem projetos.</p>",
                        unsafe_allow_html=True)


st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ÁREAS FUNCIONAIS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sc">', unsafe_allow_html=True)
is_areas = section_open("areas", "Áreas Funcionais — Performance Consolidada")
area_fn = {"Compras": get_proj_compras, "Vendas": get_proj_vendas}
if is_areas:
    st.markdown(render_macro_table(areas_view), unsafe_allow_html=True)

for a in areas_view:
    if not is_areas:
        break
    with st.expander(f"＋  Ver projetos de {a['nome']}", expanded=False):
        fn = area_fn.get(a["nome"])
        proj = fn(D) if fn else []
        n = len(proj)
        if is_bsw:
            proj = [x for x in proj if x["tipo"] == "BSW"]
        if proj:
            proj_va, pilar_html_a = projetos_por_pilar_html(proj, key_prefix=f"area_{a['nome']}")
            st.markdown(f"<p style='font-size:11px;color:{SILVER};margin:4px 0 8px;'>"
                        f"<b>{len(proj_va)}</b> de {n} projetos</p>", unsafe_allow_html=True)
            st.markdown(pilar_html_a, unsafe_allow_html=True)
        else:
            st.markdown("<p style='color:#999;font-size:12px;'>Sem projetos.</p>",
                        unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ── CLASSIFICAÇÃO DE GANHOS ────────────────────────────────────────────────────
st.markdown('<div class="sc">', unsafe_allow_html=True)
is_class = section_open("classificacao","Classificação de Ganhos — Impacto no DRE",default_open=False)
if is_class:
    cc1,cc2,cc3,cc4,cc5 = st.columns(5)
    ganhos = [
        (cc1,NAVY,   "🔵","BSW","Benchmark de peso bruto. Redução de MP — impacto direto no DRE.","✓ DRE",GREEN),
        (cc2,GREEN,  "🔥","Redução de Custo","Elimina custo direto na operação. Reduz GGF no DRE.","✓ DRE",GREEN),
        (cc3,AMBER,  "⚡","Kaizen / GR","Produtividade recorrente apurada. Entra no DRE quando validado.","✓ DRE",GREEN),
        (cc4,"#512DA8","↷","C. Evitado","MO realocada internamente — não reduz GGF no DRE.","↷ Não DRE",SILVER),
        (cc5,"#0D47A1","🏦","Cap. de Giro","Reduz estoque / melhora caixa. Impacto no balanço, não no DRE.","↷ Não DRE",SILVER),
    ]
    for col,cor,icon,titulo,texto,dre,dcor in ganhos:
        with col:
            st.markdown(f"""<div style="border:2px solid {cor};border-radius:8px;padding:12px 14px;height:100%;">
              <div style="font-weight:700;color:{cor};margin-bottom:4px;font-size:12px;">{icon} {titulo}</div>
              <div style="font-size:10px;color:#444;line-height:1.5;margin-bottom:6px;">{texto}</div>
              <div style="font-size:10px;font-weight:700;color:{dcor};">{dre}</div>
            </div>""", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ── RANKING ────────────────────────────────────────────────────────────────────
st.markdown('<div class="sc">', unsafe_allow_html=True)
is_rank = section_open("ranking","Ranking de Projetos — Todos os Pilares",default_open=False)
if is_rank:
    rk1,rk2,rk3 = st.columns([2,2,1])
    with rk1:
        f_uni = st.multiselect("Unidade:", sorted({r["uni"] for r in ranking}),
                               default=[], placeholder="Todas", key="rk_uni")
    with rk2:
        f_st  = st.multiselect("Status:", sorted({r["status"] for r in ranking if r["status"]}),
                               default=[], placeholder="Todos", key="rk_st")
    with rk3:
        n_lin = st.number_input("Linhas:", 5, 200, 25, 5)
    pf = ranking
    if f_uni: pf = [r for r in pf if r["uni"] in f_uni]
    if f_st:  pf  = [r for r in pf if r["status"] in f_st]
    st.markdown(f"<p style='font-size:11px;color:{SILVER};margin-bottom:6px;'>"
                f"Exibindo {min(int(n_lin),len(pf))} de {len(pf)} projetos</p>",
                unsafe_allow_html=True)
    rows_rk = "".join(f"""<tr>
      <td style="text-align:center;color:{SILVER};font-weight:700;font-size:11px;">{r['pos']}</td>
      <td style="font-weight:600;font-size:11px;">{r['uni']}</td>
      <td style="font-size:11px;">{r['nome']}</td>
      <td>{bdg_st(r['status'])}</td>
      <td>{bdg_custos(r['custos'])}</td>
      <td style="text-align:right;font-size:11px;">{fmt_brl(r['prev26'])}</td>
      <td style="text-align:right;font-size:11px;">{fmt_brl(r['prev_mo'])}</td>
      <td style="text-align:right;font-weight:700;color:{GREEN};font-size:11px;">{fmt_brl(r['real'])}</td>
    </tr>""" for r in pf[:int(n_lin)])
    st.markdown(th("#","Unidade","Projeto","Status","Custos",
                   "Previsto 2026","Previsto Momento","Real DRE")+rows_rk+"</tbody></table>",
                unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ── GAP ────────────────────────────────────────────────────────────────────────
st.markdown('<div class="sc">', unsafe_allow_html=True)
is_gap = section_open("gap","GAP — Projetos Aguardando Validação de Custos",
                      default_open=False, accent_color=AMBER)
if is_gap:
    st.markdown(f'<p style="font-size:11px;color:{SILVER};margin-bottom:10px;">'
                f'Projetos com valor projetado mas ainda sem validação do depto de Custos.</p>',
                unsafe_allow_html=True)
    gap = [r for r in ranking if r["custos"] not in ("OK","Não Ok","NOK","Não OK") and r["prev26"]>0]
    by_uni = {}
    for r in gap: by_uni[r["uni"]] = by_uni.get(r["uni"],0)+r["prev26"]
    tot_gap = sum(by_uni.values()) if by_uni else 0
    rows_gap = "".join(f"""<tr>
      <td style="font-weight:600;">{u}</td>
      <td style="text-align:right;color:{AMBER};font-weight:600;">{fmt_brl(v)}</td>
      <td style="text-align:right;">{v/tot_gap*100:.1f}%</td>
    </tr>""" for u,v in sorted(by_uni.items(),key=lambda x:-x[1]))
    rows_gap += f"""<tr class="tr-tot">
      <td>TOTAL GAP</td>
      <td style="text-align:right;color:{AMBER};">{fmt_brl(tot_gap)}</td>
      <td style="text-align:right;">100%</td>
    </tr>"""
    st.markdown(f"<p style='font-size:11px;color:{SILVER};'>{len(gap)} projetos aguardam validação</p>",
                unsafe_allow_html=True)
    st.markdown(th("Unidade",f'<span style="color:{AMBER}">Previsto 2026 (não validado)</span>',
                   "% do Gap")+rows_gap+"</tbody></table>",
                unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ── VALIDAÇÃO DE CUSTOS POR UNIDADE/DEPARTAMENTO ───────────────────────────────
st.markdown('<div class="sc">', unsafe_allow_html=True)
is_valcust = section_open("valcust", "Validação de Custos por Unidade/Departamento",
                          default_open=False, accent_color=AMBER)
if is_valcust:
    st.markdown(f"""<div class="nota" style="margin-top:0;">
      <b>Como ler:</b>&nbsp; <b style="color:{GREEN};">Validado</b> = Custos marcou "OK".&nbsp;
      <b style="color:{RED};">Não Validado</b> = Custos marcou "Não Ok"/"NOK".&nbsp;
      <b style="color:{SILVER};">Em Branco</b> = Custos ainda não se posicionou. Dentro de "Em Branco", cada
      projeto tem um checklist de 3 etapas (A3/Estrutura desenvolvido, Memória de Cálculo desenvolvido,
      Formalizado com Dep. de Custos):&nbsp;
      <b style="color:{AMBER};">⏳ Falta Custos aprovar</b> = as 3 etapas estão marcadas e o projeto entra no DRE
      (a unidade já fez a parte dela, falta só Custos validar).&nbsp;
      <b style="color:{NAVY};">📋 Falta unidade enviar</b> = ainda falta pelo menos 1 das 3 etapas.
    </div>""", unsafe_allow_html=True)

    ORDEM_UNIDADES = UNIDADE_SHEETS_PLANTAS + ["Compras", "Vendas", "Corporativo"]
    itens_vc = [u for u in ORDEM_UNIDADES if u in status_custos_view] + \
               [u for u in status_custos_view if u not in ORDEM_UNIDADES]

    rows_vc = ""
    for u in itens_vc:
        s = status_custos_view[u]
        pct_v = s["validado"] / s["total"] if s["total"] > 0 else 0.0
        rows_vc += f"""<tr style="border-bottom:1px solid #EEF0F3;">
          <td style="padding:10px 12px;font-weight:600;font-size:12px;">{u}</td>
          <td style="padding:10px 12px;text-align:center;font-size:12px;">{s['total']}</td>
          <td style="padding:10px 12px;text-align:center;font-size:12px;color:{GREEN};font-weight:700;">{s['validado']}</td>
          <td style="padding:10px 12px;text-align:center;font-size:12px;color:{RED};font-weight:700;">{s['nao_validado']}</td>
          <td style="padding:10px 12px;text-align:center;font-size:12px;color:{SILVER};font-weight:700;">{s['em_branco']}</td>
          <td style="padding:10px 12px;text-align:center;font-size:12px;color:{AMBER};font-weight:700;">{s['falta_custos']}</td>
          <td style="padding:10px 12px;text-align:center;font-size:12px;color:{NAVY};font-weight:700;">{s['falta_unidade']}</td>
          <td style="padding:10px 12px;">{pbar_html(pct_v)}</td>
        </tr>"""

    tot_total = sum(s["total"] for s in status_custos_view.values())
    tot_val   = sum(s["validado"] for s in status_custos_view.values())
    tot_nval  = sum(s["nao_validado"] for s in status_custos_view.values())
    tot_branco= sum(s["em_branco"] for s in status_custos_view.values())
    tot_fc    = sum(s["falta_custos"] for s in status_custos_view.values())
    tot_fu    = sum(s["falta_unidade"] for s in status_custos_view.values())
    pct_tot   = tot_val / tot_total if tot_total > 0 else 0.0
    rows_vc += f"""<tr class="tr-tot">
      <td style="padding:10px 12px;font-size:12px;">TOTAL</td>
      <td style="padding:10px 12px;text-align:center;font-size:12px;">{tot_total}</td>
      <td style="padding:10px 12px;text-align:center;font-size:12px;color:{GREEN};">{tot_val}</td>
      <td style="padding:10px 12px;text-align:center;font-size:12px;color:{RED};">{tot_nval}</td>
      <td style="padding:10px 12px;text-align:center;font-size:12px;color:{SILVER};">{tot_branco}</td>
      <td style="padding:10px 12px;text-align:center;font-size:12px;color:{AMBER};">{tot_fc}</td>
      <td style="padding:10px 12px;text-align:center;font-size:12px;color:{NAVY};">{tot_fu}</td>
      <td style="padding:10px 12px;">{pbar_html(pct_tot)}</td>
    </tr>"""

    header_vc = "".join(
        f'<th style="background:{NAVY};color:white;padding:10px 12px;font-size:11px;font-weight:600;text-align:left;">{c}</th>'
        for c in ["Unidade / Departamento", "Total", "✓ Validado", "✗ Não Validado", "Em Branco",
                  "⏳ Falta Custos Aprovar", "📋 Falta Unidade Enviar", "% Validado"]
    )
    st.markdown(
        f'<table style="width:100%;border-collapse:collapse;">'
        f'<thead><tr>{header_vc}</tr></thead><tbody>{rows_vc}</tbody></table>',
        unsafe_allow_html=True
    )
    st.markdown(f"<p style='font-size:10px;color:{SILVER};margin-top:8px;'>"
                f"{n_em_branco} projetos ainda sem posição de Custos — {n_falta_custos} já prontos "
                f"aguardando aprovação e {n_falta_unidade} ainda precisam ser finalizados pela unidade. "
                f"Corporativo não tem extração de projeto a projeto na planilha atual, por isso não aparece aqui.</p>",
                unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ── FOOTER ─────────────────────────────────────────────────────────────────────
st.markdown(f"""<div style="text-align:center;padding:16px 0;border-top:1px solid #EEF0F3;margin-top:8px;">
  <span style="font-size:11px;color:{SILVER};">
    Dashboard Executivo · Grupo Delga 2026 · Gestão Estratégica de Projetos e Redução de Custos
  </span>
</div>""", unsafe_allow_html=True)
_,cft = st.columns([5,1])
with cft:
    if st.button("🚪 Sair", key="logout"):
        st.session_state["auth"]=False; st.rerun()
