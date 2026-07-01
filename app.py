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
    'Redução de custo', 'Redução de Custo', 'Redução de Custo ',
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
.dh-b{{margin-left:auto;background:{RED};color:white;font-size:11px;
       font-weight:600;padding:5px 14px;border-radius:20px;white-space:nowrap;
       letter-spacing:.3px;}}

/* ── KPI CARDS ── */
.kpi-wrap{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:20px;}}
.kpi-6{{grid-template-columns:repeat(6,1fr);}}
.kpi-card{{background:white;border-radius:12px;padding:18px 20px;
           border-left:4px solid {NAVY};
           box-shadow:0 1px 4px rgba(28,43,74,.06),0 4px 16px rgba(28,43,74,.04);
           transition:box-shadow .2s;}}
.kpi-card:hover{{box-shadow:0 2px 8px rgba(28,43,74,.1),0 8px 24px rgba(28,43,74,.06);}}
.kpi-card.cr{{border-left-color:{RED};}}
.kpi-card.cg{{border-left-color:{GREEN};}}
.kpi-card.ca{{border-left-color:{AMBER};}}
.kpi-card.cs{{border-left-color:{SILVER};}}
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
    if pd.isna(v) or v is None or str(v)=="nan": return "—"
    try:
        if isinstance(v,(datetime.datetime,datetime.date)): return v.strftime("%m/%Y")
        s=str(v)
        if " " in s: return pd.to_datetime(s).strftime("%m/%Y")
        return s[:7]
    except: return str(v)[:7]

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
        meta     =safe(df.iloc[6,3]),
        portfolio=safe(df.iloc[6,5]),
        prev2026 =safe(df.iloc[6,6]),
        validado =safe(df.iloc[6,7]),
        real     =safe(df.iloc[6,9]),
        pct_ating=safe(df.iloc[6,11]),
        inic     =int(safe(df.iloc[6,13])),
    )

def extract_plantas(d):
    df = d["u5"]
    cfg = [("Diadema",4,"Diadema"),("Ferraz",5,"Ferraz"),
           ("São Leopoldo",6,"São Leopoldo"),("Jarinu",7,"Jarinu"),("Anchieta",8,"Anchieta")]
    res=[]
    for nome,col,sh in cfg:
        df_p = d.get(sh, pd.DataFrame())
        p2026 = safe(df_p.iloc[4,3]) if not df_p.empty else 0
        res.append(dict(nome=nome,sheet=sh,
            meta=safe(df.iloc[22,col]),prev=safe(df.iloc[23,col]),
            prev2026=p2026,val=safe(df.iloc[25,col]),
            real=safe(df.iloc[26,col]),pct=safe(df.iloc[27,col])))
    return res

def extract_areas(d):
    df = d["u5"]
    cfg = [("Corporativo",9,"Corporativo"),("Compras",10,"Compras "),("Vendas",11,"Vendas")]
    res=[]
    for nome,col,sh in cfg:
        df_a = d.get(sh, pd.DataFrame())
        p2026 = safe(df_a.iloc[4,4]) if (not df_a.empty and sh!="Corporativo") else 0
        res.append(dict(nome=nome,sheet=sh,
            meta=safe(df.iloc[22,col]),prev=safe(df.iloc[23,col]),
            prev2026=p2026,val=safe(df.iloc[25,col]),
            real=safe(df.iloc[26,col]),pct=safe(df.iloc[27,col])))
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
    # v12: col21=label, cols 22-33=monthly, rows deslocados +1
    # row54=Previsto, row55=Real, row57=Acum.Prev, row58=Acum.Real, row59=Projeção Meta
    max_col = min(34, df.shape[1])
    def row(r): return [safe(df.iloc[r,c]) for c in range(22, max_col)]
    # Pad to 12 if needed
    def pad(lst): return (lst + [0]*12)[:12]
    return dict(meses=meses,
                prev     =pad(row(54)),
                real     =pad(row(55)),
                acum_prev=pad(row(57)),
                acum_real=pad(row(58)),
                proj_meta=pad(row(59)))

# ── EXTRAÇÃO DE PROJETOS — FUNÇÃO CENTRAL ─────────────────────────────────────
def extract_projetos(df, start_row, col_tipo=0, col_nome=2, col_resp=5,
                     col_termino=7, col_custos=12, col_saving=13,
                     col_status=14, col_onde=15, col_data_lib=16,
                     col_prev_real=18,  # col com 'Previsto'/'Real'
                     col_total_ano=36): # col com Total Ano
    """
    Extrai projetos de uma aba usando a lógica:
      - Linha Previsto: col_tipo in VALID_TIPOS + col_nome preenchido + col_prev_real='Previsto'
      - Linha Real: row+1, col_prev_real='Real', col_total_ano = valor real acumulado
    Para antes de qualquer tipo não reconhecido (seção SPD, etc.)
    """
    res = []
    i = start_row
    max_row = min(start_row + 600, df.shape[0] - 1)
    while i <= max_row:
        tipo = str(df.iloc[i, col_tipo]).strip()
        nome = str(df.iloc[i, col_nome]).strip()
        c_pr = str(df.iloc[i, col_prev_real]).strip() if df.shape[1] > col_prev_real else ""

        if tipo in VALID_TIPOS and nome not in ("", "nan") and c_pr == "Previsto":
            # Linha Previsto — Total Ano
            tot_prev = safe(df.iloc[i, col_total_ano])

            # Linha Real (row+1)
            tot_real = 0.0
            impede = ""
            if i+1 <= max_row:
                c_pr_next = str(df.iloc[i+1, col_prev_real]).strip() if df.shape[1] > col_prev_real else ""
                if c_pr_next == "Real":
                    tot_real = safe(df.iloc[i+1, col_total_ano])

            # Onde está parado (col15) e Data de liberação (col16)
            onde_parado = ""
            data_lib    = ""
            if col_onde is not None and df.shape[1] > col_onde:
                v_onde = df.iloc[i, col_onde]
                if pd.notna(v_onde) and str(v_onde).strip() not in ("", "nan"):
                    onde_parado = str(v_onde).strip()
            if col_data_lib is not None and df.shape[1] > col_data_lib:
                v_dl = df.iloc[i, col_data_lib]
                if pd.notna(v_dl) and str(v_dl).strip() not in ("", "nan"):
                    data_lib = fmt_date(v_dl)

            res.append(dict(
                tipo       = tipo,
                nome       = nome,
                resp       = str(df.iloc[i, col_resp]).strip() if pd.notna(df.iloc[i, col_resp]) else "—",
                termino    = fmt_date(df.iloc[i, col_termino]),
                previsto   = tot_prev,
                real_ano   = tot_real,
                val_custos = str(df.iloc[i, col_custos]).strip() if pd.notna(df.iloc[i, col_custos]) else "",
                val_saving = safe(df.iloc[i, col_saving]),
                status     = str(df.iloc[i, col_status]).strip() if pd.notna(df.iloc[i, col_status]) else "",
                onde_parado= onde_parado,
                data_lib   = data_lib,
                entra_dre  = is_dre(tipo),
            ))
            i += 2  # pula Previsto + Real
        elif tipo not in ("", "nan") and tipo not in VALID_TIPOS:
            break   # seção SPD ou outro bloco diferente — para
        else:
            i += 1
    return res

def get_proj_planta(d, sheet_key):
    df = d.get(sheet_key)
    if df is None: return []
    # Plantas v9: col0=tipo, col2=nome, col5=resp, col7=term,
    #             col12=custos, col13=saving, col14=status,
    #             col15=onde_parado, col16=data_lib,
    #             col18=Previsto/Real, col36=Total Ano
    return extract_projetos(df, start_row=54,
        col_tipo=0, col_nome=2, col_resp=5, col_termino=7,
        col_custos=12, col_saving=13, col_status=14,
        col_onde=15, col_data_lib=16,
        col_prev_real=18, col_total_ano=36)

def get_proj_compras(d):
    df = d.get("Compras ")
    if df is None: return []
    # Compras v9: col0=tipo,col3=nome,col5=resp,col7=term,col12=custos,col13=saving,
    #             col14=status,col15=onde,col16=data_lib,col19=Prev/Real,col36=TotalAno
    return extract_projetos(df, start_row=30,
        col_tipo=0, col_nome=3, col_resp=5, col_termino=7,
        col_custos=12, col_saving=13, col_status=14,
        col_onde=15, col_data_lib=16,
        col_prev_real=19, col_total_ano=37)

def get_proj_vendas(d):
    df = d.get("Vendas")
    if df is None: return []
    # Vendas v9: col0=tipo,col1=nome,col4=resp,col6=term,col11=custos,col12=saving,
    #            col13=status,col17=Prev/Real,col35=TotalAno
    # start_row=32 (projetos começam na row32, não 36)
    return extract_projetos(df, start_row=32,
        col_tipo=0, col_nome=1, col_resp=4, col_termino=6,
        col_custos=11, col_saving=12, col_status=13,
        col_onde=None, col_data_lib=None,
        col_prev_real=17, col_total_ano=35)

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
    cfg = {
        "Acumulado Previsto":dict(data=ev["acum_prev"],color=NAVY, dash="solid"),
        "Acumulado Real":    dict(data=ev["acum_real"],color=GREEN,dash="solid"),
        "Projeção da Meta":  dict(data=ev["proj_meta"],color=RED,  dash="dash"),
        "Previsto Mensal":   dict(data=ev["prev"],     color="#4A90D9",dash="dot"),
        "Real Mensal":       dict(data=ev["real"],     color="#28A745",dash="dot"),
    }
    fig = go.Figure()
    for s in series:
        c = cfg[s]
        fig.add_trace(go.Scatter(x=ev["meses"],y=c["data"],mode="lines+markers",name=s,
            line=dict(color=c["color"],width=2.5,dash=c["dash"]),
            marker=dict(size=6,color=c["color"]),
            hovertemplate=f"<b>{s}</b><br>%{{x}}: R$ %{{y:,.0f}}<extra></extra>"))
    fig.update_layout(
        xaxis=dict(showgrid=True,gridcolor="#F0F4F8"),
        yaxis=dict(tickformat=",.0f",showgrid=True,gridcolor="#F0F4F8",title="R$"),
        legend=dict(orientation="h",y=1.05,x=0.5,xanchor="center",font=dict(size=11)),
        margin=dict(l=80,r=20,t=50,b=40),height=400,
        paper_bgcolor="white",plot_bgcolor="white",
        hovermode="x unified",font=dict(family="Inter"))
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
               "V.Previsto","V.Validado","V.Real","Custos","Status",
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
    return (th("Pilar","Qtd","V. Previsto","V. Real (Acum.)") + rows + "</tbody></table>")

# Cabeçalho macro-tabela
MC_COLS = ["Unidade / Área","Meta 2026","Previsto Total",
           f'<span style="color:{AMBER}">Previsto 2026</span>',
           f'<span style="color:{TEAL}">Validado</span>',
           f'<span style="color:{GREEN}">Real DRE</span>',
           "% Meta","Status"]

def mc_header():
    ths = "".join(f'<th style="background:{NAVY};color:white;padding:9px 11px;'
                  f'font-size:11px;font-weight:600;">{c}</th>' for c in MC_COLS)
    return f'<table class="mct"><thead><tr>{ths}</tr></thead></table>'

def mc_row(it):
    return f"""<table class="mct"><tbody><tr>
      <td style="font-weight:600;min-width:130px;">{it['nome']}</td>
      <td>{fmt_brl(it['meta'])}</td>
      <td>{fmt_brl(it['prev'])}</td>
      <td style="color:{AMBER};">{fmt_brl(it['prev2026'])}</td>
      <td style="color:{TEAL};">{fmt_brl(it['val'])}</td>
      <td style="color:{GREEN};font-weight:600;">{fmt_brl(it['real'])}</td>
      <td>{pbar_html(it['pct'])}</td>
      <td>{bdg_status(it['pct'])}</td>
    </tr></tbody></table>"""

def mc_total(items):
    tm=tp=tp26=tv=tr=0
    for it in items:
        tm+=it["meta"];tp+=it["prev"];tp26+=it["prev2026"];tv+=it["val"];tr+=it["real"]
    pt = tr/tm if tm>0 else 0
    return f"""<table class="mct"><tbody><tr class="mc-tot">
      <td><b>TOTAL</b></td>
      <td><b>{fmt_brl(tm)}</b></td><td><b>{fmt_brl(tp)}</b></td>
      <td style="color:{AMBER};"><b>{fmt_brl(tp26)}</b></td>
      <td style="color:{TEAL};"><b>{fmt_brl(tv)}</b></td>
      <td style="color:{GREEN};font-weight:700;"><b>{fmt_brl(tr)}</b></td>
      <td>{pbar_html(pt)}</td><td></td>
    </tr></tbody></table>"""

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
  <div class="dh-b">Jun / 2026</div>
</div>""", unsafe_allow_html=True)

# ADMIN UPLOAD
with st.expander("🔐 Administrador — Atualizar Planilha"):
    arquivo = st.file_uploader("Nova versão (.xlsx)", type=["xlsx"], key="up")
    if arquivo:
        b = arquivo.read(); save_bytes(b)
        st.cache_data.clear()
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
ev      = extract_evolucao(D)
ranking = extract_ranking(D)

meta=kpis["meta"]; portfolio=kpis["portfolio"]; prev2026=kpis["prev2026"]
validado=kpis["validado"]; real=kpis["real"]; pct_ating=kpis["pct_ating"]

# ── KPI CARDS ──────────────────────────────────────────────────────────────────
# ── KPI CARDS ──────────────────────────────────────────────────────────────────
cob = portfolio/meta*100 if meta>0 else 0
pp  = prev2026/portfolio*100 if portfolio>0 else 0
pv  = validado/prev2026*100 if prev2026>0 else 0

def kpi(cls,lbl,vb,sub,det):
    return (f'<div class="kpi-card {cls}"><div class="kpi-l">{lbl}</div>'
            f'<div class="kpi-v">{vb}</div><div class="kpi-s">{sub}</div>'
            f'<div class="kpi-d">{det}</div></div>')

st.markdown(f"""<div class="kpi-wrap kpi-6">
  {kpi("","Meta Anual do Grupo (2026)",fmt_mi(meta),fmt_brl(meta),"Objetivo 2026 — 100%")}
  {kpi("cs","Portfólio Previsto (Anualizado)",fmt_mi(portfolio),fmt_brl(portfolio),f"{cob:.1f}% da meta coberta")}
  {kpi("ca","Previsto 2026",fmt_mi(prev2026),fmt_brl(prev2026),f"{pp:.1f}% do portfólio total")}
  {kpi("","Validado por Custos (2026)",fmt_mi(validado),fmt_brl(validado),f"{pv:.1f}% do Previsto 2026")}
  {kpi("cg","Retorno Real (DRE) (2026)",fmt_mi(real),fmt_brl(real),f"{pct_ating*100:.1f}% de atingimento")}
  {kpi("cr","Extra DRE (Até o Momento)",fmt_mi(extra_dre),fmt_brl(extra_dre),"Ganho fora do DRE acumulado")}
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
    series_all = ["Acumulado Previsto","Acumulado Real","Projeção da Meta","Previsto Mensal","Real Mensal"]
    sel = st.multiselect("Séries:", series_all, default=series_all[:3], key="ev_sel")
    if sel:
        st.plotly_chart(chart_evolucao(ev,sel), use_container_width=True, config={"displayModeBar":False})
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
        st.plotly_chart(chart_funnel(kpis), use_container_width=True, config={"displayModeBar":False})
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
    Constrói resumo de pilares agregando todas as abas com colunas corretas.
    Valores batem com os KPIs globais da planilha 5 Unidades.
    """
    from collections import defaultdict
    qtd_g=defaultdict(int); prev_g=defaultdict(float)
    val_g=defaultdict(float); real_g=defaultdict(float); dre_g={}
    def _add(proj_list):
        for p in proj_list:
            k=PILARES_EXIBE.get(p["tipo"],p["tipo"])
            qtd_g[k]+=1
            prev_g[k]+=p["previsto"]    # Total Ano Previsto
            val_g[k] +=p["val_saving"]  # Saving Validado (col13/col12)
            real_g[k]+=p["real_ano"]    # Total Ano Real (linha Real)
            dre_g[k]  =is_dre(p["tipo"])
    for sh in ["Diadema","Ferraz","São Leopoldo","Jarinu","Anchieta"]:
        _add(get_proj_planta(D, sh))
    _add(get_proj_compras(D))   # col_total_ano=37 corrigido
    _add(get_proj_vendas(D))    # start_row=32 corrigido
    ORDER=["BSW","Kaizen","Kaizen - Ganho Recorrente","Kaizen - Custo Evitado",
           "Kaizen - Capital de Giro","Redução de Custo","Você Resolve",
           "Estratégia Comercial","Meta Executiva"]
    res=[]
    for k in ORDER:
        if k in qtd_g:
            res.append(dict(nome=k,qtd=qtd_g[k],prev=prev_g[k],
                            val=val_g[k],real=real_g[k],dre=dre_g.get(k,True)))
    for k in sorted(qtd_g):
        if k not in ORDER:
            res.append(dict(nome=k,qtd=qtd_g[k],prev=prev_g[k],
                            val=val_g[k],real=real_g[k],dre=dre_g.get(k,True)))
    return res

if is_pil:
    cp1, cp2 = st.columns([5, 4])
    with cp1:
        t1, t2, t3, _ = st.columns([1,1,1,4])
        with t1: show_prev = st.toggle("Previsto", value=True,  key="tog_prev")
        with t2: show_val  = st.toggle("Validado", value=True,  key="tog_val")
        with t3: show_real = st.toggle("Real DRE", value=False, key="tog_real")
        p_grupo = build_pilares_grupo(hash(fb))  # fonte única para gráfico E tabela
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
        st.markdown(th("Pilar","Qtd","V. Previsto (Ano)","V. Validado","V. Real (DRE)")+rows_p+"</tbody></table>",
                    unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PLANTAS INDUSTRIAIS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sc">', unsafe_allow_html=True)
st.markdown(f'<span class="st">Plantas Industriais — Performance Consolidada</span>',
            unsafe_allow_html=True)
st.markdown(mc_header(), unsafe_allow_html=True)

for p in plantas:
    st.markdown(mc_row(p), unsafe_allow_html=True)
    with st.expander(f"＋  Ver projetos de {p['nome']}", expanded=False):
        proj = get_proj_planta(D, p["sheet"])
        n = len(proj)
        if proj:
            proj_v = render_proj_filtros(proj, key_prefix=f"plt_{p['nome']}")
            st.markdown(f"<p style='font-size:11px;color:{SILVER};margin:4px 0 8px;'>"
                        f"<b>{len(proj_v)}</b> de {n} projetos</p>", unsafe_allow_html=True)
            st.markdown(proj_table_html(proj_v), unsafe_allow_html=True)
            st.markdown("<hr style='margin:12px 0;border-color:#EEF0F3;'>", unsafe_allow_html=True)
            st.markdown(f'<p style="font-size:10px;font-weight:700;color:{NAVY};margin-bottom:6px;">Resumo por Pilar — {p['nome']}</p>',
                        unsafe_allow_html=True)
            st.markdown(pilar_resumo_html(proj), unsafe_allow_html=True)
        else:
            st.markdown("<p style='color:#999;font-size:12px;'>Sem projetos.</p>",
                        unsafe_allow_html=True)

st.markdown(mc_total(plantas), unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ÁREAS FUNCIONAIS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sc">', unsafe_allow_html=True)
st.markdown(f'<span class="st">Áreas Funcionais — Performance Consolidada</span>',
            unsafe_allow_html=True)
st.markdown(mc_header(), unsafe_allow_html=True)

area_fn = {"Compras": get_proj_compras, "Vendas": get_proj_vendas}

for a in areas:
    st.markdown(mc_row(a), unsafe_allow_html=True)
    with st.expander(f"＋  Ver projetos de {a['nome']}", expanded=False):
        fn = area_fn.get(a["nome"])
        proj = fn(D) if fn else []
        n = len(proj)
        if proj:
            proj_va = render_proj_filtros(proj, key_prefix=f"area_{a['nome']}")
            st.markdown(f"<p style='font-size:11px;color:{SILVER};margin:4px 0 8px;'>"
                        f"<b>{len(proj_va)}</b> de {n} projetos</p>", unsafe_allow_html=True)
            st.markdown(proj_table_html(proj_va), unsafe_allow_html=True)
            st.markdown("<hr style='margin:12px 0;border-color:#EEF0F3;'>", unsafe_allow_html=True)
            st.markdown(f'<p style="font-size:10px;font-weight:700;color:{NAVY};margin-bottom:6px;">Resumo por Pilar — {a['nome']}</p>',
                        unsafe_allow_html=True)
            st.markdown(pilar_resumo_html(proj), unsafe_allow_html=True)
        else:
            st.markdown("<p style='color:#999;font-size:12px;'>Sem projetos.</p>",
                        unsafe_allow_html=True)

st.markdown(mc_total(areas), unsafe_allow_html=True)
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
