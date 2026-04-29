"""
Dashboard de Controle de Matéria-Prima — Bobinas BSW
Lê o arquivo Excel diretamente do SharePoint via Microsoft Graph API
ou permite upload manual do arquivo.
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import io
from datetime import datetime

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Dashboard Bobinas BSW",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# ESTILO CSS CUSTOMIZADO (tema industrial escuro)
# ============================================================
st.markdown("""
<style>
    .stApp { background-color: #0A1628; }
    header[data-testid="stHeader"] { background-color: #0A1628; }
    section[data-testid="stSidebar"] { background-color: #0F1B2D; }
    div[data-testid="stMetric"] {
        background-color: #162236;
        border: 1px solid #1E3A5F;
        border-radius: 8px;
        padding: 16px;
    }
    div[data-testid="stMetric"] label { color: #B0BEC5 !important; }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #00D4FF !important;
        font-family: 'Consolas', monospace;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px; background-color: #0F1B2D; border-radius: 8px; padding: 4px;
    }
    .stTabs [data-baseweb="tab"] { color: #B0BEC5; border-radius: 6px; }
    .stTabs [aria-selected="true"] { background-color: #1E3A5F !important; color: #00D4FF !important; }
    h1, h2, h3 { color: #00D4FF !important; }
    h4, h5, h6 { color: #B0BEC5 !important; }
    p, span, li { color: #CFD8DC; }
    .stDataFrame { border: 1px solid #1E3A5F; border-radius: 8px; }
    hr { border-color: #1E3A5F; }
    .js-plotly-plot .plotly .main-svg { background-color: #162236 !important; }
    div[data-testid="stFileUploader"] {
        background-color: #162236; border: 2px dashed #1E3A5F; border-radius: 12px; padding: 16px;
    }
    div[data-testid="stFileUploader"]:hover { border-color: #00D4FF; }
    div[data-testid="stSidebar"] .stRadio label { color: #B0BEC5 !important; }
    /* Unidade selector buttons */
    .unidade-btn {
        background-color: #162236; border: 1px solid #1E3A5F; border-radius: 8px;
        padding: 12px 16px; text-align: center; cursor: pointer; transition: all 0.3s;
    }
    .unidade-btn:hover { border-color: #00D4FF; }
    .unidade-btn.active { border-color: #00D4FF; background-color: #1E3A5F; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# CORES DO TEMA
# ============================================================
COLORS = {
    "cyan": "#00D4FF", "amber": "#FFB800", "emerald": "#00E676",
    "coral": "#FF6B6B", "purple": "#A78BFA", "teal": "#4DD0E1",
    "bg_card": "#162236", "bg_dark": "#0F1B2D",
    "border": "#1E3A5F", "text_light": "#B0BEC5", "text_white": "#ECEFF1",
}

CHART_COLORS = [
    "#00D4FF", "#FFB800", "#00E676", "#A78BFA", "#FF6B6B",
    "#4DD0E1", "#FFD54F", "#69F0AE", "#B39DDB", "#FF8A80",
    "#80DEEA", "#FFF176", "#A5D6A7", "#CE93D8", "#EF9A9A",
]

PLOTLY_LAYOUT = dict(
    paper_bgcolor="#162236",
    plot_bgcolor="#162236",
    font=dict(color="#B0BEC5", family="Arial"),
    margin=dict(l=40, r=40, t=50, b=40),
    legend=dict(
        bgcolor="rgba(22,34,54,0.8)", bordercolor="#1E3A5F",
        borderwidth=1, font=dict(color="#B0BEC5"),
    ),
)

# ============================================================
# FUNÇÕES DE CONEXÃO COM SHAREPOINT
# ============================================================
@st.cache_data(ttl=300)
def get_access_token():
    tenant_id = st.secrets["TENANT_ID"]
    client_id = st.secrets["CLIENT_ID"]
    client_secret = st.secrets["CLIENT_SECRET"]
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()["access_token"]


@st.cache_data(ttl=300)
def load_data_from_sharepoint():
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    site_domain = st.secrets["SHAREPOINT_DOMAIN"]
    site_path = st.secrets["SHAREPOINT_SITE_PATH"]
    file_path = st.secrets["SHAREPOINT_FILE_PATH"]
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_domain}:/sites/{site_path}"
    site_resp = requests.get(site_url, headers=headers)
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]
    file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/content"
    file_resp = requests.get(file_url, headers=headers)
    file_resp.raise_for_status()
    excel_bytes = io.BytesIO(file_resp.content)
    df_controle = smart_read_excel(excel_bytes, "Controle")
    excel_bytes.seek(0)
    df_formulas = pd.read_excel(excel_bytes, sheet_name="Formulas")
    return df_controle, df_formulas


def smart_read_excel(excel_bytes, sheet_name):
    """Lê uma aba do Excel detectando automaticamente a linha do cabeçalho."""
    df = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=0)
    unnamed_count = sum(1 for c in df.columns if str(c).startswith('Unnamed'))
    if unnamed_count > len(df.columns) * 0.5:
        excel_bytes.seek(0)
        df_raw = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=None)
        for row_idx in range(min(5, len(df_raw))):
            row_vals = [str(v).replace('\n', ' ').strip() for v in df_raw.iloc[row_idx] if pd.notna(v)]
            if any('Código' in v or 'Bobina' in v or 'NECESSIDADE' in v or 'Tipo' in v for v in row_vals):
                excel_bytes.seek(0)
                df = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=row_idx)
                break
    return df


def load_data_from_upload(uploaded_file):
    excel_bytes = io.BytesIO(uploaded_file.getvalue())
    df_controle = smart_read_excel(excel_bytes, "Controle")
    excel_bytes.seek(0)
    df_formulas = pd.read_excel(excel_bytes, sheet_name="Formulas")
    return df_controle, df_formulas


def process_data(df_raw):
    """Processa e limpa os dados do controle."""
    df = df_raw.copy()
    df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]

    col_codigo = [c for c in df.columns if 'Código' in c and 'Bobina' in c]
    if col_codigo:
        df = df[df[col_codigo[0]].notna() & (df[col_codigo[0]] != '')]

    nec_cols = {
        'jan': [c for c in df.columns if 'Janeiro' in c and 'MÉDIA' not in c.upper()],
        'fev': [c for c in df.columns if 'Fevereiro' in c and 'MÉDIA' not in c.upper()],
        'mar': [c for c in df.columns if ('Março' in c or 'Marco' in c) and 'MÉDIA' not in c.upper()],
        'abr': [c for c in df.columns if 'Abril' in c and 'MÉDIA' not in c.upper()],
        'mai': [c for c in df.columns if 'Maio' in c and 'MÉDIA' not in c.upper()],
        'media': [c for c in df.columns if 'MÉDIA' in c.upper() and 'FEV' in c.upper()],
    }

    col_names = {}
    for key, cols in nec_cols.items():
        if cols:
            col_names[key] = cols[0]

    for key, col in col_names.items():
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df, col_names


def parse_formulas(df_formulas):
    """Extrai dados estruturados da aba Formulas.
    
    A aba tem duas seções:
    - Linhas 0-3: Resumo por unidade Delga (Ferraz, Diadema, Jarinu, Sul)
    - Linha 4: Total
    - Linhas 7+: Detalhamento por Usina
    """
    unidades = []
    usinas = []
    
    # Seção 1: Unidades (linhas 0-3, antes do 'Total')
    for i in range(min(10, len(df_formulas))):
        row = df_formulas.iloc[i]
        nome = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        if nome.lower() == 'total':
            break
        if not nome or nome.lower() in ['nan', 'usinas', '']:
            continue
            
        try:
            bobinas = int(float(row.iloc[1])) if pd.notna(row.iloc[1]) else 0
        except (ValueError, TypeError):
            continue
        try:
            peso_total = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
        except (ValueError, TypeError):
            peso_total = 0
        try:
            peso_analisado = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
        except (ValueError, TypeError):
            peso_analisado = 0
        try:
            pct_raw = row.iloc[4]
            if pd.notna(pct_raw):
                pct = float(pct_raw)
                pct = pct * 100 if pct <= 1 else pct
            else:
                pct = 0
        except (ValueError, TypeError):
            pct = 0
        try:
            ganho = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0
        except (ValueError, TypeError):
            ganho = 0
            
        unidades.append({
            'unidade': nome,
            'bobinas': bobinas,
            'peso_total': peso_total,
            'peso_analisado': peso_analisado,
            'pct': pct,
            'ganho': ganho,
        })
    
    # Seção 2: Usinas (após a linha que contém "Usinas" na coluna 0)
    usina_start = None
    for i in range(len(df_formulas)):
        val = str(df_formulas.iloc[i, 0]).strip().lower() if pd.notna(df_formulas.iloc[i, 0]) else ''
        if val == 'usinas':
            usina_start = i + 1
            break
    
    if usina_start:
        for i in range(usina_start, len(df_formulas)):
            row = df_formulas.iloc[i]
            nome = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            if nome.lower() in ['total', '']:
                if nome.lower() == 'total':
                    break
                continue
            if nome.lower() == 'nan' or nome == '0':
                continue
            try:
                bobinas = int(float(row.iloc[1])) if pd.notna(row.iloc[1]) else 0
            except (ValueError, TypeError):
                continue
            try:
                peso = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
            except (ValueError, TypeError):
                peso = 0
            try:
                pct_repr = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
            except (ValueError, TypeError):
                pct_repr = 0
            try:
                ganho_usina = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0
            except (ValueError, TypeError):
                ganho_usina = 0
                
            usinas.append({
                'usina': nome,
                'bobinas': bobinas,
                'peso': peso,
                'pct_representacao': pct_repr * 100 if pct_repr <= 1 else pct_repr,
                'ganho': ganho_usina,
            })
    
    df_unidades = pd.DataFrame(unidades) if unidades else pd.DataFrame()
    df_usinas = pd.DataFrame(usinas) if usinas else pd.DataFrame()
    
    return df_unidades, df_usinas


# ============================================================
# FUNÇÕES DE GRÁFICOS
# ============================================================
def create_area_chart(df, col_names):
    """Gráfico de evolução mensal."""
    meses = ['Jan/2026', 'Fev/2026', 'Mar/2026', 'Abr/2026', 'Mai/2026']
    keys = ['jan', 'fev', 'mar', 'abr', 'mai']
    valores = []
    for k in keys:
        if k in col_names:
            valores.append(df[col_names[k]].sum())
        else:
            valores.append(0)

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=meses, y=valores,
        fill='tozeroy', fillcolor='rgba(0,212,255,0.15)',
        line=dict(color=COLORS["cyan"], width=3),
        mode='lines+markers',
        marker=dict(size=8, color=COLORS["cyan"]),
        hovertemplate='%{x}<br><b>%{y:,.0f} ton</b><extra></extra>',
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Evolução da Necessidade Mensal", font=dict(size=16, color=COLORS["cyan"])),
        yaxis=dict(title="Toneladas", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        height=400,
    )
    return fig


def create_pie_chart(df, col_media, title, group_col, top_n=10):
    """Gráfico de pizza/donut."""
    df_valid = df[df[group_col].notna() & (df[group_col].astype(str).str.strip() != '')].copy()
    if len(df_valid) == 0:
        return go.Figure().update_layout(**PLOTLY_LAYOUT, title=dict(text=title))
    dist = df_valid.groupby(group_col)[col_media].sum().sort_values(ascending=False).head(top_n)
    fig = go.Figure(data=[go.Pie(
        labels=dist.index, values=dist.values, hole=0.45,
        marker=dict(colors=CHART_COLORS[:len(dist)]),
        textinfo='percent+label',
        textfont=dict(size=11, color="#ECEFF1"),
        hovertemplate='%{label}<br><b>%{value:,.1f} ton</b><br>%{percent}<extra></extra>',
    )])
    fig.update_layout(**PLOTLY_LAYOUT, title=dict(text=title, font=dict(size=16, color=COLORS["cyan"])), height=400)
    return fig


def create_bar_chart(df, col_media, title, group_col, top_n=15, color=None, orientation='h'):
    """Gráfico de barras."""
    df_valid = df[df[group_col].notna() & (df[group_col].astype(str).str.strip() != '')].copy()
    if len(df_valid) == 0:
        fig = go.Figure()
        fig.update_layout(**PLOTLY_LAYOUT, title=dict(text=title, font=dict(size=16, color=COLORS["cyan"])))
        return fig

    dist = df_valid.groupby(group_col)[col_media].sum().sort_values(ascending=True).tail(top_n)

    if orientation == 'h':
        fig = go.Figure(data=[go.Bar(
            x=dist.values, y=dist.index, orientation='h',
            marker=dict(color=color or COLORS["cyan"], line=dict(width=0)),
            hovertemplate='%{y}<br><b>%{x:,.1f} ton</b><extra></extra>',
        )])
    else:
        fig = go.Figure(data=[go.Bar(
            x=dist.index, y=dist.values, orientation='v',
            marker=dict(color=color or COLORS["cyan"], line=dict(width=0)),
            hovertemplate='%{x}<br><b>%{y:,.1f} ton</b><extra></extra>',
        )])

    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text=title, font=dict(size=16, color=COLORS["cyan"])),
        height=max(400, top_n * 30),
        yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F",
                   title="Toneladas" if orientation == 'h' else None),
    )
    return fig


def create_thickness_chart(df, col_media):
    """Gráfico de distribuição por faixa de espessura."""
    esp_col = [c for c in df.columns if 'Esp' in c and 'mm' in c]
    if not esp_col:
        return None
    df_temp = df.copy()
    df_temp['esp_num'] = pd.to_numeric(df_temp[esp_col[0]], errors='coerce')
    df_temp = df_temp[df_temp['esp_num'].notna()]
    if len(df_temp) == 0:
        return None
    bins = [0, 1, 2, 4, 6, 8, 10, 15, 20, 50]
    labels = ['0-1mm', '1-2mm', '2-4mm', '4-6mm', '6-8mm', '8-10mm', '10-15mm', '15-20mm', '20+mm']
    df_temp['faixa'] = pd.cut(df_temp['esp_num'], bins=bins, labels=labels, right=True)
    dist = df_temp.groupby('faixa', observed=True)[col_media].sum().sort_index()
    dist = dist[dist > 0]
    if len(dist) == 0:
        return None
    fig = go.Figure(data=[go.Bar(
        x=[str(x) for x in dist.index], y=dist.values,
        marker=dict(color=CHART_COLORS[:len(dist)], line=dict(width=0)),
        hovertemplate='%{x}<br><b>%{y:,.1f} ton</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Distribuição por Faixa de Espessura", font=dict(size=16, color=COLORS["cyan"])),
        yaxis=dict(title="Toneladas", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        height=400,
    )
    return fig


def create_progress_chart(df_unidades):
    """Gráfico de progresso de análise por unidade."""
    if len(df_unidades) == 0:
        return go.Figure().update_layout(**PLOTLY_LAYOUT)
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Peso Total',
        x=df_unidades['unidade'], y=df_unidades['peso_total'],
        marker=dict(color='#546E7A'),
        hovertemplate='%{x}<br>Peso Total: <b>%{y:,.1f} ton</b><extra></extra>',
    ))
    fig.add_trace(go.Bar(
        name='Peso Analisado',
        x=df_unidades['unidade'], y=df_unidades['peso_analisado'],
        marker=dict(color=COLORS["cyan"]),
        hovertemplate='%{x}<br>Analisado: <b>%{y:,.1f} ton</b><extra></extra>',
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Peso Total vs Analisado por Unidade", font=dict(size=16, color=COLORS["cyan"])),
        barmode='group',
        yaxis=dict(title="Toneladas", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        height=400,
    )
    return fig


def create_usinas_chart(df_usinas, top_n=15):
    """Gráfico de barras das usinas."""
    if len(df_usinas) == 0:
        return go.Figure().update_layout(**PLOTLY_LAYOUT)
    
    df_sorted = df_usinas.nlargest(top_n, 'peso')
    df_sorted = df_sorted.sort_values('peso', ascending=True)
    
    fig = go.Figure(data=[go.Bar(
        x=df_sorted['peso'], y=df_sorted['usina'], orientation='h',
        marker=dict(color=COLORS["emerald"], line=dict(width=0)),
        hovertemplate='%{y}<br><b>%{x:,.1f} ton</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Top Usinas por Peso", font=dict(size=16, color=COLORS["cyan"])),
        height=max(400, min(top_n, len(df_sorted)) * 30),
        yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F", title="Toneladas"),
    )
    return fig


# ============================================================
# APLICAÇÃO PRINCIPAL
# ============================================================
def main():
    # SIDEBAR
    with st.sidebar:
        st.markdown("""
        <div style="text-align:center; padding:16px 0;">
            <span style="font-size:40px;">🏭</span>
            <h2 style="margin:8px 0 4px 0; font-size:18px; color:#00D4FF !important;">Dashboard BSW</h2>
            <p style="color:#546E7A; font-size:12px; margin:0;">Controle de Matéria-Prima</p>
        </div>
        <hr style="border-color:#1E3A5F; margin:8px 0 16px 0;">
        """, unsafe_allow_html=True)

        st.markdown("#### Fonte de Dados")
        fonte = st.radio(
            "Selecione como carregar os dados:",
            ["📤 Upload Manual", "☁️ SharePoint (automático)"],
            index=0,
            help="Use o upload manual enquanto o acesso ao SharePoint não estiver liberado."
        )

        uploaded_file = None
        if "Upload" in fonte:
            st.markdown("---")
            st.markdown("#### Enviar Arquivo Excel")
            uploaded_file = st.file_uploader(
                "Arraste ou clique para enviar o arquivo",
                type=["xlsx", "xls"],
                help="Envie o arquivo 'Controle Resumo - Base BSW.xlsx'",
            )
            if uploaded_file:
                st.success(f"Arquivo carregado: **{uploaded_file.name}**")
            else:
                st.info("Aguardando arquivo...")

        st.markdown("---")
        st.markdown("""
        <div style="padding:8px; background:#162236; border-radius:8px; border:1px solid #1E3A5F; margin-top:8px;">
            <p style="color:#546E7A; font-size:11px; margin:0;">
                <b style="color:#B0BEC5;">Dica:</b> Quando o acesso ao SharePoint for liberado, 
                selecione "SharePoint" e o dashboard atualizará automaticamente 
                sempre que o arquivo for modificado.
            </p>
        </div>
        """, unsafe_allow_html=True)

    # HEADER
    st.markdown("""
    <div style="display:flex; align-items:center; gap:16px; margin-bottom:8px;">
        <div style="background:#162236; border:1px solid #1E3A5F; border-radius:12px; padding:12px; display:flex; align-items:center; justify-content:center;">
            <span style="font-size:28px;">🏭</span>
        </div>
        <div>
            <h1 style="margin:0; font-size:28px; color:#00D4FF !important;">Controle de Matéria-Prima</h1>
            <p style="margin:0; color:#546E7A; font-family:Consolas,monospace; font-size:13px;">BOBINAS BSW — JAN A MAI / 2026</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # CARREGAR DADOS
    df_raw = None
    df_formulas = None

    if "Upload" in fonte:
        if uploaded_file is None:
            st.markdown("""
            <div style="text-align:center; padding:80px 20px; background:#162236; border:2px dashed #1E3A5F; border-radius:16px; margin:40px auto; max-width:600px;">
                <span style="font-size:64px;">📊</span>
                <h2 style="color:#00D4FF !important; margin:16px 0 8px 0;">Envie seu arquivo Excel</h2>
                <p style="color:#546E7A; font-size:14px;">Use o painel lateral à esquerda para fazer upload do arquivo<br><b style="color:#B0BEC5;">Controle Resumo - Base BSW.xlsx</b></p>
            </div>
            """, unsafe_allow_html=True)
            st.stop()
        else:
            try:
                df_raw, df_formulas = load_data_from_upload(uploaded_file)
            except Exception as e:
                st.error(f"Erro ao ler o arquivo: {str(e)}")
                st.info("Verifique se o arquivo possui as abas 'Controle' e 'Formulas'.")
                st.stop()
    else:
        try:
            with st.spinner("Conectando ao SharePoint e carregando dados..."):
                df_raw, df_formulas = load_data_from_sharepoint()
        except Exception as e:
            st.error(f"Erro ao conectar ao SharePoint: {str(e)}")
            st.info("Verifique as credenciais ou use o Upload Manual no painel lateral.")
            st.stop()

    # Processar dados
    df, col_names = process_data(df_raw)
    df_unidades, df_usinas = parse_formulas(df_formulas)

    col_media = col_names.get('media', '')
    if not col_media:
        st.error("Coluna de necessidade média não encontrada no arquivo.")
        st.stop()

    # Timestamp
    st.markdown(f"""
    <p style="text-align:right; color:#546E7A; font-size:12px; font-family:Consolas,monospace;">
        Última atualização: {datetime.now().strftime('%d/%m/%Y %H:%M')}
    </p>
    """, unsafe_allow_html=True)

    # ============================================================
    # KPIs GERAIS
    # ============================================================
    total_bobinas = df_unidades['bobinas'].sum() if len(df_unidades) > 0 else 0
    total_peso = df_unidades['peso_total'].sum() if len(df_unidades) > 0 else 0
    total_peso_analisado = df_unidades['peso_analisado'].sum() if len(df_unidades) > 0 else 0
    total_ganho = df_unidades['ganho'].sum() if len(df_unidades) > 0 else 0
    nec_media_total = df[col_media].sum()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("Total de Bobinas", f"{total_bobinas:,}".replace(",", "."))
    with k2:
        st.metric("Peso Total", f"{total_peso:,.0f} ton".replace(",", "."))
    with k3:
        st.metric("Peso Analisado", f"{total_peso_analisado:,.0f} ton".replace(",", "."))
    with k4:
        st.metric("Ganho Potencial", f"R$ {total_ganho:,.0f}".replace(",", "."))

    # ============================================================
    # SELETOR DE UNIDADE — KPIs por unidade
    # ============================================================
    if len(df_unidades) > 0:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### Detalhamento por Unidade")
        
        unidade_names = ["Todas"] + df_unidades['unidade'].tolist()
        selected_unidade = st.selectbox(
            "Selecione a unidade:",
            unidade_names,
            index=0,
            key="unidade_selector"
        )
        
        if selected_unidade == "Todas":
            u_bobinas = total_bobinas
            u_peso = total_peso
            u_analisado = total_peso_analisado
            u_ganho = total_ganho
            u_pct = (total_peso_analisado / total_peso * 100) if total_peso > 0 else 0
        else:
            row_u = df_unidades[df_unidades['unidade'] == selected_unidade].iloc[0]
            u_bobinas = row_u['bobinas']
            u_peso = row_u['peso_total']
            u_analisado = row_u['peso_analisado']
            u_ganho = row_u['ganho']
            u_pct = row_u['pct']
        
        uk1, uk2, uk3, uk4, uk5 = st.columns(5)
        with uk1:
            st.metric("Bobinas", f"{u_bobinas:,}".replace(",", "."))
        with uk2:
            st.metric("Peso Total", f"{u_peso:,.1f} ton".replace(",", "."))
        with uk3:
            st.metric("Peso Analisado", f"{u_analisado:,.1f} ton".replace(",", "."))
        with uk4:
            st.metric("% Concluído", f"{u_pct:.1f}%")
        with uk5:
            st.metric("Ganho (R$)", f"R$ {u_ganho:,.0f}".replace(",", "."))

    st.markdown("<br>", unsafe_allow_html=True)

    # ============================================================
    # ABAS
    # ============================================================
    tab1, tab2, tab3 = st.tabs(["📊 Visão Geral", "🔍 Análises", "💰 Financeiro"])

    # ABA 1: VISÃO GERAL
    with tab1:
        col_a, col_b = st.columns([2, 1])
        with col_a:
            fig_area = create_area_chart(df, col_names)
            st.plotly_chart(fig_area, use_container_width=True)
        with col_b:
            tipo_col = [c for c in df.columns if c.strip() == 'Tipo']
            if tipo_col:
                fig_tipo = create_pie_chart(df, col_media, "Distribuição por Tipo", tipo_col[0])
                st.plotly_chart(fig_tipo, use_container_width=True)

        col_c, col_d = st.columns(2)
        with col_c:
            fig_esp = create_thickness_chart(df, col_media)
            if fig_esp:
                st.plotly_chart(fig_esp, use_container_width=True)
        with col_d:
            unidade_col = [c for c in df.columns if 'Unidade' in c and 'Delga' in c]
            if unidade_col:
                fig_unid = create_pie_chart(df, col_media, "Distribuição por Unidade Delga", unidade_col[0], 6)
                st.plotly_chart(fig_unid, use_container_width=True)

        # Usinas da aba Formulas
        if len(df_usinas) > 0:
            fig_usinas = create_usinas_chart(df_usinas, 15)
            st.plotly_chart(fig_usinas, use_container_width=True)

        # Tabela Top 15 Bobinas
        st.markdown("### Top 15 Bobinas por Necessidade Média")
        codigo_col = [c for c in df.columns if 'Código' in c and 'Bobina' in c]
        if codigo_col:
            display_cols = [codigo_col[0]]
            tipo_col = [c for c in df.columns if c.strip() == 'Tipo']
            if tipo_col:
                display_cols.append(tipo_col[0])
            if 'Projeto' in df.columns:
                display_cols.append('Projeto')
            display_cols.append(col_media)

            top15 = df.nlargest(15, col_media)[display_cols].copy()
            col_rename = {codigo_col[0]: 'Código', col_media: 'Necessidade Média (ton)'}
            top15 = top15.rename(columns=col_rename)
            top15['Necessidade Média (ton)'] = top15['Necessidade Média (ton)'].round(1)
            top15 = top15.reset_index(drop=True)
            top15.index = top15.index + 1
            st.dataframe(top15, use_container_width=True, height=560)

    # ABA 2: ANÁLISES
    with tab2:
        if len(df_unidades) > 0:
            fig_prog = create_progress_chart(df_unidades)
            st.plotly_chart(fig_prog, use_container_width=True)

            # Tabela de progresso
            st.markdown("### Progresso de Análise por Unidade")
            df_display = df_unidades.copy()
            df_display.columns = ['Unidade', 'Bobinas', 'Peso Total (ton)', 'Peso Analisado (ton)', '% Concluído', 'Ganho (R$)']
            df_display['Peso Total (ton)'] = df_display['Peso Total (ton)'].round(1)
            df_display['Peso Analisado (ton)'] = df_display['Peso Analisado (ton)'].round(1)
            df_display['% Concluído'] = df_display['% Concluído'].apply(lambda x: f"{x:.1f}%")
            df_display['Ganho (R$)'] = df_display['Ganho (R$)'].apply(lambda x: f"R$ {x:,.0f}".replace(",", "."))
            st.dataframe(df_display, use_container_width=True, hide_index=True)
        else:
            st.info("Dados de análise não encontrados na aba Formulas.")

        # Distribuição por unidade e beneficiador
        st.markdown("### Necessidade por Unidade e Beneficiador")
        unidade_col = [c for c in df.columns if 'Unidade' in c and 'Delga' in c]
        benef_col = [c for c in df.columns if 'Beneficiador' in c]
        if unidade_col or benef_col:
            col_g, col_h = st.columns(2)
            with col_g:
                if unidade_col:
                    fig_unid2 = create_bar_chart(df, col_media, "Necessidade por Unidade Delga", unidade_col[0], 6, COLORS["cyan"])
                    st.plotly_chart(fig_unid2, use_container_width=True)
            with col_h:
                if benef_col:
                    fig_benef = create_bar_chart(df, col_media, "Necessidade por Beneficiador", benef_col[0], 10, COLORS["teal"])
                    st.plotly_chart(fig_benef, use_container_width=True)

        # Classificação ABC
        abc_col = [c for c in df.columns if c.strip().upper() == 'ABC']
        if abc_col:
            st.markdown("### Classificação ABC")
            df_abc = df[df[abc_col[0]].notna() & (df[abc_col[0]].astype(str).str.strip() != '')].copy()
            if len(df_abc) > 0:
                abc_dist = df_abc.groupby(abc_col[0])[col_media].agg(['sum', 'count']).sort_values('sum', ascending=False)
                abc_dist.columns = ['Necessidade Total (ton)', 'Qtd Bobinas']
                abc_dist['Necessidade Total (ton)'] = abc_dist['Necessidade Total (ton)'].round(1)
                st.dataframe(abc_dist, use_container_width=True)

    # ABA 3: FINANCEIRO
    with tab3:
        if len(df_unidades) > 0:
            col_i, col_j = st.columns(2)
            with col_i:
                # Gráfico de ganho por unidade (pizza)
                df_ganho_pie = df_unidades[df_unidades['ganho'] > 0]
                if len(df_ganho_pie) > 0:
                    fig_ganho = go.Figure(data=[go.Pie(
                        labels=df_ganho_pie['unidade'], values=df_ganho_pie['ganho'], hole=0.45,
                        marker=dict(colors=[COLORS["cyan"], COLORS["amber"], COLORS["emerald"], COLORS["coral"], COLORS["purple"]]),
                        textinfo='percent+label',
                        textfont=dict(size=11, color="#ECEFF1"),
                        hovertemplate='%{label}<br><b>R$ %{value:,.0f}</b><br>%{percent}<extra></extra>',
                    )])
                    fig_ganho.update_layout(
                        **PLOTLY_LAYOUT,
                        title=dict(text="Ganho Financeiro por Unidade", font=dict(size=16, color=COLORS["cyan"])),
                        height=400,
                    )
                    st.plotly_chart(fig_ganho, use_container_width=True)
                else:
                    st.info("Nenhum ganho financeiro registrado ainda. Os dados aparecerão conforme as análises forem concluídas.")
            
            with col_j:
                # Ganho por unidade em barras
                if df_unidades['ganho'].sum() > 0:
                    df_ganho_bar = df_unidades[df_unidades['ganho'] > 0].sort_values('ganho', ascending=True)
                    fig_ganho_bar = go.Figure(data=[go.Bar(
                        x=df_ganho_bar['ganho'], y=df_ganho_bar['unidade'], orientation='h',
                        marker=dict(color=COLORS["amber"]),
                        hovertemplate='%{y}<br><b>R$ %{x:,.0f}</b><extra></extra>',
                    )])
                    fig_ganho_bar.update_layout(
                        **PLOTLY_LAYOUT,
                        title=dict(text="Ganho Financeiro por Unidade (R$)", font=dict(size=16, color=COLORS["cyan"])),
                        xaxis=dict(title="R$", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
                        yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
                        height=400,
                    )
                    st.plotly_chart(fig_ganho_bar, use_container_width=True)
                else:
                    st.info("Ganho financeiro aparecerá conforme as análises forem concluídas.")

            # Usinas com ganho
            if len(df_usinas) > 0 and df_usinas['ganho'].sum() > 0:
                st.markdown("### Ganho Financeiro por Usina")
                df_usinas_ganho = df_usinas[df_usinas['ganho'] > 0].sort_values('ganho', ascending=True)
                fig_usina_ganho = go.Figure(data=[go.Bar(
                    x=df_usinas_ganho['ganho'], y=df_usinas_ganho['usina'], orientation='h',
                    marker=dict(color=COLORS["emerald"]),
                    hovertemplate='%{y}<br><b>R$ %{x:,.0f}</b><extra></extra>',
                )])
                fig_usina_ganho.update_layout(
                    **PLOTLY_LAYOUT,
                    title=dict(text="Ganho por Usina (R$)", font=dict(size=16, color=COLORS["cyan"])),
                    xaxis=dict(title="R$", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
                    yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
                    height=max(400, len(df_usinas_ganho) * 30),
                )
                st.plotly_chart(fig_usina_ganho, use_container_width=True)

            # Resumo financeiro
            st.markdown("### Resumo Financeiro por Unidade")
            df_fin = df_unidades[['unidade', 'bobinas', 'peso_total', 'peso_analisado', 'pct', 'ganho']].copy()
            df_fin.columns = ['Unidade', 'Bobinas', 'Peso Total (ton)', 'Peso Analisado (ton)', '% Concluído', 'Ganho (R$)']
            df_fin['Peso Total (ton)'] = df_fin['Peso Total (ton)'].round(1)
            df_fin['Peso Analisado (ton)'] = df_fin['Peso Analisado (ton)'].round(1)
            df_fin['% Concluído'] = df_fin['% Concluído'].apply(lambda x: f"{x:.1f}%")
            df_fin['Ganho (R$)'] = df_fin['Ganho (R$)'].apply(lambda x: f"R$ {x:,.0f}".replace(",", "."))
            st.dataframe(df_fin, use_container_width=True, hide_index=True)
        else:
            st.info("Dados financeiros não encontrados na aba Formulas.")


if __name__ == "__main__":
    main()
