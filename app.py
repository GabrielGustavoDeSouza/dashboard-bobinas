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
    /* Fundo geral */
    .stApp {
        background-color: #0A1628;
    }
    /* Header */
    header[data-testid="stHeader"] {
        background-color: #0A1628;
    }
    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #0F1B2D;
    }
    /* Cards de métricas */
    div[data-testid="stMetric"] {
        background-color: #162236;
        border: 1px solid #1E3A5F;
        border-radius: 8px;
        padding: 16px;
    }
    div[data-testid="stMetric"] label {
        color: #B0BEC5 !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #00D4FF !important;
        font-family: 'Consolas', monospace;
    }
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #0F1B2D;
        border-radius: 8px;
        padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        color: #B0BEC5;
        border-radius: 6px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E3A5F !important;
        color: #00D4FF !important;
    }
    /* Títulos */
    h1, h2, h3 {
        color: #00D4FF !important;
    }
    h4, h5, h6 {
        color: #B0BEC5 !important;
    }
    p, span, li {
        color: #CFD8DC;
    }
    /* Containers */
    div[data-testid="stVerticalBlock"] > div[data-testid="stExpander"] {
        background-color: #162236;
        border: 1px solid #1E3A5F;
        border-radius: 8px;
    }
    /* Dataframes */
    .stDataFrame {
        border: 1px solid #1E3A5F;
        border-radius: 8px;
    }
    /* Divider */
    hr {
        border-color: #1E3A5F;
    }
    /* Plotly charts background */
    .js-plotly-plot .plotly .main-svg {
        background-color: #162236 !important;
    }
    /* Upload area */
    div[data-testid="stFileUploader"] {
        background-color: #162236;
        border: 2px dashed #1E3A5F;
        border-radius: 12px;
        padding: 16px;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #00D4FF;
    }
    /* Sidebar radio */
    div[data-testid="stSidebar"] .stRadio label {
        color: #B0BEC5 !important;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# CORES DO TEMA
# ============================================================
COLORS = {
    "cyan": "#00D4FF",
    "amber": "#FFB800",
    "emerald": "#00E676",
    "coral": "#FF6B6B",
    "purple": "#A78BFA",
    "teal": "#4DD0E1",
    "bg_card": "#162236",
    "bg_dark": "#0F1B2D",
    "border": "#1E3A5F",
    "text_light": "#B0BEC5",
    "text_white": "#ECEFF1",
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
        bgcolor="rgba(22,34,54,0.8)",
        bordercolor="#1E3A5F",
        borderwidth=1,
        font=dict(color="#B0BEC5"),
    ),
    xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
    yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
)

# ============================================================
# FUNÇÕES DE CONEXÃO COM SHAREPOINT
# ============================================================
@st.cache_data(ttl=300)
def get_access_token():
    """Obtém token de acesso via Microsoft Graph API."""
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
    """Baixa o Excel do SharePoint e retorna DataFrames."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    site_domain = st.secrets["SHAREPOINT_DOMAIN"]
    site_path = st.secrets["SHAREPOINT_SITE_PATH"]
    file_path = st.secrets["SHAREPOINT_FILE_PATH"]

    # 1. Obter site ID
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_domain}:/sites/{site_path}"
    site_resp = requests.get(site_url, headers=headers)
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]

    # 2. Baixar arquivo
    file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/content"
    file_resp = requests.get(file_url, headers=headers)
    file_resp.raise_for_status()

    # 3. Ler Excel
    excel_bytes = io.BytesIO(file_resp.content)
    df_controle = smart_read_excel(excel_bytes, "Controle")
    excel_bytes.seek(0)
    df_formulas = pd.read_excel(excel_bytes, sheet_name="Formulas")

    return df_controle, df_formulas


def smart_read_excel(excel_bytes, sheet_name):
    """Lê uma aba do Excel detectando automaticamente a linha do cabeçalho."""
    # Tenta ler normalmente primeiro
    df = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=0)
    
    # Se todas as colunas são 'Unnamed', o cabeçalho real está em outra linha
    unnamed_count = sum(1 for c in df.columns if str(c).startswith('Unnamed'))
    if unnamed_count > len(df.columns) * 0.5:
        # Procurar a linha que contém o cabeçalho real (até as primeiras 5 linhas)
        excel_bytes.seek(0)
        df_raw = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=None)
        for row_idx in range(min(5, len(df_raw))):
            row_vals = [str(v).replace('\n', ' ').strip() for v in df_raw.iloc[row_idx] if pd.notna(v)]
            # Verificar se esta linha parece um cabeçalho
            if any('Código' in v or 'Bobina' in v or 'NECESSIDADE' in v or 'Tipo' in v for v in row_vals):
                excel_bytes.seek(0)
                df = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=row_idx)
                break
    
    return df


def load_data_from_upload(uploaded_file):
    """Lê o Excel do upload manual."""
    excel_bytes = io.BytesIO(uploaded_file.getvalue())
    df_controle = smart_read_excel(excel_bytes, "Controle")
    excel_bytes.seek(0)
    df_formulas = pd.read_excel(excel_bytes, sheet_name="Formulas")
    return df_controle, df_formulas


def process_data(df_raw):
    """Processa e limpa os dados do controle."""
    df = df_raw.copy()
    df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]

    # Remover linhas vazias
    col_codigo = [c for c in df.columns if 'Código' in c and 'Bobina' in c]
    if col_codigo:
        df = df[df[col_codigo[0]].notna() & (df[col_codigo[0]] != '')]

    # Colunas de necessidade
    nec_cols = {
        'jan': [c for c in df.columns if 'Janeiro' in c],
        'fev': [c for c in df.columns if 'Fevereiro' in c],
        'mar': [c for c in df.columns if 'Março' in c or 'Marco' in c],
        'abr': [c for c in df.columns if 'Abril' in c],
        'mai': [c for c in df.columns if 'Maio' in c],
        'media': [c for c in df.columns if 'MÉDIA' in c.upper() and 'FEV' in c.upper()],
    }

    col_names = {}
    for key, cols in nec_cols.items():
        if cols:
            col_names[key] = cols[0]

    return df, col_names


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
        fill='tozeroy',
        fillcolor='rgba(0,212,255,0.15)',
        line=dict(color=COLORS["cyan"], width=3),
        mode='lines+markers',
        marker=dict(size=8, color=COLORS["cyan"]),
        hovertemplate='%{x}<br><b>%{y:,.0f} ton</b><extra></extra>',
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Evolução da Necessidade Mensal", font=dict(size=16, color=COLORS["cyan"])),
        yaxis_title="Toneladas",
        height=400,
    )
    return fig


def create_pie_chart(df, col_media, title, group_col, top_n=10):
    """Gráfico de pizza/donut."""
    dist = df.groupby(group_col)[col_media].sum().sort_values(ascending=False).head(top_n)
    fig = go.Figure(data=[go.Pie(
        labels=dist.index,
        values=dist.values,
        hole=0.45,
        marker=dict(colors=CHART_COLORS[:len(dist)]),
        textinfo='percent+label',
        textfont=dict(size=11, color="#ECEFF1"),
        hovertemplate='%{label}<br><b>%{value:,.1f} ton</b><br>%{percent}<extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text=title, font=dict(size=16, color=COLORS["cyan"])),
        height=400,
        showlegend=True,
    )
    return fig


def create_bar_chart(df, col_media, title, group_col, top_n=15, color=None, orientation='h'):
    """Gráfico de barras horizontal."""
    dist = df.groupby(group_col)[col_media].sum().sort_values(ascending=True).tail(top_n)
    fig = go.Figure(data=[go.Bar(
        x=dist.values if orientation == 'h' else dist.index,
        y=dist.index if orientation == 'h' else dist.values,
        orientation=orientation,
        marker=dict(
            color=color or COLORS["cyan"],
            line=dict(width=0),
        ),
        hovertemplate='%{y}<br><b>%{x:,.1f} ton</b><extra></extra>' if orientation == 'h'
        else '%{x}<br><b>%{y:,.1f} ton</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text=title, font=dict(size=16, color=COLORS["cyan"])),
        height=max(400, top_n * 30),
        yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F", title="Toneladas" if orientation == 'h' else None),
    )
    return fig


def create_thickness_chart(df, col_media):
    """Gráfico de distribuição por faixa de espessura."""
    esp_col = [c for c in df.columns if 'Esp' in c and 'mm' in c]
    if not esp_col:
        return None

    df_temp = df.copy()
    df_temp['esp_num'] = pd.to_numeric(df_temp[esp_col[0]], errors='coerce')
    bins = [0, 1, 2, 4, 6, 8, 10, 15, 20]
    labels = ['0-1mm', '1-2mm', '2-4mm', '4-6mm', '6-8mm', '8-10mm', '10-15mm', '15-20mm']
    df_temp['faixa'] = pd.cut(df_temp['esp_num'], bins=bins, labels=labels, right=True)
    dist = df_temp.groupby('faixa', observed=True)[col_media].sum().sort_index()

    fig = go.Figure(data=[go.Bar(
        x=[str(x) for x in dist.index],
        y=dist.values,
        marker=dict(
            color=CHART_COLORS[:len(dist)],
            line=dict(width=0),
        ),
        hovertemplate='%{x}<br><b>%{y:,.1f} ton</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Distribuição por Faixa de Espessura", font=dict(size=16, color=COLORS["cyan"])),
        yaxis_title="Toneladas",
        height=400,
    )
    return fig


def create_progress_chart(df_formulas):
    """Gráfico de progresso de análise por unidade."""
    unidades = []
    for i in range(min(4, len(df_formulas))):
        row = df_formulas.iloc[i]
        if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip():
            unidades.append({
                'unidade': str(row.iloc[0]),
                'bobinas': int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0,
                'peso_total': float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0,
                'peso_analisado': float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
                'pct': float(row.iloc[4]) * 100 if pd.notna(row.iloc[4]) else 0,
                'ganho': float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0,
            })

    df_u = pd.DataFrame(unidades)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Peso Total',
        x=df_u['unidade'],
        y=df_u['peso_total'],
        marker=dict(color='#546E7A'),
        hovertemplate='%{x}<br>Peso Total: <b>%{y:,.1f} ton</b><extra></extra>',
    ))
    fig.add_trace(go.Bar(
        name='Peso Analisado',
        x=df_u['unidade'],
        y=df_u['peso_analisado'],
        marker=dict(color=COLORS["cyan"]),
        hovertemplate='%{x}<br>Analisado: <b>%{y:,.1f} ton</b><extra></extra>',
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Peso Total vs Analisado por Unidade", font=dict(size=16, color=COLORS["cyan"])),
        barmode='group',
        yaxis_title="Toneladas",
        height=400,
    )
    return fig, df_u


def create_ganho_pie(df_u):
    """Gráfico de ganho financeiro por unidade."""
    df_ganho = df_u[df_u['ganho'] > 0]
    if len(df_ganho) == 0:
        return None
    fig = go.Figure(data=[go.Pie(
        labels=df_ganho['unidade'],
        values=df_ganho['ganho'],
        hole=0.45,
        marker=dict(colors=[COLORS["cyan"], COLORS["amber"], COLORS["emerald"], COLORS["coral"]]),
        textinfo='percent+label',
        textfont=dict(size=11, color="#ECEFF1"),
        hovertemplate='%{label}<br><b>R$ %{value:,.0f}</b><br>%{percent}<extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Ganho Financeiro por Unidade", font=dict(size=16, color=COLORS["cyan"])),
        height=400,
    )
    return fig


# ============================================================
# APLICAÇÃO PRINCIPAL
# ============================================================
def main():
    # ============================================================
    # SIDEBAR — Fonte de dados
    # ============================================================
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

    # ============================================================
    # HEADER
    # ============================================================
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

    # ============================================================
    # CARREGAR DADOS
    # ============================================================
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
    # KPIs
    # ============================================================
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("Total de Bobinas", f"{len(df):,}".replace(",", "."))
    with k2:
        st.metric("Necessidade Média/Mês", f"{df[col_media].sum():,.0f} ton".replace(",", "."))
    with k3:
        st.metric("Clientes Distintos", f"{df['Cliente'].nunique()}")
    with k4:
        ganho_total = df_formulas.iloc[4, 5] if len(df_formulas) > 4 and pd.notna(df_formulas.iloc[4, 5]) else 0
        st.metric("Ganho Potencial", f"R$ {ganho_total:,.0f}".replace(",", "."))

    st.markdown("<br>", unsafe_allow_html=True)

    # ============================================================
    # ABAS
    # ============================================================
    tab1, tab2, tab3 = st.tabs(["📊 Visão Geral", "🔍 Análises", "💰 Financeiro"])

    # ----------------------------------------------------------
    # ABA 1: VISÃO GERAL
    # ----------------------------------------------------------
    with tab1:
        col_a, col_b = st.columns([2, 1])
        with col_a:
            fig_area = create_area_chart(df, col_names)
            st.plotly_chart(fig_area, use_container_width=True)
        with col_b:
            fig_tipo = create_pie_chart(df, col_media, "Distribuição por Tipo", "Tipo")
            st.plotly_chart(fig_tipo, use_container_width=True)

        col_c, col_d = st.columns(2)
        with col_c:
            fig_clientes = create_bar_chart(df, col_media, "Top 15 Clientes (Distintos)", "Cliente", 15, COLORS["cyan"])
            st.plotly_chart(fig_clientes, use_container_width=True)
        with col_d:
            usina_col = [c for c in df.columns if 'Usina' in c and 'Refer' in c]
            if usina_col:
                fig_usinas = create_bar_chart(df, col_media, "Top 10 Usinas", usina_col[0], 10, COLORS["emerald"])
                st.plotly_chart(fig_usinas, use_container_width=True)

        col_e, col_f = st.columns(2)
        with col_e:
            fig_esp = create_thickness_chart(df, col_media)
            if fig_esp:
                st.plotly_chart(fig_esp, use_container_width=True)
        with col_f:
            fig_proj = create_bar_chart(df, col_media, "Top 12 Projetos", "Projeto", 12, COLORS["purple"])
            st.plotly_chart(fig_proj, use_container_width=True)

        # Tabela Top 15 Bobinas
        st.markdown("### Top 15 Bobinas por Necessidade Média")
        codigo_col = [c for c in df.columns if 'Código' in c and 'Bobina' in c]
        if codigo_col:
            top15 = df.nlargest(15, col_media)[[codigo_col[0], 'Tipo', 'Cliente', 'Projeto', col_media]].copy()
            top15.columns = ['Código', 'Tipo', 'Cliente', 'Projeto', 'Necessidade Média (ton)']
            top15['Necessidade Média (ton)'] = top15['Necessidade Média (ton)'].round(1)
            top15 = top15.reset_index(drop=True)
            top15.index = top15.index + 1
            st.dataframe(top15, use_container_width=True, height=560)

    # ----------------------------------------------------------
    # ABA 2: ANÁLISES
    # ----------------------------------------------------------
    with tab2:
        fig_prog, df_u = create_progress_chart(df_formulas)
        st.plotly_chart(fig_prog, use_container_width=True)

        # Tabela de progresso
        st.markdown("### Progresso de Análise por Unidade")
        df_display = df_u.copy()
        df_display.columns = ['Unidade', 'Bobinas', 'Peso Total (ton)', 'Peso Analisado (ton)', '% Concluído', 'Ganho (R$)']
        df_display['Peso Total (ton)'] = df_display['Peso Total (ton)'].round(1)
        df_display['Peso Analisado (ton)'] = df_display['Peso Analisado (ton)'].round(1)
        df_display['% Concluído'] = df_display['% Concluído'].round(1)
        df_display['Ganho (R$)'] = df_display['Ganho (R$)'].apply(lambda x: f"R$ {x:,.0f}".replace(",", "."))
        st.dataframe(df_display, use_container_width=True, hide_index=True)

        # Distribuição por unidade
        st.markdown("### Necessidade por Unidade Delga")
        unidade_col = [c for c in df.columns if 'Unidade' in c and 'Delga' in c]
        if unidade_col:
            col_g, col_h = st.columns(2)
            with col_g:
                fig_unid = create_pie_chart(df, col_media, "Distribuição por Unidade", unidade_col[0], 5)
                st.plotly_chart(fig_unid, use_container_width=True)
            with col_h:
                benef_col = [c for c in df.columns if 'Beneficiador' in c]
                if benef_col:
                    fig_benef = create_bar_chart(df, col_media, "Necessidade por Beneficiador", benef_col[0], 10, COLORS["teal"])
                    st.plotly_chart(fig_benef, use_container_width=True)

    # ----------------------------------------------------------
    # ABA 3: FINANCEIRO
    # ----------------------------------------------------------
    with tab3:
        col_i, col_j = st.columns(2)
        with col_i:
            fig_ganho = create_ganho_pie(df_u)
            if fig_ganho:
                st.plotly_chart(fig_ganho, use_container_width=True)
        with col_j:
            # Ganho por usina
            usina_data = []
            for i in range(7, min(40, len(df_formulas))):
                row = df_formulas.iloc[i]
                if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() and str(row.iloc[0]) != 'Total':
                    ganho_val = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0
                    if ganho_val > 0:
                        usina_data.append({
                            'usina': str(row.iloc[0]),
                            'ganho': ganho_val,
                        })

            if usina_data:
                df_usina_ganho = pd.DataFrame(usina_data).sort_values('ganho', ascending=True)
                fig_usina_g = go.Figure(data=[go.Bar(
                    x=df_usina_ganho['ganho'],
                    y=df_usina_ganho['usina'],
                    orientation='h',
                    marker=dict(color=COLORS["amber"]),
                    hovertemplate='%{y}<br><b>R$ %{x:,.0f}</b><extra></extra>',
                )])
                fig_usina_g.update_layout(
                    **PLOTLY_LAYOUT,
                    title=dict(text="Ganho Financeiro por Usina", font=dict(size=16, color=COLORS["cyan"])),
                    xaxis_title="R$",
                    height=400,
                )
                st.plotly_chart(fig_usina_g, use_container_width=True)

        # Resumo financeiro
        st.markdown("### Resumo Financeiro por Unidade")
        if len(df_u) > 0:
            df_fin = df_u[['unidade', 'peso_total', 'peso_analisado', 'pct', 'ganho']].copy()
            df_fin.columns = ['Unidade', 'Peso Total (ton)', 'Peso Analisado (ton)', '% Concluído', 'Ganho (R$)']
            df_fin['Peso Total (ton)'] = df_fin['Peso Total (ton)'].round(1)
            df_fin['Peso Analisado (ton)'] = df_fin['Peso Analisado (ton)'].round(1)
            df_fin['% Concluído'] = df_fin['% Concluído'].round(1)
            df_fin['Ganho (R$)'] = df_fin['Ganho (R$)'].apply(lambda x: f"R$ {x:,.0f}".replace(",", "."))
            st.dataframe(df_fin, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
