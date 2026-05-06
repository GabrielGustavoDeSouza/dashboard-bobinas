"""
Dashboard de Controle de Matéria-Prima — Bobinas BSW
- Visitantes veem os dados automaticamente (sem upload)
- Admin atualiza os dados via upload protegido por senha
- Dados persistem no GitHub (não somem quando o app dorme)
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import io
import base64
from datetime import datetime

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Grupo Delga | Dashboard Bobinas BSW",
    page_icon="🔵",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# ESTILO CSS CUSTOMIZADO (identidade Grupo Delga)
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    .stApp { background-color: #080E1A; font-family: 'Inter', sans-serif; }
    header[data-testid="stHeader"] { background-color: #080E1A; }
    section[data-testid="stSidebar"] { background-color: #0C1425; border-right: 1px solid #1A2744; }
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #0F1A2E 0%, #132040 100%);
        border: 1px solid #1A2744;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    }
    div[data-testid="stMetric"] label { color: #8899B0 !important; font-weight: 500 !important; text-transform: uppercase; font-size: 11px !important; letter-spacing: 0.5px; }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #FFFFFF !important;
        font-family: 'Inter', sans-serif;
        font-weight: 700;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px; background-color: #0C1425; border-radius: 10px; padding: 4px; border: 1px solid #1A2744;
    }
    .stTabs [data-baseweb="tab"] { color: #8899B0; border-radius: 8px; font-weight: 500; }
    .stTabs [aria-selected="true"] { background-color: #1400FF !important; color: #FFFFFF !important; }
    h1, h2, h3 { color: #FFFFFF !important; font-family: 'Inter', sans-serif; font-weight: 700; }
    h4, h5, h6 { color: #8899B0 !important; font-family: 'Inter', sans-serif; font-weight: 600; }
    p, span, li { color: #B8C8DC; }
    .stDataFrame { border: 1px solid #1A2744; border-radius: 10px; overflow: hidden; }
    hr { border-color: #1A2744; }
    div[data-testid="stFileUploader"] {
        background-color: #0F1A2E; border: 2px dashed #1A2744; border-radius: 12px; padding: 16px;
    }
    div[data-testid="stFileUploader"]:hover { border-color: #1400FF; }
    div[data-testid="stSidebar"] .stRadio label { color: #8899B0 !important; }
    .stSelectbox label { color: #8899B0 !important; }
    .stButton > button[kind="primary"] { background-color: #1400FF !important; border: none; }
    .stButton > button[kind="primary"]:hover { background-color: #2010FF !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# CORES POR UNIDADE DELGA (padronizadas)
# ============================================================
UNIDADE_COLORS = {
    "Ferraz":  "#1E88E5",
    "Diadema": "#43A047",
    "Jarinu":  "#FB8C00",
    "Sul":     "#8E24AA",
}

COLORS = {
    "cyan": "#4DA3FF", "amber": "#FFB800", "emerald": "#00E676",
    "coral": "#FF6B6B", "purple": "#A78BFA", "teal": "#4DD0E1",
    "blue_delga": "#1400FF", "blue_light": "#4DA3FF",
    "bg_card": "#0F1A2E", "bg_dark": "#0C1425",
    "border": "#1A2744", "text_light": "#8899B0", "text_white": "#ECEFF1",
}

CHART_COLORS = [
    "#4DA3FF", "#FFB800", "#00E676", "#A78BFA", "#FF6B6B",
    "#4DD0E1", "#FFD54F", "#69F0AE", "#B39DDB", "#FF8A80",
    "#80DEEA", "#FFF176", "#A5D6A7", "#CE93D8", "#EF9A9A",
]

PLOTLY_LAYOUT = dict(
    paper_bgcolor="#0F1A2E",
    plot_bgcolor="#0F1A2E",
    font=dict(color="#8899B0", family="Inter, Arial"),
    margin=dict(l=40, r=40, t=50, b=40),
    legend=dict(
        bgcolor="rgba(15,26,46,0.9)", bordercolor="#1A2744",
        borderwidth=1, font=dict(color="#8899B0"),
    ),
)

# ============================================================
# CONFIGURAÇÃO DO GITHUB (para persistência de dados)
# ============================================================
GITHUB_REPO = "GabrielGustavoDeSouza/dashboard-bobinas"
GITHUB_DATA_PATH = "data/dados_atuais.xlsx"
GITHUB_BRANCH = "main"

ADMIN_PASSWORD = "M@ster"


def get_github_token():
    """Obtém o token do GitHub dos secrets do Streamlit."""
    try:
        return st.secrets["GITHUB_TOKEN"]
    except (KeyError, FileNotFoundError):
        return None


def get_unidade_color(nome):
    """Retorna a cor padronizada da unidade Delga."""
    for key, color in UNIDADE_COLORS.items():
        if key.lower() in str(nome).lower():
            return color
    return COLORS["cyan"]


def get_unidade_colors_list(names):
    """Retorna lista de cores para uma lista de nomes de unidades."""
    return [get_unidade_color(n) for n in names]


# ============================================================
# FUNÇÕES DE PERSISTÊNCIA (GitHub API)
# ============================================================
@st.cache_data(ttl=120)
def load_data_from_github():
    """Carrega o arquivo Excel salvo no repositório GitHub."""
    token = get_github_token()
    if not token:
        return None, None

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_DATA_PATH}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    params = {"ref": GITHUB_BRANCH}

    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 200:
        content = response.json()
        file_content = base64.b64decode(content["content"])
        excel_bytes = io.BytesIO(file_content)
        df_controle = smart_read_excel(excel_bytes, "Controle")
        excel_bytes.seek(0)
        df_formulas = pd.read_excel(excel_bytes, sheet_name="Formulas")
        return df_controle, df_formulas
    return None, None


def save_data_to_github(file_bytes, filename):
    """Salva o arquivo Excel no repositório GitHub (cria ou atualiza)."""
    token = get_github_token()
    if not token:
        return False, "Token do GitHub não configurado nos Secrets."

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_DATA_PATH}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
    }

    # Verificar se o arquivo já existe (para obter o SHA)
    response = requests.get(url, headers=headers, params={"ref": GITHUB_BRANCH})
    sha = None
    if response.status_code == 200:
        sha = response.json()["sha"]

    # Codificar o arquivo em base64
    content_b64 = base64.b64encode(file_bytes).decode("utf-8")

    # Montar o payload
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    payload = {
        "message": f"Atualização de dados: {filename} ({now})",
        "content": content_b64,
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha

    # Enviar
    response = requests.put(url, headers=headers, json=payload)
    if response.status_code in [200, 201]:
        # Limpar o cache para que os novos dados sejam carregados
        load_data_from_github.clear()
        return True, "Dados atualizados com sucesso!"
    else:
        return False, f"Erro ao salvar: {response.status_code} - {response.json().get('message', '')}"


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
    """Extrai dados estruturados da aba Formulas."""
    unidades = []
    usinas = []

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
            'unidade': nome, 'bobinas': bobinas, 'peso_total': peso_total,
            'peso_analisado': peso_analisado, 'pct': pct, 'ganho': ganho,
        })

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
                'usina': nome, 'bobinas': bobinas, 'peso': peso,
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
    meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai']
    keys = ['jan', 'fev', 'mar', 'abr', 'mai']
    valores = []
    for k in keys:
        if k in col_names:
            valores.append(round(df[col_names[k]].sum(), 1))
        else:
            valores.append(0)

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=meses, y=valores,
        fill='tozeroy', fillcolor='rgba(20,0,255,0.12)',
        line=dict(color=COLORS["cyan"], width=3),
        mode='lines+markers',
        marker=dict(size=10, color=COLORS["cyan"]),
        name='Necessidade (ton)',
        hovertemplate='%{x}/2026<br><b>%{y:,.0f} ton</b><extra></extra>',
    ))
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Evolução da Necessidade Mensal (ton)", font=dict(size=16, color=COLORS["cyan"])),
        yaxis=dict(title="Toneladas", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        height=400,
    )
    return fig


def create_unidade_pie_chart(df, col_media):
    """Gráfico de pizza por Unidade Delga com cores padronizadas."""
    unidade_col = [c for c in df.columns if 'Unidade' in c and 'Delga' in c]
    if not unidade_col:
        return None
    df_valid = df[df[unidade_col[0]].notna() & (df[unidade_col[0]].astype(str).str.strip() != '')].copy()
    if len(df_valid) == 0:
        return None
    dist = df_valid.groupby(unidade_col[0])[col_media].sum().sort_values(ascending=False)
    colors = get_unidade_colors_list(dist.index)
    fig = go.Figure(data=[go.Pie(
        labels=[str(x) for x in dist.index],
        values=dist.values.tolist(),
        hole=0.45,
        marker=dict(colors=colors),
        textinfo='percent+label',
        textfont=dict(size=12, color="#ECEFF1"),
        hovertemplate='%{label}<br><b>%{value:,.1f} ton</b><br>%{percent}<extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Distribuição por Unidade Delga", font=dict(size=16, color=COLORS["cyan"])),
        height=400,
    )
    return fig


def create_tipo_pie_chart(df, col_media):
    """Gráfico de pizza por Tipo de bobina."""
    tipo_col = [c for c in df.columns if c.strip() == 'Tipo']
    if not tipo_col:
        return None
    df_valid = df[df[tipo_col[0]].notna() & (df[tipo_col[0]].astype(str).str.strip() != '')].copy()
    if len(df_valid) == 0:
        return None
    dist = df_valid.groupby(tipo_col[0])[col_media].sum().sort_values(ascending=False).head(10)
    fig = go.Figure(data=[go.Pie(
        labels=[str(x) for x in dist.index],
        values=dist.values.tolist(),
        hole=0.45,
        marker=dict(colors=CHART_COLORS[:len(dist)]),
        textinfo='percent+label',
        textfont=dict(size=12, color="#ECEFF1"),
        hovertemplate='%{label}<br><b>%{value:,.1f} ton</b><br>%{percent}<extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Distribuição por Tipo de Bobina", font=dict(size=16, color=COLORS["cyan"])),
        height=400,
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
    labels = ['0-1', '1-2', '2-4', '4-6', '6-8', '8-10', '10-15', '15-20', '20+']
    df_temp['faixa'] = pd.cut(df_temp['esp_num'], bins=bins, labels=labels, right=True)
    dist = df_temp.groupby('faixa', observed=True)[col_media].sum().sort_index()
    dist = dist[dist > 0]
    if len(dist) == 0:
        return None
    fig = go.Figure(data=[go.Bar(
        x=[str(x) for x in dist.index],
        y=dist.values.tolist(),
        marker=dict(color=CHART_COLORS[:len(dist)]),
        hovertemplate='%{x} mm<br><b>%{y:,.1f} ton</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Distribuição por Faixa de Espessura (mm)", font=dict(size=16, color=COLORS["cyan"])),
        yaxis=dict(title="Toneladas", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(title="Espessura (mm)", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        height=400,
    )
    return fig


def create_progress_chart(df_unidades):
    """Gráfico de progresso de análise por unidade com cores padronizadas."""
    if len(df_unidades) == 0:
        return None
    unidades = df_unidades['unidade'].tolist()
    peso_total = df_unidades['peso_total'].tolist()
    peso_analisado = df_unidades['peso_analisado'].tolist()
    colors = get_unidade_colors_list(unidades)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Peso Total',
        x=unidades, y=peso_total,
        marker=dict(color=colors, opacity=0.4),
        hovertemplate='%{x}<br>Peso Total: <b>%{y:,.1f} ton</b><extra></extra>',
    ))
    fig.add_trace(go.Bar(
        name='Peso Analisado',
        x=unidades, y=peso_analisado,
        marker=dict(color=colors, opacity=1.0),
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
        return None
    df_sorted = df_usinas.nlargest(top_n, 'peso')
    df_sorted = df_sorted.sort_values('peso', ascending=True)
    fig = go.Figure(data=[go.Bar(
        x=df_sorted['peso'].tolist(),
        y=df_sorted['usina'].tolist(),
        orientation='h',
        marker=dict(color=COLORS["teal"]),
        hovertemplate='%{y}<br><b>%{x:,.1f} ton</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Top Usinas por Peso (ton)", font=dict(size=16, color=COLORS["cyan"])),
        height=max(400, min(top_n, len(df_sorted)) * 32),
        yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F", title="Toneladas"),
    )
    return fig


def create_bar_chart(df, col_media, title, group_col, top_n=15, color=None):
    """Gráfico de barras horizontal genérico."""
    df_valid = df[df[group_col].notna() & (df[group_col].astype(str).str.strip() != '')].copy()
    if len(df_valid) == 0:
        return None
    dist = df_valid.groupby(group_col)[col_media].sum().sort_values(ascending=True).tail(top_n)
    fig = go.Figure(data=[go.Bar(
        x=dist.values.tolist(),
        y=[str(x) for x in dist.index],
        orientation='h',
        marker=dict(color=color or COLORS["cyan"]),
        hovertemplate='%{y}<br><b>%{x:,.1f} ton</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text=title, font=dict(size=16, color=COLORS["cyan"])),
        height=max(400, min(top_n, len(dist)) * 32),
        yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F", title="Toneladas"),
    )
    return fig


def create_unidade_bar_chart(df, col_media):
    """Gráfico de barras por unidade Delga com cores padronizadas."""
    unidade_col = [c for c in df.columns if 'Unidade' in c and 'Delga' in c]
    if not unidade_col:
        return None
    df_valid = df[df[unidade_col[0]].notna() & (df[unidade_col[0]].astype(str).str.strip() != '')].copy()
    if len(df_valid) == 0:
        return None
    dist = df_valid.groupby(unidade_col[0])[col_media].sum().sort_values(ascending=True)
    colors = get_unidade_colors_list(dist.index)
    fig = go.Figure(data=[go.Bar(
        x=dist.values.tolist(),
        y=[str(x) for x in dist.index],
        orientation='h',
        marker=dict(color=colors),
        hovertemplate='%{y}<br><b>%{x:,.1f} ton</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Necessidade por Unidade Delga (ton)", font=dict(size=16, color=COLORS["cyan"])),
        height=350,
        yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        xaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F", title="Toneladas"),
    )
    return fig


def create_ganho_unidade_chart(df_unidades):
    """Gráfico de ganho financeiro por unidade com cores padronizadas."""
    if len(df_unidades) == 0:
        return None
    df_g = df_unidades[df_unidades['ganho'] > 0].copy()
    if len(df_g) == 0:
        return None
    df_g = df_g.sort_values('ganho', ascending=True)
    colors = get_unidade_colors_list(df_g['unidade'])
    fig = go.Figure(data=[go.Bar(
        x=df_g['ganho'].tolist(),
        y=df_g['unidade'].tolist(),
        orientation='h',
        marker=dict(color=colors),
        hovertemplate='%{y}<br><b>R$ %{x:,.0f}</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Ganho Financeiro por Unidade (R$)", font=dict(size=16, color=COLORS["cyan"])),
        xaxis=dict(title="R$", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        height=350,
    )
    return fig


def create_ganho_pie_chart(df_unidades):
    """Gráfico de pizza do ganho financeiro por unidade."""
    if len(df_unidades) == 0:
        return None
    df_g = df_unidades[df_unidades['ganho'] > 0].copy()
    if len(df_g) == 0:
        return None
    colors = get_unidade_colors_list(df_g['unidade'])
    fig = go.Figure(data=[go.Pie(
        labels=df_g['unidade'].tolist(),
        values=df_g['ganho'].tolist(),
        hole=0.45,
        marker=dict(colors=colors),
        textinfo='percent+label',
        textfont=dict(size=12, color="#ECEFF1"),
        hovertemplate='%{label}<br><b>R$ %{value:,.0f}</b><br>%{percent}<extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Ganho Financeiro por Unidade", font=dict(size=16, color=COLORS["cyan"])),
        height=400,
    )
    return fig


def create_ganho_usinas_chart(df_usinas):
    """Gráfico de ganho financeiro por usina."""
    if len(df_usinas) == 0:
        return None
    df_g = df_usinas[df_usinas['ganho'] > 0].copy()
    if len(df_g) == 0:
        return None
    df_g = df_g.sort_values('ganho', ascending=True)
    fig = go.Figure(data=[go.Bar(
        x=df_g['ganho'].tolist(),
        y=df_g['usina'].tolist(),
        orientation='h',
        marker=dict(color=COLORS["emerald"]),
        hovertemplate='%{y}<br><b>R$ %{x:,.0f}</b><extra></extra>',
    )])
    fig.update_layout(
        **PLOTLY_LAYOUT,
        title=dict(text="Ganho por Usina (R$)", font=dict(size=16, color=COLORS["cyan"])),
        xaxis=dict(title="R$", gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        yaxis=dict(gridcolor="#1E3A5F", zerolinecolor="#1E3A5F"),
        height=max(400, len(df_g) * 30),
    )
    return fig


# ============================================================
# HELPER: renderizar gráfico com theme=None
# ============================================================
def render_chart(fig):
    """Renderiza um gráfico Plotly no Streamlit com theme=None para evitar override de cores."""
    if fig is not None:
        st.plotly_chart(fig, use_container_width=True, theme=None)
        return True
    return False


# ============================================================
# APLICAÇÃO PRINCIPAL
# ============================================================
def main():
    # SIDEBAR
    with st.sidebar:
        st.image("logo_delga.png", use_container_width=True)
        st.markdown("""
        <div style="text-align:center; padding:8px 0 16px 0;">
            <p style="color:#8899B0; font-size:11px; margin:0; text-transform:uppercase; letter-spacing:1.5px; font-weight:600;">Controle de Matéria-Prima</p>
        </div>
        <hr style="border-color:#1A2744; margin:0 0 16px 0;">
        """, unsafe_allow_html=True)

        # ── ÁREA ADMIN (protegida por senha) ──
        st.markdown("#### 🔐 Área do Administrador")
        with st.expander("Atualizar Dados (requer senha)", expanded=False):
            senha = st.text_input("Senha:", type="password", key="admin_pwd")
            if senha == ADMIN_PASSWORD:
                st.success("Acesso liberado!")
                admin_file = st.file_uploader(
                    "Envie o Excel atualizado:",
                    type=["xlsx", "xls"],
                    key="admin_upload",
                    help="O arquivo será salvo e ficará disponível para todos os visitantes.",
                )
                if admin_file:
                    if st.button("📤 Salvar e Publicar Dados", type="primary"):
                        with st.spinner("Salvando dados no servidor..."):
                            file_bytes = admin_file.getvalue()
                            success, msg = save_data_to_github(file_bytes, admin_file.name)
                        if success:
                            st.success(f"✅ {msg}")
                            st.info("Os dados já estão disponíveis para todos os visitantes!")
                            st.balloons()
                        else:
                            st.error(f"❌ {msg}")
            elif senha and senha != ADMIN_PASSWORD:
                st.error("Senha incorreta.")

        st.markdown("---")

        # Legenda de cores das unidades
        st.markdown("#### Cores por Unidade")
        for unidade, cor in UNIDADE_COLORS.items():
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:8px;margin:4px 0;">'
                f'<div style="width:16px;height:16px;border-radius:4px;background:{cor};"></div>'
                f'<span style="color:#B0BEC5;font-size:13px;">{unidade}</span></div>',
                unsafe_allow_html=True,
            )

        st.markdown("---")
        st.markdown("""
        <div style="padding:10px; background:linear-gradient(135deg, #0F1A2E 0%, #132040 100%); border-radius:10px; border:1px solid #1A2744; margin-top:8px;">
            <p style="color:#5A7090; font-size:11px; margin:0; line-height:1.6;">
                <b style="color:#8899B0;">📋 Rotina:</b> Dados atualizados toda segunda-feira.<br>
                <b style="color:#8899B0;">👥 Visitantes:</b> Visualizam automaticamente os dados mais recentes.
            </p>
        </div>
        """, unsafe_allow_html=True)

    # HEADER
    st.markdown("""
    <div style="display:flex; align-items:center; gap:16px; margin-bottom:16px; padding-bottom:16px; border-bottom:1px solid #1A2744;">
        <div style="background:linear-gradient(135deg, #1400FF 0%, #0A00AA 100%); border-radius:12px; padding:14px; display:flex; align-items:center; justify-content:center; box-shadow:0 4px 16px rgba(20,0,255,0.3);">
            <span style="font-size:24px; color:white; font-weight:800; font-family:'Inter',sans-serif;">BSW</span>
        </div>
        <div>
            <h1 style="margin:0; font-size:26px; color:#FFFFFF !important; font-weight:800; letter-spacing:-0.5px;">Controle de Matéria-Prima</h1>
            <p style="margin:4px 0 0 0; color:#5A7090; font-size:12px; font-weight:500; letter-spacing:0.5px;">BOBINAS BSW — JAN A MAI / 2026 | GRUPO DELGA</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ============================================================
    # CARREGAR DADOS (prioridade: GitHub > SharePoint > vazio)
    # ============================================================
    df_raw = None
    df_formulas = None

    # Tentar carregar do GitHub (dados persistidos)
    try:
        df_raw, df_formulas = load_data_from_github()
    except Exception:
        df_raw, df_formulas = None, None

    # Se não tem dados no GitHub, tentar SharePoint
    if df_raw is None:
        try:
            df_raw, df_formulas = load_data_from_sharepoint()
        except Exception:
            df_raw, df_formulas = None, None

    # Se não tem dados de nenhuma fonte
    if df_raw is None:
        st.markdown("""
        <div style="text-align:center; padding:80px 20px; background:linear-gradient(135deg, #0F1A2E 0%, #132040 100%); border:2px dashed #1A2744; border-radius:16px; margin:40px auto; max-width:600px;">
            <span style="font-size:64px;">📊</span>
            <h2 style="color:#FFFFFF !important; margin:16px 0 8px 0;">Aguardando Dados</h2>
            <p style="color:#5A7090; font-size:14px;">
                Nenhum dado disponível ainda.<br><br>
                <b style="color:#8899B0;">Administrador:</b> Use a área "Atualizar Dados" no painel lateral<br>
                para enviar o arquivo Excel pela primeira vez.
            </p>
        </div>
        """, unsafe_allow_html=True)
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
    total_bobinas = int(df_unidades['bobinas'].sum()) if len(df_unidades) > 0 else 0
    total_peso = float(df_unidades['peso_total'].sum()) if len(df_unidades) > 0 else 0
    total_peso_analisado = float(df_unidades['peso_analisado'].sum()) if len(df_unidades) > 0 else 0
    total_ganho = float(df_unidades['ganho'].sum()) if len(df_unidades) > 0 else 0

    k1, k2, k3 = st.columns(3)
    with k1:
        st.metric("Peso Total (MP)", f"{total_peso:,.0f} ton".replace(",", "."))
    with k2:
        st.metric("Peso Analisado (MP)", f"{total_peso_analisado:,.0f} ton".replace(",", "."))
    with k3:
        st.metric("Ganho Potencial", f"R$ {total_ganho:,.0f}".replace(",", "."))

    # ============================================================
    # SELETOR DE UNIDADE
    # ============================================================
    if len(df_unidades) > 0:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### Detalhamento por Unidade")

        unidade_names = ["Todas"] + df_unidades['unidade'].tolist()
        selected_unidade = st.selectbox(
            "Selecione a unidade:", unidade_names, index=0, key="unidade_selector"
        )

        if selected_unidade == "Todas":
            u_bobinas = total_bobinas
            u_peso = total_peso
            u_analisado = total_peso_analisado
            u_ganho = total_ganho
            u_pct = (total_peso_analisado / total_peso * 100) if total_peso > 0 else 0
        else:
            row_u = df_unidades[df_unidades['unidade'] == selected_unidade].iloc[0]
            u_bobinas = int(row_u['bobinas'])
            u_peso = float(row_u['peso_total'])
            u_analisado = float(row_u['peso_analisado'])
            u_ganho = float(row_u['ganho'])
            u_pct = float(row_u['pct'])

        uk1, uk2, uk3, uk4 = st.columns(4)
        with uk1:
            st.metric("Peso Total (MP)", f"{u_peso:,.1f} ton".replace(",", "."))
        with uk2:
            st.metric("Peso Analisado (MP)", f"{u_analisado:,.1f} ton".replace(",", "."))
        with uk3:
            st.metric("% Concluído", f"{u_pct:.1f}%")
        with uk4:
            st.metric("Ganho (R$)", f"R$ {u_ganho:,.0f}".replace(",", "."))

    st.markdown("<br>", unsafe_allow_html=True)

    # ============================================================
    # ABAS
    # ============================================================
    tab1, tab2, tab3 = st.tabs(["📊 Visão Geral", "🔍 Análises", "💰 Financeiro"])

    # ── ABA 1: VISÃO GERAL ──
    with tab1:
        col_a, col_b = st.columns([2, 1])
        with col_a:
            render_chart(create_area_chart(df, col_names))
        with col_b:
            fig_tipo = create_tipo_pie_chart(df, col_media)
            if not render_chart(fig_tipo):
                st.info("Coluna 'Tipo' não encontrada.")

        col_c, col_d = st.columns(2)
        with col_c:
            fig_esp = create_thickness_chart(df, col_media)
            if not render_chart(fig_esp):
                st.info("Coluna de espessura não encontrada.")
        with col_d:
            fig_unid = create_unidade_pie_chart(df, col_media)
            if not render_chart(fig_unid):
                st.info("Coluna 'Unidade Delga' não encontrada.")

        # Usinas
        fig_usinas = create_usinas_chart(df_usinas, 15)
        if not render_chart(fig_usinas):
            st.info("Dados de usinas não encontrados na aba Formulas.")

        # Tabela Top 15 Bobinas
        st.markdown("### Top 15 Bobinas por Necessidade Média")
        codigo_col = [c for c in df.columns if 'Código' in c and 'Bobina' in c]
        if codigo_col:
            display_cols = [codigo_col[0]]
            tipo_col_list = [c for c in df.columns if c.strip() == 'Tipo']
            if tipo_col_list:
                display_cols.append(tipo_col_list[0])
            unidade_col_list = [c for c in df.columns if 'Unidade' in c and 'Delga' in c]
            if unidade_col_list:
                display_cols.append(unidade_col_list[0])
            if 'Projeto' in df.columns:
                display_cols.append('Projeto')
            display_cols.append(col_media)

            top15 = df.nlargest(15, col_media)[display_cols].copy()
            rename_map = {codigo_col[0]: 'Código Bobina', col_media: 'Necessidade Média (ton)'}
            if unidade_col_list:
                rename_map[unidade_col_list[0]] = 'Unidade'
            top15 = top15.rename(columns=rename_map)
            top15['Necessidade Média (ton)'] = top15['Necessidade Média (ton)'].round(1)
            top15 = top15.reset_index(drop=True)
            top15.index = top15.index + 1
            st.dataframe(top15, use_container_width=True, height=560)

    # ── ABA 2: ANÁLISES ──
    with tab2:
        if len(df_unidades) > 0:
            fig_prog = create_progress_chart(df_unidades)
            if not render_chart(fig_prog):
                st.info("Sem dados de progresso.")

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

        st.markdown("### Necessidade por Unidade e Beneficiador")
        col_g, col_h = st.columns(2)
        with col_g:
            fig_unid2 = create_unidade_bar_chart(df, col_media)
            if not render_chart(fig_unid2):
                st.info("Coluna 'Unidade Delga' não encontrada.")
        with col_h:
            benef_col = [c for c in df.columns if 'Beneficiador' in c]
            if benef_col:
                fig_benef = create_bar_chart(df, col_media, "Necessidade por Beneficiador (ton)", benef_col[0], 10, COLORS["teal"])
                if not render_chart(fig_benef):
                    st.info("Sem dados de beneficiador.")
            else:
                st.info("Coluna 'Beneficiador' não encontrada.")

        abc_col = [c for c in df.columns if c.strip().upper() == 'ABC']
        if abc_col:
            st.markdown("### Classificação ABC")
            df_abc = df[df[abc_col[0]].notna() & (df[abc_col[0]].astype(str).str.strip() != '')].copy()
            if len(df_abc) > 0:
                abc_dist = df_abc.groupby(abc_col[0])[col_media].agg(['sum', 'count']).sort_values('sum', ascending=False)
                abc_dist.columns = ['Necessidade Total (ton)', 'Qtd Bobinas']
                abc_dist['Necessidade Total (ton)'] = abc_dist['Necessidade Total (ton)'].round(1)
                st.dataframe(abc_dist, use_container_width=True)

    # ── ABA 3: FINANCEIRO ──
    with tab3:
        if len(df_unidades) > 0:
            has_ganho = df_unidades['ganho'].sum() > 0

            if has_ganho:
                col_i, col_j = st.columns(2)
                with col_i:
                    render_chart(create_ganho_pie_chart(df_unidades))
                with col_j:
                    render_chart(create_ganho_unidade_chart(df_unidades))

                if len(df_usinas) > 0 and df_usinas['ganho'].sum() > 0:
                    st.markdown("### Ganho Financeiro por Usina")
                    render_chart(create_ganho_usinas_chart(df_usinas))
            else:
                st.info(
                    "Nenhum ganho financeiro registrado ainda. "
                    "Os dados aparecerão conforme as análises forem concluídas na planilha."
                )

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
