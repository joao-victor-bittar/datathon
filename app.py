import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Dashboard Passos Mágicos | PEDE 2022–2024",
    page_icon="🪄",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    </style>
""", unsafe_allow_html=True)

# Lista oficial e ordenada das fases conforme a estrutura da Passos Mágicos
FASES_OFICIAIS = [
    'Alfa', 
    'Fase 1', 
    'Fase 2', 
    'Fase 3', 
    'Fase 4', 
    'Fase 5', 
    'Fase 6', 
    'Fase 7', 
    'Fase 8'
]

# ==============================================================================
# 2. PIPELINE DE CARREGAMENTO E TRATAMENTO DOS DADOS (CACHED)
# ==============================================================================
@st.cache_data
def load_and_process_data():
    BASE_DIR = Path(__file__).resolve().parent
    file_path = BASE_DIR / "BASE DE DADOS PEDE 2024 - DATATHON.xlsx"
    
    if not file_path.exists():
        file_path = Path("BASE DE DADOS PEDE 2024 - DATATHON.xlsx")

    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo '{file_path.name}': {e}")
        st.stop()

    frames = []

    map_years = {
        'PEDE2022': 2022,
        'PEDE2023': 2023,
        'PEDE2024': 2024
    }

    for sheet, year in map_years.items():
        if sheet in xls.sheet_names:
            df_temp = pd.read_excel(xls, sheet_name=sheet)
            df_temp.columns = df_temp.columns.astype(str).str.strip()
            
            mapped_targets = set()
            rename_dict = {}

            for col in df_temp.columns:
                col_upper = col.upper()

                if ('PONTO' in col_upper or 'VIRADA' in col_upper) and 'Atingiu_PV' not in mapped_targets:
                    rename_dict[col] = 'Atingiu_PV'
                    mapped_targets.add('Atingiu_PV')
                elif 'DEFASAGEM' in col_upper and 'Defasagem' not in mapped_targets:
                    rename_dict[col] = 'Defasagem'
                    mapped_targets.add('Defasagem')
                elif 'PEDRA' in col_upper and 'Pedra' not in mapped_targets:
                    rename_dict[col] = 'Pedra'
                    mapped_targets.add('Pedra')
                elif 'FASE' in col_upper and 'Fase' not in mapped_targets:
                    rename_dict[col] = 'Fase'
                    mapped_targets.add('Fase')
                elif 'INDE' in col_upper and 'INDE' not in mapped_targets:
                    rename_dict[col] = 'INDE'
                    mapped_targets.add('INDE')
                elif 'IDA' in col_upper and 'IDA' not in mapped_targets:
                    rename_dict[col] = 'IDA'
                    mapped_targets.add('IDA')
                elif 'IEG' in col_upper and 'IEG' not in mapped_targets:
                    rename_dict[col] = 'IEG'
                    mapped_targets.add('IEG')
                elif 'IPV' in col_upper and 'IPV' not in mapped_targets:
                    rename_dict[col] = 'IPV'
                    mapped_targets.add('IPV')
                elif 'IAN' in col_upper and 'IAN' not in mapped_targets:
                    rename_dict[col] = 'IAN'
                    mapped_targets.add('IAN')

            df_temp = df_temp.rename(columns=rename_dict)
            df_temp['Ano'] = year

            df_temp = df_temp.loc[:, ~df_temp.columns.duplicated()]

            target_cols = ['RA', 'Nome', 'Turma', 'Fase', 'Pedra', 'Ano', 'INDE', 'IAN', 'IDA', 
                           'IEG', 'IAA', 'IPS', 'IPP', 'IPV', 'Defasagem', 'Atingiu_PV']
            
            for c in target_cols:
                if c not in df_temp.columns:
                    df_temp[c] = None

            frames.append(df_temp[target_cols])

    df_pede = pd.concat(frames, ignore_index=True)

    # Padronização e normalização da coluna 'Fase'
    df_pede['Fase'] = df_pede['Fase'].astype(str).str.strip()
    
    # Mapeamento para garantir formato padronizado ("Alfa", "Fase 1", "Fase 2", etc.)
    def padronizar_fase(fase_str):
        f = fase_str.lower()
        if 'alfa' in f or f == '0':
            return 'Alfa'
        for i in range(1, 9):
            if f == str(i) or f == f'fase {i}' or f == f'fase_{i}':
                return f'Fase {i}'
        return fase_str

    df_pede['Fase'] = df_pede['Fase'].apply(padronizar_fase)

    # Filtrar estritamente apenas as Fases Oficiais (Alfa até Fase 8)
    df_pede = df_pede[df_pede['Fase'].isin(FASES_OFICIAIS)].copy()

    # Converter coluna 'Fase' para Categórica para manter a ordem lógica (Alfa -> Fase 1 -> ... -> Fase 8)
    df_pede['Fase'] = pd.Categorical(df_pede['Fase'], categories=FASES_OFICIAIS, ordered=True)

    # Limpeza de Pedra
    df_pede['Pedra'] = df_pede['Pedra'].fillna('Não Informado').astype(str).str.strip()

    numeric_cols = ['INDE', 'IAN', 'IDA', 'IEG', 'IAA', 'IPS', 'IPP', 'IPV', 'Defasagem']
    for c in numeric_cols:
        df_pede[c] = pd.to_numeric(df_pede[c], errors='coerce')

    def classificar_defasagem(val):
        if pd.isna(val): return "Sem Informação"
        if val <= -3: return "Defasagem Severa (≤ -3)"
        elif -2 <= val <= -1: return "Defasagem Moderada (-2 a -1)"
        elif val == 0: return "Nível Adequado (0)"
        else: return "Avançado (> 0)"

    df_pede['Status_Defasagem'] = df_pede['Defasagem'].apply(classificar_defasagem)
    
    return df_pede

df = load_and_process_data()

# ==============================================================================
# 3. FILTROS NA BARRA LATERAL (SIDEBAR)
# ==============================================================================
st.sidebar.image("https://passosmagicos.org.br/wp-content/uploads/2020/10/logo-passos-magicos.png", width=180)
st.sidebar.title("Filtros Interativos")

anos_disponiveis = sorted(df['Ano'].unique())
selected_years = st.sidebar.multiselect("Anos Letivos", anos_disponiveis, default=anos_disponiveis)

# Garante que as fases fiquem ordenadas do Alfa ao 8 na sidebar
fases_disponiveis = [f for f in FASES_OFICIAIS if f in df['Fase'].unique()]
selected_fases = st.sidebar.multiselect("Fases Pedagógicas", fases_disponiveis, default=fases_disponiveis)

pedras_disponiveis = sorted([p for p in df['Pedra'].unique() if p not in ['', 'nan', 'None']], key=str)
selected_pedras = st.sidebar.multiselect("Classificação (Pedra)", pedras_disponiveis, default=pedras_disponiveis)

df_filtered = df[
    (df['Ano'].isin(selected_years)) &
    (df['Fase'].isin(selected_fases)) &
    (df['Pedra'].isin(selected_pedras))
]

# ==============================================================================
# 4. PAINEL PRINCIPAL & KPIS EXECUTIVOS
# ==============================================================================
st.title("🪄 Associação Passos Mágicos")
st.markdown("### Dashboard Executivo de Impacto e Desempenho Educacional (PEDE 2022–2024)")
st.markdown("---")

col_kpi1, col_kpi2, col_kpi3, col_kpi4, col_kpi5 = st.columns(5)

with col_kpi1:
    st.metric("Total de Alunos/Anos", f"{len(df_filtered):,}".replace(",", "."))

with col_kpi2:
    media_inde = df_filtered['INDE'].mean()
    st.metric("Média INDE", f"{media_inde:.2f}" if pd.notna(media_inde) else "N/A")

with col_kpi3:
    media_ida = df_filtered['IDA'].mean()
    st.metric("Média Acadêmica (IDA)", f"{media_ida:.2f}" if pd.notna(media_ida) else "N/A")

with col_kpi4:
    media_ieg = df_filtered['IEG'].mean()
    st.metric("Média Engajamento (IEG)", f"{media_ieg:.2f}" if pd.notna(media_ieg) else "N/A")

with col_kpi5:
    defasagem_severa = (df_filtered['Defasagem'] <= -3).sum()
    pct_severa = (defasagem_severa / len(df_filtered) * 100) if len(df_filtered) > 0 else 0
    st.metric("Defasagem Severa (≤-3)", f"{pct_severa:.1f}%")

st.markdown("---")

# ==============================================================================
# 5. ABAS ANÁLITICAS
# ==============================================================================
tab_visao_geral, tab_defasagem, tab_desempenho, tab_engajamento = st.tabs([
    "📊 Visão Geral & Evolução", 
    "📈 Evolução da Defasagem (IAN)", 
    "📚 Desempenho Acadêmico (IDA)", 
    "⚡ Engajamento & Ponto de Virada"
])

# ------------------------------------------------------------------------------
# ABA 1: VISÃO GERAL
# ------------------------------------------------------------------------------
with tab_visao_geral:
    st.subheader("Evolução Histórica dos Indicadores Globais")
    
    df_evolucao = df_filtered.groupby('Ano')[['INDE', 'IDA', 'IEG', 'IPV']].mean().reset_index()
    
    fig_evol = px.line(
        df_evolucao, 
        x='Ano', 
        y=['INDE', 'IDA', 'IEG', 'IPV'],
        markers=True,
        title="Média Anual dos Indicadores Globais",
        labels={'value': 'Nota Média', 'variable': 'Indicador'},
        color_discrete_sequence=px.colors.qualitative.Set2
    )
    fig_evol.update_xaxes(dtick=1)
    fig_evol.update_layout(hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig_evol, use_container_width=True)

    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        fig_pedra = px.histogram(
            df_filtered, 
            x='Pedra', 
            color='Ano', 
            barmode='group',
            title="Distribuição de Alunos por Classificação (Pedra)",
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        st.plotly_chart(fig_pedra, use_container_width=True)

    with col_g2:
        fig_fase = px.histogram(
            df_filtered, 
            x='Fase', 
            color='Ano', 
            barmode='group',
            title="Distribuição por Fase Pedagógica (Alfa a Fase 8)",
            color_discrete_sequence=px.colors.qualitative.Safe,
            category_orders={'Fase': FASES_OFICIAIS}
        )
        st.plotly_chart(fig_fase, use_container_width=True)

# ------------------------------------------------------------------------------
# ABA 2: EVOLUÇÃO DA DEFASAGEM (IAN)
# ------------------------------------------------------------------------------
with tab_defasagem:
    st.subheader("Redução da Defasagem Idade-Série (IAN)")
    st.caption("Evolução da distribuição dos alunos por nível de defasagem pedagógica ao longo dos anos.")

    df_def_ano = df_filtered.groupby(['Ano', 'Status_Defasagem']).size().reset_index(name='Quantidade')
    df_def_ano['Proporção (%)'] = df_def_ano.groupby('Ano')['Quantidade'].transform(lambda x: (x / x.sum()) * 100)

    fig_def = px.bar(
        df_def_ano, 
        x='Ano', 
        y='Proporção (%)', 
        color='Status_Defasagem',
        title="Proporção Relativa do Status de Defasagem por Ano (%)",
        text_auto='.1f',
        color_discrete_map={
            "Defasagem Severa (≤ -3)": "#d9534f",
            "Defasagem Moderada (-2 a -1)": "#f0ad4e",
            "Nível Adequado (0)": "#5bc0de",
            "Avançado (> 0)": "#5cb85c"
        }
    )
    fig_def.update_xaxes(dtick=1)
    st.plotly_chart(fig_def, use_container_width=True)

# ------------------------------------------------------------------------------
# ABA 3: DESEMPENHO ACADÊMICO (IDA)
# ------------------------------------------------------------------------------
with tab_desempenho:
    st.subheader("Análise de Transição e Desempenho Acadêmico por Fase")
    
    df_fase_ida = df_filtered.groupby('Fase', observed=False)['IDA'].mean().reset_index()
    fig_curva = px.line(
        df_fase_ida, 
        x='Fase', 
        y='IDA', 
        markers=True,
        title="Média do Índice Acadêmico (IDA) por Fase Pedagógica (Alfa até Fase 8)",
        labels={'IDA': 'Média IDA', 'Fase': 'Fase Pedagógica'},
        category_orders={'Fase': FASES_OFICIAIS}
    )
    fig_curva.update_traces(line_color='#e74c3c', line_width=3)
    st.plotly_chart(fig_curva, use_container_width=True)

# ------------------------------------------------------------------------------
# ABA 4: ENGAJAMENTO & PONTO DE VIRADA
# ------------------------------------------------------------------------------
with tab_engajamento:
    st.subheader("Relação entre Engajamento (IEG) e Resultados (IDA)")

    col_e1, col_e2 = st.columns(2)

    with col_e1:
        fig_scat = px.scatter(
            df_filtered, 
            x='IEG', 
            y='IDA', 
            color='Pedra',
            hover_data=['Nome', 'Fase', 'Ano'],
            title="Dispersão: Engajamento vs. Desempenho Acadêmico",
            trendline="ols",
            opacity=0.7
        )
        st.plotly_chart(fig_scat, use_container_width=True)

    with col_e2:
        df_pv = df_filtered[df_filtered['Atingiu_PV'].notna()]
        if not df_pv.empty:
            fig_pv = px.box(
                df_pv, 
                x='Atingiu_PV', 
                y='IEG', 
                color='Atingiu_PV',
                title="Distribuição do Engajamento (IEG) por Atingimento do Ponto de Virada",
                labels={'Atingiu_PV': 'Atingiu Ponto de Virada?', 'IEG': 'IEG'}
            )
            st.plotly_chart(fig_pv, use_container_width=True)
        else:
            st.info("Para visualizar a comparação oficial do Ponto de Virada, certifique-se de incluir o ano de 2022 nos filtros laterais.")

# ==============================================================================
# 6. RODAPÉ
# ==============================================================================
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #7f8c8d;'>"
    "Dashboard Passos Mágicos | Dataviz & Analytics"
    "</div>", 
    unsafe_allow_html=True
)
