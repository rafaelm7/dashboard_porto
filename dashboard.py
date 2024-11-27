import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
import os
from pathlib import Path

# Configuração da página
st.set_page_config(page_title="Dashboard Comércio Exterior", layout="wide", initial_sidebar_state="expanded")
st.title("Dashboard de Análise de Comércio Exterior")

# Caminho do banco de dados relativo ao diretório do script
DB_PATH = Path(__file__).parent / "comercio_exterior.sqlite"

# Verificação da existência do arquivo
if not os.path.exists(DB_PATH):
    st.error("Arquivo do banco de dados não encontrado!")
    st.stop()

# Conexão com o banco de dados
@st.cache_data
def carregar_dados():
    conn = sqlite3.connect(DB_PATH)
    query = """
    SELECT Fluxo, Ano, Países, "UF do Produto" as UF, URF, 
           "Código Seção" as Cod_Secao, "Descrição Seção" as Desc_Secao,
           Via, "Código SH6" as Cod_SH6, "Descrição SH6" as Desc_SH6,
           "Valor US$ FOB" as Valor_FOB
    FROM comercio_exterior
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

# Carregando os dados
try:
    df = carregar_dados()
except Exception as e:
    st.error(f"Erro ao carregar o banco de dados: {e}")
    st.stop()

# Sidebar para filtros
st.sidebar.header("Filtros")

# Filtros principais
col1_side, col2_side = st.sidebar.columns(2)

with col1_side:
    anos_selecionados = st.multiselect(
        "Ano",
        options=sorted(df['Ano'].unique()),
        default=df['Ano'].max()
    )

with col2_side:
    fluxos_selecionados = st.multiselect(
        "Fluxo",
        options=sorted(df['Fluxo'].unique()),
        default=df['Fluxo'].unique()
    )

paises_selecionados = st.sidebar.multiselect(
    "Países",
    options=sorted(df['Países'].unique()),
    default=[]
)

ufs_selecionadas = st.sidebar.multiselect(
    "UF do Produto",
    options=sorted(df['UF'].unique()),
    default=[]
)

# Novo filtro URF
urf_selecionadas = st.sidebar.multiselect(
    "URF",
    options=sorted(df['URF'].unique()),
    default=[]
)

secoes_selecionadas = st.sidebar.multiselect(
    "Seção",
    options=sorted(df['Desc_Secao'].unique()),
    default=[]
)

# Aplicando filtros
df_filtrado = df.copy()

if anos_selecionados:
    df_filtrado = df_filtrado[df_filtrado['Ano'].isin(anos_selecionados)]
if fluxos_selecionados:
    df_filtrado = df_filtrado[df_filtrado['Fluxo'].isin(fluxos_selecionados)]
if paises_selecionados:
    df_filtrado = df_filtrado[df_filtrado['Países'].isin(paises_selecionados)]
if ufs_selecionadas:
    df_filtrado = df_filtrado[df_filtrado['UF'].isin(ufs_selecionadas)]
if urf_selecionadas:
    df_filtrado = df_filtrado[df_filtrado['URF'].isin(urf_selecionadas)]
if secoes_selecionadas:
    df_filtrado = df_filtrado[df_filtrado['Desc_Secao'].isin(secoes_selecionadas)]

# Métricas principais
st.subheader("Métricas Principais")
col1, col2, col3, col4 = st.columns(4)

with col1:
    valor_total = df_filtrado['Valor_FOB'].sum()
    st.metric("Valor Total FOB (USD)", f"${valor_total:,.2f}")

with col2:
    n_paises = df_filtrado['Países'].nunique()
    st.metric("Número de Países", f"{n_paises:,}")

with col3:
    n_produtos = df_filtrado['Cod_SH6'].nunique()
    st.metric("Número de Produtos", f"{n_produtos:,}")

with col4:
    n_ufs = df_filtrado['UF'].nunique()
    st.metric("Número de UFs", f"{n_ufs:,}")

# Visualizações
st.subheader("Análises Gráficas")

# Tab para diferentes visualizações
tab1, tab2, tab3 = st.tabs(["Análise Temporal", "Análise Geográfica", "Análise por Produto"])

with tab1:
    # Gráfico de evolução temporal
    fig_temporal = px.line(
        df_filtrado.groupby(['Ano', 'Fluxo'])['Valor_FOB'].sum().reset_index(),
        x='Ano',
        y='Valor_FOB',
        color='Fluxo',
        title="Evolução do Valor FOB por Ano e Fluxo"
    )
    st.plotly_chart(fig_temporal, use_container_width=True)

with tab2:
    col1, col2 = st.columns(2)
    
    with col1:
        # Top 10 países
        top_paises = df_filtrado.groupby('Países')['Valor_FOB'].sum().sort_values(ascending=True).tail(10)
        fig_paises = px.bar(
            top_paises,
            orientation='h',
            title="Top 10 Países por Valor FOB"
        )
        st.plotly_chart(fig_paises, use_container_width=True)
    
    with col2:
        # Distribuição por UF
        fig_ufs = px.bar(
            df_filtrado.groupby('UF')['Valor_FOB'].sum().sort_values(ascending=False).reset_index(),
            x='UF',
            y='Valor_FOB',
            title="Valor FOB por UF"
        )
        st.plotly_chart(fig_ufs, use_container_width=True)

with tab3:
    col1, col2 = st.columns(2)
    
    with col1:
        # Top 10 seções
        top_secoes = df_filtrado.groupby('Desc_Secao')['Valor_FOB'].sum().sort_values(ascending=True).tail(10)
        fig_secoes = px.bar(
            top_secoes,
            orientation='h',
            title="Top 10 Seções por Valor FOB"
        )
        st.plotly_chart(fig_secoes, use_container_width=True)
    
    with col2:
        # Top 10 produtos (SH6)
        top_produtos = df_filtrado.groupby('Desc_SH6')['Valor_FOB'].sum().sort_values(ascending=False).head(10)
        fig_produtos = px.bar(
            top_produtos,
            orientation='h',
            title="Top 10 Produtos (SH6) por Valor FOB"
        )
        st.plotly_chart(fig_produtos, use_container_width=True)

# Tabela de dados
st.subheader("Dados Detalhados")
st.dataframe(
    df_filtrado.sort_values('Valor_FOB', ascending=False),
    hide_index=True
)

# Download dos dados filtrados
if st.button("Download dos dados filtrados (CSV)"):
    csv = df_filtrado.to_csv(index=False)
    st.download_button(
        label="Clique para download",
        data=csv,
        file_name="comercio_exterior_filtrado.csv",
        mime="text/csv"
    )