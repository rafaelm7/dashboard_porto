import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
import os
from pathlib import Path
import io
import plotly.graph_objects as go
import pycountry
from unidecode import unidecode
import streamlit.components.v1 as components
import json

@st.cache_data(ttl=3600)  # Cache por 1 hora
def criar_mapa_cores_produtos(produtos):
    """
    Cria um mapeamento fixo de cores para todos os produtos,
    garantindo que cada produto tenha uma cor única.
    
    Args:
        produtos (list): Lista de produtos únicos
        
    Returns:
        dict: Dicionário com produtos e suas cores correspondentes
    """
    # Combinar paletas de cores mais eficientemente
    paletas_base = (
        px.colors.qualitative.Set3 +
        px.colors.qualitative.Pastel +
        px.colors.qualitative.Safe +
        px.colors.qualitative.Plotly +
        px.colors.qualitative.D3
    )
    
    n_cores_necessarias = len(produtos)
    
    # Se as paletas base são suficientes, usar diretamente
    if len(paletas_base) >= n_cores_necessarias:
        cores_finais = paletas_base[:n_cores_necessarias]
    else:
        # Gerar cores adicionais de forma mais eficiente
        cores_adicionais = []
        for i in range(n_cores_necessarias - len(paletas_base)):
            h = (i * 137.508) % 360  # Número áureo para distribuição
            s = 0.7 + (i % 3) * 0.1  # Saturação entre 70% e 90%
            l = 0.45 + (i % 5) * 0.05  # Luminosidade entre 45% e 65%
            
            # Converter HSL para RGB e depois para hex
            import colorsys
            rgb = colorsys.hls_to_rgb(h/360, l, s)
            cor_hex = '#{:02x}{:02x}{:02x}'.format(
                int(rgb[0]*255),
                int(rgb[1]*255),
                int(rgb[2]*255)
            )
            cores_adicionais.append(cor_hex)
        
        cores_finais = paletas_base + cores_adicionais
    
    # Criar o mapeamento
    return dict(zip(produtos, cores_finais))

def aplicar_filtros(df, filtros):
    """
    Aplica múltiplos filtros ao DataFrame de forma otimizada.
    
    Args:
        df (pd.DataFrame): DataFrame a ser filtrado
        filtros (dict): Dicionário com colunas e valores para filtrar
        
    Returns:
        pd.DataFrame: DataFrame filtrado
    """
    mask = pd.Series(True, index=df.index)
    for coluna, valores in filtros.items():
        if valores:  # Só aplica o filtro se houver valores selecionados
            mask &= df[coluna].isin(valores)
    return df[mask]

def format_big_number(value):
    """Formata números grandes para usar K, M e B"""
    suffixes = {1e9: 'B', 1e6: 'M', 1e3: 'K'}
    for size, suffix in suffixes.items():
        if abs(value) >= size:
            return f"{value/size:.1f}{suffix}"
    return f"{value:.1f}"

def format_currency(value):
    """Formata valores monetários em K, M ou B"""
    suffixes = {1e9: 'B', 1e6: 'M', 1e3: 'K'}
    for size, suffix in suffixes.items():
        if value >= size:
            return f'${value/size:.2f}{suffix}'
    return f'${value:.0f}'

def criar_mapa_paises():
    """Cria um dicionário de mapeamento para nomes de países PT-BR -> EN"""
    paises_map = {}
    for country in pycountry.countries:
        nome_pt = unidecode(country.name.lower())
        paises_map[nome_pt] = country.name
    return paises_map

# Função auxiliar para plotly_chart com retorno de seleção
def plotly_chart(fig, use_container_width=True, key=None):
    """Versão otimizada da função plotly_chart"""
    fig.update_layout(
        updatemenus=[dict(
            type="buttons",
            showactive=False,
            buttons=[dict(
                label="Reset Selection",
                method="update",
                args=[{"selectedpoints": None}]
            )]
        )],
        template='plotly_dark',
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        hoverlabel=dict(
            bgcolor='white',
            font_color='black',
            font_size=12
        )
    )
    
    components.html(
        fig.to_html(
            include_plotlyjs=True,
            config={'displayModeBar': True}
        ),
        height=600,
        width=None if use_container_width else 800
    )
    
    # Retornar dados da seleção via session_state
    return st.session_state.get('plotly_selected_data', None)

# Adicionar handler para mensagens do JavaScript
if 'plotly_selected_data' not in st.session_state:
    st.session_state.plotly_selected_data = None

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
@st.cache_data(ttl=3600)  # Cache por 1 hora
def carregar_dados():
    query = """
    SELECT 
        Fluxo, Ano, Países, "UF do Produto" as UF, URF, 
        "Código Seção" as Cod_Secao, "Descrição Seção" as Desc_Secao,
        Via, "Código SH6" as Cod_SH6, "Descrição SH6" as Desc_SH6,
        "Valor US$ FOB" as Valor_FOB
    FROM comercio_exterior
    """
    with sqlite3.connect(DB_PATH) as conn:
        return pd.read_sql_query(query, conn)

# Carregando os dados
try:
    df = carregar_dados()
except Exception as e:
    st.error(f"Erro ao carregar o banco de dados: {e}")
    st.stop()

# Criar o mapeamento de cores uma única vez
MAPA_CORES_PRODUTOS = criar_mapa_cores_produtos(sorted(df['Desc_SH6'].unique()))

# Antes dos filtros, adicionar um container para armazenar os filtros selecionados
if 'filtros_ativos' not in st.session_state:
    st.session_state.filtros_ativos = {
        'Ano': [],
        'Fluxo': [],
        'Países': [],
        'UF': [],
        'URF': [],
        'Desc_Secao': [],
        'Desc_SH6': []
    }

# Sidebar para filtros
st.sidebar.header("Filtros")

# Filtros principais
col1_side, col2_side = st.sidebar.columns(2)

with col1_side:
    anos_selecionados = st.multiselect(
        "Ano",
        options=sorted(df['Ano'].unique()),
        default=st.session_state.filtros_ativos['Ano']
    )

with col2_side:
    fluxos_selecionados = st.multiselect(
        "Fluxo",
        options=sorted(df['Fluxo'].unique()),
        default=st.session_state.filtros_ativos['Fluxo']
    )

paises_selecionados = st.sidebar.multiselect(
    "Países",
    options=sorted(df['Países'].unique()),
    default=st.session_state.filtros_ativos['Países']
)

ufs_selecionadas = st.sidebar.multiselect(
    "UF do Produto",
    options=sorted(df['UF'].unique()),
    default=st.session_state.filtros_ativos['UF']
)

urf_selecionadas = st.sidebar.multiselect(
    "URF",
    options=sorted(df['URF'].unique()),
    default=st.session_state.filtros_ativos['URF']
)

secoes_selecionadas = st.sidebar.multiselect(
    "Seção",
    options=sorted(df['Desc_Secao'].unique()),
    default=st.session_state.filtros_ativos['Desc_Secao']
)

sh6_selecionados = st.sidebar.multiselect(
    "Produto (SH6)",
    options=sorted(df['Desc_SH6'].unique()),
    default=st.session_state.filtros_ativos['Desc_SH6']
)

# Botão para aplicar filtros
if st.sidebar.button('Aplicar Filtros', type='primary'):
    # Atualizar filtros ativos
    st.session_state.filtros_ativos = {
        'Ano': anos_selecionados,
        'Fluxo': fluxos_selecionados,
        'Países': paises_selecionados,
        'UF': ufs_selecionadas,
        'URF': urf_selecionadas,
        'Desc_Secao': secoes_selecionadas,
        'Desc_SH6': sh6_selecionados
    }
    # Forçar rerun para atualizar a visualização
    st.rerun()

# Botão para limpar filtros
if st.sidebar.button('Limpar Filtros'):
    # Resetar filtros ativos
    st.session_state.filtros_ativos = {
        'Ano': [],
        'Fluxo': [],
        'Países': [],
        'UF': [],
        'URF': [],
        'Desc_Secao': [],
        'Desc_SH6': []
    }
    # Forçar rerun para atualizar a visualização
    st.rerun()

# Aplicar filtros usando os valores armazenados em session_state
filtros = st.session_state.filtros_ativos
df_filtrado = aplicar_filtros(df, filtros)

# Mostrar filtros ativos
if any(filtros.values()):
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Filtros Ativos")
    for campo, valores in filtros.items():
        if valores:
            st.sidebar.markdown(f"**{campo}:** {', '.join(map(str, valores))}")

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
    df_temporal = df_filtrado.groupby(['Ano', 'Fluxo'])['Valor_FOB'].sum().reset_index()
    
    # Calcular o valor formatado para o hover
    df_temporal['Valor_FOB_Format'] = df_temporal['Valor_FOB'].apply(format_big_number)
    
    # Calcular os valores min e max para o eixo Y
    y_min = df_temporal['Valor_FOB'].min()
    y_max = df_temporal['Valor_FOB'].max()
    y_range = y_max - y_min
    
    # Criar valores para o eixo Y (6 pontos igualmente espaçados)
    y_ticks = [y_min + (y_range * i / 5) for i in range(6)]
    y_tick_texts = [format_big_number(val) for val in y_ticks]
    
    fig_temporal = px.line(
        df_temporal,
        x='Ano',
        y='Valor_FOB',
        color='Fluxo',
        title="Evolução do Valor FOB por Ano e Fluxo",
        template='plotly_dark',
        labels={'Valor_FOB': 'Valor FOB', 'Ano': 'Ano', 'Fluxo': 'Fluxo'}
    )
    
    # Limpar os traces automáticos
    fig_temporal.data = []
    
    # Personalizar o layout
    fig_temporal.update_layout(
        height=500,
        hovermode='x unified',
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=0.01,
            bgcolor='rgba(0,0,0,0.3)'
        ),
        yaxis=dict(
            ticktext=y_tick_texts,  # Usar os textos pré-calculados
            tickvals=y_ticks,       # Usar os valores pré-calculados
            gridcolor='rgba(128,128,128,0.2)',
            title_font=dict(size=14),
            tickfont=dict(size=12)
        ),
        xaxis=dict(
            gridcolor='rgba(128,128,128,0.2)',
            title_font=dict(size=14),
            tickfont=dict(size=12),
            dtick=1
        ),
        title=dict(
            font=dict(size=16),
            y=0.95
        ),
        margin=dict(l=60, r=30, t=50, b=50)
    )
    
    # Personalizar as linhas com cores específicas
    colors = {'Exportação': '#636EFA', 'Importação': '#EF553B'}
    
    for fluxo in df_temporal['Fluxo'].unique():
        df_fluxo = df_temporal[df_temporal['Fluxo'] == fluxo]
        
        fig_temporal.add_trace(
            go.Scatter(
                x=df_fluxo['Ano'],
                y=df_fluxo['Valor_FOB'],
                name=fluxo,
                mode='lines+markers',
                line=dict(width=3, color=colors[fluxo]),
                marker=dict(size=8, color=colors[fluxo]),
                hovertemplate="<b>Ano: %{x}</b><br>" +
                             f"{fluxo}: %{{text}}<br>" +
                             "<extra></extra>",
                text=df_fluxo['Valor_FOB_Format']
            )
        )
        
        # Adicionar rótulo no ponto final
        ultimo_valor = df_fluxo.iloc[-1]
        fig_temporal.add_annotation(
            x=ultimo_valor['Ano'],
            y=ultimo_valor['Valor_FOB'],
            text=ultimo_valor['Valor_FOB_Format'],
            showarrow=True,
            arrowhead=0,
            ax=40,
            ay=-40 if fluxo == 'Exportação' else 40,
            font=dict(size=12),
            bgcolor='rgba(0,0,0,0.5)',
            bordercolor='rgba(255,255,255,0.3)',
            borderwidth=1,
            borderpad=4
        )
    
    # Exibir o gráfico
    st.plotly_chart(fig_temporal, use_container_width=True)
    
    st.markdown("""
    <div style='background-color: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px;'>
        <small>
        Este gráfico mostra a evolução temporal dos valores FOB de importação e exportação ao longo dos anos.
        As linhas representam as tendências de cada fluxo, com valores formatados para melhor visualização.
        Os pontos finais são destacados com rótulos para facilitar a interpretação dos valores mais recentes.
        </small>
    </div>
    """, unsafe_allow_html=True)

with tab2:
    st.subheader("Distribuição Global do Valor FOB")
    
    # Atualizar o dicionário de países (agora com 198 países)
    pais_map = {
        'Afeganistão': 'Afghanistan', 'África do Sul': 'South Africa', 'Albânia': 'Albania',
        'Alemanha': 'Germany', 'Andorra': 'Andorra', 'Angola': 'Angola',
        'Antígua e Barbuda': 'Antigua and Barbuda', 'Arábia Saudita': 'Saudi Arabia',
        'Argélia': 'Algeria', 'Argentina': 'Argentina', 'Armênia': 'Armenia',
        'Austrália': 'Australia', 'Áustria': 'Austria', 'Azerbaijão': 'Azerbaijan',
        'Bahamas': 'Bahamas', 'Bangladesh': 'Bangladesh', 'Barbados': 'Barbados',
        'Barein': 'Bahrain', 'Bélgica': 'Belgium', 'Belize': 'Belize',
        'Benin': 'Benin', 'Bielorrússia': 'Belarus', 'Bolívia': 'Bolivia',
        'Bósnia e Herzegovina': 'Bosnia and Herzegovina', 'Botsuana': 'Botswana',
        'Brasil': 'Brazil', 'Brunei': 'Brunei', 'Bulgária': 'Bulgaria',
        'Burkina Faso': 'Burkina Faso', 'Burundi': 'Burundi', 'Butão': 'Bhutan',
        'Cabo Verde': 'Cape Verde', 'Camarões': 'Cameroon', 'Camboja': 'Cambodia',
        'Canadá': 'Canada', 'Catar': 'Qatar', 'Cazaquistão': 'Kazakhstan',
        'Chade': 'Chad', 'Chile': 'Chile', 'China': 'China', 'Chipre': 'Cyprus',
        'Cingapura': 'Singapore', 'Colômbia': 'Colombia', 'Comores': 'Comoros',
        'Congo': 'Congo', 'Coreia do Norte': 'North Korea', 'Coreia do Sul': 'South Korea',
        'Costa do Marfim': 'Ivory Coast', 'Costa Rica': 'Costa Rica', 'Croácia': 'Croatia',
        'Cuba': 'Cuba', 'Dinamarca': 'Denmark', 'Djibuti': 'Djibouti', 'Dominica': 'Dominica',
        'Egito': 'Egypt', 'El Salvador': 'El Salvador', 'Emirados Árabes Unidos': 'United Arab Emirates',
        'Equador': 'Ecuador', 'Eritreia': 'Eritrea', 'Eslováquia': 'Slovakia',
        'Eslovênia': 'Slovenia', 'Espanha': 'Spain', 'Estados Unidos': 'United States',
        'Estônia': 'Estonia', 'Eswatini': 'Eswatini', 'Etiópia': 'Ethiopia',
        'Fiji': 'Fiji', 'Filipinas': 'Philippines', 'Finlândia': 'Finland',
        'França': 'France', 'Gabão': 'Gabon', 'Gâmbia': 'Gambia', 'Gana': 'Ghana',
        'Geórgia': 'Georgia', 'Granada': 'Grenada', 'Grécia': 'Greece',
        'Guatemala': 'Guatemala', 'Guiana': 'Guyana', 'Guiné': 'Guinea',
        'Guiné Equatorial': 'Equatorial Guinea', 'Guiné-Bissau': 'Guinea-Bissau',
        'Haiti': 'Haiti', 'Honduras': 'Honduras', 'Hungria': 'Hungary',
        'Iêmen': 'Yemen', 'Índia': 'India', 'Indonésia': 'Indonesia',
        'Irã': 'Iran', 'Iraque': 'Iraq', 'Irlanda': 'Ireland',
        'Islândia': 'Iceland', 'Israel': 'Israel', 'Itália': 'Italy',
        'Jamaica': 'Jamaica', 'Japão': 'Japan', 'Jordânia': 'Jordan',
        'Kuwait': 'Kuwait', 'Laos': 'Laos', 'Lesoto': 'Lesotho',
        'Letônia': 'Latvia', 'Líbano': 'Lebanon', 'Libéria': 'Liberia',
        'Líbia': 'Libya', 'Liechtenstein': 'Liechtenstein', 'Lituânia': 'Lithuania',
        'Luxemburgo': 'Luxembourg', 'Macedônia do Norte': 'North Macedonia',
        'Madagascar': 'Madagascar', 'Malásia': 'Malaysia', 'Malaui': 'Malawi',
        'Maldivas': 'Maldives', 'Mali': 'Mali', 'Malta': 'Malta',
        'Marrocos': 'Morocco', 'Maurício': 'Mauritius', 'Mauritânia': 'Mauritania',
        'México': 'Mexico', 'Mianmar': 'Myanmar', 'Micronésia': 'Micronesia',
        'Moçambique': 'Mozambique', 'Moldávia': 'Moldova', 'Mônaco': 'Monaco',
        'Mongólia': 'Mongolia', 'Montenegro': 'Montenegro', 'Namíbia': 'Namibia',
        'Nauru': 'Nauru', 'Nepal': 'Nepal', 'Nicarágua': 'Nicaragua',
        'Níger': 'Niger', 'Nigéria': 'Nigeria', 'Noruega': 'Norway',
        'Nova Zelândia': 'New Zealand', 'Omã': 'Oman', 'Países Baixos': 'Netherlands',
        'Palau': 'Palau', 'Panamá': 'Panama', 'Papua Nova Guiné': 'Papua New Guinea',
        'Paquistão': 'Pakistan', 'Paraguai': 'Paraguay', 'Peru': 'Peru',
        'Polônia': 'Poland', 'Portugal': 'Portugal', 'Quênia': 'Kenya',
        'Quirguistão': 'Kyrgyzstan', 'Reino Unido': 'United Kingdom',
        'República Centro-Africana': 'Central African Republic',
        'República Democrática do Congo': 'Democratic Republic of the Congo',
        'República Dominicana': 'Dominican Republic', 'República Tcheca': 'Czech Republic',
        'Romênia': 'Romania', 'Ruanda': 'Rwanda', 'Rússia': 'Russia',
        'Salomão': 'Solomon Islands', 'Samoa': 'Samoa', 'San Marino': 'San Marino',
        'Santa Lúcia': 'Saint Lucia', 'São Cristóvão e Nevis': 'Saint Kitts and Nevis',
        'São Tomé e Príncipe': 'Sao Tome and Principe',
        'São Vicente e Granadinas': 'Saint Vincent and the Grenadines',
        'Seicheles': 'Seychelles', 'Senegal': 'Senegal', 'Serra Leoa': 'Sierra Leone',
        'Sérvia': 'Serbia', 'Síria': 'Syria', 'Somália': 'Somalia',
        'Sri Lanka': 'Sri Lanka', 'Sudão': 'Sudan', 'Sudão do Sul': 'South Sudan',
        'Suécia': 'Sweden', 'Suíça': 'Switzerland', 'Suriname': 'Suriname',
        'Tadjiquistão': 'Tajikistan', 'Tailândia': 'Thailand', 'Taiwan': 'Taiwan',
        'Tanzânia': 'Tanzania', 'Timor-Leste': 'Timor-Leste', 'Togo': 'Togo',
        'Tonga': 'Tonga', 'Trinidad e Tobago': 'Trinidad and Tobago',
        'Tunísia': 'Tunisia', 'Turcomenistão': 'Turkmenistan', 'Turquia': 'Turkey',
        'Tuvalu': 'Tuvalu', 'Ucrânia': 'Ukraine', 'Uganda': 'Uganda',
        'Uruguai': 'Uruguay', 'Uzbequistão': 'Uzbekistan', 'Vanuatu': 'Vanuatu',
        'Vaticano': 'Vatican City', 'Venezuela': 'Venezuela', 'Vietnã': 'Vietnam',
        'Zâmbia': 'Zambia', 'Zimbábue': 'Zimbabwe',
        'Hong Kong': 'Hong Kong',
        'Macau': 'Macao',
        'Taiwan, Província da China': 'Taiwan',
        'Coreia, República da': 'South Korea',
        'Irã, República Islâmica do': 'Iran',
        'República Democrática Popular do Laos': 'Laos',
        'Vietnã': 'Vietnam',
        'Estado da Palestina': 'Palestine',
        'Síria, República Árabe da': 'Syria',
        'Brunei Darussalam': 'Brunei',
        'Ilhas Virgens Britânicas': 'British Virgin Islands',
        'Ilhas Cayman': 'Cayman Islands',
        'São Martinho (Países Baixos)': 'Sint Maarten',
        'Curaçao': 'Curacao',
        'Guadalupe': 'Guadeloupe',
        'Martinica': 'Martinique',
        'Porto Rico': 'Puerto Rico',
        'Guiana Francesa': 'French Guiana',
        'Saara Ocidental': 'Western Sahara',
        'Ilha de Man': 'Isle of Man',
        'Ilhas Faroe': 'Faroe Islands',
        'Groenlândia': 'Greenland',
        'Guam': 'Guam',
        'Nova Caledônia': 'New Caledonia',
        'Polinésia Francesa': 'French Polynesia',
        'Samoa Americana': 'American Samoa',
        'Territórios Franceses do Sul': 'French Southern Territories',
        'República da Macedônia do Norte': 'North Macedonia',
        'Kosovo': 'Kosovo',
        'Território Britânico do Oceano Índico': 'British Indian Ocean Territory',
        'Mayotte': 'Mayotte',
        'Reunião': 'Reunion',
        'Santa Helena': 'Saint Helena',
        'Svalbard e Jan Mayen': 'Svalbard and Jan Mayen',
        'Ilhas Malvinas': 'Falkland Islands'
    }
    
    # Preparar dados para o mapa
    df_mapa = df_filtrado.groupby('Países')['Valor_FOB'].sum().reset_index()
    df_mapa['Países_EN'] = df_mapa['Países'].map(pais_map).fillna(df_mapa['Países'])
    df_mapa['Valor_FOB_Format'] = df_mapa['Valor_FOB'].apply(format_currency)
    
    # Criar função para formatar os valores da régua
    def format_colorbar_tick(value):
        """Formata os valores da régua do mapa para um formato mais conciso"""
        if value >= 1e9:
            return f"${value/1e9:.2f}B"
        elif value >= 1e6:
            return f"${value/1e6:.2f}M"
        elif value >= 1e3:
            return f"${value/1e3:.2f}K"
        return f"${value:.2f}"
    
    # Criar mapa
    fig_mapa = px.choropleth(
        df_mapa,
        locations='Países_EN',
        locationmode='country names',
        color='Valor_FOB',
        hover_name='Países',
        hover_data={
            'Países_EN': False,
            'Valor_FOB': False,
            'Valor_FOB_Format': True
        },
        color_continuous_scale='RdBu',
        template='plotly_dark'
    )
    
    # Configurar o hover template
    fig_mapa.update_traces(
        hovertemplate="<b>%{hovertext}</b><br>" +
                     "Valor FOB: %{customdata[0]}<br>" +
                     "<extra></extra>",
        customdata=df_mapa[['Valor_FOB_Format']]
    )
    
    # Calcular os valores dos ticks da régua
    max_valor = df_mapa['Valor_FOB'].max()
    tick_values = [i * max_valor/4 for i in range(5)]  # 5 pontos na régua
    
    # Configurar layout do mapa
    fig_mapa.update_layout(
        geo=dict(
            showframe=True,
            showcoastlines=True,
            projection_type='natural earth',
            coastlinecolor='Gray',
            countrycolor='Gray',
            showland=True,
            landcolor='rgba(50, 50, 50, 0.8)',
            showocean=True,
            oceancolor='rgba(30, 30, 30, 0.8)',
            showcountries=True,
            bgcolor='rgba(0,0,0,0)'
        ),
        height=600,
        margin=dict(l=0, r=0, t=30, b=0),
        paper_bgcolor='rgba(0,0,0,0)',
        coloraxis_colorbar=dict(
            title='Valor FOB',
            ticktext=[format_colorbar_tick(val) for val in tick_values],
            tickvals=tick_values,
            len=0.8,
            thickness=20,
            tickfont=dict(size=12)
        )
    )
    
    st.plotly_chart(fig_mapa, use_container_width=True)

    # Gráficos de análise geográfica
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Top Países por Valor FOB")
        n_paises = st.selectbox(
            "Número de países a exibir",
            options=[10, 20, 50, 100],
            key="n_paises"
        )
        
        df_paises = (df_filtrado.groupby('Países')['Valor_FOB']
                    .sum()
                    .sort_values(ascending=False)
                    .head(n_paises)
                    .reset_index())
        
        df_paises['Valor_FOB_Format'] = df_paises['Valor_FOB'].apply(format_currency)
        
        # Calcular os valores dos ticks
        max_valor_paises = df_paises['Valor_FOB'].max()
        tick_values_paises = [i * max_valor_paises/5 for i in range(6)]
        
        fig_paises = go.Figure()
        fig_paises.add_trace(
            go.Bar(
                x=df_paises['Valor_FOB'],
                y=df_paises['Países'],
                orientation='h',
                text=df_paises['Valor_FOB_Format'],
                textposition='outside',
                marker=dict(
                    color='rgba(99, 110, 250, 0.8)',
                    line=dict(color='rgba(99, 110, 250, 1.0)', width=2)
                ),
                hovertemplate="<b>%{y}</b><br>" +
                             "Valor FOB: %{text}<br>" +
                             "<extra></extra>"
            )
        )
        
        fig_paises.update_layout(
            xaxis=dict(
                title="Valor FOB",
                ticktext=[format_currency(val) for val in tick_values_paises],
                tickvals=tick_values_paises,
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(128,128,128,0.2)',
            ),
            yaxis=dict(title=""),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            height=500,
            margin=dict(l=10, r=10, t=30, b=10),
            hoverlabel=dict(
                bgcolor='white',
                font_color='black',
                font_size=12
            )
        )
        
        st.plotly_chart(fig_paises, use_container_width=True)
        
        st.markdown("""
        <div style='background-color: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px;'>
            <small>
            Este gráfico apresenta os principais países ordenados por valor FOB total.
            As barras mostram a contribuição de cada país para o comércio exterior,
            permitindo identificar os parceiros comerciais mais significativos.
            </small>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.subheader("Top URF por Valor FOB")
        n_urf = st.selectbox(
            "Número de URFs a exibir",
            options=[10, 20, 50, 100],
            key="n_urf"
        )
        
        df_urf = (df_filtrado.groupby('URF')['Valor_FOB']
                 .sum()
                 .sort_values(ascending=False)
                 .head(n_urf)
                 .reset_index())
        
        df_urf['Valor_FOB_Format'] = df_urf['Valor_FOB'].apply(format_currency)
        
        # Calcular os valores dos ticks
        max_valor_urf = df_urf['Valor_FOB'].max()
        tick_values_urf = [i * max_valor_urf/5 for i in range(6)]
        
        fig_urf = go.Figure()
        fig_urf.add_trace(
            go.Bar(
                x=df_urf['Valor_FOB'],
                y=df_urf['URF'],
                orientation='h',
                text=df_urf['Valor_FOB_Format'],
                textposition='outside',
                marker=dict(
                    color='rgba(99, 110, 250, 0.8)',
                    line=dict(color='rgba(99, 110, 250, 1.0)', width=2)
                ),
                hovertemplate="<b>%{y}</b><br>" +
                             "Valor FOB: %{text}<br>" +
                             "<extra></extra>"
            )
        )
        
        fig_urf.update_layout(
            xaxis=dict(
                title="Valor FOB",
                ticktext=[format_currency(val) for val in tick_values_urf],
                tickvals=tick_values_urf,
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(128,128,128,0.2)',
            ),
            yaxis=dict(title=""),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            height=500,
            margin=dict(l=10, r=10, t=30, b=10),
            hoverlabel=dict(
                bgcolor='white',
                font_color='black',
                font_size=12
            )
        )
        
        st.plotly_chart(fig_urf, use_container_width=True)
        
        st.markdown("""
        <div style='background-color: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px;'>
            <small>
            Este gráfico mostra as principais Unidades da Receita Federal (URFs) por valor FOB.
            Cada barra representa o volume total de operações processadas por cada URF,
            auxiliando na identificação dos principais pontos de entrada e saída de mercadorias.
            </small>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("Distribuição de Produtos por URF")
    
    # Controles de seleção
    col_urf_controls = st.columns([1, 1, 2])
    
    with col_urf_controls[0]:
        n_urf_geo = st.selectbox(
            "Número de URFs",
            options=[5, 10, 15, 20],
            key="n_urf_geo"
        )
    
    with col_urf_controls[1]:
        n_produtos_urf = st.selectbox(
            "Produtos por URF",
            options=[5, 10, 15],
            index=0,
            key="n_produtos_urf"
        )
    
    # Preparar dados para o gráfico
    top_urfs = (df_filtrado.groupby('URF')['Valor_FOB']
               .sum()
               .sort_values(ascending=True)
               .tail(n_urf_geo)
               .index)
    
    df_urf_stacked = df_filtrado[df_filtrado['URF'].isin(top_urfs)]
    
    # Para cada URF, pegar os top N produtos
    dfs_urf_produtos = []
    for urf in top_urfs:
        df_urf = df_urf_stacked[df_urf_stacked['URF'] == urf]
        top_produtos_urf = (df_urf.groupby('Desc_SH6')['Valor_FOB']
                            .sum()
                            .sort_values(ascending=False)
                            .head(n_produtos_urf)
                            .reset_index())
        top_produtos_urf['URF'] = urf
        dfs_urf_produtos.append(top_produtos_urf)
    
    df_urf_plot = pd.concat(dfs_urf_produtos)
    
    # Calcular os valores dos ticks antes de criar o gráfico
    max_valor_urf = df_urf_plot['Valor_FOB'].max()
    tick_values_urf = [i * max_valor_urf/5 for i in range(6)]
    
    # Criar o gráfico
    fig_urf_stacked = go.Figure()
    
    # Adicionar uma barra para cada produto
    produtos_unicos_urf = df_urf_plot['Desc_SH6'].unique()
    
    for i, produto in enumerate(produtos_unicos_urf):
        df_produto = df_urf_plot[df_urf_plot['Desc_SH6'] == produto]
        
        hover_text = [
            f"<b>URF:</b> {urf}<br>" +
            f"<b>Produto:</b> {produto}<br>" +
            f"<b>Valor:</b> {format_currency(valor)}"
            for urf, valor in zip(df_produto['URF'], df_produto['Valor_FOB'])
        ]
        
        fig_urf_stacked.add_trace(go.Bar(
            name=produto[:50] + '...' if len(produto) > 50 else produto,
            y=df_produto['URF'],
            x=df_produto['Valor_FOB'],
            orientation='h',
            hovertext=hover_text,
            hoverinfo='text',
            marker_color=MAPA_CORES_PRODUTOS[produto]  # Usar a cor fixa do mapeamento
        ))
    
    # Atualizar o layout
    fig_urf_stacked.update_layout(
        barmode='stack',
        height=max(400, n_urf_geo * 40),
        margin=dict(l=20, r=20, t=30, b=20),
        xaxis=dict(
            ticktext=[format_currency(val) for val in tick_values_urf],
            tickvals=tick_values_urf,
            title="Valor FOB",
            showgrid=True,
            gridwidth=1,
            gridcolor='rgba(128,128,128,0.2)',
        ),
        yaxis=dict(
            title="",
            categoryorder='total ascending'
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.3,
            xanchor="center",
            x=0.5,
            bgcolor='rgba(255, 255, 255, 0.1)'
        ),
        showlegend=True,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        hoverlabel=dict(
            bgcolor='white',
            font_color='black',
            font_size=12
        )
    )
    
    # Exibir o gráfico
    st.plotly_chart(fig_urf_stacked, use_container_width=True)
    
    # Adicionar legenda explicativa
    st.markdown("""
    <div style='background-color: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px;'>
        <small>
        Este gráfico mostra a distribuição dos principais produtos por URF (Unidade da Receita Federal).
        Cada barra representa uma URF, e as cores diferentes mostram a contribuição de cada produto.
        Passe o mouse sobre as barras para ver os detalhes.
        </small>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("Comparação de Produtos entre URFs")
    
    # Controles para seleção
    col_comp_controls = st.columns([1, 1, 1])
    
    with col_comp_controls[0]:
        urf_1 = st.selectbox(
            "URF 1",
            options=sorted(df_filtrado['URF'].unique()),
            key="urf_1"
        )
    
    with col_comp_controls[1]:
        urf_2 = st.selectbox(
            "URF 2",
            options=sorted(df_filtrado['URF'].unique()),
            key="urf_2"
        )
    
    with col_comp_controls[2]:
        min_valor = st.number_input(
            "Valor FOB Mínimo (USD)",
            min_value=0,
            value=1000000,
            step=1000000,
            format="%d"
        )
    
    # Preparar dados para comparação
    df_urf1 = df_filtrado[df_filtrado['URF'] == urf_1].groupby('Desc_SH6')['Valor_FOB'].sum()
    df_urf2 = df_filtrado[df_filtrado['URF'] == urf_2].groupby('Desc_SH6')['Valor_FOB'].sum()
    
    # Encontrar produtos em comum
    produtos_comuns = set(df_urf1.index) & set(df_urf2.index)
    
    # Criar DataFrame com produtos em comum
    df_comparacao = pd.DataFrame({
        'Produto': list(produtos_comuns),
        'Valor_URF1': [df_urf1[prod] for prod in produtos_comuns],
        'Valor_URF2': [df_urf2[prod] for prod in produtos_comuns]
    })
    
    # Filtrar por valor mínimo
    df_comparacao = df_comparacao[
        (df_comparacao['Valor_URF1'] >= min_valor) |
        (df_comparacao['Valor_URF2'] >= min_valor)
    ]
    
    # Calcular diferença percentual
    df_comparacao['Diferenca_Percentual'] = (
        (df_comparacao['Valor_URF1'] - df_comparacao['Valor_URF2']) /
        ((df_comparacao['Valor_URF1'] + df_comparacao['Valor_URF2']) / 2) * 100
    )
    
    # Criar o gráfico de dispersão
    fig_comparacao = go.Figure()
    
    # Adicionar linha diagonal de referência
    max_valor = max(df_comparacao['Valor_URF1'].max(), df_comparacao['Valor_URF2'].max())
    fig_comparacao.add_trace(go.Scatter(
        x=[0, max_valor],
        y=[0, max_valor],
        mode='lines',
        name='Linha de Igualdade',
        line=dict(dash='dash', color='gray'),
        hoverinfo='skip'
    ))
    
    # Adicionar os pontos
    hover_text = [
        f"<b>Produto:</b> {prod}<br>" +
        f"<b>{urf_1}:</b> {format_currency(val1)}<br>" +
        f"<b>{urf_2}:</b> {format_currency(val2)}<br>" +
        f"<b>Diferença:</b> {diff:.1f}%"
        for prod, val1, val2, diff in zip(
            df_comparacao['Produto'],
            df_comparacao['Valor_URF1'],
            df_comparacao['Valor_URF2'],
            df_comparacao['Diferenca_Percentual']
        )
    ]
    
    fig_comparacao.add_trace(go.Scatter(
        x=df_comparacao['Valor_URF1'],
        y=df_comparacao['Valor_URF2'],
        mode='markers',
        name='Produtos',
        marker=dict(
            size=10,
            color=df_comparacao['Diferenca_Percentual'],
            colorscale='RdBu',
            colorbar=dict(
                title='Diferença %',
                ticksuffix='%'
            ),
            showscale=True
        ),
        hovertext=hover_text,
        hoverinfo='text'
    ))
    
    # Atualizar layout
    fig_comparacao.update_layout(
        title=f"Comparação de Valores FOB entre {urf_1} e {urf_2}",
        xaxis=dict(
            title=f"Valor FOB - {urf_1}",
            type='log',
            showgrid=True,
            gridwidth=1,
            gridcolor='rgba(128,128,128,0.2)',
        ),
        yaxis=dict(
            title=f"Valor FOB - {urf_2}",
            type='log',
            showgrid=True,
            gridwidth=1,
            gridcolor='rgba(128,128,128,0.2)',
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        height=600,
        showlegend=True,
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=0.01,
            bgcolor='rgba(255,255,255,0.1)'
        ),
        hoverlabel=dict(
            bgcolor='white',
            font_color='black',
            font_size=12
        )
    )
    
    # Exibir o gráfico
    st.plotly_chart(fig_comparacao, use_container_width=True)
    
    # Adicionar explicação
    st.markdown("""
    <div style='background-color: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px;'>
        <small>
        Este gráfico permite comparar os valores FOB de produtos comercializados entre duas URFs:
        <ul>
            <li>Cada ponto representa um produto comercializado por ambas URFs</li>
            <li>A cor indica a diferença percentual entre as URFs</li>
            <li>Pontos acima da linha diagonal indicam maior valor na URF 2</li>
            <li>Pontos abaixo da linha diagonal indicam maior valor na URF 1</li>
            <li>A escala logarítmica permite visualizar melhor as diferenças em diferentes ordens de magnitude</li>
        </ul>
        </small>
    </div>
    """, unsafe_allow_html=True)

with tab3:
    st.subheader("Análise por Produto")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Top Seções por Valor FOB")
        n_secoes = st.selectbox(
            "Número de seções a exibir",
            options=[10, 20, 50, 100],
            key="n_secoes"
        )
        
        # Top N seções
        df_secoes = (df_filtrado.groupby('Desc_Secao')['Valor_FOB']
                    .sum()
                    .sort_values(ascending=False)
                    .head(n_secoes)
                    .reset_index())
        
        df_secoes['Valor_FOB_Format'] = df_secoes['Valor_FOB'].apply(format_currency)
        
        # Calcular os valores dos ticks
        max_valor_secoes = df_secoes['Valor_FOB'].max()
        tick_values_secoes = [i * max_valor_secoes/5 for i in range(6)]
        
        fig_secoes = go.Figure()
        fig_secoes.add_trace(
            go.Bar(
                x=df_secoes['Valor_FOB'],
                y=df_secoes['Desc_Secao'],
                orientation='h',
                text=df_secoes['Valor_FOB_Format'],
                textposition='outside',
                marker=dict(
                    color='rgba(99, 110, 250, 0.8)',
                    line=dict(color='rgba(99, 110, 250, 1.0)', width=2)
                ),
                hovertemplate="<b>%{y}</b><br>" +
                             "Valor FOB: %{text}<br>" +
                             "<extra></extra>"
            )
        )
        
        fig_secoes.update_layout(
            xaxis=dict(
                title="Valor FOB",
                ticktext=[format_currency(val) for val in tick_values_secoes],
                tickvals=tick_values_secoes,
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(128,128,128,0.2)',
            ),
            yaxis=dict(
                title="",
                tickfont=dict(size=10)
            ),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            height=500,
            margin=dict(l=10, r=120, t=30, b=10),
            hoverlabel=dict(
                bgcolor='white',
                font_color='black',
                font_size=12
            )
        )
        
        st.plotly_chart(fig_secoes, use_container_width=True)
        
        st.markdown("""
        <div style='background-color: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px;'>
            <small>
            Este gráfico apresenta as principais seções de produtos por valor FOB.
            As seções são categorias amplas que agrupam produtos similares,
            permitindo uma visão macro da distribuição do comércio exterior por tipo de mercadoria.
            </small>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.subheader("Top Produtos por Valor FOB")
        n_produtos = st.selectbox(
            "Número de produtos a exibir",
            options=[10, 20, 50, 100],
            key="n_produtos"
        )
        
        # Top N produtos
        df_produtos = (df_filtrado.groupby('Desc_SH6')['Valor_FOB']
                      .sum()
                      .sort_values(ascending=False)
                      .head(n_produtos)
                      .reset_index())
        
        df_produtos['Valor_FOB_Format'] = df_produtos['Valor_FOB'].apply(format_currency)
        
        # Calcular os valores dos ticks
        max_valor_produtos = df_produtos['Valor_FOB'].max()
        tick_values_produtos = [i * max_valor_produtos/5 for i in range(6)]
        
        fig_produtos = go.Figure()
        fig_produtos.add_trace(
            go.Bar(
                x=df_produtos['Valor_FOB'],
                y=df_produtos['Desc_SH6'],
                orientation='h',
                text=df_produtos['Valor_FOB_Format'],
                textposition='outside',
                marker=dict(
                    color='rgba(99, 110, 250, 0.8)',
                    line=dict(color='rgba(99, 110, 250, 1.0)', width=2)
                ),
                hovertemplate="<b>%{y}</b><br>" +
                             "Valor FOB: %{text}<br>" +
                             "<extra></extra>"
            )
        )
        
        fig_produtos.update_layout(
            xaxis=dict(
                title="Valor FOB",
                ticktext=[format_currency(val) for val in tick_values_produtos],
                tickvals=tick_values_produtos,
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(128,128,128,0.2)',
            ),
            yaxis=dict(
                title="",
                tickfont=dict(size=10)
            ),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            height=500,
            margin=dict(l=10, r=120, t=30, b=10),
            hoverlabel=dict(
                bgcolor='white',
                font_color='black',
                font_size=12
            )
        )
        
        st.plotly_chart(fig_produtos, use_container_width=True)
        
        st.markdown("""
        <div style='background-color: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px;'>
            <small>
            Este gráfico mostra os principais produtos específicos (códigos SH6) por valor FOB.
            Permite identificar os produtos individuais mais relevantes no comércio exterior,
            oferecendo uma visão detalhada das mercadorias mais comercializadas.
            </small>
        </div>
        """, unsafe_allow_html=True)
    
    # Adicionar após os gráficos existentes
    st.markdown("---")
    st.subheader("Contribuição dos Principais Produtos por País")
    
    # Controles de seleção
    col_controls = st.columns([1, 1, 2])
    
    with col_controls[0]:
        n_paises_stacked = st.selectbox(
            "Número de países",
            options=[5, 10, 15, 20],
            key="n_paises_stacked"
        )
    
    with col_controls[1]:
        n_produtos_stacked = st.selectbox(
            "Produtos por país",
            options=[5, 10, 15],
            index=0,
            key="n_produtos_stacked"
        )
    
    # Preparar dados para o gráfico
    top_paises = (df_filtrado.groupby('Países')['Valor_FOB']
                 .sum()
                 .sort_values(ascending=True)
                 .tail(n_paises_stacked)
                 .index)
    
    df_stacked = df_filtrado[df_filtrado['Países'].isin(top_paises)]
    
    # Para cada país, pegar os top N produtos
    dfs_produtos = []
    for pais in top_paises:
        df_pais = df_stacked[df_stacked['Países'] == pais]
        top_produtos = (df_pais.groupby('Desc_SH6')['Valor_FOB']
                       .sum()
                       .sort_values(ascending=False)
                       .head(n_produtos_stacked)
                       .reset_index())
        top_produtos['Países'] = pais
        dfs_produtos.append(top_produtos)
    
    df_plot = pd.concat(dfs_produtos)
    
    # Calcular os valores dos ticks antes de criar o gráfico
    max_valor = df_plot['Valor_FOB'].max()
    tick_values = [i * max_valor/5 for i in range(6)]

    # Criar o gráfico
    fig_stacked = go.Figure()
    
    # Adicionar uma barra para cada produto
    produtos_unicos = df_plot['Desc_SH6'].unique()
    
    for i, produto in enumerate(produtos_unicos):
        df_produto = df_plot[df_plot['Desc_SH6'] == produto]
        
        hover_text = [
            f"<b>País:</b> {pais}<br>" +
            f"<b>Produto:</b> {produto}<br>" +
            f"<b>Valor:</b> {format_currency(valor)}"
            for pais, valor in zip(df_produto['Países'], df_produto['Valor_FOB'])
        ]
        
        fig_stacked.add_trace(go.Bar(
            name=produto[:50] + '...' if len(produto) > 50 else produto,
            y=df_produto['Países'],
            x=df_produto['Valor_FOB'],
            orientation='h',
            hovertext=hover_text,
            hoverinfo='text',
            marker_color=MAPA_CORES_PRODUTOS[produto]  # Usar a cor fixa do mapeamento
        ))
    
    # Atualizar o layout
    fig_stacked.update_layout(
        barmode='stack',
        height=max(400, n_paises_stacked * 40),
        margin=dict(l=20, r=20, t=30, b=20),
        xaxis=dict(
            ticktext=[format_currency(val) for val in tick_values],
            tickvals=tick_values,
            title="Valor FOB",
            showgrid=True,
            gridwidth=1,
            gridcolor='rgba(128,128,128,0.2)',
        ),
        yaxis=dict(
            title="",
            categoryorder='total ascending'
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.3,
            xanchor="center",
            x=0.5,
            bgcolor='rgba(255, 255, 255, 0.1)'
        ),
        showlegend=True,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        hoverlabel=dict(
            bgcolor='white',
            font_color='black',
            font_size=12
        )
    )
    
    # Exibir o gráfico
    st.plotly_chart(fig_stacked, use_container_width=True)
    
    # Adicionar legenda explicativa
    st.markdown("""
    <div style='background-color: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px;'>
        <small>
        Este gráfico mostra a contribuição dos principais produtos para cada país selecionado.
        Cada barra representa um país, e as cores diferentes mostram a contribuição de cada produto.
        Passe o mouse sobre as barras para ver os detalhes.
        </small>
    </div>
    """, unsafe_allow_html=True)

# Modificar a parte do download para Excel
# Substituir a parte final do código onde está o download
st.subheader("Dados Detalhados")
st.dataframe(
    df_filtrado.sort_values('Valor_FOB', ascending=False),
    hide_index=True
)

# Download em Excel
if st.button("Download dos dados filtrados (Excel)"):
    # Criar um buffer para o arquivo Excel
    buffer = io.BytesIO()
    
    # Criar o arquivo Excel
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_filtrado.to_excel(writer, sheet_name='Dados', index=False)
        
        # Ajustar as colunas automaticamente
        worksheet = writer.sheets['Dados']
        for i, col in enumerate(df_filtrado.columns):
            column_len = max(df_filtrado[col].astype(str).apply(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)
    
    # Preparar o download
    buffer.seek(0)
    st.download_button(
        label="Clique para download",
        data=buffer,
        file_name="comercio_exterior_filtrado.xlsx",
        mime="application/vnd.ms-excel"
    )
