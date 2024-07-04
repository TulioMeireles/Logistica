# Importando as bibliotecas
import pandas as pd
import plotly_express as px
import streamlit as st
import folium
from streamlit_folium import folium_static 
from PIL import Image
from fpdf import FPDF
import os
import pyodbc
from streamlit_extras.metric_cards import style_metric_cards

# 1 - CRIANDO O LAYOUT
st.set_page_config(page_title='Tulio Meireles', page_icon='chart_with_upwards_trend', layout='wide')
imagem = Image.open("Logistica.png")
st.image(imagem, use_column_width=True, caption="LOGISTICA DE ENTREGA DE PRODUTOS")


# 2 - CONFIGURANDO A CONEXÃO COM O BANCO

server = 'DESKTOP-D3SPOTC\SQLEXPRESS'
database = 'Powerbi'
usuario = 'sa'
senha = '245316'

def load_data():
    try:
        conn = pyodbc.connect('Driver={ODBC Driver 11 for SQL Server};'
                              f'Server={server};'
                              f'Username={usuario};'
                              f'Password={senha};'
                              f'Database={database};'
                              'Trusted_Connection=yes;')

        # Script de consulta da tabela logistica
        consulta = '''SELECT *
                      FROM Logistica
                      ORDER BY [Data entrega]'''

        # Carregando a tabela
        dados = pd.read_sql_query(consulta, conn)
        conn.close()

        return dados  # Retornando os dados carregados

    except pyodbc.Error as e:
        st.error(f"Erro ao carregar os dados do banco: {e}")
        return pd.DataFrame()  # Retorna um DataFrame vazio em caso de erro
        # ou outra ação adequada ao tratamento de erro desejado


# 3 - OBDENTO OS DADOS DA FUNÇÃO load_data()
df = load_data()

# 4 - CRIAÇÃO DO SIDEBAR E BOTÕES

with st.sidebar:
    # Logo da Loja
    logo = Image.open('TM.png')
    st.sidebar.image(logo, use_column_width=True, caption='Logistica de Produtos')
    st.header('', divider='orange')

    # Criando o botão de seleção [Periodo]
    st.info("Periodo", icon="📆")
    st.write('Periodo da entregas: (Fevereiro/2019 á Dezembro/2021)')
    opcao1 = ["Selecione o periodo"] + list(df['Periodo'].unique())
    bt_periodo = st.sidebar.selectbox('', opcao1)
    st.header('', divider='green')

    # Criando o botão de seleção [Destino]
    st.info("Destino", icon="🛣️")
    opcao2 = ["Selecione o destino"] + list(df['Destino'].unique())
    bt_destino = st.sidebar.selectbox('Destino das entregas: ', opcao2)
    st.header('', divider='blue')

    # Criando o botão de seleção [Status]
    st.info("Status", icon="🚚")
    opcao3 = ['Selecione o status'] + list(df['Status'].unique())
    bt_status = st.radio(label='Status no prazo ou atrasado:', options=opcao3)

# 5 - CRIANDO OS FILTROS
df_status = df.loc[(df['Status'] == bt_status) & (df['Periodo'] == bt_periodo)]
df_total = df.loc[(df['Status'] == bt_status) & (df['Periodo'] == bt_periodo) & (df['Destino'] == bt_destino)]
df_destino = df.loc[(df['Status'] == bt_status) & (df['Destino'] == bt_destino) & (df['Periodo'] == bt_periodo)][['Latitude', 'Longitude']]



# 6 - CRIAÇÃO DAS COLUNAS
col1, col2, col3 = st.columns(3)
col4, col5 = st.columns(2)
col6, col7 = st.columns(2)
col8, col9 = st.columns(2)
col10, col11 = st.columns(2)


# SEQUÊNCIA DE CORES DAS BARRAS

color1 = ["#FD5A68", "#72B2E4", "#D06814", "#008996", "#31fe3d"]
color2 = ["#FD5A68", "#72B2E4", "#D06814", "#008996", "#31fe3d"] 
color3 = ["#a4ff2b", "#ea1c02", "#82143f", "#17c105", "#f839ac", "#161570"]
color4 = ["#FD5A68", "#72B2E4", "#D06814", "#008996"] 

# 7 - CRIANDO OS GRÁFICOS

# Card total
def coluna_total():
    
    total = df_total['Faturamento'].sum()
    col1.info('Total R$', icon='💲')
    col1.metric(label='Total das entregas', value="R$ " + f'{total:,.2f}')
    style_metric_cards(background_color="#4682B4", border_left_color="#9b2b2b")

coluna_total()


# Card total de entregas
def coluna_entrega():
    
    entregas = df_total['Pedido'].count()
    col2.info('Total de entregas', icon='🚛')
    col2.metric(label='Qtde', value=int(entregas))

coluna_entrega()

# Card quantidade de produtos
def coluna_produtos():
    
    produtos = df_total['Itens'].sum()
    col3.info('Total Produtos', icon='📦')
    col3.metric(label='Qtde', value=int(produtos))

coluna_produtos()

# Gráfico de barra entregas por motorista
def entregas_motoristas():
    
    col4.info('Entregas por Motorista', icon='📊')
    entregas = df_total.groupby('Motorista')[['Pedido']].count().reset_index()
    fig_entregas = px.bar(entregas,
                          y='Pedido',
                          x='Motorista',
                          text_auto='values',
                          color='Motorista',
                          color_discrete_sequence=color1)
    fig_entregas.update_traces(textfont_size=12, textangle=0, textposition="inside", cliponaxis=False)
    col4.plotly_chart(fig_entregas, use_container_width=True)

entregas_motoristas()

# Gráfico faturamento por motorista
def faturamento_motorista():
        
    col5.info('Faturamento por Motorista', icon='🍕')
    faturamento = df_total.groupby('Motorista')[['Faturamento']].sum().reset_index()
    fig_motorista = px.pie(faturamento,
                           values='Faturamento',
                           names='Motorista',
                           color_discrete_sequence=color2,
                           hole=.3)
    # Formatação dos números no gráfico de pizza
    fig_motorista.update_traces(textinfo='label+value+percent',
                                textposition='inside',  # Posicionamento do texto dentro das fatias
                                pull=[0.1, 0.1, 0.1, 0, 0.1, 0]) 
    
    col5.plotly_chart(fig_motorista, use_container_width=True)
    
faturamento_motorista()

# Gráfico faturamento por clientes
def faturamento_cliente():
    
    col6.info('Faturamento por Clientes', icon='🍕')
    faturamento = df_total.groupby('Cliente')[['Faturamento']].sum().reset_index()
    fig_faturamento = px.pie(faturamento,
                             values='Faturamento',
                             names='Cliente',
                             color_discrete_sequence=color2,
                             hole=.3)
    
    fig_faturamento.update_traces(textinfo='label+value',
                                  pull=[0.1, 0.1, 0.1, 0.1, 0.1, 0.1])
    
    col6.plotly_chart(fig_faturamento, use_container_width=True)

faturamento_cliente()

# Gráfico de faturamento de entregas dos estados
def entregas_estado():
        
    col7.info('Faturamento de entregas dos estados', icon='📉')
    estados = df_status.groupby('Destino')[['Faturamento']].sum().reset_index()
    fig_estados = px.bar(estados,
                         y='Destino',
                         x='Faturamento',
                         text_auto='.2s',
                         color='Destino',
                         orientation='h',
                         color_discrete_sequence=color3)

    fig_estados.update_traces(textfont_size=12, textangle=0, textposition="inside", cliponaxis=False)
    fig_estados.update_xaxes(showgrid=True, gridwidth=0.5, gridcolor="lightgray",
                             showline=True, linewidth=0.5, linecolor="lightgray")
    fig_estados.update_yaxes(showgrid=True, gridwidth=0.5, gridcolor="lightgray",
                             showline=True, linewidth=0.5, linecolor="lightgray")
    
    col7.plotly_chart(fig_estados, use_container_width=True)
    
entregas_estado()

# Gráfico de linha entregas de pedidos por data
def pedidos_data():
    
    col8.info('Entregas de pedidos por data', icon='〰️')
    linha = df_total.groupby('Data entrega')[['Pedido']].count().reset_index()
    fig_line = px.line(linha,
                       x="Data entrega",
                       y="Pedido",
                       text='Pedido',
                       markers=True)

    fig_line.update_traces(textposition="bottom right", line=dict(color="#ea1c02", width=2))
    fig_line.update_xaxes(showgrid=True, gridwidth=0.5, gridcolor="lightgray",
                          showline=True, linewidth=0.5, linecolor="lightgray")
    fig_line.update_yaxes(showgrid=True, gridwidth=0.5, gridcolor="lightgray",
                          showline=True, linewidth=0.5, linecolor="lightgray")
    
    col8.plotly_chart(fig_line, use_container_width=True)
    
pedidos_data()

# Gráfico de barra motivo das devoluções dos produtos
def devolucao_produtos():
    
    col9.info('Motivo das devoluções dos produtos', icon='📶')
    motivo = df_total.groupby('Motivo devolucao')[['Qtde devolucao']].sum().reset_index()
    fig_motivo = px.bar(motivo,
                        x="Motivo devolucao",
                        y="Qtde devolucao",
                        color="Motivo devolucao",
                        text_auto='values',
                        color_discrete_sequence=color4,

                        pattern_shape='Qtde devolucao',
                        pattern_shape_sequence=['.', 'x', '+'])

    fig_motivo.update_xaxes(showline=True, linewidth=0.5, linecolor="lightgray")
    fig_motivo.update_yaxes(showline=True, linewidth=0.5, linecolor="lightgray")

    fig_motivo.update_traces(textposition='outside')
    col9.plotly_chart(fig_motivo, use_container_width=True)

devolucao_produtos()



# Gráfico de barra motivo das devoluções dos produtos
def avaliacao():
    
    col10.info('Avaliação de devoluções', icon='📶')
    motivo = df_total.groupby('Avaliacao')[['Qtde devolucao']].count().reset_index()
    fig_motivo = px.bar(motivo,
                        x="Avaliacao",
                        y="Qtde devolucao",
                        color="Avaliacao",
                        text_auto='values',
                        color_discrete_sequence=color4,

                        pattern_shape='Qtde devolucao',
                        pattern_shape_sequence=['.', 'x', '+'])

    fig_motivo.update_xaxes(showline=True, linewidth=0.5, linecolor="lightgray")
    fig_motivo.update_yaxes(showline=True, linewidth=0.5, linecolor="lightgray")

    fig_motivo.update_traces(textposition='outside')
    col10.plotly_chart(fig_motivo, use_container_width=True)

avaliacao()

# Tabela logistica
with col11:
        
    st.sidebar.info('Tabela logistica', icon='📅')
    st.info('Relatório da logistica', icon='📅')
    tabela = st.sidebar.multiselect('Selecione as colunas:', df_total.columns.to_list(), default=[])

    if tabela:  # Verifica se a lista de colunas não está vazia
        # Filtra o DataFrame para exibir apenas as colunas selecionadas
        df_display = df_total[tabela]
        if not df_display.empty:
            st.write(df_display)

            # Botão para exportar o DataFrame para um arquivo .csv
            if st.button('Exportar como CSV'):
                filename = "Dataframe.csv"

                # Seu DataFrame
                data = df_display
                df_display = pd.DataFrame(data)

                # Obtendo o caminho para a pasta de downloads do usuário
                path_csv = os.path.join(os.path.expanduser("~"), "Downloads", filename)

                # Salvando o DataFrame como um arquivo CSV na pasta de downloads
                df_display.to_csv(path_csv, index=False, sep=';')
                st.success(f'DataFrame exportado com sucesso como {path_csv}')

            # Botão para exportar o DataFrame para um arquivo .xlsx
            if st.button('Exportar como Excel'):
                filename = "Dataframe.xlsx"

                # Seu DataFrame
                data = df_display
                df_display = pd.DataFrame(data)

                # Obtendo o caminho para a pasta de downloads do usuário
                path_xlsx = os.path.join(os.path.expanduser("~"), "Downloads", filename)

                # Salvando o DataFrame para um arquivo Excel (.xlsx)
                df_display.to_excel(path_xlsx, index=False)
                st.success(f'DataFrame exportado com sucesso como {path_xlsx}')


            # Função para gerar o PDF
            def generate_pdf(dataframe):
                pdf = FPDF()
                pdf.set_auto_page_break(auto=1, margin=10)
                pdf.add_page()
                pdf.set_font("Arial", "B", size=12)
                pdf.cell(200, 10, "Análise de logística de produtos", ln=True, align="C")

                col_names = dataframe.columns.values.tolist()
                data = dataframe.values.tolist()

                # Imprimir os cabeçalhos das colunas
                pdf.set_font("Arial", "B", size=10)
                for col_name in col_names:
                    pdf.cell(30, 8, col_name, border=1, align='C')
                pdf.ln()

                # Imprimir os dados do DataFrame
                pdf.set_font("Arial", size=8)
                for row in data:
                    for item in row:
                        pdf.cell(30, 8, str(item), border=1)
                    pdf.ln()

                return pdf


            if st.button('Exportar como PDF', key='pdf'):
                pdf_filename = "Dataframe.pdf"
                # Gera o PDF
                pdf = generate_pdf(df_display)

                # Obtendo o caminho para a pasta de downloads do usuário
                path_pdf = os.path.join(os.path.expanduser("~"), "Downloads", pdf_filename)

                # Salvando o PDF na pasta de downloads
                pdf.output(path_pdf)
                st.success(f'DataFrame exportado com sucesso como {path_pdf}')

# Mapa de localização das lojas
def localizacao_loja():
    
    st.info('Localização das lojas', icon='🗺️')

# Verificando se as coordenadas são encontradas
    if not df_destino.empty:
# Criando um mapa com base nas coordenadas da cidade selecionada
        m = folium.Map(location=[df_destino['Latitude'].iloc[0], df_destino['Longitude'].iloc[0]], zoom_start=12, width='100%' , height='100%')

# Adicionando um marcador com a localização da cidade selecionada
        folium.Marker(
            [df_destino['Latitude'].iloc[0], 
             df_destino['Longitude'].iloc[0]],
             popup='localizacao',
             tooltip='localizacao'
        ).add_to(m)

# Exibindo o mapa no Streamlit
        folium_static(m)
    else:
         st.warning("O destino ainda não foi selecionado.")    

localizacao_loja()