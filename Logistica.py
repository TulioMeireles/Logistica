# Importando as bibliotecas
import pandas as pd
import plotly_express as px
import streamlit as st
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
@st.cache_resource
def load_data():
    try:
        conn = pyodbc.connect('Driver={ODBC Driver 11 for SQL Server};'
                              'Server=DESKTOP-D3SPOTC\SQLEXPRESS;'
                              'Username=sa;'
                              'Password=245316;'
                              'Database=Powerbi;'
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
    st.write('Entregas entre o periodo: (Fevereiro-2019 á Dezembro-2021)')
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

df_local = df.loc[df['Destino'] == bt_destino][['Latitude', 'Longitude', 'Destino']]

# 6 - CRIAÇÃO DAS COLUNAS
col1, col2, col3, col4 = st.columns(4)
col5, col6 = st.columns(2)
col7, col8 = st.columns(2)
col9, col10 = st.columns(2)
col11, col12 = st.columns(2)

# 7 - CRIANDO OS GRÁFICOS

# Card total
with col1:
    total = df_total['Faturamento'].sum()
    st.info('Total R$', icon='💲')
    st.metric(label='Total das entregas', value="R$ " + f'{total:,.2f}')
    style_metric_cards(background_color="#4682B4", border_left_color="#9b2b2b")

# Card média
with col2:
    media = df_total['Faturamento'].mean()
    st.info('Média R$', icon='🔢')
    st.metric(label='Média das entregas', value="R$ " + f'{media:,.2f}')

# Card total de entregas
with col3:
    entregas = df_total['Pedido'].count()
    st.info('Total de entregas', icon='🚛')
    st.metric(label='Qtde', value=int(entregas))

# Card quantidade de produtos
with col4:
    produtos = df_total['Itens'].sum()
    st.info('Total Produtos', icon='📦')
    st.metric(label='Qtde', value=int(produtos))

# Gráfico de barra entregas por motorista
with col5:
    color1 = ["#FD5A68", "#72B2E4", "#D06814", "#008996", "#31fe3d"]
    st.info('Entregas por Motorista', icon='📊')
    entregas = df_total.groupby('Motorista')[['Pedido']].count().reset_index()
    fig_entregas = px.bar(entregas,
                          y='Pedido',
                          x='Motorista',
                          text_auto='values',
                          color='Motorista',
                          color_discrete_sequence=color1)
    fig_entregas.update_traces(textfont_size=12, textangle=0, textposition="inside", cliponaxis=False)
    col5.plotly_chart(fig_entregas, use_container_width=True)

# Gráfico faturamento por motorista
with col6:
    color2 = ["#FD5A68", "#72B2E4", "#D06814", "#008996", "#31fe3d"]
    st.info('Faturamento por Motorista', icon='🍕')
    faturamento = df_total.groupby('Motorista')[['Faturamento']].sum().reset_index()
    fig_motorista = px.pie(faturamento,
                           values='Faturamento',
                           names='Motorista',
                           color_discrete_sequence=color2,
                           hole=.3)
    # Formatação dos números no gráfico de pizza
    fig_motorista.update_traces(textinfo='label+value+percent',
                                textposition='inside')  # Posicionamento do texto dentro das fatias
    col6.plotly_chart(fig_motorista, use_container_width=True)

# Gráfico faturamento por clientes
with col7:
    color2 = ["#FD5A68", "#72B2E4", "#D06814", "#008996", "#31fe3d"]
    st.info('Faturamento por Cliente', icon='🍕')
    faturamento = df_total.groupby('Cliente')[['Faturamento']].sum().reset_index()
    fig_faturamento = px.pie(faturamento,
                             values='Faturamento',
                             names='Cliente',
                             color_discrete_sequence=color2,
                             hole=.3)
    fig_faturamento.update_traces(textinfo='label+value')
    col7.plotly_chart(fig_faturamento, use_container_width=True)

# Gráfico de faturamento de entregas por estados
with col8:
    color3 = ["#a4ff2b", "#ea1c02", "#82143f", "#17c105", "#f839ac", "#161570"]
    st.info('Faturamento de entregas por estados', icon='📉')
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
    col8.plotly_chart(fig_estados, use_container_width=True)

# Gráfico de linha entregas de pedidos por data
with col9:
    st.info('Entregas de pedidos por data', icon='〰️')
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
    col9.plotly_chart(fig_line, use_container_width=True)

# Gráfico de barra motivo das devoluções dos produtos
with col10:
    color4 = ["#FD5A68", "#72B2E4", "#D06814", "#008996"]
    st.info('Motivo das devoluções dos produtos', icon='📶')
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
    col10.plotly_chart(fig_motivo, use_container_width=True)

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

# (9) Mapa de localização das lojas
with col12:
    # Criando o filtro para adicionar a coluna faturamento (soma das vendas por estado)
    faturamento_estado = df_total.groupby('Destino')['Faturamento'].sum().reset_index()
    df_local = pd.merge(df_local, faturamento_estado, on='Destino')  # Corrigindo o nome da coluna para o merge
    

    # Criando o mapa com Plotly Express
    st.info("Localização da loja", icon="🗺️")
    fig = px.scatter_mapbox(df_local,
                            lat='Latitude',
                            lon='Longitude',
                            hover_name='Destino',
                            color='Destino',
                            size='Faturamento',
                            size_max=10)

    # Configura o layout do mapa
    fig.update_layout(
        mapbox_style="open-street-map",  # Estilo do mapa (pode ser "carto-positron", "carto-darkmatter", etc.)
        mapbox_zoom=6,  # Zoom inicial do mapa
        mapbox_center={"lat": df_local['Latitude'].mean(), "lon": df_local['Longitude'].mean()}  # Centro do mapa
    )

    # Exibe o mapa
    col12.plotly_chart(fig, use_container_width=True)

# Rodapé
st.subheader("", divider="green")
st.markdown("<p style='text-align: center;'><b>Desenvolvido por: Tulio Meireles</b></p>", unsafe_allow_html=True)
