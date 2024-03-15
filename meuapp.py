import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from bokeh.plotting import figure
from bokeh.models import HoverTool, ColumnDataSource
from bokeh.models import TabPanel, Tabs
from bokeh.models import CustomJS, RangeSlider
from bokeh.layouts import row, column



df_cpm = None
excel_file_path = ""
report_file_path = ""

st.set_page_config(page_title="Pricing Plot", layout="wide")

def carrega_excel():
    global df_cpm, excel_file_path        
    uploaded_file = st.file_uploader('Insira um arquivo por vez do formato planilha (xlsx/xls/csv):', type=(["xlsx", "csv", "xls"]))
    
    if uploaded_file:
        # Leitura e atribuição do arquivo excel ao arquivo de upload.
        excel_file_path = uploaded_file
        df_cpm = pd.read_excel(uploaded_file, sheet_name='CPM', na_filter=False)


def gerar_relatorio():
    
    carrega_excel()
    
    global report_file_path
    if df_cpm is None:
        st.warning("Por favor, carregue um arquivo do formato especificado.")
        return

       
    table = df_cpm.loc[:, "Unnamed: 0":"Unnamed: 10"]
    table2 = table.dropna(axis=0, how='all').dropna(axis=1, how='all')
    table2 = table2.iloc[0:50, :]

    table2['Unnamed: 7'] = pd.to_numeric(table2['Unnamed: 7'], errors='coerce')
    table2['Unnamed: 9'] = pd.to_numeric(table2['Unnamed: 9'], errors='coerce')
    filtered_data = table2[table2['Unnamed: 9'] < 100]

    # Criar a origem dos dados
    source = ColumnDataSource(data=dict(x=filtered_data['Unnamed: 7'], y=filtered_data['Unnamed: 9'], labels=filtered_data['Unnamed: 1'], preco=filtered_data['Unnamed: 8']))
                        
    # Criar o gráfico de dispersão
    p = figure(title="Gráfico de Dispersão - CPM x Nro de Ouvintes por Minuto", tools="pan,box_zoom,reset,save,zoom_in, zoom_out", background_fill_color='lightblue', width=900, height=600)
    points = p.scatter('x', 'y', size=10, source=source, fill_alpha=0.6, line_color=None, legend_label="Veículos")

    range_slider_cpm = RangeSlider(start=0.00, end=90.00, value=(0,90), step=.1, title="CPM", width=200, margin=10)

    range_slider_cpm.js_on_change('value', CustomJS(args=dict(c=p.y_range), code="""
                                                    c.start = this.value[0];                                      
                                                    c.end = this.value[1];
                                                    """)
                                )

    range_slider_audiencia = RangeSlider(start=0.00, end=90000.00, value=(0,90000), step=.1, title="Audiência", width=200, margin=10)

    range_slider_audiencia.js_on_change('value', CustomJS(args=dict(a=p.x_range), code="""
                                                        a.start = this.value[0];
                                                        a.end = this.value[1];                                                                                                            
                                                        """)
                                        )

    layout = (row(column(range_slider_audiencia, range_slider_cpm), p))
    tabCPM = TabPanel(child=layout, title="CPM")

    # Ativa o autohide para mostrar a barra de ferramentas somente quando o mouse estiver em cima dela.
    p.toolbar.autohide = True

    # Adicionar ferramenta de dica interativa
    hover = HoverTool(renderers=[points], tooltips=[("Veículo", "@labels"), ("CPM", "@y{sep='.'}"), ("Audiência", "@x{0}"), ("Preço 30'", "@preco")])
    p.add_tools(hover)

    # Adicionar rótulos com contorno
    texts = []
    text_offset = 0.015                                                              # Ajuste manual da posição vertical do texto

    for i, txt in enumerate(table2['Unnamed: 1']):
        color = None

        # Definir cores com base no nome
        if "CBN" in txt:
            color = "red"
        elif "GAZ" in txt:
            color = "blue"
        elif "LIT" in txt:
            color = "brown"
        elif "MIX" in txt:
            color = "green"
        elif "TV" in txt:
            color = "black"
        text = p.text(
            x=float(table2['Unnamed: 7'].iloc[i]) + 0.1,
            y=float(table2['Unnamed: 9'].iloc[i]) - text_offset,
            text=[str(txt)],
            text_font_size="7pt",
            text_align="left",
            angle= 0.6,
            text_baseline="bottom",
            text_color=color)                                                    # Definir a cor da fonte
        texts.append(text)

    # Configurar rótulos dos eixos
    p.xaxis.axis_label = "Audiência (%)"
    p.yaxis.axis_label = "CPM (%)"


    #Lê a aba Tx Ocupação (somente uma tabela) e pula a linha de índice 1.
    df2 = pd.read_excel(excel_file_path, sheet_name='Tx Ocupação', na_filter=False)
    table = df2.loc[:, "Unnamed: 1":"Unnamed: 12"]
    table3 = table.dropna(axis=0, how='all').dropna(axis=1, how='all')                                               # Remove valores nulos de linhas e colunas.            
    table3 = table3.iloc[0:75, :]                                                                                    # Percorrendo linhas da tabela.

    table3['Unnamed: 7'] = pd.to_numeric(table3['Unnamed: 7']*100, errors='coerce', downcast='float')
    table3['Unnamed: 12'] = pd.to_numeric(table3['Unnamed: 12']*100, errors='coerce', downcast='float')

    # Cria a origem dos dados
    source2 = ColumnDataSource(data=dict(x=table3['Unnamed: 7'], y=table3['Unnamed: 12'], labels=table3['Unnamed: 1'] + ' / ' + table3['Unnamed: 2']))

    # Cria o gráfico de dispersão
    p2 = figure(title="Gráfico de Dispersão - Ocupação (Faixa) x Desconto (Faixa)", tools="pan,box_zoom,reset,save,zoom_in, zoom_out", background_fill_color='lightblue', width=900, height=600)
    points = p2.scatter('x', 'y', size=10, source=source2, fill_alpha=0.6, line_color=None, legend_label="Veículos")

    range_slider_ocupacao = RangeSlider(start=0.00, end=100.00, value=(0,100), step=.1, title="Ocupação", width=200, margin=10)

    range_slider_ocupacao.js_on_change('value', CustomJS(args=dict(o=p2.x_range) ,code="""
                                                o.start = this.value[0];
                                                o.end = this.value[1];                                                                                                                     
                                                """
                                                )
    )

    range_slider_desconto = RangeSlider(start=0.00, end=100.00, value=(0,100), step=.1, title="Desconto", width=200, margin=10)

    range_slider_desconto.js_on_change('value', CustomJS(args=dict(d=p2.y_range), code="""
                                                d.start = this.value[0];
                                                d.end = this.value[1];                                                                                                                    
                                                """
                                                )
    )

    layout2 = (row(column(range_slider_ocupacao, range_slider_desconto), p2))
    tabDesc = TabPanel(child=layout2, title="Ocupação/Desconto")

    # Ativa o autohide para mostrar a barra de ferramentas somente quando o mouse estiver em cima dela.
    p2.toolbar.autohide = True

    # Adiciona a ferramenta de dica interativa
    hover = HoverTool(renderers=[points], tooltips=[("Veículo", "@labels"), ("Ocupação (Faixa)", "@x{0.00}%"), ("Desconto (Faixa)", "@y{0.00}%")])
    p2.add_tools(hover)

    # Adiciona rótulos com contorno
    texts = []
    text_offset = 0.015                                                              # Ajuste manual da posição vertical do texto

    for i, txt in enumerate(table3['Unnamed: 1'] + ' / ' + table3["Unnamed: 2"]):
        color = None

        # Definir cores com base no nome
        if "LITORAL" in txt:
            color = "red"
        elif "GAZETA" in txt:
            color = "blue"
        text = p2.text(
            x= table3['Unnamed: 7'].iloc[i] + 0.1,
            y= table3['Unnamed: 12'].iloc[i] - text_offset,
            text=[str(txt)],
            text_font_size="7pt",
            text_align="left",
            angle= 0.6,
            text_baseline="bottom",
            text_color=color)                                                    # Definir a cor da fonte
        texts.append(text)

    # Configurar rótulos dos eixos
    p2.xaxis.axis_label = "Ocupação (Faixa) %"
    p2.yaxis.axis_label = "Desconto (Faixa) %"

    abas = Tabs(tabs=[tabCPM, tabDesc]) 
    
    with st.container():                                                                # Container que exibe o html do gráfico.      
        HtmlFile = open("relatorio.html", 'r', encoding='utf-8')
        source_code = HtmlFile.read()
        st.write("---")
        components.html(source_code, width=1300, height=1000)
  
        

with st.container():
    st.subheader("Gráficos de CPM x Audiência / Ocupação x Desconto")
    st.write("Informações sobre os veículos da Rede.")

def abrir_relatorio():
    gerar_relatorio()

with st.container():
    st.write("---")
    abrir_relatorio()

