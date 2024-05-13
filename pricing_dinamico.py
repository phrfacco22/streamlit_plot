import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
from bokeh.plotting import figure, show
from bokeh.models import HoverTool, ColumnDataSource
from bokeh.models import TabPanel, Tabs
from bokeh.models import CustomJS, RangeSlider, CheckboxGroup
from bokeh.layouts import row, column



df = None
excel_file_path = ""
report_file_path = ""

st.set_page_config(page_title="Pricing Plot", layout="wide")

def carrega_excel():
    global df, excel_file_path        
    uploaded_file = st.file_uploader('Insira um arquivo por vez do formato planilha (xlsx/xls/csv):', type=(["xlsx", "csv", "xls"]))
    
    if uploaded_file:
        # Leitura e atribuição do arquivo excel ao arquivo de upload.
        excel_file_path = uploaded_file
        df = pd.read_excel(uploaded_file)


def gerar_relatorio():
    
    carrega_excel()
    
    global report_file_path
    if df is None:
        st.warning("Por favor, carregue um arquivo do formato especificado.")
        return

    df_cpm = pd.read_excel(excel_file_path, sheet_name='CPM', na_filter=True)
    table = df_cpm.loc[:, "Unnamed: 1":"Unnamed: 10"]
    table2 = table.dropna(axis=0, how='all').dropna(axis=1, how='all')
    table2 = table2.iloc[1:50, :]

    table2['Unnamed: 7'] = pd.to_numeric(table2['Unnamed: 7'], errors='coerce')
    table2['Unnamed: 9'] = pd.to_numeric(table2['Unnamed: 9'], errors='coerce')

    # Criar a origem dos dados
    source = ColumnDataSource(data=dict(x=table2['Unnamed: 7'], y=table2['Unnamed: 9'], labels=table2['Unnamed: 1'], preco=table2['Unnamed: 8']))
    
    source_x = source.data['x']
    source_y = source.data['y']
    source_labels = source.data['labels']
    source_preco = source.data['preco']
    
                        
    # Criar o gráfico de dispersão de cpm/audiência
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
    
    # Lista de praças
    lista_praca = list(table2['Unnamed: 1'])

     # Legendas dos pontos plotados no gráfico.
    labels = LabelSet(x='x', y='y', text='labels', x_offset=2.5, y_offset=2.5, source=source, text_font_size="7pt", text_color='brown', angle= 0.6)

    # Checkbox Group para seleção de praças
    checkbox_praca = CheckboxGroup(labels=lista_praca)

    checkbox_praca.js_on_change('active', CustomJS(args=dict(source=source, source_x=source_x, source_y=source_y, source_labels=source_labels, source_preco=source_preco), code="""
        
        const selected_indices = cb_obj.active;
           
		if (selected_indices.length == 0) {
                source.data['x'] = source_x;
                source.data['y'] = source_y;
                source.data['labels'] = source_labels; 
                source.data['preco'] = source_preco;
		} else if (selected_indices.length > 0) {
            source.data = {
                'x': source_x.filter((x, i) => selected_indices.includes(i)),
                'y': source_y.filter((y, i) => selected_indices.includes(i)),
                'labels': source_labels.filter((label, i) => selected_indices.includes(i)),
                'preco': source_preco.filter((preco, i) => selected_indices.includes(i)),
            };
        }
  
        source.change.emit();
    """))
    
    layout = (row(column(range_slider_cpm, range_slider_audiencia, checkbox_praca), p))
    tabCPM = TabPanel(child=layout, title="CPM")

    # Ativa o autohide para mostrar a barra de ferramentas somente quando o mouse estiver em cima dela.
    p.toolbar.autohide = True

    # Adicionar ferramenta de dica interativa
    hover_audiencia = HoverTool(renderers=[points], tooltips=[("Veículo", "@labels"), ("CPM", "@y{sep='.'}"), ("Audiência", "@x{0}"), ("Preço 30'", "@preco")])
    p.add_tools(hover_audiencia) 
    
    # Configurar rótulos dos eixos
    p.xaxis.axis_label = "Audiência (%)"
    p.yaxis.axis_label = "CPM (%)"

    
    #Lê a aba Tx Ocupação (somente uma tabela) e pula a linha de índice 1.
    df2 = pd.read_excel(excel_file_path, sheet_name='Tx Ocupação')
    table = df2.loc[:, "Unnamed: 1":"Unnamed: 12"]
    table3 = table.dropna(thresh=6).dropna(axis=1, how='all')                                                                       # Remove linhas com valores nulos de linhas e colunas.            
    table3 = table3.iloc[3:75]                                                                                                      # Percorrendo linhas da tabela.
    mask = (table3['Unnamed: 1'] == 'Emissora')
    table3 = table3.loc[~mask]
    
    table3['Unnamed: 7'] = pd.to_numeric(table3['Unnamed: 7'] * 100, errors='coerce', downcast='float')
    table3['Unnamed: 12'] = pd.to_numeric(table3['Unnamed: 12'] * 100, errors='coerce', downcast='float')
    
    
    # Cria a origem dos dados
    source2 = ColumnDataSource(data=dict(x=table3['Unnamed: 7'], y=table3['Unnamed: 12'], labels=table3['Unnamed: 1'], cpm=table2['Unnamed: 9'], audiencia=table2['Unnamed: 7']))

    source_x = source2.data['x']
    source_y = source2.data['y']
    source_labels = source2.data['labels']
    source_cpm = source2.data['cpm']
    source_audiencia = source2.data['audiencia']


    # Cria o gráfico de dispersão
    p2 = figure(title="Gráfico de Dispersão - Ocupação (Faixa) x Desconto (Faixa)", tools="pan,box_zoom,reset,save,zoom_in, zoom_out", background_fill_color='lightblue', width=900, height=600)
    points = p2.scatter('x', 'y', size=10, source=source2, fill_alpha=0.6, line_color=None, legend_label="Veículos")

    range_slider_ocupacao = RangeSlider(start=0.00, end=100.00, value=(0,100), step=.1, title="Ocupação", width=200, margin=10)
    range_slider_ocupacao.js_on_change('value', CustomJS(args=dict(o=p2.x_range) ,code="""
            o.start = this.value[0];
            o.end = this.value[1];
    """))

    range_slider_desconto = RangeSlider(start=0.00, end=100.00, value=(0,100), step=.1, title="Desconto", width=200, margin=10)
    range_slider_desconto.js_on_change('value', CustomJS(args=dict(d=p2.y_range), code="""
            d.start = this.value[0];
            d.end = this.value[1];
    """))
    
    labels_ocupacao = list(table3['Unnamed: 1'] + ' ' + table3['Unnamed: 2'])
    
    # Checkbox Group para seleção de praças
    checkbox_praca_ocupacao = CheckboxGroup(labels=labels_ocupacao)

    # Define função callback para grupo checkbox
    checkbox_praca_ocupacao.js_on_change('active', CustomJS(args=dict(source=source2, source_x=source_x, source_y=source_y, source_labels=source_labels, source_cpm=source_cpm, source_audiencia=source_audiencia), code="""
        
        const selected_indices = cb_obj.active;
           
		if (selected_indices.length === 0) {
                source.data['x'] = source_x;
                source.data['y'] = source_y;
                source.data['labels'] = source_labels; 
                source.data['cpm'] = source_cpm;
                source.data['audiencia'] = source_audiencia;
        } else if (selected_indices.length > 0) {
            source.data = {
                'x': source_x.filter((x, i) => selected_indices.includes(i)),
                'y': source_y.filter((y, i) => selected_indices.includes(i)),
                'labels': source_labels.filter((label, i) => selected_indices.includes(i)),
                'cpm': source_cpm.filter((cpm, i) => selected_indices.includes(i)),
                'audiencia': source_audiencia.filter((audiencia, i) => selected_indices.includes(i)),
            };
        }

        // Trigger source.change event to update the plot
        source.change.emit();
"""))


    # Ativa o autohide para mostrar a barra de ferramentas somente quando o mouse estiver em cima dela.
    p2.toolbar.autohide = True
    
    # Adiciona a ferramenta de dica interativa
    hover_ocupacao = HoverTool(renderers=[points], tooltips=[("Veículo", "@labels"), ("Ocupação (Faixa)", "@x{0.00}%"), ("Desconto (Faixa)", "@y{0.00}%"), ("Audiência", "@audiencia{0}"), ("CPM", "@cpm{0.0}")])
    p2.add_tools(hover_ocupacao)
    
    # Adiciona rótulos com contorno
    texts = []
    text_offset = 0.015                                                              # Ajuste manual da posição vertical do texto

    labels_ocupacao = list(source_labels)
    
    # Legendas dos pontos plotados no gráfico.
    labels_ocup = LabelSet(x='x', y='y', text='labels', x_offset=2.5, y_offset=2.5, source=source2, text_font_size="7pt", text_color='blue', angle= 0.6)
       
    # Configurar rótulos dos eixos
    p2.xaxis.axis_label = "Ocupação (Faixa) %"
    p2.yaxis.axis_label = "Desconto (Faixa) %"
   
    layout2 = (row(column(range_slider_ocupacao, range_slider_desconto, checkbox_praca_ocupacao), p2))
    tabDesc = TabPanel(child=layout2, title="Ocupação/Desconto")

    abas = Tabs(tabs=[tabCPM, tabDesc])
    
    with st.container():                                                                # Container que exibe o html do gráfico.      
        HtmlFile = open("relatorio.html", 'r', encoding='utf-8')
        source_code = HtmlFile.read()
        st.write("---")
        components.html(source_code, width=1200, height=1000)
  
        

with st.container():
    st.subheader("Gráficos de CPM x Audiência / Ocupação x Desconto")
    st.write("Informações sobre os veículos da Rede.")

def abrir_relatorio():
    gerar_relatorio()

with st.container():
    st.write("---")
    abrir_relatorio()

