import streamlit as st
import pandas as pd
# import locale
from io import BytesIO
from PIL import Image
import base64
import xlrd

# Definindo o locale para o padrão brasileiro
# locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
# Tente definir a localidade para pt_BR.UTF-8
# try:
#     locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
# except locale.Error:
#     st.warning("Localidade 'pt_BR.utf8' não suportada. Usando localidade padrão.")
#     locale.setlocale(locale.LC_ALL, '')


# Configurar o layout da página para largura ampla
st.set_page_config(layout="wide", page_title="Conciliação de Inventário Rotativo", page_icon = './logo2.png')

# Função para carregar a imagem
def carregar_imagem(caminho_imagem):
    with open(caminho_imagem, "rb") as img_file:
        img_bytes = img_file.read()
        encoded_img = base64.b64encode(img_bytes).decode()
        return encoded_img

# Caminho da imagem (altere para o caminho correto da sua imagem)
caminho_imagem_topo = "C:/Users/AmaraNzero/Documents/AmaraBrasil/vs_code/ConciliadorInventarioRotativo/image.png"
caminho_imagem_rodape = "./logo_branca.png"

# Adicionando uma imagem centralizada acima do título com tamanho reduzido
st.image('./pagina-logo.jpg',width=1480, use_column_width=False)

# Função para converter imagem para base64
# def img_to_bytes(img_path):
#     img = Image.open(img_path)
#     buffer = BytesIO()
#     img.save(buffer, format="PNG")
#     img_str = base64.b64encode(buffer.getvalue()).decode()
#     return img_str

# # Caminho da imagem
# image_path = 'C:/Users/AmaraNzero/Documents/AmaraBrasil/vs_code/ConciliadorInventarioRotativo/pagina-logo.jpg'
# image_base64 = img_to_bytes(image_path)

# Centralizando a imagem usando HTML
# st.markdown(f"""
#     <div style='text-align: center;'>
#         <img src='data:image/png;base64,{image_base64}' width='300'>
#     </div>
#     """, unsafe_allow_html=True)  
# st.header('Hello')


st.header('')

hide_st_style = '''
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
                </style>
'''

st.markdown(hide_st_style, unsafe_allow_html=True)

# Customização do título com HTML e CSS
st.markdown("""
    <h1 style='text-align: center; color: #009000;'>
        Conciliação de Inventário Rotativo
    </h1>
    """, unsafe_allow_html=True)


# Modelagem df1:

def modelagem_df1(df):
    colunas_selecionadas = [
        'Centro', 'Depósito', 'Lote', 'Material', 'Texto breve de material',
        'UM básica', 'Utilização livre', 'Val.utiliz.livre', 'Bloqueado', 'Val.estoque bloq.'
    ]
    
    # Filtrar o DataFrame apenas pelas colunas desejadas
    df_selecionado = df[colunas_selecionadas]

    # Converter a coluna 'Material' para string no DataFrame selecionado
    df_selecionado['Material'] = df_selecionado['Material'].astype(str).str.replace('.0', '')


    # Agregar os valores
    df_agg = df.groupby('Material').agg({
        'Utilização livre': 'sum',
        'Val.utiliz.livre': 'sum',
        'Bloqueado': 'sum',
        'Val.estoque bloq.': 'sum',
        'UM básica': 'first'
    }).reset_index()    

    # Converter a coluna 'Material' para string
    df_agg['Material'] = df_agg['Material'].astype(str).str.replace('.0', '')


   # Trocar vírgula por ponto apenas nas colunas 'Val.utliz.livre' e 'Val.estoque bloq.'
    df_agg['Val.utiliz.livre'] = df_agg['Val.utiliz.livre'].astype(str).str.replace(',', '.').astype(float)
    df_agg['Val.estoque bloq.'] = df_agg['Val.estoque bloq.'].astype(str).str.replace(',', '.').astype(float)
    
    # Selecionar apenas as colunas desejadas após a agregação
    df_agg = df_agg[['Material', 'UM básica', 'Utilização livre', 'Val.utiliz.livre',  'Bloqueado', 'Val.estoque bloq.']]

    # Mesclar o DataFrame agregado com o DataFrame original
    df_final = pd.merge(df_agg, df_selecionado[['Material', 'Texto breve de material']], on='Material', how='left')

     # Remover duplicatas
    df_final = df_final.drop_duplicates()

    # Organizar a ordem das colunas
    colunas_organizadas = [
        'Material', 'Texto breve de material', 'UM básica', 'Utilização livre', 'Val.utiliz.livre', 'Bloqueado', 'Val.estoque bloq.'
    ]
    df_final = df_final.reindex(columns=colunas_organizadas)
    
    return df_final

# Função para verificar as colunas
def verificar_colunas(df, colunas_esperadas):
    colunas_faltantes = [col for col in colunas_esperadas if col not in df.columns]
    if colunas_faltantes:
        raise ValueError(f"Colunas faltantes: {colunas_faltantes}")
    
# Instruções para o usuário
# st.write("Faça o upload de uma planilha para visualizar os dados tratados.")

# Upload do arquivo
uploaded_file1 = st.file_uploader("Carregue a planilha SAP MB52", type=['xlsx', 'xls'])

# Verificação se o arquivo foi carregado
if uploaded_file1:
    # Leitura da planilha 1
    df1 = pd.read_excel(uploaded_file1)

       
    # Tratar a planilha 1
    df_final = modelagem_df1(df1)
    
    # Exibição da planilha final
   # st.write("SAP MB52 modelada:")
    #st.write(df_final)
else: 
    st.warning("Por favor, faça o upload da planilha SAP")

def modelagem_df2(df):
    colunas_selecionadas_2 = [
        'Endereço', 'Produto', 'Descrição', 'Empenhada', 'Bloqueado', 'Qualidade', 'Saldo'
    ]
    
    # Filtrar o DataFrame apenas pelas colunas desejadas
    df_selecionado = df[colunas_selecionadas_2]

    # Criar a coluna 'Saldo wms'
    df_selecionado['Saldo wms'] = df_selecionado['Empenhada'] + df_selecionado['Bloqueado'] + df_selecionado['Qualidade'] + df_selecionado['Saldo']

    # Converter a coluna 'Produto' para string
    df_selecionado['Produto'] = df_selecionado['Produto'].astype(str).str.replace('.0', '')

    # Agregar os valores por produto
    df_agregado = df_selecionado.groupby('Produto').agg({
        'Empenhada': 'sum',
        'Saldo wms': 'sum',
        'Descrição': 'first'  # Manter a descrição do primeiro registro
    }).reset_index()

    return df_agregado
# Função ajustada modelagem_df2 com verificações e logs

# # Função para padrão 2 wms
def modelagem_df3(df):
    # renomear as colunas para o padrão original
    df.rename(columns={
        'Cód. Produto': 'Produto',
        'Desc. Produto': 'Descrição',
        'Qtde Empenhada': 'Empenhada',
        'Qtde Bloqueada': 'Bloqueado',
        'Qtde Bloq. Qualidade': 'Qualidade',
        'Qtde Saldo': 'Saldo'
    }, inplace = True)

    colunas_selecionas_2 = [
        'Endereço', 'Produto', 'Descrição',  'Empenhada', 'Bloqueado', 'Qualidade', 'Saldo'
    ]

    # Criar o dataframe com as colunas selecionadas
    df_selecionado = df[colunas_selecionas_2]

    # Alterar o tipo de dado da coluna 'Produto' para float64
    df_selecionado['Produto'] = df_selecionado['Produto'].astype('float64')

    # Criar a coluna 'Saldo wms'
    df_selecionado['Saldo wms'] = df_selecionado['Empenhada'] + df_selecionado['Bloqueado'] + df_selecionado['Qualidade'] + df_selecionado['Saldo']

    # Converter a coluna 'Produto' para string
    df_selecionado['Produto'] = df_selecionado['Produto'].astype(str).str.replace('.0', '')

    # Agregar os valores por produto
    df_agregado = df_selecionado.groupby('Produto').agg({
        'Empenhada': 'sum',
        'Saldo wms': 'sum',
        'Descrição': 'first'  # Manter a descrição do primeiro registro
    }).reset_index()

    return df_agregado


# Instruções para o usuário
#st.write("Faça o upload de uma planilha para visualizar os dados tratados.")

# Upload do arquivo
uploaded_file2 = st.file_uploader("Carregue a planilha WMS Sintético", type=['xlsx', 'xls'])

# Verificação se o arquivo foi carregado
if uploaded_file2:
    # Leitura da planilha 1
    df2 = pd.read_excel(uploaded_file2)
    
    # Tratar a planilha 1
    df2_final = modelagem_df1(df1)
                                                                                            
    # Exibição da planilha final
   # st.write("SAP MB52 modelada:")
    #st.write(df_final)
else: 
    st.warning("Por favor, faça o upload da planilha WMS")

# Verificação se os arquivos foram carregados
if uploaded_file1 and uploaded_file2:
    try:
        # Leitura das planilhas
        df1 = pd.read_excel(uploaded_file1)
        
        # Tratamento das planilhas
        df1_final = modelagem_df1(df1)

        # Lendo uploaded_file2 como um DataFrame pandas
          # Pula as duas primeiras linhas
        df2 = pd.read_excel(uploaded_file2, skiprows=11)  # Pula as duas primeiras linhas
        
        # Identificar o padrão da planilha WMS para determinar qual função de modelagem usar
        if 'Produto' not in df2.columns and 'Descrição' not in df2.columns:
           df2 = pd.read_excel(uploaded_file2, skiprows=2)  # Usa modelagem_df3 para o padrão específico da WMS
           df2_final = modelagem_df3(df2)
        else:
            df2 = pd.read_excel(uploaded_file2, skiprows=11)  # Usa modelagem_df3 para o padrão específico da WMS
            df2_final = modelagem_df2(df2)
        
        # # Tratamento das planilhas
        # df1_final = modelagem_df1(df1)
        # df2_final = modelagem_df2(df2)
    
        # Mesclar os DataFrames usando outer join
        df_conciliacao = pd.merge(df1_final, df2_final, left_on='Material', right_on='Produto', how='outer')

        # Renomear a coluna de descrição da planilha 2 para um nome único
        df_conciliacao.rename(columns={'Descrição': 'Descrição_wms'}, inplace=True)

        # Concatenar a coluna de descrição
        df_conciliacao['Descrição'] = df_conciliacao['Descrição_wms'].fillna(df_conciliacao['Texto breve de material'])

        # Remover a coluna de descrição da planilha 2, já que foi concatenada com sucesso
        df_conciliacao.drop(columns=['Descrição_wms'], inplace=True)

        # Reordenar as colunas
        colunas_ordenadas = ['Produto', 'Descrição', 'Empenhada', 'Saldo wms']
        df2_final = df2_final[colunas_ordenadas]
        
        # Etapa 1: Coluna de Diferenças
        df_conciliacao['Diferenças'] = df_conciliacao['Saldo wms'] - df_conciliacao['Utilização livre']

        # Etapa 2: Criar coluna de Valor Unit
        df_conciliacao['Valor Unit'] = df_conciliacao['Val.utiliz.livre'] / df_conciliacao['Utilização livre']

        # Etapa 3: Criar coluna de Saldo 
        df_conciliacao['Saldo SAP'] = df_conciliacao['Utilização livre'] + df_conciliacao['Bloqueado']

        # Etapa 4: Criar coluna de Diferenças R$
        df_conciliacao['Diferenças R$'] = df_conciliacao['Diferenças'] * df_conciliacao['Valor Unit']

        # Etapa 5: Criar coluna de Local DIF
        df_conciliacao.loc[df_conciliacao['Diferenças'] == 0, 'Local DIF'] = 'Sem Divergência'
        df_conciliacao.loc[df_conciliacao['Diferenças'] > 0, 'Local DIF'] = 'SAP'
        df_conciliacao.loc[df_conciliacao['Diferenças'] < 0, 'Local DIF'] = 'WMS'
        df_conciliacao.loc[df_conciliacao['Material'].isnull(), 'Local DIF'] = 'NC / SAP'
        df_conciliacao.loc[df_conciliacao['Saldo wms'].isnull(), 'Local DIF'] = 'NC / WMS'

        # Etapa 6: Criar coluna de Sobra / Falta
        df_conciliacao.loc[df_conciliacao['Local DIF'] == 'Sem Divergência', 'Sobra / Falta'] = 'Sem Divergência'
        df_conciliacao.loc[df_conciliacao['Local DIF'] == 'SAP', 'Sobra / Falta'] = 'Sobra WMS'
        df_conciliacao.loc[df_conciliacao['Local DIF'] == 'WMS', 'Sobra / Falta'] = 'Falta WMS'
        df_conciliacao.loc[df_conciliacao['Material'].isnull(), 'Sobra / Falta'] = 'NC / SAP'
        df_conciliacao.loc[df_conciliacao['Saldo wms'].isnull(), 'Sobra / Falta'] = 'NC / WMS'

        df_conciliacao = df_conciliacao[['Material', 'Texto breve de material', 'UM básica', 'Utilização livre', 'Val.utiliz.livre', 'Bloqueado', 'Val.estoque bloq.', 'Produto',
                                'Descrição', 'Empenhada', 'Saldo wms', 'Diferenças', 'Valor Unit', 'Saldo SAP', 'Diferenças R$', 'Local DIF', 'Sobra / Falta']]
        # Criar a planilha de Diferenças SAP
        df_diferencas_sap = df_conciliacao[df_conciliacao['Sobra / Falta'] == 'Sobra WMS']

        remover_colunas = ['Produto', 'Descrição', 'Saldo wms', 'Diferenças', 'Diferenças R$', 'Empenhada']
        df_diferencas_sap = df_diferencas_sap.drop(columns=remover_colunas)

        df_diferencas_sap['Material'] = df_diferencas_sap['Material'].astype(str).str.replace('.0', '')
        df1['Material'] = df1['Material'].astype(str).str.replace('.0', '')

        #Mesclar df_diferencas_sap com df_final1
        df_final_diferencas = pd.merge(df_diferencas_sap, df1, how='inner', on=['Material'])

        remover_colunas_2 = ['Utilização livre_x', 'Val.utiliz.livre_x', 'Bloqueado_x', 'Val.estoque bloq._x', 'Saldo SAP',	'Local DIF',
                            'Sobra / Falta', 'Denominação depósito', 'Nº estoque especial', 'Valor Unit', 'UM básica_y', 'Texto breve de material_y']
        
        df_final_diferencas = df_final_diferencas.drop(columns=remover_colunas_2)

        df_final_diferencas = df_final_diferencas.rename(columns={'Texto breve de material_x': 'Texto breve de material', 'UM básica_x': 'UMB', 'Utilização livre_y': 'Utilização livre',
                                                          'Val.utiliz.livre_y': 'Val. Utiliz.Livre', 'Bloqueado_y': 'Bloqueado', 'Val.estoque bloq._y': 'Val.estoque bloq'})
        
        df_final_diferencas = df_final_diferencas[['Material', 'Texto breve de material', 'UMB','Centro', 'Depósito', 'Lote', 'Utilização livre', 'Val. Utiliz.Livre', 'Bloqueado', 'Val.estoque bloq']]


        # Criação da planilha Diferenças WMS:
        df_diferencas_WMS = df_conciliacao[df_conciliacao['Sobra / Falta'] == 'Falta WMS']

        # Removendo algumas colunas:
        remover_colunas_wms_1 = ['Material', 'Texto breve de material', 'Utilização livre', 'Val.utiliz.livre', 'Bloqueado', 'Val.estoque bloq.', 'Saldo SAP']

        df_diferencas_WMS = df_diferencas_WMS.drop(columns=remover_colunas_wms_1)

        # Mudando o timpo de dados da coluna
        df2['Produto'] = df2['Produto'].astype(str).str.replace('.0', '')

        df_final_dif_wms = pd.merge(df_diferencas_WMS, df2, how = 'inner', on = ['Produto'])

        # Selecionando as colunas para o DataFrame final 
        df_final_dif_wms = df_final_dif_wms[['Produto', 'Descrição_x', 'Endereço', 'Empenhada_y', 'Bloqueado', 'Qualidade', 'Saldo']]

        df_final_dif_wms['Saldo WMS'] = df_final_dif_wms['Empenhada_y'] + df_final_dif_wms['Bloqueado'] + df_final_dif_wms['Qualidade'] + df_final_dif_wms['Saldo']

        df_final_dif_wms = df_final_dif_wms.rename(columns={'Descrição_x': 'Descrição', 'Empenhada_y': 'Empenhada'})

        # Função para converter DataFrame em Excel e retornar BytesIO
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            processed_data = output.getvalue()
            return processed_data

        # Converter DataFrames em Excel
        conciliacao_xlsx = to_excel(df_conciliacao)
        diferencas_sap_xlsx = to_excel(df_final_diferencas)
        diferencas_wms_xlsx = to_excel(df_final_dif_wms)

        # Botões para download
        st.download_button(label="Download Conciliação Geral", data=conciliacao_xlsx, file_name="Conciliação Geral.xlsx")
        st.download_button(label="Download Diferenças (-) WMS vs SAP", data=diferencas_sap_xlsx, file_name="Diferenças (-) WMS vs SAP.xlsx")
        st.download_button(label="Download Diferenças(+) WMS vs SAP", data=diferencas_wms_xlsx, file_name="Diferenças (+) WMS vs SAP.xlsx")

        # Exibir as tabelas no Streamlit
        st.header("Conciliação Geral")
        st.write(df_conciliacao)

        st.header("Diferenças (-) WMS vs SAP")
        st.write(df_final_diferencas)

        st.header("Diferenças (+) WMS vs SAP")
        st.write(df_final_dif_wms)
    
    except Exception as e:
        st.error(f"Erro ao processar os arquivos: {e}")

# Adicionando o rodapé
st.markdown("""
    <style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #009000;
        color: white;
        text-align: left;
        padding: 10px;
        font-size: 14px;
    }
    </style>
    <div class="footer">
        <p><b> Amara Net Zero Brasil</b> - Versão 00.01</p>
    </div>
    """, unsafe_allow_html=True)


# Carregar a imagem e adicioná-la ao rodapé da página
imagem_base64_rodape = carregar_imagem(caminho_imagem_rodape)
st.markdown(
    f"""
    <style>
    .footer-img {{
        position: fixed;
        bottom: 0;
        width: 90%;
        text-align: center;
    }}
    .footer-img img {{
        width: 9%;
        max-width: 300px;
    }}
    </style>
    <div class="footer-img">
        <img src="data:image/png;base64,{imagem_base64_rodape}" alt="logo">
    </div>
    """,
    unsafe_allow_html=True
)


