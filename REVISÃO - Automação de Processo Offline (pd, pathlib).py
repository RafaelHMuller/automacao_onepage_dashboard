#!/usr/bin/env python
# coding: utf-8

# # Projeto: Automação de Indicadores
# 
# ##### Objetivo: 
# Criar uma automação de um processo feito no computador (offline).
# 
# ##### Descrição:
# 
# Imagine que você trabalha em uma grande rede de lojas de roupa com 25 lojas espalhadas por todo o Brasil.
# 
# Todo dia, pela manhã, a equipe de análise de dados calcula os chamados One Pages e envia para o gerente de cada loja o OnePage da sua loja, bem como todas as informações usadas no cálculo dos indicadores.
# 
# Um One Page é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores de cada loja e permitir em 1 página (daí o nome OnePage) tanto a comparação entre diferentes lojas, quanto quais indicadores aquela loja conseguiu cumprir naquele dia ou não.
# 
# Exemplo de OnePage:

# ![title](onepage.png)

# O seu papel, como Analista de Dados, é conseguir criar um processo da forma mais automática possível para:
# - Salvar as planilhas de cada loja dentro de uma pasta da loja com a data da planilha, a fim de criar um histórico de backup;
# - Calcular o OnePage de cada loja e enviar um email para o gerente de cada loja com o seu OnePage no corpo do e-mail e também o arquivo completo com os dados da sua respectiva loja em anexo (o e-mail a ser enviado para o Gerente de cada loja deve seguir como exemplo a imagem Exemplo.jpg.); 
# - Enviar ainda um e-mail para a diretoria (informações no arquivo Emails.xlsx) com rankings das melhores lojas em termos de faturamento, um ranking do dia e outro ranking anual. Além disso, no corpo do e-mail, deve ressaltar qual foi a melhor e a pior loja do dia e do ano;
# - Disponibilizar para a diretoria uma página na internet com gráficos interativos (dash) que traduzem as informações enviadas pelo email;

# ##### Indicadores do OnePage
# 
# - Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
# - Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4
# - Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500
# 
# Obs: Cada indicador deve ser calculado no dia e no ano. O indicador do dia deve ser o do último dia disponível na planilha de Vendas (a data mais recente)
# 
# Obs2: Dica para o caracter do sinal verde e vermelho: pegue o caracter desse site (https://fsymbols.com/keyboard/windows/alt-codes/list/) e formate com html

# In[10]:


import pandas as pd
from pathlib import Path
import win32com.client as win32
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from dash import Dash, html, dcc


# In[14]:


# importar a base de dados
local_atual = Path.cwd()
caminho_dados = fr'{local_atual}\Bases de Dados'

vendas = pd.read_excel(fr'{caminho_dados}\Vendas.xlsx')
lojas = pd.read_csv(fr'{caminho_dados}\Lojas.csv', sep=';', encoding='latin1')
emails = pd.read_excel(fr'{caminho_dados}\Emails.xlsx')

display(vendas, lojas, emails)


#    - Salvar as planilhas de cada loja dentro de uma pasta específica da loja com a data da planilha, a fim de criar um histórico de backup;

# In[3]:


vendas = vendas.merge(lojas, on='ID Loja')

# criar uma pasta de backup em comum onde todas as pastas das lojas serão inseridas
try:
    Path(local_atual / 'Backup Arquivos Lojas - Revisão').mkdir()
except:
    pass

# criar uma pasta para cada loja
caminho_backup = Path(local_atual / 'Backup Arquivos Lojas - Revisão')
for loja in vendas['Loja'].unique():
    try:
        Path(caminho_backup / f'{loja}').mkdir()
    except:
        pass
    
# criar uma planilha excel para cada loja; este processo é feito todos os dias, ou seja, na data mais atual
hoje = vendas['Data'].max().strftime('%d-%m-%Y')
ano = hoje.split('-')[2]
print(hoje, ano)

dicionario_dfs = {}

for loja in vendas['Loja'].unique():
    dicionario_dfs[loja] = vendas.loc[vendas['Loja']==loja, :]
    caminho_dfs_lojas = Path(rf'C:\Users\W10\Desktop\Python\Arquivos_estudo\Projetos\Projetos_hashtag\Projeto 1 - Automações de Processo\Backup Arquivos Lojas - Revisão\{loja}')
    nome_df_loja = f'{loja}_{hoje}.xlsx'
    dicionario_dfs[loja].to_excel(caminho_dfs_lojas / nome_df_loja) 


# - Calcular o OnePage de cada loja e enviar um email para o gerente respectivo com o seu OnePage no corpo do e-mail e também o arquivo excel completo com os dados da sua loja em anexo (o e-mail a ser enviado para o Gerente de cada loja deve seguir como exemplo a imagem Exemplo.jpg.); 

# In[4]:


# calcular os indicadores de cada loja
hoje = vendas['Data'].max()

for loja in dicionario_dfs:
    faturamento_ano = dicionario_dfs[loja]['Valor Final'].sum()
    faturamento_dia = dicionario_dfs[loja].loc[dicionario_dfs[loja]['Data']==hoje, 'Valor Final'].sum()
    diversidade_ano = len(dicionario_dfs[loja]['Produto'].unique())
    diversidade_dia = len(dicionario_dfs[loja].loc[dicionario_dfs[loja]['Data']==hoje, 'Produto'].unique())
    ticketmedio_ano = dicionario_dfs[loja]['Valor Final'].mean()
    ticketmedio_dia = dicionario_dfs[loja].loc[dicionario_dfs[loja]['Data']==hoje, 'Valor Final'].mean()

# definir os indicadores do OnePage
meta_faturamento_ano = 1650000
meta_faturamento_dia = 1000
meta_diversidade_ano = 120
meta_diversidade_dia = 4
meta_ticketmedio_ano = 500
meta_ticketmedio_dia = 500


# In[5]:


# enviar email com OnePage para cada gerente de loja
# para a criação da tabela, pesquisa-se no google: table html
def cores_OnePage(indicador_loja, meta):
    if indicador_loja >= meta:
        return 'green'
    else:
        return 'red'


for loja in dicionario_dfs:
    email_gerente = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    nome_gerente = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    hoje_format = hoje.strftime('%d-%m-%Y')
    
    cor_fat_dia = cores_OnePage(faturamento_dia, meta_faturamento_dia)
    cor_div_dia = cores_OnePage(diversidade_dia, meta_diversidade_dia)
    cor_tic_dia = cores_OnePage(ticketmedio_dia, meta_ticketmedio_dia)
    
    if loja == 'Iguatemi Esplanada': # TESTE !!!!!!
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = f'{email_gerente}'
        mail.Subject = f'OnePage {loja} - {hoje_format}'
        mail.HTMLBody = f'''<p>Bom dia, {nome_gerente}.
        <p></p>
        <h2>OnePage do dia {hoje_format}</h2>
        <table style="width:50%">
          <tr>
            <th style="text-align: center">Indicador</th>
            <th style="text-align: center">Valor Dia</th>
            <th style="text-align: center">Meta Dia</th>
            <th style="text-align: center">Cenário Dia</th>
          </tr>
          <tr>
            <td style="text-align: center">Faturamento</td>
            <td style="text-align: center">{faturamento_dia}</td>
            <td style="text-align: center">{meta_faturamento_dia}</td>
            <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
          </tr>
          <tr>
            <td style="text-align: center">Diversidade de Produtos</td>
            <td style="text-align: center">{diversidade_dia}</td>
            <td style="text-align: center">{meta_diversidade_dia}</td>
            <td style="text-align: center"><font color="{cor_div_dia}">◙</font></td>
          </tr>
           <tr>
            <td style="text-align: center">Ticket Médio por Produto</td>
            <td style="text-align: center">{ticketmedio_dia}</td>
            <td style="text-align: center">{meta_ticketmedio_dia}</td>
            <td style="text-align: center"><font color="{cor_tic_dia}">◙</font></td>
          </tr>
        </table>
        </body>
        </html>
        <p></p>
        <p>Segue em anexo a planilha com os dados do dia para mais detalhes.</p>
        <p></p>
        <p>Att,</p>
        <p>Rafael Muller.</p>
        '''
        nome_df_loja = f'{loja}_{hoje_format}.xlsx'
        caminho_anexo = Path(rf'C:\Users\W10\Desktop\Python\Arquivos_estudo\Projetos\Projetos_hashtag\Projeto 1 - Automações de Processo\Backup Arquivos Lojas - Revisão\{loja}\{nome_df_loja}')
        mail.Attachments.Add(str(caminho_anexo)) # str() define todo o caminho como texto, não como objeto do pathlib
        mail.Send()
        
        print(f'Email da loja {loja} enviado!')


# - Enviar ainda um e-mail para a diretoria (informações no arquivo Emails.xlsx) com rankings das melhores lojas em termos de faturamento, sendo um ranking do dia e outro do ano. Além disso, no corpo do e-mail, deve-se ressaltar qual foi a melhor e a pior loja do dia e do ano;
# - Disponibilizar para a diretoria uma página na internet com gráficos interativos (dash) que traduzem as informações enviadas pelo email;

# In[6]:


# ranking faturamento dia
df_hoje = vendas.loc[vendas['Data']==hoje, :]
df_lojas_hoje = df_hoje[['Loja', 'Valor Final']].groupby('Loja').sum()
df_lojas_hoje = df_lojas_hoje.sort_values(by='Valor Final', ascending=False)
df_lojas_hoje['Média'] = df_lojas_hoje['Valor Final'].mean()
display(df_lojas_hoje)

# ranking faturamento ano
df_lojas_ano = vendas[['Loja', 'Valor Final']].groupby('Loja').sum()
df_lojas_ano = df_lojas_ano.sort_values(by='Valor Final', ascending=False)
df_lojas_ano['Média'] = df_lojas_ano['Valor Final'].mean()
display(df_lojas_ano)


# In[7]:


# criação dos gráficos para o dashboard - PLOTLY
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# gráfico do faturamento do DIA

trace1 = go.Bar(
    x=df_lojas_hoje.index,
    y=df_lojas_hoje['Valor Final'],
    name='Total Faturamento Dia',
    marker=dict(color='rgb(34,139,34)' )
) # Esse primeiro traçado gera o gráfico de barras. O argumento x você coloca uma lista com o nome de cada coluna; no y você coloca o valor de cada coluna; o name é a legenda da sua coluna e o marker é a cor das colunas.

trace2 = go.Scatter(
    x=df_lojas_hoje.index,
    y=df_lojas_hoje['Média'],
    name='Média Faturamento Dia',
    yaxis='y2',
    marker=dict(color='rgb(255,0,255)' )
) # Esse segundo traçado vai desenhar a linha do gráfico. Nos argumentos x e y você vai colocar as repectivas listas de valores; o name é a legenda do gráfico, e esse yaxis está dizendo que você quer utilizar o eixo y do lado direito  do gráfico

fig1 = make_subplots(specs=[[{"secondary_y": False}]])  
fig1.add_trace(trace1)
fig1.add_trace(trace2,secondary_y=False)
fig1.update_layout(yaxis_range=[1550000,1800000])
fig1['layout'].update(height = 600, width = 1000, xaxis=dict(
      tickangle=-90
    )) # Essas configurações finais servem para juntar os dois gráficos e fazer os últimos ajustes
fig1.show()

# gráfico do faturamento do ANO

trace1 = go.Bar(
    x=df_lojas_ano.index,
    y=df_lojas_ano['Valor Final'],
    name='Total Faturamento Ano',
    marker=dict(color='rgb(34,139,34)' )
) # Esse primeiro traçado gera o gráfico de barras. O argumento x você coloca uma lista com o nome de cada coluna; no y você coloca o valor de cada coluna; o name é a legenda da sua coluna e o marker é a cor das colunas.

trace2 = go.Scatter(
    x=df_lojas_ano.index,
    y=df_lojas_ano['Média'],
    name='Média Faturamento Ano',
    yaxis='y2',
    marker=dict(color='rgb(255,0,255)' )
) # Esse segundo traçado vai desenhar a linha do gráfico. Nos argumentos x e y você vai colocar as repectivas listas de valores; o name é a legenda do gráfico, e esse yaxis está dizendo que você quer utilizar o eixo y do lado direito  do gráfico

fig2 = make_subplots(specs=[[{"secondary_y": False}]])  
fig2.add_trace(trace1)
fig2.add_trace(trace2,secondary_y=False)
fig2.update_layout(yaxis_range=[1550000,1800000])
fig2['layout'].update(height = 600, width = 1000, xaxis=dict(
      tickangle=-90
    )) # Essas configurações finais servem para juntar os dois gráficos e fazer os últimos ajustes
fig2.show()


# In[8]:


# criação dos gráficos para o dashboard - MATPLOTLIB, SEABORN

# gráfico do faturamento do DIA
fig, ax = plt.subplots(figsize=(15,5))
plt.title(f'Total Faturamento Dia vs. Média Dia {hoje_format}')
grafico = sns.barplot(data=df_lojas_hoje, x=df_lojas_hoje.index, y='Valor Final', ax=ax)
grafico = sns.lineplot(data=df_lojas_hoje, x=df_lojas_hoje.index, y='Média', ax=ax)
grafico.tick_params(axis='x', rotation=90) 

# gráfico do faturamento do ANO
fig, ax = plt.subplots(figsize=(15,5))
plt.title(f'Total Faturamento Ano vs. Média Ano {ano}')
grafico = sns.barplot(data=df_lojas_ano, x=df_lojas_ano.index, y='Valor Final', ax=ax)
grafico = sns.lineplot(data=df_lojas_ano, x=df_lojas_ano.index, y='Média', ax=ax)
plt.ylim(1500000, 1800000)
grafico.tick_params(axis='x', rotation=90) 


# In[9]:


# link fixo dashboard
link_dash = 'http://127.0.0.1:8050/'

# email para a diretoria
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = f'{email_gerente}'
mail.Subject = f'Ranking Faturamento das lojas do Dia {hoje_format} e do Ano {ano}'
mail.HTMLBody = f'''<p>Bom dia, diretoria.</p>
<p></p>
<p>Ranking das lojas que mais venderam hoje {hoje_format}:</p>
<p>{df_lojas_hoje.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}</p>
<p></p>
<p>Ranking das lojas que mais venderam no ano {ano}:</p>
<p>{df_lojas_ano.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}</p>
<p></p>
<p>Loja que mais vendeu no dia:<strong> {df_lojas_hoje.index[0]}</strong></p>
<p>Loja que mais vendeu no ano:<strong> {df_lojas_ano.index[0]}</strong></p>
<p></p>
<p>Segue o link para visualização gráfica do faturamento das lojas: {link_dash}</p>
<p></p>
<p>Att,</p>
<p>Rafael Muller.</p>
'''
mail.Send()


# In[10]:


#criação do aplicativo dash
app = Dash(__name__)

# criação do layout
app.layout = html.Div(children=[
    
    html.H1(children='Dashboard do faturamento das lojas'),

    html.H3(children=f'Faturamento do dia vs. Média do dia {hoje_format}'),
    
    dcc.Graph(id='grafico_barras', figure=fig1),
    
    html.H3(children=f'Faturamento do ano vs. Média do ano {ano}'),
    
    dcc.Graph(id='grafico_barras', figure=fig2),
    
], style={'text-align':'center'})

# upload do dash na internet
if __name__ == '__main__':
    app.run_server(debug=False)


# In[ ]:




