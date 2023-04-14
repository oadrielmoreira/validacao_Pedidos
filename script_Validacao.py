#Importando bibliotecas
import pandas as pd
import numpy as np
import pyodbc
from sqlalchemy import create_engine

#Importando Planilha de Pedidos
pedido = pd.read_excel("PEDIDO.xlsx")

#Importando Planilha Unificados SC
caminho_controle ='C:/Users/a.moreira/OneDrive - RAZER BRASIL/_Controle/Controle Geral (Forecast Diretoria) 2.0.xlsx'
controle_unificados = pd.read_excel(caminho_controle, sheet_name='Unificados SC')

#Importando Planilha Geral Dist
caminho_controle ='C:/Users/a.moreira/OneDrive - RAZER BRASIL/_Controle/Controle Geral (Forecast Diretoria) 2.0.xlsx'
controle_geral = pd.read_excel(caminho_controle, sheet_name='Geral Dist')


# Buscando o Preço da Planilha Unificados SC do Controle
verifica_preco = pd.merge(pedido, controle_unificados[['PN Evergame','Retail Novo']], on='PN Evergame')
pedido['VerificaPreco'] = verifica_preco['Retail Novo']

# Fazendo a verificação de igualdade do preço
ResultadoPreco = np.where(pedido['PreçoUni'] == pedido['VerificaPreco'], 'Ok', 'VERIFICAR')
pedido['Resultado Preco'] = ResultadoPreco


# Buscando o Estoque da Planilha Unificados SC do Controle
verifica_estoque = pd.merge(pedido, controle_unificados[['PN Evergame','Estoque SC']], on='PN Evergame')
pedido['VerificaEstoque'] = verifica_estoque['Estoque SC']

# Fazendo a verificação de disponibilidade de estoque
ResultadoEstoque = np.where(pedido['Qtde'] <= pedido['VerificaEstoque'], 'Ok', 'Indisponível')
pedido['Resultado Estoque'] = ResultadoEstoque


# Buscando o Custo da Planilha Unificados SC do Controle
custo = pd.merge(pedido, controle_unificados[['PN Evergame','CM']], on='PN Evergame')
pedido['Custo'] = custo['CM']


#Calculando Margem
pedido['Margem'] = (pedido['Preco']-pedido['Custo'])/pedido['Preco']

#Calculando Total Custo Por Produto
pedido['CT Produto'] = pedido['Custo']*pedido['Qtde']


#Verificando disponibilidade do TB
pedido['PN TB'] = pedido['PN Evergame']+'.TB'

pn_tb = pedido['PN TB']
evergame_values = []
for value in pn_tb:
    evergame_series = controle_geral.loc[controle_geral['PN Evergame'] == value, 'Ever SC']
    if not evergame_series.empty:
        evergame_value = evergame_series.iloc[0]
    else:
        evergame_value = None
    evergame_values.append(evergame_value)
pedido['Estoque TB'] = evergame_values

#Verificando disponibilidade do RT
pedido['PN RT'] = pedido['PN Evergame']+'.RT'

pn_rt = pedido['PN RT']
evergame_values = []
for value in pn_rt:
    evergame_series = controle_geral.loc[controle_geral['PN Evergame'] == value, 'Ever SC']
    if not evergame_series.empty:
        evergame_value = evergame_series.iloc[0]
    else:
        evergame_value = None
    evergame_values.append(evergame_value)
pedido['Estoque RT'] = evergame_values

#Verificando disponibilidade do RS
pedido['PN RS'] = pedido['PN Evergame']+'.RS'

pn_rs = pedido['PN RS']
evergame_values = []
for value in pn_rs:
    evergame_series = controle_geral.loc[controle_geral['PN Evergame'] == value, 'Ever SC']
    if not evergame_series.empty:
        evergame_value = evergame_series.iloc[0]
    else:
        evergame_value = None
    evergame_values.append(evergame_value)
pedido['Estoque RS'] = evergame_values



#Exportando base de disponibilidade
disponibilidade = pedido.loc[:, ['PN TB','Estoque TB','PN RT','Estoque RT','PN RS','Estoque RS','Qtde','Preco']]
disponibilidade.to_excel('Disponibilidade.xlsx', index=False)



#Exportando base para SQL

server_name = '######'
database_name = '######'
username = '######'
password = '######'

connection_string = f"mssql+pyodbc://######:######@######?driver=ODBC+Driver+17+for+SQL+Server"

engine = create_engine(connection_string)

pedido.to_sql(name='PEDIDOS', con=engine, if_exists='append', index=False)



#Calculando Receita, Custo e Retorno Sobre a Venda (RSV)
receita = pedido['Subtotal'].sum()
custo = pedido['CT Produto'].sum()
rsv = ((receita - custo)/receita)
rsv = round(rsv * 100, 2)
rsv = "{}%".format(rsv)


#Criando dataframe de verificação de Disponibilidade e Compatibilidade dos preços ofertados pelo vendedor com os preços anúnciados pela empresa
verifica = pd.DataFrame({
    'PN' : pedido['PN Evergame'],
    'Margem' : pedido['Margem'],
    'Qtde' : pedido['Qtde'],
    'Estoque' : pedido['VerificaEstoque'],
    'VERIFICA E': pedido['Resultado Estoque'],
    
    'Preço' : pedido['PreçoUni'],
    'Price' : pedido['VerificaPreco'],
    'VERIFICA P' : pedido['Resultado Preco'],
})

#Printando todas as verificações necessárias para a validação de um pedido
verifica
disponibilidade
print('Retorno Sobre a Venda: ' + rsv)
