#Importando bibliotecas
import pandas as pd
import numpy as np
import pyodbc
from sqlalchemy import create_engine

#Importando Planilha de Pedidos
pedido = pd.read_excel("PEDIDO.xlsx")

#Importando Planilha Marketplace
caminho_controle ='C:/Users/a.moreira/OneDrive - RAZER BRASIL/_Controle/Controle Geral (Forecast Diretoria) 2.0.xlsx'
controle_mktplace = pd.read_excel(caminho_controle, sheet_name='Geral Mktplace')

#Importando Planilha Geral Dist
controle_geral = pd.read_excel(caminho_controle, sheet_name='Geral Dist')

# Buscando e Validando o Preço
for i in range(len(pedido)):
    if pedido.loc[i,'CD'] == "SC":
        verifica_preco = pd.merge(pedido, controle_geral[['PN Evergame','Retail SC']], on='PN Evergame')
        pedido['VerificaPreco'] = verifica_preco['Retail SC']
    else:
        verifica_preco = pd.merge(pedido, controle_geral[['PN Evergame','Retail SP']], on='PN Evergame')
        pedido['VerificaPreco'] = verifica_preco['Retail SP']
        
# Fazendo a verificação de igualdade do preço
ResultadoPreco = np.where(pedido['PreçoUni'] == pedido['VerificaPreco'], 'Ok', 'VERIFICAR')
pedido['Resultado Preco'] = ResultadoPreco

# Buscando e Verificando disponibilidade do Estoque
for i in range(len(pedido)):
    if pedido.loc[i,'CD'] == "SC":
        verifica_estoque = pd.merge(pedido, controle_geral[['PN Evergame','Ever SC']], on='PN Evergame')
        pedido['VerificaEstoque'] = verifica_estoque['Ever SC']
    else:
        verifica_estoque = pd.merge(pedido, controle_geral[['PN Evergame','Empire SP']], on='PN Evergame')
        pedido['VerificaEstoque'] = verifica_estoque['Empire SP']

        
# Fazendo a verificação de disponibilidade de estoque
ResultadoEstoque = np.where(pedido['Qtde'] <= pedido['VerificaEstoque'], 'Ok', 'Indisponível')
pedido['Resultado Estoque'] = ResultadoEstoque

#Buscando o Custos
for i in range(len(pedido)):
    if pedido.loc[i,'CD'] == "SC":
        custo = pd.merge(pedido, controle_geral[['PN Evergame','Custo Ever SC']], on='PN Evergame')
        pedido['Custo'] = custo['Custo Ever SC']
    else:
        verifica_estoque = pd.merge(pedido, controle_geral[['PN Evergame','Custo Empire SP']], on='PN Evergame')
        pedido['Custo'] = verifica_estoque['Custo Empire SP']
        
#Calculando Margem
pedido['Margem'] = (pedido['Preco']-pedido['Custo'])/pedido['Preco']

#Calculando Total Custo Por Produto
pedido['CT Produto'] = pedido['Custo']*pedido['Qtde']




#Exportando base para SQL

server_name = '######'
database_name = '######'
username = '######'
password = '######'

connection_string = f"mssql+pyodbc://######:######@######?driver=ODBC+Driver+17+for+SQL+Server"

engine = create_engine(connection_string)

pedido.to_sql(name='PEDIDOS', con=engine, if_exists='append', index=False)
