import pandas as pd
import numpy as np


df = pd.read_excel ('.\\original\\nomearquivo.xlsx',encoding = "ISO-8859-1")

df['Produto'] = df['Orçamento']

df['Orçamento'] = df['Orçamento'].astype(str)
df2 = df

#orçamento
ultimo_pedido = 0
for index, value in df2['Orçamento'].items():
    if value.isnumeric():
        ultimo_pedido = value
    else:
        df2['Orçamento'][index] = ultimo_pedido
df2['Orçamento'] = df2['Orçamento'].astype(int)

#clientes
df2['Cliente'] = df2['Cliente'].astype(str)
ultimo_cliente = ''
for index, value in df2['Cliente'].items():
    if value != 'nan':
        ultimo_cliente = value
    else:
        df2['Cliente'][index] = ultimo_cliente

#vendedor
df2['Vendedor'] = df2['Vendedor'].astype(str)
ultimo_vendedor = ''
for index, value in df2['Vendedor'].items():
    if value != 'nan':
        ultimo_vendedor = value
    else:
        df2['Vendedor'][index] = ultimo_vendedor

#data do orçamento
df2['Data do orçamento'] = df2['Data do orçamento'].astype(str)
ultima_data = ''
for index, value in df2['Data do orçamento'].items():
    if value != 'NaT':
        ultima_data = value
    else:
        df2['Data do orçamento'][index] = ultima_data

#situação
df2['Situação'] = df2['Situação'].astype(str)
ultima_situacao = ''
for index, value in df2['Situação'].items():
    if value != 'nan':
        ultima_situacao = value
    else:
        df2['Situação'][index] = ultima_situacao

#valor total
df2['Valor Total'] = df2['Valor Total'].astype(str)
valor_total = ''
for index, value in df2['Valor Total'].items():
    if value != 'nan':
        valor_tota = value
    else:
        df2['Valor Total'][index] = valor_total

#peso
df2['Peso Bruto'] = df2['Peso Bruto'].astype(str)
peso = ''
for index, value in df2['Peso Bruto'].items():
    if value != 'nan':
        peso = value
    else:
        df2['Peso Bruto'][index] = peso
df2['Peso Bruto'] = df2['Peso Bruto'].astype(float)

#endereço
df2['Endereço'] = df2['Endereço'].astype(str)
endereco = ''
for index, value in df2['Endereço'].items():
    if value != 'nan':
        endereco = value
    else:
        df2['Endereço'][index] = endereco

#numero
df2['Numero'] = df2['Numero'].astype(str)
numero = ''
for index, value in df2['Numero'].items():
    if value != 'nan':
        numero = value
    else:
        df2['Numero'][index] = numero

#cidade
df2['Cidade'] = df2['Cidade'].astype(str)
cidade = ''
for index, value in df2['Cidade'].items():
    if value != 'nan':
        cidade = value
    else:
        df2['Cidade'][index] = cidade

#estado
df2['Estado'] = df2['Estado'].astype(str)
estado = ''
for index, value in df2['Estado'].items():
    if value != 'nan':
        estado = value
    else:
        df2['Estado'][index] = estado

#bairro
df2['Bairro'] = df2['Bairro'].astype(str)
bairro = ''
for index, value in df2['Bairro'].items():
    if value != 'nan':
        bairro = value
    else:
        df2['Bairro'][index] = bairro

#cep
df2['CEP'] = df2['CEP'].astype(str)
cep = ''
for index, value in df2['CEP'].items():
    if value != 'nan':
        cep = value
    else:
        df2['CEP'][index] = cep

#parcelas
df2['Total de parcelas'] = df2['Total de parcelas'].astype(str)
parcelas = ''
for index, value in df2['Total de parcelas'].items():
    if value != 'nan':
        parcelas = value
    else:
        df2['Total de parcelas'][index] = parcelas
df2['Total de parcelas'] = df2['Total de parcelas'].astype(float)

#transportadora
df2['Transportadora'] = df2['Transportadora'].astype(str)
transportadora = ''
for index, value in df2['Transportadora'].items():
    if value != 'nan':
        transportadora = value
    else:
        df2['Transportadora'][index] = transportadora

#qtde. de produtos
df2['Qtde. Produtos'] = df2['Qtde. Produtos'].astype(str)
qtd_produto = ''
for index, value in df2['Qtde. Produtos'].items():
    if value != 'nan':
        qtd_produto = value
    else:
        df2['Qtde. Produtos'][index] = qtd_produto

#total de unidades
df2['Total Unidades'] = df2['Total Unidades'].astype(str)
total_unidades = ''
for index, value in df2['Total Unidades'].items():
    if value != 'nan':
        total_unidades = value
    else:
        df2['Total Unidades'][index] = total_unidades
df2['Total Unidades'] = df2['Total Unidades'].astype(float)

#forma de pagamento
df2['Forma de Pagamento'] = df2['Forma de Pagamento'].astype(str)
forma_pagamento = ''
for index, value in df2['Forma de Pagamento'].items():
    if value != 'nan':
        forma_pagamento = value
    else:
        df2['Forma de Pagamento'][index] = forma_pagamento

#cortar linhas
df2 = df2[df2['Produto'] != 'Produto']
df2 = df2[df2['Produto'] != 'Quantidade de orçamentos']
df2 = df2[df2['Produto'] != 'Valor Total']

#renomear colunas
df2.columns = ['Orçamento', 'Unidade', 'Qtde', 'Valor Unitário', 'IPI', 'Valor Total', ' Cliente', 'Endereço', 'Numero' , 'Bairro', 'Cidade',
              'Estado', 'CEP', 'Vendedor', 'Valor Total', 'Total de parcelas', 'Peso Bruto' ,'Transportadora', 'Data do orçamento', 'Qtde. Produtos', 'Total Unidades',
              'Situação', 'Forma de Pagamento', 'Produto']

#apagar linhas
df2 = df2[df2['Unidade'].notna()]

df2['Data do orçamento'] = pd.to_datetime(df['Data do orçamento'])

#salvar arquivo
df2.to_excel('.\\manipulado\\maiscimentao_manipulado.xls', encoding='utf-8',index=False)