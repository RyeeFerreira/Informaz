import pandas as pd

# Fazendo uma observação sobre o projeto, tem algumas tarefas que eu fiz 2 versões por não entender completamente o que era pedido.

arquivo_excel = 'Case_Infomaz_Base_de_Dados.xlsx'

cadastro_produtos = pd.read_excel(arquivo_excel, sheet_name='Cadastro Produtos', header=1)
cadastro_clientes = pd.read_excel(arquivo_excel, sheet_name='Cadastro Clientes', header=1)
cadastro_de_estoque = pd.read_excel(arquivo_excel, sheet_name='Cadastro de Estoque', header=1)
cadastro_de_fornecedores = pd.read_excel(arquivo_excel, sheet_name='Cadastro Fornecedores', header=1)
transacao_vendas = pd.read_excel(arquivo_excel, sheet_name='Transações Vendas', header=1)

transacao_vendas['MES'] = transacao_vendas['DATA NOTA'].dt.to_period('M')

# 1) Calcule o valor total de venda dos produtos por categoria. Utilize as tabelas CADASTRO_PRODUTOS e TRANSACOES_VENDAS.

vendas_produtos = transacao_vendas.merge(cadastro_produtos, on='ID PRODUTO')

vendas_produtos['VALOR_TOTAL'] = vendas_produtos['VALOR ITEM'] * vendas_produtos['QTD ITEM']

metrica_1 = (
    vendas_produtos
    .groupby('CATEGORIA')['VALOR_TOTAL']
    .sum()
    .reset_index()
    .sort_values(by='VALOR_TOTAL', ascending=False)
)
print("\n=== VALOR TOTAL DE VENDAS POR CATEGORIA ===")
print(metrica_1.to_string(index=False))

# 2) Calcule a margem dos produtos subtraindo o valor unitario pelo valor de venda. Utilize as tabelas CADASTRO_ESTOQUE e TRANSACOES_VENDAS.

cadastro_de_estoque['ID PRODUTO'] = cadastro_de_estoque['ID ESTOQUE'] - 4000

cadastro_de_estoque['VALOR UNITARIO'] = cadastro_de_estoque['VALOR ESTOQUE'] / cadastro_de_estoque['QTD ESTOQUE']

margem_produtos = transacao_vendas.merge(
    cadastro_de_estoque,
    on='ID PRODUTO',
    how='left'
)

margem_produtos['MARGEM UNITARIA'] = margem_produtos['VALOR ITEM'] - margem_produtos['VALOR UNITARIO']
margem_produtos['MARGEM TOTAL'] = margem_produtos['MARGEM UNITARIA'] * margem_produtos['QTD ITEM']

metrica_2 = margem_produtos[[
    'ID PRODUTO',
    'VALOR ITEM',
    'VALOR UNITARIO',
    'MARGEM UNITARIA',
    'MARGEM TOTAL'
]]

print("\n=== MARGEM DOS PRODUTOS ===")
print(metrica_2.head(10))
# caso queira exibir o ranking completo, descomente a linha abaixo
# print(metrica_2.to_string(index=False))

# 3) Calcule um ranking de clientes por quantidade de produtos comprados por mês. Utilize as tabelas CADASTRO_CLIENTES e TRANSACOES_VENDAS.

ranking_cliente = transacao_vendas.merge(cadastro_clientes, on='ID CLIENTE')

ranking_cliente['MES'] = ranking_cliente['DATA NOTA'].dt.to_period('M')

metrica_3 = (ranking_cliente
    .groupby(['MES', 'ID CLIENTE', 'NOME CLIENTE'])['QTD ITEM']
    .sum()
    .reset_index()
    .sort_values(['QTD ITEM'], ascending=[False])
)
print("\n=== RANKING DE CLIENTES POR QUANTIDADE DE PRODUTOS COMPRADOS POR MÊS ===")
print(metrica_3.head(10))   

top_cliente_por_mes = (
    metrica_3
    .loc[
        metrica_3
        .groupby('MES')['QTD ITEM']
        .idxmax()
    ]
    .sort_values('MES')
    .reset_index(drop=True)
)

print("\n=== TOP CLIENTE DE CADA MÊS ===")
print(top_cliente_por_mes.to_string(index=False))


# 4) Calcule um ranking de fornecedores por quantidade de estoque disponivel por mês. Utilize as tabelas CADASTRO_FORNECEDORES e CADASTRO_ESTOQUE.

cadastro_de_estoque['DATA ESTOQUE'] = pd.to_datetime(
    cadastro_de_estoque['DATA ESTOQUE'],
    dayfirst=True, errors='coerce'
)

ranking_fornecedor = (
    cadastro_de_estoque
    .merge(cadastro_de_fornecedores, on='ID FORNECEDOR', how='left')
)

ranking_fornecedor['MES'] = ranking_fornecedor['DATA ESTOQUE'].dt.to_period('M')

# Versão 1: Ranking dos Fornecedores pelo maior estoque agregado em um mês

agregado = (
    ranking_fornecedor
    .groupby(['MES', 'ID FORNECEDOR', 'NOME FORNECEDOR'], as_index=False)['QTD ESTOQUE']
    .sum()
)

ranking_geral = (
    agregado
    .groupby(['ID FORNECEDOR', 'NOME FORNECEDOR'], as_index=False)['QTD ESTOQUE']
    .max()
    .sort_values('QTD ESTOQUE', ascending=False)
    .reset_index(drop=True)
)

print("\n=== RANKING GERAL DOS FORNECEDORES PELO MAIOR ESTOQUE EM UM MÊS ===")
print(ranking_geral.head(10))
# caso queira exibir o ranking completo, descomente a linha abaixo
# print(ranking_geral.to_string(index=False))


# Versão 2: Ranking dos Fornecedores de cada mes

top1_por_mes = (
    agregado
    .loc[
        agregado
        .groupby('MES')['QTD ESTOQUE']
        .idxmax()
    ]
    .sort_values('MES')
    .reset_index(drop=True)
)

print("\n === TOP 1 de cada mês (fornecedor com maior estoque agregado) ===")
print(top1_por_mes.head(10))
# caso queira exibir o ranking completo, descomente a linha abaixo
# print(top1_por_mes.to_string(index=False))

# 5) Calcule um ranking de produtos por quantidade de venda por mês. Utilize a tabela TRANSACOES_VENDAS.

transacao_vendas['DATA NOTA'] = pd.to_datetime(
    transacao_vendas['DATA NOTA'],
    dayfirst=True, errors='coerce'
)

produto_QTD = (
    transacao_vendas
    .groupby(['MES', 'ID PRODUTO'], as_index=False)['QTD ITEM']
    .sum()
)

ranking_produto_mais_comprado = (
    produto_QTD
    .groupby(['ID PRODUTO'], as_index=False)['QTD ITEM']
    .max()
    .sort_values( 'QTD ITEM', ascending=[False])
    .reset_index(drop=True)
)

print("\n=== RANKING DE PRODUTOS POR QUANTIDADE DE VENDAS POR MÊS ===")
print(ranking_produto_mais_comprado.head(10))
# caso queira exibir o ranking completo, descomente a linha abaixo
# print(ranking_produto_mais_comprado.to_string(index=False))

# Versão 2: Top produto mais comprado de cada mês
top_produto_por_mes = (
    produto_QTD
    .loc[
        produto_QTD
        .groupby('MES')['QTD ITEM']
        .idxmax()
    ]
    .sort_values('MES')
    .reset_index(drop=True)
)

print("\n=== TOP PRODUTO MAIS COMPRADO DE CADA MÊS ===")
print(top_produto_por_mes.head(10))
# caso queira exibir o ranking completo, descomente a linha abaixo
# print(top_produto_por_mes.to_string(index=False))


# 6) Calcule um ranking de produtos por valor de venda por mês. Utilize a tabela TRANSACOES_VENDAS.

transacao_vendas['VALOR_TOTAL'] = transacao_vendas['VALOR ITEM'] * transacao_vendas['QTD ITEM']

produto_valor = (
    transacao_vendas
    .groupby(['MES', 'ID PRODUTO'], as_index=False)['VALOR_TOTAL']
    .sum()
)

ranking_produto_mais_vendido = (
    produto_valor
    .groupby(['MES', 'ID PRODUTO'], as_index=False)['VALOR_TOTAL']
    .max()
    .sort_values( 'VALOR_TOTAL', ascending=[False])
    .reset_index(drop=True)
)
print("\n=== RANKING DE PRODUTOS POR VALOR DE VENDAS POR MÊS ===")
print(ranking_produto_mais_vendido.head(10))
# caso queira exibir o ranking completo, descomente a linha abaixo
# print(ranking_produto_mais_vendido.to_string(index=False))

# Versão 2: Top produto com maior valor de venda de cada mês
top_produto_valor_por_mes = (
    produto_valor
    .loc[
        produto_valor
        .groupby('MES')['VALOR_TOTAL']
        .idxmax()
    ]
    .sort_values('MES')
    .reset_index(drop=True)
)

print("\n=== TOP PRODUTO COM MAIOR VALOR DE VENDA DE CADA MÊS ===")
print(top_produto_valor_por_mes.head(10))
# print(top_produto_valor_por_mes.to_string(index=False))

# 7) Calcule a média de valor de venda por categoria de produto por mês. Utiliza as tabelas CADASTRO_PRODUTOS e TRANSACOES_VENDAS.



valor_de_categoria_por_mes = (
    vendas_produtos
    .groupby(['CATEGORIA', 'MES'], as_index=False)['VALOR_TOTAL']
    .sum()
)

# print(valor_de_categoria_por_mes.head(10))

media_de_categoria_por_mes = (
    valor_de_categoria_por_mes
    .groupby(['CATEGORIA'], as_index=False)['VALOR_TOTAL']
    .mean()
    .sort_values('VALOR_TOTAL', ascending=False)
    )

media_de_categoria_por_mes['VALOR_TOTAL'] = media_de_categoria_por_mes['VALOR_TOTAL'].map('{:,.2f}'.format).str.replace(',', 'X').str.replace('.', ',').str.replace('X', '.')

print("\n=== MÉDIA DE VALOR DE VENDAS POR CATEGORIA POR MÊS ===")
print(media_de_categoria_por_mes.to_string(index=False))

# 8) Calcule um ranking de margem de lucro por categoria

margem_por_categoria = ( 
    margem_produtos.merge(cadastro_produtos, on='ID PRODUTO')
    .groupby(['CATEGORIA'], as_index=False)['MARGEM TOTAL']
    .sum()
    .sort_values('MARGEM TOTAL', ascending=False)
    )

print("\n=== RANKING DE MARGEM POR CATEGORIA ===")
print(margem_por_categoria.to_string(index=False))

# 9) Liste produtos comprados por clientes

clientes = transacao_vendas.merge(cadastro_clientes, on='ID CLIENTE')
produtos = clientes.merge(cadastro_produtos, on='ID PRODUTO')

produtos_comprados = (
    produtos
    .groupby(['ID PRODUTO', 'NOME PRODUTO' , 'NOME CLIENTE', 'ID CLIENTE'], as_index=False)['QTD ITEM']
    .sum()
    .sort_values(['ID CLIENTE', 'QTD ITEM'], ascending=[True, False])
)
print("\n=== PRODUTOS COMPRADOS POR CLIENTES ===")
print(produtos_comprados.head(10))
# caso queira exibir o ranking completo, descomente a linha abaixo
# print(produtos_comprados.to_string(index=False))


# 10) Ranking de produtos por quantidade de estoque

estoque_produtos = cadastro_de_estoque.merge(cadastro_produtos, on='ID ESTOQUE')

ranking_produto = (
    estoque_produtos
    .groupby(['ID ESTOQUE', 'NOME PRODUTO'], as_index=False)['QTD ESTOQUE']
    .sum()
    .sort_values('QTD ESTOQUE', ascending=False)
    )

print("\n=== RANKING DE PRODUTOS POR QUANTIDADE DE ESTOQUE ===")
print(ranking_produto.head(10))
# caso queira exibir o ranking completo, descomente a linha abaixo
# print(ranking_produto.to_string(index=False))
