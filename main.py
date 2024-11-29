
## Set

import pandas as pd

caminho_pend = r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\PEND.XLSX'
caminho_pend_geral = r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\pend geral.XLSX'
caminho_zstok = r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\zstok.XLSX'
caminho_fup = r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\fup.XLSX'
caminho_fob = r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\ZTMM069.XLSX'
caminho_pecas = r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\ZTMM085.XLSX'
caminho_atuais = r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\atuais.xlsx'
caminho_classificacao =r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\Classificação das Peças.xlsx'

## Pendências

pendencias = pd.read_excel(caminho_pend)

pendencias.columns

# Substitua a lista pelos nomes das colunas que você deseja visualizar
colunas_selecionadas = pendencias[[ 'Criado por',"Canal"]]

# Exibir as colunas selecionadas
print(colunas_selecionadas)

num_linhas = pendencias.shape[0]

# Exibir o número de linhas
print(f'O número de linhas em merged_table é: {num_linhas}')

import pandas as pd

# Se você já não tiver feito isso, crie a tabela dinâmica como mostrado no código fornecido
tabela_dinamica_pend = pd.pivot_table(pendencias, values='Qtd.pendente', index='Material', columns='Centro', aggfunc='sum', fill_value=0)
tabela_dinamica_pend = tabela_dinamica_pend.add_suffix('_zva70')

# Exibir a tabela dinâmica com as colunas adicionadas
print(tabela_dinamica_pend)

tabela_dinamica_pend2 = pd.pivot_table(pendencias, values='Denominação', index='Material', aggfunc='first', fill_value=0)

print(tabela_dinamica_pend2)

tabela_dinamica_pend3 = pd.pivot_table(pendencias, values='Criado por', index='Material', aggfunc=lambda x: '/'.join(x), fill_value='')
tabela_dinamica_pend4 = pd.pivot_table(pendencias, values='Canal', index='Material', aggfunc=lambda x: '/'.join(str(v) for v in x), fill_value='')

print(tabela_dinamica_pend3)
print(tabela_dinamica_pend4)

# Realizar left join entre as tabelas usando a coluna "Material" como chave
merged_table_pend = tabela_dinamica_pend.merge(tabela_dinamica_pend2, how='left', left_index=True, right_index=True)
merged_table_pend2 = merged_table_pend.merge(tabela_dinamica_pend3, how='left', left_index=True, right_index=True)
_table_pend = merged_table_pend2.merge(tabela_dinamica_pend4, how='left', left_index=True, right_index=True)
# Preencher valores nulos com zero
_table_pend = _table_pend.fillna(0)

print(_table_pend)

"""## Pendencia geral"""

pendencias_geral = pd.read_excel(caminho_pend_geral)

pendencias_geral.head

pendencias_geral.columns

tabela_dinamica_pendente = pd.pivot_table(pendencias_geral, values='Qtd.pendente', index='Material', aggfunc='sum', fill_value=0)

tabela_dinamica_pendente = tabela_dinamica_pendente.rename(columns={'Qtd.pendente': 'Pendências Gerais'})

print(tabela_dinamica_pendente)

"""## zstok"""

zstok = pd.read_excel(caminho_zstok)

zstok.columns

# Criar a tabela dinâmica mantendo 'Cen.' como texto
tabela_dinamica_zstok = pd.pivot_table(zstok, values='Utiliz.livre', index='Material', columns='Cen.', aggfunc='sum', fill_value=0)

tabela_dinamica_zstok = tabela_dinamica_zstok.add_suffix('_zstok')


# Exibir a tabela dinâmica
print(tabela_dinamica_zstok)

tabela_dinamica_zstok_qualidade = pd.pivot_table(zstok, values='Contr.qualid.', index='Material', aggfunc='sum', fill_value=0)

# Exibir a tabela dinâmica
print(tabela_dinamica_zstok_qualidade)

"""## fup

"""

fup = pd.read_excel(caminho_fup)

fup.columns

# Selecione apenas as colunas relevantes do DataFrame 'fup'
fup_relevantes = fup[['Material', 'Qtde Pedido', 'Data total']].copy()

# Converter 'Data total' para datetime e sobrescrever a coluna original
fup_relevantes['Data total'] = pd.to_datetime(fup_relevantes['Data total'], errors='coerce')

# Criar uma nova coluna 'MesAno' no formato desejado
fup_relevantes['MesAno'] = fup_relevantes['Data total'].dt.strftime('%b/%y')

# Criar a tabela dinâmica
tabela_dinamica_fup = pd.pivot_table(fup_relevantes, values='Qtde Pedido', index='Material', columns='Data total', aggfunc='sum', fill_value=0)

# Exibir a tabela dinâmica
print(tabela_dinamica_fup)

# Exibir apenas a coluna 'Material' da tabela dinâmica
coluna_material = tabela_dinamica_fup.index
print(coluna_material)

"""> Bloco com recuo

## ZTMM069
"""

fob = pd.read_excel(caminho_fob)

fob.head

fob.columns

# Criar a tabela dinâmica mantendo 'Cen.' como texto
tabela_dinamica_fob = pd.pivot_table(fob, values='Montante', index='Material', aggfunc='sum', fill_value=0)

# Exibir a tabela dinâmica
print(tabela_dinamica_fob)

"""## ZTMM085

"""

pecas = pd.read_excel(caminho_pecas)

pecas.head

pecas.columns

# Criar a tabela dinâmica mantendo 'Cen.' como texto
tabela_dinamica_pecas = pd.pivot_table(pecas, values='Modelos Lista Técnica', index='Material', aggfunc='sum', fill_value=0)

# Exibir a tabela dinâmica
print(tabela_dinamica_pecas)



"""## codigos atuais"""

atuais = pd.read_excel(caminho_atuais)

atuais

atuais.columns

resultados_filtrados = atuais[atuais['Atual'].isin(coluna_material)]

tabela_resultados = pd.DataFrame(resultados_filtrados)

# Exibir o novo DataFrame
print(tabela_resultados)

tabela_resultados = tabela_resultados.set_index(tabela_resultados.columns[0])

# Exibir o novo DataFrame
print(tabela_resultados)





"""## Classificação de peças"""

classificacao = pd.read_excel(caminho_classificacao, sheet_name='consumable parts')

classificacao.head

classificacao.columns

# Criar a tabela dinâmica mantendo 'Cen.' como texto
tabela_dinamica_classificacao = pd.pivot_table(classificacao, values='Durabilidade', index='Código', aggfunc='first', fill_value=0)


# Exibir a tabela dinâmica
print(tabela_dinamica_classificacao)

"""## juntando tabelas

juntando pendencias
"""

# Realizar left join entre as tabelas usando a coluna "Material" como chave
merged_table_pend = _table_pend.merge(tabela_dinamica_pendente, how='left', left_index=True, right_index=True)

# Preencher valores nulos com zero
merged_table_pend = merged_table_pend.fillna(0)

print(merged_table_pend)

merged_table_pend.columns

"""juntando zstok"""

# Realizar left join entre as tabelas usando a coluna "Material" como chave
merged_table_zstok = tabela_dinamica_zstok.merge(tabela_dinamica_zstok_qualidade, how='left', left_index=True, right_index=True)

# Preencher valores nulos com zero
merged_table_zstok = merged_table_zstok.fillna(0)

print(merged_table_zstok)

"""juntando pend com zstok"""

# Realizar left join entre as tabelas usando a coluna "Material" como chave
merged_table_1 = merged_table_pend.merge(merged_table_zstok, how='left', left_index=True, right_index=True)

# Preencher valores nulos com zero
merged_table_1 = merged_table_1.fillna(0)

print(merged_table_1)

merged_table_1.columns

"""colocando o fup"""

# Realizar left join entre as tabelas usando a coluna "Material" como chave
merged_table_2 = merged_table_1.merge(tabela_dinamica_fup, how='left', left_index=True, right_index=True)

# Preencher valores nulos com zero
merged_table_2 = merged_table_2.fillna(0)

print(merged_table_2)

merged_table_2.columns

# Renomear colunas com sufixo "_x"
#merged_table_2 = merged_table_2.rename(columns=lambda x: x.replace('_x', '_zva70'))

# Renomear colunas com sufixo "_y"
#merged_table_2 = merged_table_2.rename(columns=lambda x: x.replace('_y', '_zstok'))

# Realizar left join entre as tabelas usando a coluna "Material" como chave
merged_table_3 = merged_table_2.merge(tabela_dinamica_fob, how='left', left_index=True, right_index=True)

# Preencher valores nulos com zero
merged_table_3 = merged_table_3.fillna(0)

print(merged_table_3)

# Realizar left join entre as tabelas usando a coluna "Material" como chave
merged_table_4 = merged_table_3.merge(tabela_dinamica_pecas, how='left', left_index=True, right_index=True)

# Preencher valores nulos com zero
merged_table_4 = merged_table_4.fillna(0)

print(merged_table_4)

# Realizar left join entre as tabelas usando a coluna "Material" como chave
merged_table_5 = merged_table_4.merge(tabela_resultados, how='left', left_index=True, right_index=True)

# Preencher valores nulos com zero
merged_table_5 = merged_table_5.fillna(0)

print(merged_table_5)

# Realizar left join entre as tabelas usando a coluna "Material" como chave
merged_table_6 = merged_table_5.merge(tabela_dinamica_classificacao, how='left', left_index=True, right_index=True)

# Preencher valores nulos com zero
merged_table_6 = merged_table_6.fillna(0)

print(merged_table_6)

merged_table_6.columns

# Supondo que você tenha um DataFrame chamado 'merged_table_2'
# Certifique-se de ajustar isso conforme necessário com o nome correto do seu DataFrame

# Especificar o caminho do arquivo Excel de saída
caminho_saida = r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Mapa.xlsx'

# Exportar o DataFrame para o arquivo Excel
merged_table_6.to_excel(caminho_saida, index=True)

merged_table_6.head

merged_table_6 = pd.read_excel(r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\Mapa.xlsx')

ztsd017_data = pd.read_excel(r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\ztsd017.XLSX')

merged_table_6.columns = merged_table_6.columns.str.strip()
ztsd017_data.columns = ztsd017_data.columns.str.strip()

# Função para encontrar o material no Mapa em uma das colunas e retornar o valor de estoque
def buscar_estoque(material, df_estoque):
    # Verificar se o material existe no ztsd017 e pegar o valor de 'Quant estoque 09/2024'
    if material in df_estoque['Material'].values:
        return df_estoque.loc[df_estoque['Material'] == material, 'Quant estoque 24/10/2024'].values[0]
    return 0  # Se não encontrar o material, retornar 0

# Função para verificar o material ou seus similares e buscar o estoque correspondente
def buscar_estoque_similar(row, df_estoque):
    # Lista de colunas de materiais para verificar (Material, Atual, Similar1, Similar2, Similar3, Similar4)
    materiais = [row['Material'], row['Atual'], row['Similar1'], row['Similar2'], row['Similar3'], row['Similar4']]

    # Procurar em cada material e similar
    for material in materiais:
        estoque = buscar_estoque(material, df_estoque)
        if estoque > 0:
            return estoque  # Se encontrar estoque, retornar o valor

    return 0  # Retornar 0 se não encontrar estoque

# Aplicar a função no DataFrame do Mapa para trazer o estoque da planilha ztsd017
merged_table_6['Quant estoque'] = merged_table_6.apply(lambda row: buscar_estoque_similar(row, ztsd017_data), axis=1)

# Especificar o caminho do arquivo Excel de saída
caminho_saida = r'C:\Users\fsp_adolpho.salvador\Desktop\Konica Minolta\Desktop Cloud - Documentos\Desktop\Py\Pendencias\Arquivos\Mapa2.xlsx'

# Exportar o DataFrame para o arquivo Excel
merged_table_6.to_excel(caminho_saida, index=True)

