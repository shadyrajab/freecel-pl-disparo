import pandas as pd
from datetime import datetime
import re

def formatar_cep(cep):
  string_cep = str(cep)
  cep_splited = string_cep.split('.')
  cep = cep_splited[0]

  return cep

def formatar_endereco(n):
  n1 = str(n).split(' ')[0].split('-')[0]
  if n1 == 'SN' or n1 == 'SN1' or n1 == 'S/N':
    return 'SN'
  if n1.isnumeric():
    return n1
  else:
    endereco_final = re.sub('[^0-9]', '', n1)
    return endereco_final

cidade = input('Nome do município: ')

df_vivo = pd.read_excel(f'vivo_empresas_{cidade}.xlsx', na_values=None)
df_empresas = pd.read_excel(f'empresas_aqui_{cidade}.xlsx', na_values=None)

df_vivo.dropna(subset=['CEP', 'NUM'], inplace=True)
df_empresas.dropna(subset=['CEP', 'Número', 'CNPJ', 'Telefone 1'], inplace=True)

df_vivo.drop(axis=1, columns=['ARMARIO', 'CAIXA', 'FIBRA', 'CAPACIDADE CAIXA', 'OCUPACAO CAIXA', 'RELEVANCIA'], inplace=True)
df_empresas.drop(axis=1, columns=['Telefone 2'], inplace=True)

df_empresas.drop_duplicates(subset=['CNPJ', 'Telefone 1', 'Razão', 'Fantasia'])

df_vivo['NUM'] = df_vivo['NUM'].apply(lambda endereco: formatar_endereco(endereco))
df_empresas['Número'] = df_empresas['Número'].apply(lambda endereco: formatar_endereco(endereco))

df_vivo['CEP'] = df_vivo['CEP'].apply(lambda cep: formatar_cep(cep))
df_empresas['CEP'] = df_empresas['CEP'].apply(lambda cep: formatar_cep(cep))

df_vivo['CEP_ENDERECO'] = df_vivo['CEP'].map(str) + '-' + df_vivo['NUM']
df_empresas['CEP_ENDERECO'] = df_empresas['CEP'].map(str) + '-' + df_empresas['Número']

new_data_frame = pd.merge(df_empresas, df_vivo, how='left', on='CEP_ENDERECO')

new_data_frame.drop_duplicates(subset=['Razão', 'CNPJ', 'Telefone 1'], inplace=True)
new_data_frame.dropna(subset=['CNPJ', 'Telefone 1', 'Razão'])
new_data_frame.drop_duplicates(subset=['Telefone 1'], inplace=True)
new_data_frame.drop(axis=1, columns=['CEP_y', 'UF_y', 'NUM', 'TIPO', 'CEP_ENDERECO', 'Nome do Sócio', 'BAIRRO', 'CIDADE'], inplace=True)

new_data_frame.rename(columns={'UF_x': 'UF', 'CEP_x': 'CEP', 'Telefone 1': 'Telefone', 'LOGRADOURO': 'Logradouro'}, inplace=True)

new_data_frame = new_data_frame[['CNPJ', 'Razão', 'Fantasia', 'UF', 'Cidade', 'Endereço', 'Tipo', 'Bairro', 'Logradouro', 'Complemento', 'Número', 'CEP', 'Telefone', 'E-mail', 'Situação Cad.', 'Data Situação Cad.', 'Data Início Atv.']]

now = datetime.now()

new_data_frame.to_excel(f'./cidades/{cidade}/com cobertura {cidade} - {now.strftime('%d-%m-%Y')}.xlsx', index=False)
com_cobertura = pd.read_excel(f'./cidades/{cidade}/com cobertura {cidade} - {now.strftime('%d-%m-%Y')}.xlsx', na_values=None)

com_cobertura.dropna(subset=['Número', 'Telefone', 'CEP', 'CNPJ'], inplace=True)
com_cobertura.to_excel(f'./cidades/{cidade}/com cobertura {cidade} - {now.strftime('%d-%m-%Y')}.xlsx', index=False)