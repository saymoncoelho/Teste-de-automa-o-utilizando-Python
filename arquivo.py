import pandas as pd

# Carregar o arquivo Excel
file_path = 'C:/Users/saymo/OneDrive/Documentos/Projetos/Faturamento_original.xlsx'
df = pd.read_excel(file_path)

# Mapeamento de valores para as colunas de dólar e euro
dolar_values = {
    '0': 0.178830, '1': 0.178830, '2': 0.178830,
    '3': 0.176835, '4': 0.176835, '5': 0.176835,
    '6': 0.176196, '7': 0.176196, '8': 0.176196,
    '9': 0.17883, '10': 0.17883, '11': 0.178830,
    '12': 0.18386, '13': 0.18386, '14': 0.18386
}

euro_values = {
    '0': 0.1663, '1': 0.1663, '2': 0.1663,
    '3': 0.1644, '4': 0.1644, '5': 0.1644,
    '6': 0.1637, '7': 0.1637, '8': 0.1637,
    '9': 0.1666, '10': 0.1666, '11': 0.1666,
    '12': 0.1679, '13': 0.1679, '14': 0.1679
}


# Adicionar coluna 'O Dólar do fechamento do dia'
df['O Dólar do fechamento do dia'] = df.index.map(str).map(dolar_values)

# Adicionar coluna 'O Euro do fechamento do dia'
df['O Euro do fechamento do dia'] = df.index.map(str).map(euro_values)

df['Faturamento em Dólar'] = df['Faturamento em real'] * df['O Dólar do fechamento do dia']

df['Faturamento em Euro'] = df['Faturamento em real'] * df['O Euro do fechamento do dia']

new_file_path = 'C:/Users/saymo/OneDrive/Documentos/Projetos/Faturamento_novo.xlsx'
df.to_excel(new_file_path, index=False)

print("Colunas adicionadas e valores inseridos com sucesso!")
def formatar(valor):
    return "{:,.2f}".format(valor)

caminho_arquivo = "C:/Users/saymo/OneDrive/Documentos/Projetos/Faturamento_novo.xlsx"
tabela = pd.read_excel(caminho_arquivo)

tabela['Faturamento em Dólar'] = tabela['Faturamento em Dólar'].apply(formatar)
tabela['Faturamento em Euro'] = tabela['Faturamento em Euro'].apply(formatar)

print(tabela.head())

tabela.to_excel(caminho_arquivo, index=False)
print(f"Alterações salvas no arquivo {caminho_arquivo}")