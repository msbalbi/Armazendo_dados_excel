import xlrd
import pandas as pd

# Especifique o caminho para o seu arquivo Excel .xls
excel_file_path = "/content/vendas-combustiveis-m3.xls"

# Abra o arquivo Excel com xlrd
workbook = xlrd.open_workbook(excel_file_path)

# Acesse a planilha que contém a tabela dinâmica (substitua 'Planilha1' pelo nome correto)
worksheet = workbook.sheet_by_name('Plan1')

# Defina as linhas e colunas onde a tabela dinâmica está localizada
# Substitua esses valores pelos limites de onde sua tabela dinâmica começa e termina
start_row = 53  # A primeira linha da tabela dinâmica (geralmente os cabeçalhos)
end_row = 66    # A última linha da tabela dinâmica
start_col = 1  # A primeira coluna da tabela dinâmica
end_col = 23     # A última coluna da tabela dinâmica

# Inicialize uma lista para armazenar os dados brutos
data = []

# Itere sobre as linhas e colunas para extrair os dados brutos
for rowx in range(start_row, end_row + 1):
    row_data = []
    for colx in range(start_col, end_col + 1):
        cell_value = worksheet.cell_value(rowx, colx)
        row_data.append(cell_value)
    data.append(row_data)

# Obtenha a primeira linha da planilha como rótulos de coluna
column_labels = [worksheet.cell_value(start_row - 1, colx) for colx in range(start_col, end_col + 1)]

# Crie um DataFrame do pandas com os dados brutos e rótulos de coluna
df = pd.DataFrame(data, columns=column_labels)

# Arredonde os valores no DataFrame
df = df.round(0)

# Substitua os nomes de coluna que contêm ponto flutuante (por exemplo, 2014.0 para 2014)
df.columns = df.columns.astype(str).str.split('.').str[0]

# Exiba o DataFrame
display(df)
