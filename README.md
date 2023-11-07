# Armazendo_dados_excel
Exemplo de armazenar dados de um arquivo excle .xls, que estão armazenado em uma tabela dinamica.

Um arquivo Excel no formato .xls usando a biblioteca xlrd e convertê-lo em um DataFrame do pandas. 
Abaixo a explicação passo a passo do que o código faz:

import xlrd e import pandas as pd: Importa as bibliotecas necessárias, xlrd para ler o arquivo Excel e pandas para trabalhar com DataFrames.

excel_file_path: Especifica o caminho para o arquivo Excel .xls que você deseja abrir. Você deve substituir "/content/vendas-combustiveis-m3.xls" pelo caminho correto do seu arquivo.

workbook: Abre o arquivo Excel especificado usando xlrd e armazena-o na variável workbook.

worksheet: Acesse a planilha que contém a tabela dinâmica. Neste caso, é a planilha com o nome 'Plan1'.

Defina as variáveis start_row, end_row, start_col, end_col para determinar a área da tabela dinâmica que você deseja ler.

Crie uma lista chamada data para armazenar os dados brutos da tabela dinâmica.

Itere sobre as linhas e colunas da tabela dinâmica e extraia os valores das células usando worksheet.cell_value(rowx, colx). Os valores são armazenados em row_data e depois adicionados à lista data.

column_labels: Obtenha os rótulos de coluna da primeira linha da tabela dinâmica (geralmente cabeçalhos) e armazene-os em uma lista chamada column_labels.

df: Crie um DataFrame do pandas usando a lista data como os dados brutos e column_labels como rótulos de coluna.

Arredonde os valores do DataFrame para 0 casas decimais usando df.round(0).

Renomeie as colunas que contêm pontos flutuantes (por exemplo, de 2014.0 para 2014) usando df.columns = df.columns.astype(str).str.split('.').str[0].

Exiba o DataFrame na tela usando print(df).

Este código lê os dados da tabela dinâmica em um arquivo Excel .xls, arredonda os valores e renomeia automaticamente as colunas, conforme descrito anteriormente. Certifique-se de ajustar o caminho do arquivo e as variáveis start_row, end_row, start_col e end_col de acordo com a sua própria planilha.
