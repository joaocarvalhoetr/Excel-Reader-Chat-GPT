import pandas as pd
import ast
import csv

# Caminho para o ficheiro CSV
csv_file_path = 'output.csv'

# Função para converter strings de dicionários em dicionários reais
def convert_to_dict(s):
    try:
        return ast.literal_eval(s)
    except ValueError:
        return None

# Ler o ficheiro CSV linha por linha
data = []
with open(csv_file_path, 'r', encoding='utf-8') as f:
    reader = csv.reader(f)
    for row in reader:
        if row:  # Verifica se a linha não está vazia
            # Converte a string do dicionário em um dicionário real
            dict_data = convert_to_dict(row[0])
            if dict_data:
                data.append(dict_data)

# Criar um DataFrame a partir da lista de dicionários
df = pd.DataFrame(data)

# Caminho para o ficheiro Excel de saída
excel_file_path = 'output.xlsx'

# Guardar o DataFrame num ficheiro Excel
df.to_excel(excel_file_path, index=False)

print(f'Dados exportados para {excel_file_path} com sucesso.')
