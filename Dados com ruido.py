import pandas as pd
import numpy as np
import xlsxwriter

def impd2(w, R, t, a):
    f = R / (1 + (1j * w * t) ** (1 - a))
    return f

# Parâmetros fixos
R = 800
a = 0.2
t = 10**-3
sigma_R = 0.05 * R  # Incerteza de 5% em R
sigma_t = 0.05 * t  # Incerteza de 5% em t

# Intervalo para omega
omega_values = np.linspace(300, 10000, 1000)

# Lista para armazenar os resultados
data = []

# Calculando os valores para cada omega e adicionando o ruído
for omega in omega_values:
    noisy_R = np.random.normal(R, sigma_R)  # Gerando valor de R com incerteza
    noisy_t = np.random.normal(t, sigma_t)  # Gerando valor de t com incerteza

    result = impd2(omega, noisy_R, noisy_t, a)
    data.append((omega, result.real, result.imag))

# Criando um DataFrame do pandas com os dados
df = pd.DataFrame(data, columns=['omega', 'Re', 'Im'])

# Especificando o caminho completo para o arquivo de saída
output_path = r"E:\eduro\Documents\Graduação ITA\Iniciação científica\Criação de dados para testar\dados_impedancia.xlsx"

# Criando o objeto ExcelWriter com xlsxwriter
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Escrevendo o DataFrame na planilha Sheet1
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Obtendo o objeto Worksheet
    worksheet = writer.sheets['Sheet1']

    # Criando a tabela na planilha Sheet1
    num_rows, num_cols = df.shape
    table_range = f'A1:C{num_rows+1}'  # +1 para incluir o cabeçalho
    worksheet.add_table(table_range, {'name': 'Table1', 'style': 'Table Style Medium 1'})

print(f"Dados salvos em '{output_path}'.")