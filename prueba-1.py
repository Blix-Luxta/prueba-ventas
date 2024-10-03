import pandas as pd
import openpyxl

# Leer un archivo Excel
df = pd.read_excel('Reporte Operadores s2s - Mera.xlsx', header=6, usecols="B:I")

df.columns = ['PCRC', 'OPERADOR', 'Cod. Agente', 'Ll. ACD', 'LOGUEO', 'Q. Ventas', 'vma', 'Supervisor']

df_filtrado = df[df['Ll. ACD'] > 0]

df_filtrado = df_filtrado.dropna()

print(df_filtrado)