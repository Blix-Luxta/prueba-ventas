import pandas as pd
import openpyxl

# Leer un archivo Excel
df = pd.read_excel('Reporte Operadores s2s - Mera.xlsx', header=6, usecols="B:I")
print(df.shape)
df.columns = ['PCRC', 'OPERADOR', 'Cod. Agente', 'Ll. ACD', 'LOGUEO', 'Q. Ventas', 'vma', 'Supervisor']
df['Ll. ACD'] = pd.to_numeric(df['Ll. ACD'], errors='coerce')


if 'Ll. ACD' in df:
    df_filtrado = df[df['Ll. ACD'] != 0]
    print(df_filtrado.shape)
else:
    print("Error", "La columna 'Ll. ACD' no se encuentra en los datos.")

df_filtrado = df_filtrado[df_filtrado["OPERADOR"] != "#N/D"]

print(df_filtrado.shape)

df_filtrado.to_excel('Reporte Operadores s2s - Mera.xlsx', index=False, engine='openpyxl')
