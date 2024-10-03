import pandas as pd
import openpyxl
import os
import sys
import tkinter as tk
from tkinter import messagebox

def operadores_mera_s2s(archivo_origen, archivo_destino):
    try:
        # Verificar si los archivos existen
        if not os.path.exists(archivo_origen):
            messagebox.showerror("Error", f"El archivo origen '{archivo_origen}' no existe.")
            return
        if not os.path.exists(archivo_destino):
            messagebox.showinfo("Información", f"El archivo destino '{archivo_destino}' no existe. Se creará uno vacío.")
            # Crear un libro de Excel vacío si no existe
            wb = openpyxl.Workbook()
            wb.save(archivo_destino)

        # Cargar el archivo de origen (CSV MERA) usando pandas
        hoja_origen = 'Operadores S2S'
        df_origen = pd.read_excel(archivo_origen, sheet_name=hoja_origen, engine='openpyxl')

        # Filtrar los datos
        # 1. Excluir filas con cualquier NaN (equivalente a #N/D)
        df_filtrado = df_origen.dropna()

        # 2. Excluir filas donde 'Ll. ACD' == 0
        if 'Ll. ACD' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['Ll. ACD'] != 0]
        else:
            messagebox.showerror("Error", "La columna 'Ll. ACD' no se encuentra en los datos.")
            return

        # 3. Excluir filas que contienen "(en blanco)" en cualquier celda
        df_filtrado = df_filtrado[~df_filtrado.apply(lambda row: row.astype(str).str.contains(r'\(en blanco\)', case=False).any(), axis=1)]

        # Cargar el archivo destino (Mera) usando openpyxl
        wb_destino = openpyxl.load_workbook(archivo_destino)
        
        # Renombrar la hoja 'Envío' a 'Eliminar' si existe
        if "Envío" in wb_destino.sheetnames:
            hoja_envio = wb_destino["Envío"]
            hoja_envio.title = "Eliminar"
            print("Hoja 'Envío' renombrada a 'Eliminar'.")
        else:
            print("La hoja 'Envío' no existe en el archivo destino. Se procederá a crear una nueva hoja 'Envío'.")

        # Escribir los datos filtrados en una nueva hoja llamada 'Operadores S2S' (se reemplazará si existe)
        with pd.ExcelWriter(archivo_destino, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_filtrado.to_excel(writer, sheet_name="Operadores S2S", index=False)
            print("Datos filtrados copiados a la hoja 'Operadores S2S'.")

        # Renombrar la hoja 'Operadores S2S' a 'Envío'
        wb_destino = openpyxl.load_workbook(archivo_destino)
        if "Operadores S2S" in wb_destino.sheetnames:
            hoja_nueva = wb_destino["Operadores S2S"]
            hoja_nueva.title = "Envío"
            print("Hoja 'Operadores S2S' renombrada a 'Envío'.")
        else:
            print("La hoja 'Operadores S2S' no se encontró después de la escritura.")

        # Eliminar la hoja 'Eliminar' si existe
        if "Eliminar" in wb_destino.sheetnames:
            del wb_destino["Eliminar"]
            print("Hoja 'Eliminar' eliminada.")
        else:
            print("La hoja 'Eliminar' no existe y no se puede eliminar.")

        # Guardar y cerrar el archivo destino
        wb_destino.save(archivo_destino)
        wb_destino.close()
        messagebox.showinfo("Éxito", f"Operación completada. Archivo guardado en '{archivo_destino}'.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

def main():
    # Inicializar Tkinter
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Rutas de los archivos (ajusta según tu entorno)
    directorio = os.path.dirname(os.path.abspath(sys.argv[0]))
    archivo_origen = os.path.join(directorio, "Reporte Operadores s2s - CSV MERA.xlsm")
    archivo_destino = os.path.join(directorio, "Reporte Operadores s2s - Mera.xlsx")

    operadores_mera_s2s(archivo_origen, archivo_destino)

if __name__ == "__main__":
    main()
