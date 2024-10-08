import pandas as pd
import openpyxl
import os
import sys

def operadores_mera_s2s(archivo_origen, archivo_destino):
    try:
        # Verificar si los archivos existen
        if not os.path.exists(archivo_origen):
            print("Error", f"El archivo origen '{archivo_origen}' no existe.")
            return
        if not os.path.exists(archivo_destino):
            print("Información", f"El archivo destino '{archivo_destino}' no existe. Se creará uno vacío.")
            # Crear un libro de Excel vacío si no existe
            wb = openpyxl.Workbook()
            wb.save(archivo_destino)
            
        hoja_origen = 'Operadores S2S'
        # Leer un archivo Excel
        df_origen = pd.read_excel(archivo_origen, sheet_name=hoja_origen, engine='openpyxl', header=6, usecols="B:I")
        print(df_origen.shape)
        df_origen.columns = ['PCRC', 'OPERADOR', 'Cod. Agente', 'Ll. ACD', 'LOGUEO', 'Q. Ventas', 'vma', 'Supervisor']
        df_origen['Ll. ACD'] = pd.to_numeric(df_origen['Ll. ACD'], errors='coerce')


        if 'Ll. ACD' in df_origen:
            df_filtrado = df_origen[df_origen['Ll. ACD'] != 0]
            print(df_filtrado.shape)
        else:
            print("Error", "La columna 'Ll. ACD' no se encuentra en los datos.")

        df_filtrado = df_filtrado[df_filtrado["OPERADOR"] != "#N/D"]
        df_filtrado = df_filtrado[~df_filtrado.apply(lambda row: row.astype(str).str.contains(r'\(en blanco\)', case=False).any(), axis=1)]

        print(df_filtrado.shape)

        wb_destino = openpyxl.load_workbook(archivo_destino)
            
        # Renombrar la hoja 'Envío' a 'Eliminar' si existe
        if "Envío" in wb_destino.sheetnames:
            hoja_envio = wb_destino["Envío"]
            hoja_envio.title = "Eliminar"
            wb_destino.save(archivo_destino)
            print("Hoja 'Envío' renombrada a 'Eliminar'.")
        else:
            print("La hoja 'Envío' no existe en el archivo destino.")
            
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
        print("Éxito", f"Operación completada. Archivo guardado en '{archivo_destino}'.")
        
    except Exception as e:
        print("Error", f"Ocurrió un error: {e}")


def main():

    # Rutas de los archivos (ajusta según tu entorno)
    directorio = os.path.dirname(os.path.abspath(sys.argv[0]))
    archivo_origen = os.path.join(directorio, "Reporte Operadores s2s - CSV MERA.xlsm")
    archivo_destino = os.path.join(directorio, "Reporte Operadores s2s - Mera.xlsx")

    operadores_mera_s2s(archivo_origen, archivo_destino)
    
if __name__ == "__main__":
    main()