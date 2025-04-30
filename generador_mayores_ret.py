import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def procesar_archivo():
    # Seleccionar archivo de origen
    root = tk.Tk()
    root.withdraw()
    
    archivo_origen = filedialog.askopenfilename(
        title="Seleccione el archivo Excel de origen",
        filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
    )
    
    if not archivo_origen:
        print("No se seleccionó ningún archivo.")
        return
    
    try:
        # Leer el archivo Excel
        df = pd.read_excel(archivo_origen, sheet_name='MayorDeCuentas')
        
        # Identificar las filas que son encabezados de cuenta (contienen '2.1.1.4.1.')
        es_cuenta = df['Número'].astype(str).str.contains(r'^2\.1\.1\.4\.1\.\d+$', na=False)
        cuentas = df[es_cuenta]['Número'].unique()
        
        print(f"Cuentas contables encontradas: {cuentas}")
        
        resultados = []
        
        for cuenta in cuentas:
            # Encontrar el índice de la fila que contiene la cuenta
            idx_cuenta = df[df['Número'] == cuenta].index[0]
            
            # Encontrar el próximo índice de cuenta para determinar el rango de registros
            siguiente_idx = len(df)
            for next_cuenta in cuentas:
                if df[df['Número'] == next_cuenta].index[0] > idx_cuenta:
                    siguiente_idx = df[df['Número'] == next_cuenta].index[0]
                    break
            
            # Obtener los registros entre esta cuenta y la siguiente
            registros_cuenta = df.iloc[idx_cuenta+1:siguiente_idx].copy()
            
            # Filtrar registros que contengan "Benef" pero no "CUT"
            filtro_benef = registros_cuenta['Detalle'].str.contains('Benef', na=False, case=False)
            filtro_no_cut = ~registros_cuenta['Detalle'].str.contains('CUT', na=False, case=False)
            
            registros_filtrados = registros_cuenta[filtro_benef & filtro_no_cut].copy()
            
            # Calcular suma de Haber (convertir a numérico por si hay formatos de texto)
            registros_filtrados['Haber'] = pd.to_numeric(
                registros_filtrados['Haber'].astype(str).str.replace('.', '').str.replace(',', '.'),
                errors='coerce'
            )
            suma_haber = registros_filtrados['Haber'].sum()
            
            # Agregar al resultado
            resultados.append({
                'Cuenta': cuenta,
                'Cantidad de Movimientos': len(registros_filtrados),
                'Suma Haber': suma_haber,
                'Registros': registros_filtrados[['Fecha', 'Número', 'Detalle', 'Haber']]
            })
        
        # Generar archivo de salida
        nombre_base = os.path.splitext(os.path.basename(archivo_origen))[0]
        archivo_salida = f"{nombre_base}_resultados.xlsx"
        
        with pd.ExcelWriter(archivo_salida) as writer:
            for resultado in resultados:
                # Usar el código de cuenta como nombre de hoja (reemplazando puntos por guiones)
                nombre_hoja = str(resultado['Cuenta']).replace('.', '_')[:31]
                
                # Crear DataFrames para el resumen y los registros
                resumen_df = pd.DataFrame({
                    'Cuenta': [resultado['Cuenta']],
                    'Cantidad de Movimientos': [resultado['Cantidad de Movimientos']],
                    'Suma Haber': [resultado['Suma Haber']]
                })
                
                # Escribir resumen
                resumen_df.to_excel(writer, sheet_name=nombre_hoja, index=False)
                
                # Escribir registros detallados
                resultado['Registros'].to_excel(
                    writer,
                    sheet_name=nombre_hoja,
                    startrow=len(resumen_df) + 2,
                    index=False
                )
        
        print(f"Procesamiento completado. Resultados guardados en: {archivo_salida}")
        
    except Exception as e:
        print(f"Error al procesar el archivo: {e}")
        raise  # Esto ayuda a ver el traceback completo durante el desarrollo

if __name__ == "__main__":
    procesar_archivo()