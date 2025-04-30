import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import re

# Diccionario de nombres de cuentas predefinidos
NOMBRES_CUENTAS = {
    '2.1.1.4.1.1': 'RETENCIONES IMPUESTO A LAS GANANCIAS',
    '2.1.1.4.1.14': 'RETENCIONES INGRESOS BRUTOS A PAGAR',
    '2.1.1.4.1.17': 'RETENCION SEGUROS CONTRATOS F.A',
    '2.1.1.4.1.57': 'RETENCIONES IVA',
    '2.1.1.4.1.6': 'RETENCIONES AL SUSS SERVICIO LIMPIEZA',
    '2.1.1.4.1.7': 'RETENCIONES AL SUSS OBRAS',
    '2.1.1.4.1.8': 'RETENCIONES AL SUSS RESOLUCION GENERAL',
    '2.1.1.4.1.9': 'RETENCION SUSS SEGURIDAD'
}

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
        # Leer el archivo Excel omitiendo la primera fila (encabezado combinado)
        df = pd.read_excel(archivo_origen, sheet_name='MayorDeCuentas', skiprows=1)
        
        # Identificar las filas que son encabezados de cuenta
        es_cuenta = df['Número'].astype(str).str.contains(r'^2\.1\.1\.4\.1\.\d+$', na=False)
        cuentas = df[es_cuenta]['Número'].unique()
        
        # Filtrar solo las cuentas que están en nuestro diccionario
        cuentas_validas = [c for c in cuentas if str(c) in NOMBRES_CUENTAS]
        
        print(f"Cuentas contables encontradas: {cuentas_validas}")
        
        resultados = []
        
        for cuenta in cuentas_validas:
            # Obtener el nombre de la cuenta del diccionario
            nombre_cuenta = NOMBRES_CUENTAS[str(cuenta)]
            
            # Encontrar el índice de la fila que contiene la cuenta
            idx_cuenta = df[df['Número'] == cuenta].index[0]
            
            # Encontrar el próximo índice de cuenta para determinar el rango de registros
            siguiente_idx = len(df)
            for next_cuenta in cuentas_validas:
                next_idx = df[df['Número'] == next_cuenta].index
                if len(next_idx) > 0 and next_idx[0] > idx_cuenta:
                    siguiente_idx = next_idx[0]
                    break
            
            # Obtener los registros entre esta cuenta y la siguiente
            registros_cuenta = df.iloc[idx_cuenta+1:siguiente_idx].copy()
            
            # Filtrar registros que contengan "Benef" pero no "CUT" y que tengan valor en Haber
            filtro_benef = registros_cuenta['Detalle'].str.contains('Benef', na=False, case=False)
            filtro_no_cut = ~registros_cuenta['Detalle'].str.contains('CUT', na=False, case=False)
            registros_cuenta['Haber'] = pd.to_numeric(
                registros_cuenta['Haber'].astype(str).str.replace('.', '').str.replace(',', '.'), 
                errors='coerce'
            )
            filtro_haber = registros_cuenta['Haber'] > 0
            
            registros_filtrados = registros_cuenta[filtro_benef & filtro_no_cut & filtro_haber].copy()
            
            # Si no hay movimientos, saltar esta cuenta
            if len(registros_filtrados) == 0:
                continue
            
            # Calcular suma de Haber
            suma_haber = registros_filtrados['Haber'].sum()
            
            # Agregar al resultado
            resultados.append({
                'Cuenta': cuenta,
                'NombreCuenta': nombre_cuenta,
                'Cantidad de Movimientos': len(registros_filtrados),
                'Suma Haber': suma_haber,
                'Registros': registros_filtrados[['Fecha', 'Número', 'Detalle', 'Haber']]
            })
        
        # Generar archivo de salida solo si hay resultados
        if resultados:
            nombre_base = os.path.splitext(os.path.basename(archivo_origen))[0]
            archivo_salida = f"{nombre_base}_resultados.xlsx"
            
            with pd.ExcelWriter(archivo_salida) as writer:
                for resultado in resultados:
                    # Crear nombre de hoja con código y nombre de cuenta
                    nombre_completo = f"{resultado['Cuenta']} - {resultado['NombreCuenta']}"
                    nombre_hoja = nombre_completo[:31]
                    nombre_hoja = re.sub(r'[\\/*?:[\]]', '', nombre_hoja)  # Eliminar caracteres no permitidos
                    
                    # Crear DataFrames para el resumen
                    resumen_df = pd.DataFrame({
                        'Cuenta': [nombre_completo],
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
        else:
            print("No se encontraron movimientos que cumplan los criterios especificados.")
        
    except Exception as e:
        print(f"Error al procesar el archivo: {e}")

if __name__ == "__main__":
    procesar_archivo()