import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import re
from pathlib import Path

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

def instalar_fpdf():
    """Instala automáticamente fpdf si no está disponible"""
    try:
        from fpdf import FPDF
    except ImportError:
        import sys
        import subprocess
        print("Instalando la librería fpdf2...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "fpdf2"])
        from fpdf import FPDF
    return FPDF

def extraer_beneficiario(detalle):
    """Extrae el texto desde 'Benef:' hasta 30 caracteres después"""
    if not isinstance(detalle, str):
        return ""
    
    # Buscar la posición de 'Benef:'
    pos_benef = detalle.find('Benef:')
    if pos_benef == -1:
        return detalle[:30]  # Si no encuentra 'Benef:', tomar primeros 30 caracteres
    
    # Extraer desde 'Benef:' hasta 30 caracteres después
    texto = detalle[pos_benef:pos_benef+36]  # 'Benef:' + 30 caracteres
    return texto.strip()

def crear_pdf(resultado, escritorio_path):
    """Crea un PDF apaisado con los resultados de una cuenta"""
    FPDF = instalar_fpdf()
    
    cuenta = resultado['Cuenta']
    nombre_cuenta = resultado['NombreCuenta']
    registros = resultado['Registros']
    
    # Configurar PDF en orientación apaisada (landscape)
    class PDF(FPDF):
        def __init__(self):
            super().__init__(orientation='L')  # 'L' para landscape/apaisado
            self.set_auto_page_break(auto=True, margin=15)
        
        def header(self):
            self.set_font('Arial', 'B', 12)
            self.cell(0, 10, 'Mayor de Retenciones', 0, 1, 'C')
            self.ln(5)
        
        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')
    
    pdf = PDF()
    pdf.add_page()
    
    # Título del documento
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(0, 10, f"Mayor de retenciones - {nombre_cuenta} ({cuenta})", 0, 1)
    pdf.ln(10)
    
    # Resumen
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'Resumen:', 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f"Cantidad de movimientos: {resultado['Cantidad de Movimientos']}", 0, 1)
    pdf.cell(0, 10, f"Total Haber: ${resultado['Suma Haber']:,.2f}", 0, 1)
    pdf.ln(10)
    
    # Tabla de registros (apaisada)
    pdf.set_font('Arial', 'B', 10)
    col_widths = [25, 25, 60, 30]  # Anchos de columnas ajustados para orientación apaisada
    headers = ['Fecha', 'Número', 'Beneficiario (30 chars)', 'Haber']
    
    # Encabezados de tabla
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, 1)
    pdf.ln()
    
    # Contenido de tabla
    pdf.set_font('Arial', '', 8)
    for _, row in registros.iterrows():
        # Asegurar que los valores no sean NaN
        fecha = str(row['Fecha']) if pd.notna(row['Fecha']) else ''
        numero = str(row['Número']) if pd.notna(row['Número']) else ''
        beneficiario = extraer_beneficiario(str(row['Detalle'])) if pd.notna(row['Detalle']) else ''
        haber = f"${row['Haber']:,.2f}" if pd.notna(row['Haber']) else '$0.00'
        
        pdf.cell(col_widths[0], 10, fecha, 1)
        pdf.cell(col_widths[1], 10, numero, 1)
        pdf.cell(col_widths[2], 10, beneficiario, 1)
        pdf.cell(col_widths[3], 10, haber, 1)
        pdf.ln()
    
    # Guardar PDF
    nombre_archivo = f"Mayor de retenciones - {nombre_cuenta}.pdf"
    nombre_archivo = re.sub(r'[\\/*?:"<>|]', '', nombre_archivo)  # Eliminar caracteres inválidos
    pdf_path = os.path.join(escritorio_path, nombre_archivo)
    pdf.output(pdf_path)
    return pdf_path

def procesar_archivo():
    """Procesa el archivo Excel y genera los PDFs apaisados"""
    # Configurar interfaz para seleccionar archivo
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
        # Leer archivo Excel
        df = pd.read_excel(archivo_origen, sheet_name='MayorDeCuentas', skiprows=1)
        
        # Identificar cuentas contables
        es_cuenta = df['Número'].astype(str).str.contains(r'^2\.1\.1\.4\.1\.\d+$', na=False)
        cuentas = df[es_cuenta]['Número'].unique()
        cuentas_validas = [c for c in cuentas if str(c) in NOMBRES_CUENTAS]
        
        print(f"Cuentas contables encontradas: {cuentas_validas}")
        
        # Procesar cada cuenta
        resultados = []
        for cuenta in cuentas_validas:
            nombre_cuenta = NOMBRES_CUENTAS[str(cuenta)]
            idx_cuenta = df[df['Número'] == cuenta].index[0]
            
            # Encontrar próxima cuenta
            siguiente_idx = len(df)
            for next_cuenta in cuentas_validas:
                next_idx = df[df['Número'] == next_cuenta].index
                if len(next_idx) > 0 and next_idx[0] > idx_cuenta:
                    siguiente_idx = next_idx[0]
                    break
            
            # Filtrar registros válidos
            registros_cuenta = df.iloc[idx_cuenta+1:siguiente_idx].copy()
            registros_cuenta['Haber'] = pd.to_numeric(
                registros_cuenta['Haber'].astype(str).str.replace('.', '').str.replace(',', '.'), 
                errors='coerce'
            )
            
            filtro = (
                registros_cuenta['Detalle'].str.contains('Benef', na=False, case=False) &
                ~registros_cuenta['Detalle'].str.contains('CUT', na=False, case=False) &
                (registros_cuenta['Haber'] > 0))
            
            registros_filtrados = registros_cuenta[filtro].copy()
            
            if len(registros_filtrados) > 0:
                resultados.append({
                    'Cuenta': cuenta,
                    'NombreCuenta': nombre_cuenta,
                    'Cantidad de Movimientos': len(registros_filtrados),
                    'Suma Haber': registros_filtrados['Haber'].sum(),
                    'Registros': registros_filtrados[['Fecha', 'Número', 'Detalle', 'Haber']]
                })
        
        # Generar PDFs
        if resultados:
            escritorio_path = str(Path.home() / "Desktop")
            for resultado in resultados:
                try:
                    pdf_path = crear_pdf(resultado, escritorio_path)
                    print(f"PDF creado: {pdf_path}")
                except Exception as e:
                    print(f"Error al crear PDF para {resultado['NombreCuenta']}: {str(e)}")
            
            print(f"\nProceso completado. Se generaron {len(resultados)} archivos PDF apaisados en el escritorio.")
        else:
            print("No se encontraron movimientos que cumplan los criterios especificados.")
    
    except Exception as e:
        print(f"Error al procesar el archivo: {str(e)}")

if __name__ == "__main__":
    procesar_archivo()