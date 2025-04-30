import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import re
from pathlib import Path
import locale
from decimal import Decimal, InvalidOperation

# Configurar locale para formato numérico
try:
    locale.setlocale(locale.LC_ALL, 'es_AR.UTF-8')  # Ajustar según tu región
except:
    locale.setlocale(locale.LC_ALL, '')  # Usar configuración regional por defecto

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
    
    pos_benef = detalle.find('Benef:')
    if pos_benef == -1:
        return detalle[:30]
    
    texto = detalle[pos_benef:pos_benef+36]
    return texto.strip()

def convertir_a_decimal(valor):
    """Convierte cualquier formato numérico a Decimal con precisión"""
    if pd.isna(valor):
        return Decimal('0')
    
    try:
        if isinstance(valor, (Decimal, float, int)):
            return Decimal(str(valor))
        
        str_valor = str(valor).strip()
        str_valor = str_valor.replace('$', '').replace(' ', '')
        
        # Contar separadores de miles y decimales
        puntos = str_valor.count('.')
        comas = str_valor.count(',')
        
        if puntos == 1 and comas == 1:
            # Determinar cuál es el separador decimal
            if str_valor.rfind('.') > str_valor.rfind(','):
                # Formato 1.234,56 → 1234.56
                str_valor = str_valor.replace('.', '').replace(',', '.')
            else:
                # Formato 1,234.56 → 1234.56
                str_valor = str_valor.replace(',', '')
        elif comas == 1:
            # Formato 1234,56 → 1234.56
            str_valor = str_valor.replace(',', '.')
        elif puntos == 1:
            # Podría ser 1.234 (mil doscientos treinta y cuatro) o 1.234 (uno punto dos tres cuatro)
            # Asumimos que es separador de miles si hay exactamente 3 decimales
            partes = str_valor.split('.')
            if len(partes) == 2 and len(partes[1]) == 3:
                # Formato 1.234 → 1234
                str_valor = str_valor.replace('.', '')
        
        return Decimal(str_valor)
    except (InvalidOperation, ValueError, TypeError):
        return Decimal('0')

def format_currency(value):
    """Formatea correctamente valores monetarios según locale"""
    try:
        decimal_value = convertir_a_decimal(value)
        formatted = locale.currency(decimal_value, grouping=True, symbol='$')
        
        # Asegurar formato consistente para Argentina (14.041,36)
        if ',' in formatted and '.' in formatted:
            # Si el locale no configuró correctamente los separadores
            parts = formatted.split(',')
            integer_part = parts[0].replace('.', '').replace('$', '').strip()
            decimal_part = parts[1][:2]
            formatted = f"${integer_part}.{decimal_part}"
        
        return formatted
    except:
        return '$0.00'

def crear_pdf(resultado, escritorio_path):
    """Crea un PDF apaisado con los resultados de una cuenta"""
    FPDF = instalar_fpdf()
    
    cuenta = resultado['Cuenta']
    nombre_cuenta = resultado['NombreCuenta']
    registros = resultado['Registros']
    
    # Configurar PDF en orientación apaisada (landscape)
    class PDF(FPDF):
        def __init__(self):
            super().__init__(orientation='L')
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
    
    # Formateo correcto del total
    total_formateado = format_currency(resultado['Suma Haber'])
    
    pdf.cell(0, 10, f"Cantidad de movimientos: {resultado['Cantidad de Movimientos']}", 0, 1)
    pdf.cell(0, 10, f"Total Haber: {total_formateado}", 0, 1)
    pdf.ln(10)
    
    # Tabla de registros (apaisada)
    pdf.set_font('Arial', 'B', 10)
    col_widths = [25, 25, 60, 30]
    headers = ['Fecha', 'Número', 'Beneficiario (30 chars)', 'Haber']
    
    # Encabezados de tabla
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, 1)
    pdf.ln()
    
    # Contenido de tabla
    pdf.set_font('Arial', '', 8)
    for _, row in registros.iterrows():
        fecha = str(row['Fecha']) if pd.notna(row['Fecha']) else ''
        numero = str(row['Número']) if pd.notna(row['Número']) else ''
        beneficiario = extraer_beneficiario(str(row['Detalle'])) if pd.notna(row['Detalle']) else ''
        haber = format_currency(row['Haber'])
        
        pdf.cell(col_widths[0], 10, fecha, 1)
        pdf.cell(col_widths[1], 10, numero, 1)
        pdf.cell(col_widths[2], 10, beneficiario, 1)
        pdf.cell(col_widths[3], 10, haber, 1)
        pdf.ln()
    
    # Guardar PDF
    nombre_archivo = f"Mayor de retenciones - {nombre_cuenta}.pdf"
    nombre_archivo = re.sub(r'[\\/*?:"<>|]', '', nombre_archivo)
    pdf_path = os.path.join(escritorio_path, nombre_archivo)
    pdf.output(pdf_path)
    return pdf_path

def procesar_archivo():
    """Procesa el archivo Excel y genera los PDFs"""
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
            
            # Obtener registros para esta cuenta
            registros_cuenta = df.iloc[idx_cuenta+1:siguiente_idx].copy()
            
            # Convertir Haber a Decimal para precisión (CORRECCIÓN: paréntesis faltante añadido)
            registros_cuenta['Haber'] = registros_cuenta['Haber'].apply(
                lambda x: convertir_a_decimal(x)
            )
            
            # Filtrar registros válidos
            filtro = (
                registros_cuenta['Detalle'].str.contains('Benef', na=False, case=False) &
                ~registros_cuenta['Detalle'].str.contains('CUT', na=False, case=False) &
                (registros_cuenta['Haber'] > Decimal('0'))
            )
            
            registros_filtrados = registros_cuenta[filtro].copy()
            
            if len(registros_filtrados) > 0:
                # Calcular suma con precisión decimal
                suma_haber = registros_filtrados['Haber'].sum()
                
                resultados.append({
                    'Cuenta': cuenta,
                    'NombreCuenta': nombre_cuenta,
                    'Cantidad de Movimientos': len(registros_filtrados),
                    'Suma Haber': float(suma_haber),
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
    # Verificación del formato numérico
    print("=== Prueba de formato numérico ===")
    test_values = [
        '14.041,36',    # Formato europeo
        '1.234,56',     # Formato europeo
        '1234,56',      # Formato europeo sin separador de miles
        '1234.56',      # Formato internacional sin separador de miles
        '1,234.56',     # Formato internacional
        '1234',         # Entero
        1234.56,        # Float
        'abc',          # Inválido
        None,           # Nulo
        '$ 1.234,56',   # Con símbolo de moneda
        '1.234',        # ¿Mil o uno punto dos tres cuatro?
        '1.234.567,89', # Formato europeo con múltiples separadores
    ]
    
    for val in test_values:
        print(f"Original: {val!r:<15} → Decimal: {convertir_a_decimal(val)} → Formateado: {format_currency(val)}")
    
    print("\n=== Iniciando procesamiento ===")
    procesar_archivo()