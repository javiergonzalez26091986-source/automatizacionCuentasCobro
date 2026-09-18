import requests
import io
import zipfile
import os
import re
import math
import pandas as pd
from datetime import datetime, timezone, timedelta
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
import plotly.express as px
import streamlit as st

# ==============================================================================
# 1. CONFIGURACIÓN DE LA PÁGINA Y PANEL DE ESTILOS CSS
# ==============================================================================
icono_pestana = "sergemLogo.ico" if os.path.exists("sergemLogo.ico") else "sergemLogo.png"
st.set_page_config(page_title="SERGEM - Generador automático de documentos", page_icon=icono_pestana, layout="wide")

st.markdown("""
    <style>
    [data-testid="stHeader"] { display: none !important; }
    [data-testid="stToolbar"] { display: none !important; }
    .stAppDeployButton { display: none !important; }
    #MainMenu { display: none !important; }
    footer { display: none !important; }
    div[class*="viewerBadge"], [data-testid="stAppCreatorBadge"] { display: none !important; }
    
    .stApp { background-color: #F4F6F9; }
    .block-container {
        background-color: #FFFFFF;
        padding: 2rem 3rem;
        border-radius: 12px;
        box-shadow: 0px 4px 15px rgba(0, 0, 0, 0.05);
        margin-top: 2rem;
        margin-bottom: 2rem;
    }
    h1, h2, h3, p, span, label, div { color: #1E293B; }
    
    .stSelectbox > div > div > div {
        background-color: #f8fafc !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        color: #1e293b !important;
    }
    
    div[data-testid="stTextInput"] input {
        border: 2px solid #E3000F !important;
        background-color: #fff1f2 !important;
        color: #E3000F !important;
        font-size: 1.15rem !important;
        font-weight: 800 !important;
        padding: 0.75rem !important;
        border-radius: 8px !important;
    }
    
    div.stButton > button {
        border-radius: 8px !important;
        font-weight: 800 !important;
        border: none !important;
        transition: transform 0.1s ease !important;
    }
    div.stButton > button:active { transform: scale(0.98) !important; }
    
    div.stButton > button[kind="primary"] { 
        background-color: #E3000F !important; 
        color: white !important; 
        padding: 0.75rem 1.5rem;
        font-size: 1.1rem;
    }
    div.stButton > button[kind="primary"]:hover { background-color: #B3000C !important; box-shadow: 0px 4px 10px rgba(227, 0, 15, 0.3); }
    
    div.stButton > button[kind="secondary"] { 
        background-color: #94a3b8 !important; 
        color: white !important; 
    }
    div.stButton > button[kind="secondary"]:hover { background-color: #64748b !important; }
    
    .metric-box {
        background-color: #f1f5f9; padding: 15px; border-radius: 8px; text-align: center; border: 1px solid #e2e8f0;
    }
    .caja-destacada {
        background-color: #e2e8f0; 
        padding: 20px; 
        border-radius: 10px; 
        border: 2px solid #94a3b8;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

GAS_URL = "https://script.google.com/macros/s/AKfycbzUT0gAjxyq4bCDGo06siBJya4C9OHDRwAEL_igBG_gJK19DIJCIkgmCYH819BTym1u/exec"

CODIGOS_BANCOS = {
    "BANCO DE BOGOTA": "1", "BANCO POPULAR": "2", "BANCOLOMBIA": "7",
    "DAVIVIENDA": "51", "BANCO DE OCCIDENTE": "23", "BANCO CAJA SOCIAL": "32",
    "BANCO AGRARIO": "40", "BANCO AV VILLAS": "52", "NEQUI": "1507",
    "DAVIPLATA": "1551", "BANCO W": "53", "SCOTIABANK COLPATRIA": "14",
    "TUYA S.A": "26", "BANCO FALABELLA S.A.": "62", "LULO BANK S.A.": "70"
}

def obtener_fecha_actual():
    meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    zona_colombia = timezone(timedelta(hours=-5))
    hoy = datetime.now(zona_colombia)
    return f"{hoy.day} DE {meses[hoy.month - 1]} DE {hoy.year}"

@st.cache_data(ttl=600) 
def cargar_datos(url):
    try:
        req_url = f"{url}?t={int(datetime.now().timestamp())}"
        response = requests.get(req_url, allow_redirects=True, timeout=15)
        if response.status_code != 200:
            st.error(f"⚠️ Error de respuesta de Google Apps Script: Código {response.status_code}")
            return None
        return response.json()
    except Exception as e:
        st.error(f"⚠️ Error en la conexión HTTP: {e}")
        return None

def subir_bulk_a_sheets(url, sheet_name, bulk_data, clear_first=False, replace_period=None):
    payload = {
        "sheet_name": sheet_name,
        "clear_first": clear_first,
        "replace_period": replace_period,
        "bulk_data": bulk_data
    }
    try:
        response = requests.post(url, json=payload, allow_redirects=True, timeout=90)
        if response.status_code != 200:
            return {"status": "error", "message": f"HTTP {response.status_code}: {response.text[:100]}"}
        try:
            return response.json()
        except ValueError:
            return {"status": "error", "message": f"Respuesta no válida del servidor: {response.text[:100]}"}
    except requests.exceptions.Timeout:
        return {"status": "error", "message": "El servidor de Google tardó demasiado en responder. Intenta sincronizar la base de datos para ver si los datos llegaron."}
    except Exception as e:
        return {"status": "error", "message": str(e) or "Error de conexión desconocido."}

def extraer_df_desde_excel(archivo, sheet_buscada=None):
    xls = pd.ExcelFile(archivo)
    hoja_objetivo = None
    posibles = [sheet_buscada, 'REPORTE', 'COBRO', 'PAGO', 'BASE', 'DATOS'] if sheet_buscada else ['REPORTE', 'COBRO', 'PAGO', 'BASE', 'DATOS']
    
    for nombre in posibles:
        if nombre and nombre in xls.sheet_names:
            hoja_objetivo = nombre
            break
            
    if not hoja_objetivo:
        for sh in xls.sheet_names:
            df_test = pd.read_excel(archivo, sheet_name=sh, nrows=15, header=None)
            if df_test.astype(str).apply(lambda col: col.str.contains('CEDULA|CÉDULA|IDENTIFICACION|IDENTIFICACIÓN', case=False, na=False)).any().any():
                hoja_objetivo = sh
                break
                
    if not hoja_objetivo: hoja_objetivo = xls.sheet_names[0]

    df_temp = pd.read_excel(archivo, sheet_name=hoja_objetivo, nrows=15, header=None)
    fila_header = 0
    for idx, fila in df_temp.iterrows():
        textos = fila.astype(str).str.upper().tolist()
        if any(col in textos for col in ['CÉDULA', 'CEDULA', 'CC', 'IDENTIFICACION', 'IDENTIFICACIÓN', 'NOMBRE COMPLETO', 'EMPLEADO']):
            fila_header = idx
            break

    df = pd.read_excel(archivo, sheet_name=hoja_objetivo, skiprows=fila_header)
    df.columns = df.columns.astype(str).str.strip().str.upper()
    
    col_id = None
    for col in df.columns:
        if any(alias in col for alias in ['CÉDULA', 'CEDULA', 'CC', 'IDENTIFICACION', 'IDENTIFICACIÓN']):
            col_id = col
            break
            
    if col_id:
        df = df.dropna(subset=[col_id])
        df = df[~df[col_id].astype(str).str.upper().str.contains('TOTAL')]
        df = df[df[col_id].astype(str).str.strip() != '']
        
    return df

def preparar_df_para_sheets(df_raw, cols_esperadas, periodo=""):
    df_out = pd.DataFrame(columns=cols_esperadas)
    
    for col in cols_esperadas:
        if col == 'PERIODO' and periodo:
            df_out[col] = periodo
        else:
            if col == 'TOTAL FACTURAR':
                alias = ['TOTAL', 'TOTAL FACTURAR', 'VALOR TOTAL', 'NETO', 'VALOR_TOTAL']
            elif col == 'CEDULA' or col == 'IDENTIFICACIÓN': 
                alias = [col, 'CÉDULA', 'C.C.', 'IDENTIFICACION', 'DOCUMENTO', 'CC']
            elif col == 'NOMBRE' or col == 'NOMBRE COMPLETO': 
                alias = [col, 'NOMBRES', 'CONDUCTOR', 'EMPLEADO', 'BENEFICIARIO']
            elif col == 'VALOR PAGO':
                alias = [col, 'PAGO', 'VALOR']
            elif col == 'TIPO DE VEHICULO': 
                alias = [col, 'VEHICULO', 'CATEGORIA']
            elif col == 'PUNTO DE VENTA': 
                alias = [col, 'ALMACEN', 'CLIENTE', 'PUNTO_VENTA']
            else:
                alias = [col]
            
            col_found = obtener_nombre_columna(df_raw, alias)
            if col_found:
                df_out[col] = df_raw[col_found]
            else:
                df_out[col] = ""
                
    for col in df_out.columns:
        if pd.api.types.is_datetime64_any_dtype(df_out[col]):
            df_out[col] = df_out[col].dt.strftime('%Y-%m-%d')
        elif col in ['CEDULA', 'IDENTIFICACIÓN', 'CÓDIGO INGRESO']:
            df_out[col] = df_out[col].astype(str).replace(r'\.0$', '', regex=True).replace(['nan', 'NaT', 'None'], '').str.strip()
        elif col in ['TOTAL', 'TOTAL FACTURAR', 'TARIFA', 'NETO', 'VALOR TOTAL', 'VALOR PAGO']:
            df_out[col] = pd.to_numeric(df_out[col], errors='coerce').fillna(0)
            
    raw_list = df_out.values.tolist()
    cleaned_list = []
    for row in raw_list:
        cleaned_row = []
        for val in row:
            if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
                cleaned_row.append("")
            elif pd.isna(val):
                cleaned_row.append("")
            else:
                cleaned_row.append(val)
        cleaned_list.append(cleaned_row)
            
    return cleaned_list

def limpiar_dinero(val):
    if pd.isna(val) or val == "": return 0.0
    s = str(val).upper().replace('$', '').replace(',', '').replace('.', '').replace(' ', '')
    try: return float(s)
    except: return 0.0

def limpiar_texto(txt):
    if pd.isna(txt): return ""
    return re.sub(r'\s+', ' ', str(txt).upper().strip())

def obtener_horas(row):
    for col in ['HORAS', 'CANTIDAD DE HORAS', 'CANTIDAD HORAS', 'TOTAL HORAS', 'CANTIDAD', 'NUMERO DE HORAS']:
        if col in row.index:
            try:
                val = float(row[col])
                if not pd.isna(val):
                    if val >= 40000:
                        fecha_erronea = datetime(1899, 12, 30) + timedelta(days=int(val))
                        return float(f"{fecha_erronea.day}.{fecha_erronea.month}")
                    return val
            except: pass
    return 1.0

def get_pdf_bytes(pdf_obj):
    out = pdf_obj.output()
    return out.encode('latin-1') if isinstance(out, str) else bytes(out)

# ==============================================================================
# GENERACIÓN DE ARCHIVO EXCEL PAB Y PDFS
# ==============================================================================
def generar_excel_pab(df_banco, corte_seleccionado):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "FORMATOPAB"
    
    fill_headers = PatternFill(start_color="31869B", end_color="31869B", fill_type="solid")
    font_headers = Font(color="FFFFFF", bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    
    anchos = {'A': 26, 'B': 22, 'C': 25, 'D': 21, 'E': 28, 'F': 28, 'G': 28, 'H': 28, 'I': 20, 'J': 23, 'K': 25, 'L': 19}
    for col, width in anchos.items():
        ws.column_dimensions[col].width = width

    def map_tipo_doc(t):
        t = str(t).upper()
        if 'NIT' in t: return 3
        elif 'CE' in t or 'EXTRANJER' in t: return 2
        elif 'TI' in t or 'IDENTIDAD' in t: return 4
        elif 'PP' in t or 'PASAPORTE' in t: return 5
        else: return 1
        
    fecha_app = datetime.now(timezone(timedelta(hours=-5))).strftime("%Y%m%d")
    
    headers_fila1 = ['NIT PAGADOR', 'TIPO DE PAGO', 'APLICACIÓN', 'SECUENCIA DE ENVÍO', 'NRO CUENTA A DEBITAR', 'TIPO DE CUENTA A DEBITAR', 'DESCRIPCIÓN DEL PAGO']
    for col_idx, header in enumerate(headers_fila1, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.fill = fill_headers
        cell.font = font_headers
        cell.alignment = align_center

    desc_pago = f"SERVICOS {str(corte_seleccionado).upper()}"[:40]
    valores_fila2 = [900561833, 225, 'I', 'A1', 81016173001, 'D', desc_pago]
    for col_idx, val in enumerate(valores_fila2, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = val
        cell.alignment = align_center

    headers_fila3 = ['Tipo Documento Beneficiario', 'Nit Beneficiario', 'Nombre Beneficiario ', 'Tipo Transaccion ', 'Código Banco ', 'No Cuenta Beneficiario ', 'Email ', 'Documento Autorizado ', 'Referencia ', 'Celular Beneficiario', 'ValorTransaccion ', 'Fecha de aplicación']
    for col_idx, header in enumerate(headers_fila3, start=1):
        cell = ws.cell(row=3, column=col_idx)
        cell.value = header
        cell.fill = fill_headers
        cell.font = font_headers
        cell.alignment = align_center

    fila_actual = 4
    df_banco['CODIGO_BANCO'] = df_banco["BANCO_DESTINO"].map(CODIGOS_BANCOS).fillna("")
    
    for _, row in df_banco.iterrows():
        ws.cell(row=fila_actual, column=1).value = map_tipo_doc(row['TIPO_IDENTIFICACION'])
        ws.cell(row=fila_actual, column=1).alignment = align_center
        
        nit_val = str(row['NIT_BENEFICIARIO']).replace('.0', '').replace(r'\D', '')
        try: nit_val = int(nit_val)
        except: pass
        ws.cell(row=fila_actual, column=2).value = nit_val
        ws.cell(row=fila_actual, column=3).value = str(row['NOMBRE_BENEFICIARIO']).strip()
        ws.cell(row=fila_actual, column=4).value = 37
        ws.cell(row=fila_actual, column=4).alignment = align_center
        
        cod_banco = row['CODIGO_BANCO']
        try: cod_banco = int(cod_banco)
        except: pass
        ws.cell(row=fila_actual, column=5).value = cod_banco
        ws.cell(row=fila_actual, column=5).alignment = align_center
        
        cuenta_val = str(row['NUMERO_CUENTA']).replace("'", "").replace("-", "").replace(" ", "").strip()
        try: cuenta_val = int(cuenta_val)
        except: pass
        ws.cell(row=fila_actual, column=6).value = cuenta_val
        ws.cell(row=fila_actual, column=11).value = float(row['VALOR_NETO_A_PAGAR'])
        ws.cell(row=fila_actual, column=11).number_format = '0'
        ws.cell(row=fila_actual, column=12).value = int(fecha_app)
        fila_actual += 1
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def agregar_pagina_pdf_cuenta_cobro(pdf, datos):
    pdf.add_page()
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"{datos['ciudad']} {datos['fecha_emision']}".upper(), 0, 1, 'R')
    pdf.ln(12)
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 6, "SERGEM MENSAJERIA S.A.S.", 0, 1, 'C')
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "NIT: 900.561.833-1", 0, 1, 'C')
    pdf.ln(8)
    pdf.cell(0, 6, "DEBE A:", 0, 1, 'C')
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 6, str(datos['nombre_prestador']).upper(), 0, 1, 'C')
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"{datos['tipo_documento_pab']} {datos['cedula_prestador']}", 0, 1, 'C')
    pdf.ln(8)

    pdf.set_font("helvetica", "B", 11)
    pdf.cell(80, 6, "VALOR BASE:", 0, 0)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"$ {datos['ingreso_base']:,.0f}", 0, 1)

    if datos.get('valor_dia_negociado', 0) > 0:
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(80, 6, "VALOR DÍA NEGOCIADO:", 0, 0)
        pdf.set_font("helvetica", "", 11)
        pdf.cell(0, 6, f"$ {datos['valor_dia_negociado']:,.0f}", 0, 1)

    if datos.get('fuera_perimetro', 0) > 0:
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(80, 6, "FUERA DE PERÍMETRO:", 0, 0)
        pdf.set_font("helvetica", "", 11)
        pdf.cell(0, 6, f"$ {datos['fuera_perimetro']:,.0f}", 0, 1)

    if datos.get('retefuente', 0) > 0:
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(80, 6, "MENOS RETEFUENTE (1%):", 0, 0)
        pdf.set_font("helvetica", "", 11)
        pdf.set_text_color(227, 0, 15)
        pdf.cell(0, 6, f"$ -{datos['retefuente']:,.0f}", 0, 1)
        pdf.set_text_color(0, 0, 0)

    if datos.get('ica', 0) > 0:
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(80, 6, "MENOS RETEICA (1%):", 0, 0)
        pdf.set_font("helvetica", "", 11)
        pdf.set_text_color(227, 0, 15)
        pdf.cell(0, 6, f"$ -{datos['ica']:,.0f}", 0, 1)
        pdf.set_text_color(0, 0, 0)

    pdf.ln(2)
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(80, 8, "NETO SERVICIOS:", 0, 0)
    pdf.cell(0, 8, f"$ {datos['neto_pagar']:,.0f}", 0, 1)
    pdf.ln(6)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(25, 6, "CONCEPTO:", 0, 1)
    pdf.set_font("helvetica", "", 10)
    pdf.multi_cell(0, 5, f"SERVICIO DE MENSAJERÍA PRESTADO EN EL CORTE DE {datos['corte_fechas']}, DETALLADO A CONTINUACIÓN:")
    pdf.ln(2)
    
    pdf.set_fill_color(227, 0, 15)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(85, 6, "CONDUCTOR / DETALLE", 1, 0, 'C', fill=True)
    pdf.cell(35, 6, "CÉDULA", 1, 0, 'C', fill=True)
    pdf.cell(25, 6, "CANTIDAD", 1, 0, 'C', fill=True)
    pdf.cell(45, 6, "VALOR NETO", 1, 1, 'C', fill=True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("helvetica", "", 9)
    for c in datos['conductores']:
        pdf.cell(85, 6, c['nombre_conductor'][:40], 1, 0, 'L')
        pdf.cell(35, 6, c['cedula_conductor'], 1, 0, 'C')
        pdf.cell(25, 6, f"{c['horas']:g} Horas", 1, 0, 'C')
        pdf.cell(45, 6, f"$ {c['neto_pagar']:,.0f}", 1, 1, 'R')
        
    for fpu in datos.get('fpu_items', []):
        pdf.cell(85, 6, f"F. Perímetro: {fpu['destino'][:25]}", 1, 0, 'L')
        pdf.cell(35, 6, "N/A", 1, 0, 'C')
        pdf.cell(25, 6, f"{fpu['cantidad']:g} Viaje(s)", 1, 0, 'C')
        pdf.cell(45, 6, f"$ {fpu['valor_unitario']:,.0f}", 1, 0, 'R')
        pdf.cell(0, 6, f"$ {fpu['total']:,.0f}", 1, 1, 'R')
        item_idx += 1

    pdf.ln(8)
    
    if datos.get('anticipos', 0) > 0 or datos.get('otros_descuentos', 0) > 0:
        if datos.get('anticipos', 0) > 0:
            pdf.set_font("helvetica", "B", 11)
            pdf.cell(80, 6, "MENOS ANTICIPOS:", 0, 0)
            pdf.set_font("helvetica", "", 11)
            pdf.set_text_color(227, 0, 15)
            pdf.cell(0, 6, f"$ -{datos['anticipos']:,.0f}", 0, 1)
            pdf.set_text_color(0, 0, 0)
            
        if datos.get('otros_descuentos', 0) > 0:
            pdf.set_font("helvetica", "B", 11)
            pdf.cell(80, 6, "MENOS OTROS DESCUENTOS (ARL):", 0, 0)
            pdf.set_font("helvetica", "", 11)
            pdf.set_text_color(227, 0, 15)
            pdf.cell(0, 6, f"$ -{datos['otros_descuentos']:,.0f}", 0, 1)
            pdf.set_text_color(0, 0, 0)

        pdf.ln(2)
        pdf.set_font("helvetica", "B", 12)
        pdf.set_text_color(227, 0, 15)
        pdf.cell(80, 8, "TOTAL A CONSIGNAR:", 0, 0)
        pdf.cell(0, 8, f"$ {datos['neto_final']:,.0f}", 0, 1)
        pdf.set_text_color(0, 0, 0)
        pdf.ln(6)
    
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "Autorizo me sea consignado en:", 0, 1)
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(0, 6, f"CUENTA # {datos['num_cuenta']} - {datos['tipo_cuenta'].upper()}", 0, 1)
    pdf.cell(0, 6, f"BANCO: {datos['banco'].upper()}", 0, 1)
    pdf.cell(0, 6, f"TITULAR: {datos['nombre_titular_banco']} (C.C/NIT: {datos['cedula_titular_banco']})", 0, 1)
    pdf.ln(15)
    
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "Atentamente,", 0, 1)
    pdf.ln(10)
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(80, 5, str(datos['nombre_prestador']).upper(), "T", 1, "L")
    pdf.set_font("helvetica", "", 11)
    pdf.cell(80, 5, f"{datos['tipo_documento_pab']} {datos['cedula_prestador']}", 0, 1, "L")

def agregar_pagina_pdf_doc_equivalente(pdf, datos):
    pdf.add_page()
    try:
        if os.path.exists('sergemLogo.png'):
            pdf.image('sergemLogo.png', 10, 8, w=45)
    except: pass

    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 5, "DOCUMENTO EQUIVALENTE A LA FACTURA DE VENTA", 0, 1, 'R')
    pdf.set_font('helvetica', '', 9)
    pdf.cell(0, 4, "(DECRETO 522 DE 2003)", 0, 1, 'R')
    pdf.cell(0, 4, "DOCUMENTO SOPORTE EN ADQUISICIONES A NO OBLIGADOS A FACTURAR", 0, 1, 'R')
    pdf.ln(2)
    
    pdf.set_font('helvetica', 'B', 11)
    pdf.set_text_color(227, 0, 15)
    pdf.cell(0, 6, f"CONSECUTIVO NO: {datos['id']}", 0, 1, 'R')
    pdf.set_text_color(0, 0, 0)
    pdf.ln(5)

    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(35, 5, "Fecha de Expedición:")
    pdf.set_font('helvetica', '', 9)
    pdf.cell(0, 5, datos['fecha_emision'], 0, 1)
    pdf.ln(3)

    pdf.set_fill_color(51, 51, 51)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(0, 6, " INFORMACIÓN DE LA EMPRESA (COMPRADOR)", 1, 1, 'L', fill=True)
    pdf.set_text_color(0, 0, 0)
    
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(25, 6, "Razón Social:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(100, 6, "SERGEM MENSAJERIA S.A.S.", 1)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(15, 6, "NIT:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(0, 6, "900.561.833-1", 1, 1)

    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(25, 6, "Dirección:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(60, 6, "CRA 62 9 235", 1)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(20, 6, "Teléfono:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(45, 6, "3994620", 1)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(15, 6, "Ciudad:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(0, 6, "CALI", 1, 1)
    pdf.ln(4)

    pdf.set_fill_color(51, 51, 51)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(0, 6, " DATOS DEL BENEFICIARIO / PROVEEDOR (VENDEDOR)", 1, 1, 'L', fill=True)
    pdf.set_text_color(0, 0, 0)

    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(25, 6, "Nombre:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(100, 6, datos['nombre_prestador'][:45], 1)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(20, 6, f"{datos['tipo_documento_pab']}:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(0, 6, datos['cedula_prestador'], 1, 1)
    
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(25, 6, "Ciudad:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(100, 6, datos['ciudad'], 1)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(20, 6, "Conductores:", 1)
    pdf.set_font('helvetica', '', 9)
    
    nombres_conds = ", ".join([c['nombre_conductor'] for c in datos['conductores']])
    if len(nombres_conds) > 25: nombres_conds = nombres_conds[:22] + "..."
    pdf.cell(0, 6, nombres_conds, 1, 1)
    pdf.ln(6)

    pdf.set_fill_color(227, 0, 15)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(10, 6, "Ítem", 1, 0, 'C', fill=True)
    pdf.cell(90, 6, "Concepto", 1, 0, 'C', fill=True)
    pdf.cell(20, 6, "Cantidad", 1, 0, 'C', fill=True)
    pdf.cell(35, 6, "V. Unitario", 1, 0, 'C', fill=True)
    pdf.cell(0, 6, "V. Total", 1, 1, 'C', fill=True)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('helvetica', '', 9)

    item_idx = 1
    for c in datos['conductores']:
        v_unitario = c['ingreso_base'] / c['horas'] if c['horas'] > 0 else c['ingreso_base']
        pdf.cell(10, 6, str(item_idx), 1, 0, 'C')
        pdf.cell(90, 6, f"Servicio mensajería - {c['nombre_conductor'][:25]}", 1, 0, 'L')
        pdf.cell(20, 6, f"{c['horas']:g}", 1, 0, 'C')
        pdf.cell(35, 6, f"$ {v_unitario:,.0f}", 1, 0, 'R')
        pdf.cell(0, 6, f"$ {c['ingreso_base']:,.0f}", 1, 1, 'R')
        item_idx += 1

    for fpu in datos.get('fpu_items', []):
        pdf.cell(10, 6, str(item_idx), 1, 0, 'C')
        pdf.cell(90, 6, f"Fuera Perímetro: {fpu['destino'][:20]}", 1, 0, 'L')
        pdf.cell(20, 6, f"{fpu['cantidad']:g}", 1, 0, 'C')
        pdf.cell(35, 6, f"$ {fpu['valor_unitario']:,.0f}", 1, 0, 'R')
        pdf.cell(0, 6, f"$ {fpu['total']:,.0f}", 1, 1, 'R')
        item_idx += 1

    pdf.ln(2)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(120, 6, "", 0, 0)
    pdf.cell(35, 6, "SUBTOTAL:", 1, 0, 'R')
    pdf.cell(0, 6, f"$ {datos['ingreso_bruto_total']:,.0f}", 1, 1, 'R')

    pdf.cell(120, 6, "", 0, 0)
    pdf.cell(35, 6, "IVA (19%):", 1, 0, 'R')
    pdf.cell(0, 6, "$ 0", 1, 1, 'R')

    pdf.cell(120, 6, "", 0, 0)
    pdf.cell(35, 6, "RETEIVA:", 1, 0, 'R')
    pdf.cell(0, 6, "$ 0", 1, 1, 'R')

    pdf.cell(120, 6, "", 0, 0)
    pdf.cell(35, 6, "RTE FTE (1%):", 1, 0, 'R')
    val_rte = -datos['retefuente'] if datos['retefuente'] > 0 else 0
    pdf.cell(0, 6, f"$ {val_rte:,.0f}", 1, 1, 'R')

    pdf.cell(120, 6, "", 0, 0)
    pdf.cell(35, 6, "RETEICA (1%):", 1, 0, 'R')
    val_ica = -datos['ica'] if datos['ica'] > 0 else 0
    pdf.cell(0, 6, f"$ {val_ica:,.0f}", 1, 1, 'R')

    pdf.set_fill_color(244, 246, 249)
    pdf.cell(120, 6, "", 0, 0)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(35, 6, "NETO SERVICIOS:", 1, 0, 'R', fill=True)
    pdf.cell(0, 6, f"$ {datos['neto_pagar']:,.0f}", 1, 1, 'R', fill=True)

    if datos.get('anticipos', 0) > 0:
        pdf.cell(120, 6, "", 0, 0)
        pdf.cell(35, 6, "ANTICIPOS:", 1, 0, 'R')
        pdf.set_text_color(227, 0, 15)
        pdf.cell(0, 6, f"$ -{datos['anticipos']:,.0f}", 1, 1, 'R')
        pdf.set_text_color(0, 0, 0)
        
    if datos.get('otros_descuentos', 0) > 0:
        pdf.cell(120, 6, "", 0, 0)
        pdf.cell(35, 6, "OTROS DESC (ARL):", 1, 0, 'R')
        pdf.set_text_color(227, 0, 15)
        pdf.cell(0, 6, f"$ -{datos['otros_descuentos']:,.0f}", 1, 1, 'R')
        pdf.set_text_color(0, 0, 0)
        
    if datos.get('anticipos', 0) > 0 or datos.get('otros_descuentos', 0) > 0:
        pdf.cell(120, 8, "", 0, 0)
        pdf.set_text_color(227, 0, 15)
        pdf.cell(35, 8, "TOTAL A CONSIGNAR:", 1, 0, 'R', fill=True)
        pdf.set_font('helvetica', 'B', 11)
        pdf.cell(0, 8, f"$ {datos['neto_final']:,.0f}", 1, 1, 'R', fill=True)
        pdf.set_text_color(0, 0, 0)

    pdf.ln(12)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(80, 5, "________________________________________________", 0, 1)
    pdf.cell(80, 5, "FIRMA PRESTADOR DEL SERVICIO", 0, 1)
    pdf.cell(80, 5, f"{datos['tipo_documento_pab']}: {datos['cedula_prestador']}", 0, 1)
    pdf.cell(80, 5, f"NOMBRE: {datos['nombre_prestador']}", 0, 1)

def construir_hoja_documento_equivalente_excel(ws, datos):
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="E3000F", end_color="E3000F", fill_type="solid")
    dark_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
    bold_font = Font(bold=True)
    border_thin = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    try:
        if os.path.exists('sergemLogo.png'):
            img = XLImage('sergemLogo.png')
            img.width = 150
            img.height = 60
            ws.add_image(img, 'B2')
    except: pass

    ws.merge_cells('D2:H4')
    ws['D2'] = "DOCUMENTO EQUIVALENTE A LA FACTURA DE VENTA\n(DECRETO 522 DE 2003)\nDOCUMENTO SOPORTE EN ADQUISICIONES A NO OBLIGADOS A FACTURAR"
    ws['D2'].font = Font(bold=True, size=11, color="1E293B")
    ws['D2'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    hoy = datetime.now(timezone(timedelta(hours=-5)))
    ws['B6'] = "Fecha de Expedición:"; ws['B6'].font = bold_font
    ws['C6'] = "Año:"; ws['D6'] = hoy.year; ws['E6'] = "Mes:"; ws['F6'] = f"{hoy.month:02d}"; ws['G6'] = "Día:"; ws['H6'] = f"{hoy.day:02d}"
    
    ws['G5'] = "CONSECUTIVO NO:"; ws['G5'].font = Font(bold=True, size=11, color="E3000F"); ws['G5'].alignment = right_align
    ws['H5'] = datos['id']; ws['H5'].font = Font(bold=True, size=12, color="E3000F"); ws['H5'].alignment = center_align

    ws['B9'] = " INFORMACIÓN DE LA EMPRESA (COMPRADOR)"
    ws['B9'].font = header_font; ws['B9'].fill = dark_fill
    ws.merge_cells('B9:H9')
    ws['B10'] = "Razón Social:"; ws['B10'].font = bold_font; ws['C10'] = "SERGEM MENSAJERIA S.A.S."; ws.merge_cells('C10:E10')
    ws['G10'] = "NIT:"; ws['G10'].font = bold_font; ws['H10'] = "900.561.833-1"

    ws['B13'] = " DATOS DEL BENEFICIARIO / PROVEEDOR (VENDEDOR)"
    ws['B13'].font = header_font; ws['B13'].fill = dark_fill; ws.merge_cells('B13:H13')

    ws['B14'] = "Nombre:"; ws['B14'].font = bold_font; ws['C14'] = datos['nombre_prestador']; ws.merge_cells('C14:E14')
    ws['G14'] = f"{datos['tipo_documento_pab']}:"; ws['G14'].font = bold_font; ws['H14'] = datos['cedula_prestador']
    
    ws['G16'] = "Conductores:"; ws['G16'].font = bold_font
    nombres_conds = ", ".join([c['nombre_conductor'] for c in datos['conductores']])
    if len(nombres_conds) > 25: nombres_conds = nombres_conds[:22] + "..."
    ws['H16'] = nombres_conds

    fila = 18
    for i, h in enumerate(["Ítem", "Concepto", "Cantidad", "V. Unitario", "V. Total"]):
        c = ['B', 'C', 'F', 'G', 'H'][i] + str(fila)
        ws[c] = h; ws[c].font = header_font; ws[c].fill = header_fill; ws[c].alignment = center_align; ws[c].border = border_thin
    ws.merge_cells(f'C{fila}:E{fila}')
    
    fila += 1
    item_idx = 1
    for c in datos['conductores']:
        ws[f'B{fila}'] = item_idx
        ws[f'C{fila}'] = f"Servicio mensajería - {c['nombre_conductor']}"
        ws.merge_cells(f'C{fila}:E{fila}')
        ws[f'F{fila}'] = float(c['horas']) 
        v_unitario = c['ingreso_base'] / c['horas'] if c['horas'] > 0 else c['ingreso_base']
        ws[f'G{fila}'] = v_unitario
        ws[f'H{fila}'] = c['ingreso_base']
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H']: ws[f'{col}{fila}'].border = border_thin
        ws[f'G{fila}'].number_format = '"$"#,##0'; ws[f'H{fila}'].number_format = '"$"#,##0'
        ws[f'B{fila}'].alignment = center_align; ws[f'C{fila}'].alignment = left_align; ws[f'F{fila}'].alignment = center_align
        fila += 1
        item_idx += 1

    for fpu in datos.get('fpu_items', []):
        ws[f'B{fila}'] = item_idx
        ws[f'C{fila}'] = f"Fuera Perímetro: {fpu['destino']}"
        ws.merge_cells(f'C{fila}:E{fila}')
        ws[f'F{fila}'] = fpu['cantidad']
        ws[f'G{fila}'] = fpu['valor_unitario']
        ws[f'H{fila}'] = fpu['total']
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H']: ws[f'{col}{fila}'].border = border_thin
        ws[f'G{fila}'].number_format = '"$"#,##0'; ws[f'H{fila}'].number_format = '"$"#,##0'
        ws[f'B{fila}'].alignment = center_align; ws[f'C{fila}'].alignment = left_align; ws[f'F{fila}'].alignment = center_align
        fila += 1
        item_idx += 1
        
    fila += 1
    
    totales = [
        ("SUBTOTAL:", datos['ingreso_bruto_total']), 
        ("IVA (19%):", ""), ("RETEIVA:", ""), 
        ("RTE FTE (1%):", -datos['retefuente'] if datos['retefuente']>0 else 0),
        ("RETEICA (1%):", -datos['ica'] if datos['ica']>0 else 0), 
        ("NETO SERVICIOS:", datos['neto_pagar'])
    ]
    
    if datos.get('anticipos', 0) > 0: totales.append(("MENOS ANTICIPOS:", -datos['anticipos']))
    if datos.get('otros_descuentos', 0) > 0: totales.append(("OTROS DESC (ARL):", -datos['otros_descuentos']))
    
    if datos.get('anticipos', 0) > 0 or datos.get('otros_descuentos', 0) > 0:
        totales.append(("TOTAL A CONSIGNAR:", datos['neto_final']))
    
    fila_firma = fila + len(totales) + 2
    for label, valor in totales:
        ws[f'G{fila}'] = label; ws[f'H{fila}'] = valor
        ws[f'G{fila}'].font = bold_font; ws[f'G{fila}'].alignment = right_align
        ws[f'G{fila}'].border = border_thin; ws[f'H{fila}'].border = border_thin
        if valor != "": ws[f'H{fila}'].number_format = '"$"#,##0'
        
        es_resaltado = (label == "TOTAL A CONSIGNAR:") or (label == "NETO SERVICIOS:" and datos.get('anticipos', 0) == 0 and datos.get('otros_descuentos', 0) == 0)
        if es_resaltado:
            ws[f'G{fila}'].font = Font(bold=True, color="E3000F"); ws[f'H{fila}'].font = Font(bold=True, size=12)
            ws[f'H{fila}'].fill = PatternFill(start_color="F4F6F9", end_color="F4F6F9", fill_type="solid")
        fila += 1

    ws[f'B{fila_firma}'] = "________________________________________________"
    ws[f'B{fila_firma+1}'] = "FIRMA PRESTADOR DEL SERVICIO"
    ws[f'B{fila_firma+1}'].font = bold_font
    ws[f'B{fila_firma+2}'] = f"{datos['tipo_documento_pab']}: {datos['cedula_prestador']}"
    ws[f'B{fila_firma+3}'] = f"NOMBRE: {datos['nombre_prestador']}"

    ws.column_dimensions['B'].width = 16; ws.column_dimensions['C'].width = 12; ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12; ws.column_dimensions['F'].width = 10; ws.column_dimensions['G'].width = 22; ws.column_dimensions['H'].width = 22

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT; ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.5; ws.page_margins.right = 0.5; ws.page_margins.top = 0.5; ws.page_margins.bottom = 0.5

# ==============================================================================
# PROCESO MATEMÁTICO PRINCIPAL
# ==============================================================================
def obtener_nombre_columna(df, opciones):
    for op in opciones:
        for col in df.columns:
            if limpiar_texto(col) == limpiar_texto(op):
                return col
    return None

def calcular_valores_agrupados(grupo_df, df_fuera, corte_seleccionado, col_prestador, col_ced_prestador, col_tit_banco, col_ced_banco, col_estado, col_anticipos, col_otros_desc, col_valor_dia):
    conductores = []
    fpu_items_doc = []
    conductores_procesados_fpu = set()
    
    es_nuevo = False
    suma_neto = 0
    suma_bruto = 0
    suma_fuera_bruto = 0
    suma_retefuente = 0
    suma_ica = 0
    suma_horas = 0
    
    suma_anticipos = 0
    suma_otros_desc = 0

    row_titular = grupo_df.iloc[0]
    
    tipo_doc_raw = str(row_titular.get('TIPO DE DOCUMENTO', '')).upper().strip()
    if not tipo_doc_raw or tipo_doc_raw == "NAN":
        tipo_doc_raw = str(row_titular.get('TIPO DOCUMENTO', 'CC')).upper().strip()
        
    if 'NIT' in tipo_doc_raw: tipo_doc_pab = 'NIT'
    elif 'CE' in tipo_doc_raw or 'EXTRANJ' in tipo_doc_pab: tipo_doc_pab = 'CE'
    elif 'PASAPORTE' in tipo_doc_raw or 'PP' in tipo_doc_pab: tipo_doc_pab = 'PP'
    else: tipo_doc_pab = 'CC'
    
    nombre_prestador = str(row_titular.get(col_prestador, 'S/N')).strip()
    cedula_prestador = str(row_titular.get(col_ced_prestador, '')).strip()
    
    nombre_titular_banco = str(row_titular.get(col_tit_banco, nombre_prestador)).strip()
    cedula_titular_banco = str(row_titular.get(col_ced_banco, cedula_prestador)).strip()
    
    banco = str(row_titular.get('BANCO', '')).strip()
    tipo_cuenta = str(row_titular.get('TIPO CUENTA', '')).strip()
    num_cuenta = str(row_titular.get('NO. CUENTA', '')).strip()
    ciudad_titular = str(row_titular.get('CIUDAD', '')).upper().strip()
    
    valor_dia_negociado = limpiar_dinero(row_titular.get(col_valor_dia, 0)) if col_valor_dia else 0

    for _, row in grupo_df.iterrows():
        if col_estado and str(row.get(col_estado, '')).strip().upper() == 'NUEVO':
            es_nuevo = True
            
        if col_anticipos: suma_anticipos += limpiar_dinero(row.get(col_anticipos, 0))
        if col_otros_desc: suma_otros_desc += limpiar_dinero(row.get(col_otros_desc, 0))

        ingreso_neto_esperado = limpiar_dinero(row.get('TOTAL A PAGAR', 0))
        nombre_conductor = str(row.get('CONDUCTOR', '')).strip().upper()

        ciudad = str(row.get('CIUDAD', '')).upper().strip()
        cedula_conductor = str(row.get('CÉDULA', row.get('CEDULA', ''))).strip()
        horas = obtener_horas(row)

        porcentaje_retefuente = 0.01
        porcentaje_ica = 0.01 if ciudad == 'CALI' else 0.0
        tasa_total_impuestos = porcentaje_retefuente + porcentaje_ica

        fuera_perimetro_neto = 0.0
        
        if not df_fuera.empty and nombre_conductor != "" and nombre_conductor not in conductores_procesados_fpu:
            col_cond_fuera = obtener_nombre_columna(df_fuera, ['CONDUCTOR', 'NOMBRE', 'PRESTADOR', 'NOMBRES'])
            
            if col_cond_fuera:
                df_fuera_cond = df_fuera[df_fuera[col_cond_fuera].astype(str).str.upper().str.contains(nombre_conductor, na=False, regex=False)]
                
                if not df_fuera_cond.empty:
                    conductores_procesados_fpu.add(nombre_conductor)
                    
                    col_dest = obtener_nombre_columna(df_fuera_cond, ['FUERA PERIMETRO CEDI', 'DESTINO', 'CIUDAD', 'LUGAR'])
                    col_val = obtener_nombre_columna(df_fuera_cond, ['VALOR', 'PRECIO'])
                    col_cant = obtener_nombre_columna(df_fuera_cond, ['CANTIDAD', 'CANT'])
                    
                    if col_dest and col_val and col_cant:
                        for _, f_row in df_fuera_cond.iterrows():
                            try:
                                cant = float(f_row.get(col_cant, 0))
                                if cant > 0:
                                    destino = str(f_row.get(col_dest, ''))
                                    valor_uni_neto = limpiar_dinero(f_row.get(col_val, 0))
                                    tot_neto = cant * valor_uni_neto
                                    
                                    tot_bruto = round(tot_neto / (1 - tasa_total_impuestos))
                                    val_uni_bruto = round(valor_uni_neto / (1 - tasa_total_impuestos))
                                    
                                    fuera_perimetro_neto += tot_neto
                                    suma_fuera_bruto += tot_bruto
                                    
                                    fpu_items_doc.append({
                                        'destino': destino,
                                        'cantidad': cant,
                                        'valor_unitario': val_uni_bruto,
                                        'total': tot_bruto,
                                        'neto': tot_neto
                                    })
                            except: pass

        total_neto_esperado = ingreso_neto_esperado + fuera_perimetro_neto
        ingreso_bruto_total = round(total_neto_esperado / (1 - tasa_total_impuestos)) if total_neto_esperado > 0 else 0
        retefuente = round(ingreso_bruto_total * porcentaje_retefuente)
        ica = round(ingreso_bruto_total * porcentaje_ica)

        fpu_bruto_cond = round(fuera_perimetro_neto / (1 - tasa_total_impuestos)) if fuera_perimetro_neto > 0 else 0.0
        ingreso_base_bruta = ingreso_bruto_total - fpu_bruto_cond
        
        neto_total_conductor = ingreso_bruto_total - retefuente - ica
        neto_solo_horas = neto_total_conductor - fuera_perimetro_neto

        if ingreso_neto_esperado > 0 or nombre_conductor != "":
            conductores.append({
                'nombre_conductor': nombre_conductor if nombre_conductor else nombre_prestador,
                'cedula_conductor': cedula_conductor if cedula_conductor else cedula_prestador,
                'horas': horas,
                'ingreso_base': ingreso_base_bruta,
                'neto_pagar': neto_solo_horas
            })

        suma_neto += neto_total_conductor
        suma_bruto += ingreso_bruto_total
        suma_retefuente += retefuente
        suma_ica += ica
        suma_horas += horas

    if not conductores: 
        conductores.append({
            'nombre_conductor': nombre_prestador,
            'cedula_conductor': cedula_prestador,
            'horas': suma_horas if suma_horas > 0 else 1.0,
            'ingreso_base': 0,
            'neto_pagar': 0
        })

    neto_final = suma_neto - suma_anticipos - suma_otros_desc

    return {
        'es_nuevo': es_nuevo,
        'anticipos': suma_anticipos,
        'otros_descuentos': suma_otros_desc,
        'neto_final': neto_final,
        'nombre_prestador': nombre_prestador,
        'cedula_prestador': cedula_prestador,
        'nombre_titular_banco': nombre_titular_banco,
        'cedula_titular_banco': cedula_titular_banco,
        'tipo_documento_pab': tipo_doc_pab,
        'banco': banco,
        'tipo_cuenta': tipo_cuenta,
        'num_cuenta': num_cuenta,
        'ciudad': ciudad_titular,
        'valor_dia_negociado': valor_dia_negociado,
        'conductores': conductores,
        'fpu_items': fpu_items_doc,
        'ingreso_base': suma_bruto - suma_fuera_bruto,
        'fuera_perimetro': suma_fuera_bruto,
        'ingreso_bruto_total': suma_bruto,
        'retefuente': suma_retefuente,
        'ica': suma_ica,
        'neto_pagar': suma_neto,
        'total_horas': suma_horas
    }

# ==============================================================================
# FUNCIÓN DE RENTABILIDAD CON CONEXIÓN DIRECTA A DB (GOOGLE SHEETS)
# ==============================================================================
def procesar_rentabilidad_db(df_cobro_raw, titulo_modulo, df_pagos_reales_rent, corte_a_evaluar, total_nomina_bd_rent, costo_label="Costo Nómina Real (Pagos SERGEM)"):
    try:
        col_periodo = obtener_nombre_columna(df_cobro_raw, ['PERIODO', 'CORTE'])
        
        if corte_a_evaluar == "GLOBAL":
            df_cobro = df_cobro_raw.copy()
        elif col_periodo:
            df_cobro = df_cobro_raw[df_cobro_raw[col_periodo].astype(str).str.strip().str.upper() == str(corte_a_evaluar).strip().upper()].copy()
        else:
            df_cobro = df_cobro_raw.copy()

        if df_cobro.empty:
            return

        if titulo_modulo == "LTSA":
            st.session_state['df_raw_LTSA'] = df_cobro.copy()
            
        col_ced_cobro = obtener_nombre_columna(df_cobro, ['CÉDULA', 'CEDULA', 'CC', 'C.C.', 'IDENTIFICACION', 'IDENTIFICACIÓN'])
        col_total_cobro = obtener_nombre_columna(df_cobro, ['TOTAL', 'VALOR TOTAL', 'TOTAL FACTURAR', 'NETO'])
        col_vehiculo = obtener_nombre_columna(df_cobro, ['TIPO DE VEHICULO', 'VEHICULO', 'CATEGORIA'])
        col_almacen = obtener_nombre_columna(df_cobro, ['PUNTO DE VENTA', 'ALMACEN', 'CLIENTE'])
        col_nombre_cobro = obtener_nombre_columna(df_cobro, ['NOMBRE', 'NOMBRES', 'CONDUCTOR', 'EMPLEADO', 'BENEFICIARIO']) 
        
        if not (col_ced_cobro and col_total_cobro):
            st.error(f"No se encontraron columnas de Cédula o Total en la base de datos ({titulo_modulo}).")
            return

        df_cobro = df_cobro.dropna(subset=[col_ced_cobro])
        df_cobro['_cedula_clean'] = df_cobro[col_ced_cobro].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.strip()
        df_cobro = df_cobro[df_cobro['_cedula_clean'] != '']
        
        df_cobro[col_total_cobro] = pd.to_numeric(df_cobro[col_total_cobro], errors='coerce').fillna(0)
        
        agg_dict = { col_total_cobro: 'sum' }
        if col_vehiculo: agg_dict[col_vehiculo] = 'first'
        if col_almacen: agg_dict[col_almacen] = 'first'
        if col_nombre_cobro: agg_dict[col_nombre_cobro] = 'first'
        
        cobro_agrupado = df_cobro.groupby('_cedula_clean').agg(agg_dict).reset_index()
        cobro_agrupado.rename(columns={col_total_cobro: 'TOTAL_COBRADO_LTSA'}, inplace=True)
        
        if df_pagos_reales_rent.empty and '_cedula_clean' not in df_pagos_reales_rent.columns:
            df_pagos_reales_rent = pd.DataFrame(columns=['_cedula_clean', 'VALOR_PAGADO_NETO', 'NOMBRE_EMPLEADO'])
            
        df_cruce = pd.merge(cobro_agrupado, df_pagos_reales_rent, on='_cedula_clean', how='outer')
        df_cruce['TOTAL_COBRADO_LTSA'] = df_cruce['TOTAL_COBRADO_LTSA'].fillna(0)
        df_cruce['VALOR_PAGADO_NETO'] = df_cruce['VALOR_PAGADO_NETO'].fillna(0)
        
        if col_nombre_cobro in df_cruce.columns and 'NOMBRE_EMPLEADO' in df_cruce.columns:
            df_cruce['NOMBRE_EMPLEADO'] = df_cruce['NOMBRE_EMPLEADO'].fillna(df_cruce[col_nombre_cobro])
        elif 'NOMBRE_EMPLEADO' not in df_cruce.columns and col_nombre_cobro in df_cruce.columns:
            df_cruce['NOMBRE_EMPLEADO'] = df_cruce[col_nombre_cobro]
            
        df_cruce['NOMBRE_EMPLEADO'] = df_cruce.get('NOMBRE_EMPLEADO', pd.Series(['S/N']*len(df_cruce))).fillna("S/N (Sin cobro/pago)")
        
        df_cruce['UTILIDAD_REAL_NETA'] = df_cruce['TOTAL_COBRADO_LTSA'] - df_cruce['VALOR_PAGADO_NETO']
        df_cruce['MARGEN_REAL'] = (df_cruce['UTILIDAD_REAL_NETA'] / df_cruce['TOTAL_COBRADO_LTSA'].replace(0, 1)).fillna(0)
        
        if col_vehiculo:
            df_cruce['CATEGORIA_VEHICULO'] = df_cruce.get(col_vehiculo, pd.Series(["NO DEFINIDO"]*len(df_cruce))).apply(lambda x: 
                "MOTO CARGUERO" if pd.notna(x) and "CARGUERO" in str(x).upper() 
                else "MOTO" if pd.notna(x) and "MOTO" in str(x).upper() 
                else "CARRY / CARRO" if pd.notna(x)
                else "NO IDENTIFICADO"
            )
        else:
            df_cruce['CATEGORIA_VEHICULO'] = "NO DEFINIDO"
        
        st.success(f"✅ Mostrando los **{len(df_cruce)}** registros de rentabilidad reales.")
        
        tab_emp, tab_veh, tab_alm = st.tabs(["👥 Rentabilidad por Empleado", "🛵 Rentabilidad por Vehículo", "🏢 Rentabilidad por Almacén"])
        
        format_dict = {
            'TOTAL_COBRADO_LTSA': '${:,.0f}', 'VALOR_PAGADO_NETO': '${:,.0f}',
            'UTILIDAD_REAL_NETA': '${:,.0f}', 'MARGEN_REAL': '{:.1%}'
        }

        with tab_emp:
            st.markdown("#### Detalle Real por Colaborador")
            df_show_emp = df_cruce[['_cedula_clean', 'NOMBRE_EMPLEADO', 'CATEGORIA_VEHICULO', 'TOTAL_COBRADO_LTSA', 'VALOR_PAGADO_NETO', 'UTILIDAD_REAL_NETA', 'MARGEN_REAL']].sort_values('UTILIDAD_REAL_NETA', ascending=False)
            
            styler_emp = df_show_emp.style.format(format_dict)
            try:
                if hasattr(styler_emp, 'map'):
                    styler_emp = styler_emp.map(lambda x: 'color: #E3000F' if x < 0 else 'color: #15803d', subset=['UTILIDAD_REAL_NETA'])
                else:
                    styler_emp = styler_emp.applymap(lambda x: 'color: #E3000F' if x < 0 else 'color: #15803d', subset=['UTILIDAD_REAL_NETA'])
                st.dataframe(styler_emp, hide_index=True, use_container_width=True, height=400)
            except Exception:
                st.dataframe(df_show_emp, hide_index=True, use_container_width=True, height=400)
            
        with tab_veh:
            st.markdown("#### Rentabilidad Consolidada por Vehículo")
            df_veh_agg = df_cruce.groupby('CATEGORIA_VEHICULO').agg({
                'TOTAL_COBRADO_LTSA': 'sum', 'VALOR_PAGADO_NETO': 'sum', 'UTILIDAD_REAL_NETA': 'sum'
            }).reset_index()
            df_veh_agg['MARGEN_REAL'] = (df_veh_agg['UTILIDAD_REAL_NETA'] / df_veh_agg['TOTAL_COBRADO_LTSA'].replace(0, 1)).fillna(0)
            st.dataframe(df_veh_agg.style.format(format_dict), hide_index=True, use_container_width=True)
            
        with tab_alm:
            if col_almacen:
                st.markdown("#### Rentabilidad Consolidada por Almacén")
                df_alm_agg = df_cruce.groupby(col_almacen).agg({
                    'TOTAL_COBRADO_LTSA': 'sum', 'VALOR_PAGADO_NETO': 'sum', 'UTILIDAD_REAL_NETA': 'sum'
                }).reset_index().sort_values('UTILIDAD_REAL_NETA', ascending=False)
                df_alm_agg['MARGEN_REAL'] = (df_alm_agg['UTILIDAD_REAL_NETA'] / df_alm_agg['TOTAL_COBRADO_LTSA'].replace(0, 1)).fillna(0)
                st.dataframe(df_alm_agg.style.format(format_dict), hide_index=True, use_container_width=True, height=400)
            else:
                st.info("No se encontró columna de almacén en este validador.")
        
        st.divider()
        st.markdown(f"### 💰 Gran Total ({titulo_modulo})")
        
        r1, r2, r3 = st.columns(3)
        tot_cobrado = float(df_cruce['TOTAL_COBRADO_LTSA'].sum())
        tot_pagado = float(df_cruce['VALOR_PAGADO_NETO'].sum())
        tot_utilidad = float(df_cruce['UTILIDAD_REAL_NETA'].sum())
        
        r1.metric("Facturación Cliente", f"${tot_cobrado:,.0f}")
        r2.metric(costo_label, f"${tot_pagado:,.0f}")
        r3.metric("UTILIDAD NETA", f"${tot_utilidad:,.0f}")

    except Exception as e:
        st.error(f"Error calculando la rentabilidad: {e}")

# ==============================================================================
# INTERFAZ DE USUARIO 
# ==============================================================================
col1, col2 = st.columns([1, 4])
with col1:
    try:
        if os.path.exists("sergemLogo.png"): st.image("sergemLogo.png", use_column_width=True)
        elif os.path.exists("sergemLogo_2.png"): st.image("sergemLogo_2.png", use_column_width=True)
    except: pass
with col2:
    st.title("Generador Automático de Documentos")
    st.markdown("**SERGEM Mensajería S.A.S.**")

if st.button("🔄 Sincronizar Base de Datos", key="btn_sync", type="secondary"):
    cargar_datos.clear() 
    st.cache_data.clear()
    st.rerun()

with st.spinner("Conectando con Google Sheets..."):
    data_cruda = cargar_datos(GAS_URL)

if not data_cruda:
    st.error("Error conectando a Google Sheets. Revise la URL o los permisos del Apps Script.")
    st.stop()

# ==============================================================================
# SEPARACIÓN ESTRICTA DE BASES DE DATOS
# ==============================================================================
df_pagos_completo = pd.DataFrame(data_cruda.get('pagos', []))
df_bd_maestra = pd.DataFrame(data_cruda.get('bd', []))
df_fuera = pd.DataFrame(data_cruda.get('fueras_perimetro', []))

df_rent_hist = pd.DataFrame(data_cruda.get('rent_hist', []))
df_rent_pollos = pd.DataFrame(data_cruda.get('rent_pollos', []))
df_rent_directo = pd.DataFrame(data_cruda.get('rent_directo', []))
df_rent_pollos_pago = pd.DataFrame(data_cruda.get('rent_pollos_pago', []))

if df_pagos_completo.empty:
    st.warning("No se encontraron datos en la pestaña PAGOS PERSONAL POR SERVICIOS.")
    st.stop()

df_pagos_completo.columns = df_pagos_completo.columns.str.strip().str.upper()
if not df_bd_maestra.empty: df_bd_maestra.columns = df_bd_maestra.columns.str.strip().str.upper()
if not df_fuera.empty: df_fuera.columns = df_fuera.columns.str.strip().str.upper()
if not df_rent_hist.empty: df_rent_hist.columns = df_rent_hist.columns.str.strip().str.upper()
if not df_rent_pollos.empty: df_rent_pollos.columns = df_rent_pollos.columns.str.strip().str.upper()
if not df_rent_directo.empty: df_rent_directo.columns = df_rent_directo.columns.str.strip().str.upper()
if not df_rent_pollos_pago.empty: df_rent_pollos_pago.columns = df_rent_pollos_pago.columns.str.strip().str.upper()

if 'CORTE' in df_pagos_completo.columns:
    df_pagos_completo['CORTE'] = df_pagos_completo['CORTE'].astype(str).str.strip().str.upper()

col_prestador = obtener_nombre_columna(df_pagos_completo, ['A NOMBRE DE QUIEN HACE CUENTA DE COBRO', 'NOMBRE PRESTADOR', 'A NOMBRE DE QUIEN HACE CUENTA'])
col_cedula_prestador = obtener_nombre_columna(df_pagos_completo, ['CÉDULA DE CUENTA DE COBRO', 'CEDULA DE CUENTA DE COBRO'])
col_titular_banco = obtener_nombre_columna(df_pagos_completo, ['NOMBRE TITULAR CUENTA BANCARIA', 'NOMBRE_TITULAR'])
col_cedula_banco = obtener_nombre_columna(df_pagos_completo, ['CÉDULA TITULAR', 'CEDULA TITULAR'])
col_estado = obtener_nombre_columna(df_pagos_completo, ['ESTADO', 'ESTADO_EMPLEADO'])
col_anticipos = obtener_nombre_columna(df_pagos_completo, ['ANTICIPOS', 'ANTICIPO'])
col_otros_desc = obtener_nombre_columna(df_pagos_completo, ['OTROS DESCUENTOS', 'OTROS_DESCUENTOS', 'DESCUENTOS'])
col_valor_dia = obtener_nombre_columna(df_pagos_completo, ['VALOR DIA NEGOCIADO', 'VALOR_DIA_NEGOCIADO', 'VALOR DIA'])

if not col_prestador or not col_titular_banco:
    st.error("Faltan las columnas que diferencian a quien cobra del titular del banco. Verifique sus nombres en el Sheets.")
    st.stop()

# LA LISTA GENERAL SUPERIOR VUELVE A SER NORMAL (Solo periodos de pagos)
cortes_disponibles = [c for c in df_pagos_completo['CORTE'].unique() if str(c).strip() != "" and str(c).lower() != "nan"]

st.divider()

corte_seleccionado = st.selectbox("📅 Seleccione el Corte a procesar / visualizar:", cortes_disponibles)

# El df general para el resto del programa
df_pagos_corte = df_pagos_completo[df_pagos_completo['CORTE'] == corte_seleccionado].copy()

df_pagos_corte['_ced_prestador_clean'] = df_pagos_corte[col_cedula_prestador].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.strip()
df_pagos_corte['_ced_banco_clean'] = df_pagos_corte[col_cedula_banco].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.strip()

tab_generador, tab_informes, tab_rentabilidad = st.tabs([
    "📄 Generador de Documentos", 
    "📊 Panel de Informes Gerenciales", 
    "📈 Rentabilidad Operativa"
])

# ==============================================================================
# PESTAÑA 1: GENERADOR
# ==============================================================================
with tab_generador:
    modo_trabajo = st.radio("⚙️ Modo de trabajo:", 
                            ["🗂️ Generación Masiva (Paquete Gerencial)", 
                             "👤 Vista Previa Individual",
                             "⏱️ Actualizador de Horas Automático (Excel a Drive)"], horizontal=True)

    if "Actualizador" in modo_trabajo:
        st.markdown("### ⏱️ Depurador y Actualizador de Horas (De Sistema a Drive)")
        st.info("Sube el archivo Excel biométrico. El programa construirá una tabla depurada basándose en tu pestaña BD para que la pegues en PAGOS PERSONAL POR SERVICIOS.")
        
        nuevo_corte = st.text_input("✍️ Escriba el nombre exacto del Corte a generar (Ej: 1 AL 15 AGOSTO):")
        archivo_horas = st.file_uploader("📥 Sube el reporte de horas en formato Excel (.xlsx)", type=["xlsx", "xls"])
        
        if archivo_horas and nuevo_corte:
            try:
                df_raw = pd.read_excel(archivo_horas)
                
                col_cc = obtener_nombre_columna(df_raw, ['CC', 'CEDULA', 'CÉDULA'])
                col_horas = obtener_nombre_columna(df_raw, ['TOTAL_HORAS', 'TOTAL HORAS', 'HORAS'])
                
                if col_cc and col_horas:
                    df_raw[col_cc] = pd.to_numeric(df_raw[col_cc], errors='coerce')
                    df_raw = df_raw.dropna(subset=[col_cc])
                    
                    agg_dict = {col: 'first' for col in df_raw.columns if col != col_cc and col != col_horas}
                    agg_dict[col_horas] = 'sum'
                    grouped = df_raw.groupby(col_cc, as_index=False).agg(agg_dict)
                    
                    columnas_destino = [c for c in df_bd_maestra.columns if str(c).strip() != "" and "UNNAMED" not in str(c).upper()]
                    
                    col_ced_bd = obtener_nombre_columna(df_bd_maestra, ['CÉDULA', 'CEDULA', 'C.C.', 'C.C', 'CC'])
                    col_horas_bd = obtener_nombre_columna(df_bd_maestra, ['NÚMERO DE HORAS', 'NUMERO DE HORAS', 'HORAS', 'TOTAL HORAS'])
                    col_corte_bd = obtener_nombre_columna(df_bd_maestra, ['CORTE', 'PERIODO'])
                    col_total_bd = obtener_nombre_columna(df_bd_maestra, ['TOTAL A PAGAR', 'TOTAL_A_PAGAR'])
                    col_val_hora_bd = obtener_nombre_columna(df_bd_maestra, ['VALOR HORA', 'VALOR_HORA'])
                    
                    result_rows = []
                    for _, row in grouped.iterrows():
                        cc = row[col_cc]
                        horas = row[col_horas]
                        
                        match = pd.DataFrame()
                        if col_ced_bd:
                            ced_bd = df_bd_maestra[col_ced_bd].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            ced_match = str(cc).replace(".0", "").strip()
                            match = df_bd_maestra[ced_bd == ced_match]
                        
                        new_row = {}
                        for col in columnas_destino:
                            val_final = ""
                            alias_busqueda = [col]
                            
                            if col == "CLIENTE": alias_busqueda.extend(["EMPRESA", "PUNTO_VENTA"])
                            if col == "CONDUCTOR": alias_busqueda.extend(["MENSAJERO", "NOMBRE"])
                            if col == "VALOR HORA": alias_busqueda.extend(["VALOR_HORA"])
                            if col == "ESTADO": alias_busqueda.extend(["ESTADO_EMPLEADO"])
                            if col == "TIPO DE DOCUMENTO": alias_busqueda.extend(["TIPO_DOCUMENTO", "DOCUMENTO"])
                            
                            col_raw_match = obtener_nombre_columna(df_raw, alias_busqueda)
                            if col_raw_match and pd.notna(row[col_raw_match]) and str(row[col_raw_match]).strip() != "":
                                val_final = row[col_raw_match]
                            elif not match.empty:
                                bd_row = match.iloc[0]
                                bd_col_match = obtener_nombre_columna(df_bd_maestra, alias_busqueda)
                                if bd_col_match and pd.notna(bd_row[bd_col_match]) and str(bd_row[bd_col_match]).strip() != "":
                                    val_final = bd_row[bd_col_match]
                                    
                            new_row[col] = val_final
                        
                        if col_ced_bd: new_row[col_ced_bd] = int(cc) if cc else ""
                        if col_horas_bd: new_row[col_horas_bd] = round(horas, 2)
                        if col_corte_bd: new_row[col_corte_bd] = nuevo_corte.strip().upper()
                        
                        val_hora = 0
                        if col_val_hora_bd: val_hora = limpiar_dinero(new_row.get(col_val_hora_bd, 0))
                        if col_total_bd: new_row[col_total_bd] = round(horas * val_hora, 0) if val_hora > 0 else 0
                            
                        result_rows.append(new_row)
                        
                    df_res = pd.DataFrame(result_rows)
                    st.success(f"✅ ¡Cruce Exitoso 100% Dinámico! Se extrajeron las horas y cruzaron con todas las columnas actuales de BD.")
                    
                    excel_out = io.BytesIO()
                    df_res.to_excel(excel_out, index=False, sheet_name="PAGOS PERSONAL POR SERVICIOS")
                    excel_out.seek(0)
                    
                    st.download_button(
                        label="📥 DESCARGAR BASE DEPURADA (EXCEL PARA DRIVE)",
                        data=excel_out,
                        file_name=f"Base_Depurada_SERGEM_{nuevo_corte.replace(' ', '_').replace('/', '-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    st.error("El archivo subido no tiene la estructura biométrica.")
            except Exception as e:
                st.error(f"Error procesando el archivo: {e}")

    else:
        if not df_fuera.empty:
            col_cond_fuera_test = obtener_nombre_columna(df_fuera, ['CONDUCTOR', 'NOMBRE', 'NOMBRES'])
            if not col_cond_fuera_test:
                st.warning("⚠️ **Aviso Importante:** No se detectó una columna llamada 'CONDUCTOR' en la pestaña *FUERAS PERIMETRO /ADIC*.")

        if "Individual" in modo_trabajo:
            titulares_unicos = df_pagos_corte.drop_duplicates(subset=['_ced_prestador_clean', '_ced_banco_clean'])
            
            lista_opciones = []
            for _, row in titulares_unicos.iterrows():
                if row['_ced_prestador_clean'] in ['nan', '', 'None']: continue
                lbl = f"📄 Factura: {row[col_prestador]} (NIT: {row['_ced_prestador_clean']}) ➔ 🏦 Cuenta de Cobro a nombre de {row[col_titular_banco]} (CC: {row['_ced_banco_clean']})"
                lista_opciones.append(lbl)
                
            if "mostrar_preview" not in st.session_state:
                st.session_state.mostrar_preview = False
                st.session_state.titular_actual = ""
                
            opcion_seleccionada = st.selectbox("Seleccione el documento a previsualizar:", sorted(lista_opciones))
            
            if st.button("🔍 Calcular y Previsualizar"):
                st.session_state.mostrar_preview = True
                st.session_state.titular_actual = opcion_seleccionada
                
            if st.session_state.mostrar_preview and st.session_state.titular_actual == opcion_seleccionada:
                
                try:
                    ced_prestador_seleccionada = opcion_seleccionada.split("(NIT: ")[1].split(")")[0].strip()
                    ced_banco_seleccionada = opcion_seleccionada.split("(CC: ")[1].split(")")[0].strip()
                    
                    mask_prestador = df_pagos_corte['_ced_prestador_clean'] == ced_prestador_seleccionada
                    mask_banco = df_pagos_corte['_ced_banco_clean'] == ced_banco_seleccionada
                    
                    grupo_titular = df_pagos_corte[mask_prestador & mask_banco]
                    calculos = calcular_valores_agrupados(grupo_titular, df_fuera, corte_seleccionado, col_prestador, col_cedula_prestador, col_titular_banco, col_cedula_banco, col_estado, col_anticipos, col_otros_desc, col_valor_dia)
                    
                    if not calculos:
                        st.warning("Este titular tiene saldo neto en $0 para este corte sin anticipos reportados.")
                    else:
                        st.markdown(f"### Resumen Financiero: {calculos['nombre_prestador']} (Cuenta destino: {calculos['nombre_titular_banco']})")
                        if calculos['es_nuevo']: st.info("👤 **ESTADO NUEVO DETECTADO:** Esta persona se incluirá en el paquete exclusivo de nuevos.")
                        
                        cA, cB, cC, cD = st.columns(4)
                        cA.markdown(f"<div class='metric-box'><b>BASE BRUTA</b><br>${calculos['ingreso_base']:,.0f}</div>", unsafe_allow_html=True)
                        cB.markdown(f"<div class='metric-box'><b>DESCUENTOS LEGALES</b><br>${calculos['retefuente']+calculos['ica']:,.0f}</div>", unsafe_allow_html=True)
                        cC.markdown(f"<div class='metric-box'><b>ANTICIPOS/ARL</b><br>${calculos['anticipos']+calculos['otros_descuentos']:,.0f}</div>", unsafe_allow_html=True)
                        cD.markdown(f"<div class='metric-box' style='background-color:#E3000F; color:white;'><b>TOTAL CONSIGNAR</b><br>${calculos['neto_final']:,.0f}</div>", unsafe_allow_html=True)
                        
                        if calculos['fuera_perimetro'] > 0: st.success(f"🚚 **¡Valor detectado!** Se sumaron **${calculos['fuera_perimetro']:,.0f}**.")
                        if calculos['neto_final'] <= 0: st.warning("⚠️ El total a consignar da **$0 o negativo**.")

                    datos_doc = calculos.copy()
                    datos_doc.update({'id': "", 'fecha_emision': obtener_fecha_actual(), 'corte_fechas': corte_seleccionado})
                    
                    pdf_ct = FPDF(); agregar_pagina_pdf_cuenta_cobro(pdf_ct, datos_doc)
                    pdf_eq = FPDF(); agregar_pagina_pdf_doc_equivalente(pdf_eq, datos_doc)
                    
                    colBtn1, colBtn2 = st.columns(2)
                    colBtn1.download_button("📥 Descargar Cuenta de Cobro (PDF)", data=get_pdf_bytes(pdf_ct), file_name=f"Cuenta_{calculos['cedula_prestador']}.pdf", mime="application/pdf", use_container_width=True)
                    colBtn2.download_button("📥 Descargar Doc. Equivalente (PDF)", data=get_pdf_bytes(pdf_eq), file_name=f"DocEq_{calculos['cedula_prestador']}.pdf", mime="application/pdf", use_container_width=True)
                except Exception as ex:
                    st.error(f"Error procesando los datos de esta persona: {ex}")

        elif "Masiva" in modo_trabajo:
            if st.button("🚀 Procesar Lote General", use_container_width=True, type="primary"):
                mensaje_carga = st.info(f"📥 Procesando la información y empaquetando archivos de forma inteligente...")
                
                try:
                    pagos_procesados_banco = []
                    nuevos_detectados = []
                    ceros_detectados = []
                    ignorados = count_banco = count_nuevos = count_ceros = 0
                    fecha_actual = obtener_fecha_actual() 
                    
                    pdf_ct_banco = FPDF(); pdf_eq_banco = FPDF()
                    wb_eq_banco = openpyxl.Workbook(); wb_eq_banco.remove(wb_eq_banco.active)
                    
                    pdf_ct_nuevos = FPDF(); pdf_eq_nuevos = FPDF()
                    wb_eq_nuevos = openpyxl.Workbook(); wb_eq_nuevos.remove(wb_eq_nuevos.active)
                    
                    pdf_ct_ceros = FPDF(); pdf_eq_ceros = FPDF()
                    wb_eq_ceros = openpyxl.Workbook(); wb_eq_ceros.remove(wb_eq_ceros.active)
                    
                    grupos = df_pagos_corte.groupby(['_ced_prestador_clean', '_ced_banco_clean'])
                    contador = 1
                    
                    for (ced_prestador, ced_banco), grupo_titular in grupos:
                        if ced_prestador in ['nan', '', 'None']: continue
                        
                        calculos = calcular_valores_agrupados(grupo_titular, df_fuera, corte_seleccionado, col_prestador, col_ced_prestador, col_titular_banco, col_ced_banco, col_estado, col_anticipos, col_otros_desc, col_valor_dia)
                        
                        if not calculos: continue
                        
                        datos_doc = calculos.copy()
                        datos_doc.update({'id': str(contador).zfill(3), 'fecha_emision': fecha_actual, 'corte_fechas': corte_seleccionado})
                        nombre_pestana = f"{contador}_{datos_doc['nombre_prestador'][:20]}".replace(":", "").replace("/", "-")

                        if datos_doc['es_nuevo']:
                            agregar_pagina_pdf_cuenta_cobro(pdf_ct_nuevos, datos_doc); agregar_pagina_pdf_doc_equivalente(pdf_eq_nuevos, datos_doc)
                            ws = wb_eq_nuevos.create_sheet(title=nombre_pestana); construir_hoja_documento_equivalente_excel(ws, datos_doc)
                            nuevos_detectados.append(datos_doc); count_nuevos += 1

                        if datos_doc['neto_final'] <= 0:
                            agregar_pagina_pdf_cuenta_cobro(pdf_ct_ceros, datos_doc); agregar_pagina_pdf_doc_equivalente(pdf_eq_ceros, datos_doc)
                            ws = wb_eq_ceros.create_sheet(title=nombre_pestana); construir_hoja_documento_equivalente_excel(ws, datos_doc)
                            ceros_detectados.append(datos_doc); count_ceros += 1
                        else:
                            agregar_pagina_pdf_cuenta_cobro(pdf_ct_banco, datos_doc); agregar_pagina_pdf_doc_equivalente(pdf_eq_banco, datos_doc)
                            ws = wb_eq_banco.create_sheet(title=nombre_pestana); construir_hoja_documento_equivalente_excel(ws, datos_doc)
                            
                            pagos_procesados_banco.append({
                                'TIPO_IDENTIFICACION': datos_doc['tipo_documento_pab'],
                                'NIT_BENEFICIARIO': datos_doc['cedula_titular_banco'],
                                'NOMBRE_BENEFICIARIO': datos_doc['nombre_titular_banco'],
                                'BANCO_DESTINO': datos_doc['banco'],
                                'TIPO_CUENTA': datos_doc['tipo_cuenta'],
                                'NUMERO_CUENTA': datos_doc['num_cuenta'],
                                'VALOR_NETO_A_PAGAR': datos_doc['neto_final']
                            })
                            count_banco += 1
                        contador += 1
                    
                    zip_cuentas_banco_io = io.BytesIO()
                    if count_banco > 0:
                        with zipfile.ZipFile(zip_cuentas_banco_io, "w", zipfile.ZIP_DEFLATED) as zipf: zipf.writestr("Cuentas_de_Cobro_Aprobadas.pdf", get_pdf_bytes(pdf_ct_banco))
                    zip_cuentas_banco_io.seek(0)
                    
                    zip_eq_banco_io = io.BytesIO()
                    if count_banco > 0:
                        with zipfile.ZipFile(zip_eq_banco_io, "w", zipfile.ZIP_DEFLATED) as zipf:
                            zipf.writestr("Documentos_Equivalentes_Aprobados.pdf", get_pdf_bytes(pdf_eq_banco))
                            excel_io = io.BytesIO(); wb_eq_banco.save(excel_io); excel_io.seek(0)
                            zipf.writestr("Documentos_Equivalentes_Excel.xlsx", excel_io.read())
                    zip_eq_banco_io.seek(0)
                    
                    df_banco = pd.DataFrame(pagos_procesados_banco)
                    archivo_pab_bytes = generar_excel_pab(df_banco, corte_seleccionado) if len(df_banco) > 0 else b""
                    
                    mensaje_carga.empty() 
                    st.success(f"✅ ¡Éxito! Procesamiento finalizado. **{count_banco}** pagos aprobados listos para pago en banco.")
                    st.divider()

                    if count_nuevos > 0:
                        df_nuevos = pd.DataFrame([{'NOMBRES Y APELLIDOS': d['nombre_titular_banco'], 'CÉDULA': d['cedula_titular_banco'], 'BANCO': d['banco'], 'TIPO CUENTA': d['tipo_cuenta'], 'NO. CUENTA': d['num_cuenta']} for d in nuevos_detectados])
                        excel_nuevos_io = io.BytesIO(); df_nuevos.to_excel(excel_nuevos_io, index=False, sheet_name="PERSONAL NUEVO"); excel_nuevos_io.seek(0)
                        st.error(f"🚨 **ATENCIÓN - SE DETECTARON {count_nuevos} PERSONAS NUEVAS**")
                        colN1, colN2, colN3 = st.columns(3)
                        colN1.download_button("1️⃣ 📥 Listado Excel (Para Don José)", data=excel_nuevos_io, file_name="Listado_Nuevos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        colN2.download_button("2️⃣ 📥 Cuentas de Cobro (Solo Nuevos)", data=get_pdf_bytes(pdf_ct_nuevos), file_name="Cuentas_Cobro_Nuevos.pdf", mime="application/pdf", use_container_width=True)
                        zip_eq_nuevos_io = io.BytesIO()
                        with zipfile.ZipFile(zip_eq_nuevos_io, "w", zipfile.ZIP_DEFLATED) as zipf:
                            zipf.writestr("Docs_Equivalentes_Nuevos.pdf", get_pdf_bytes(pdf_eq_nuevos)); excel_io = io.BytesIO(); wb_eq_nuevos.save(excel_io); excel_io.seek(0)
                            zipf.writestr("Docs_Equivalentes_Nuevos_Excel.xlsx", excel_io.read())
                        zip_eq_nuevos_io.seek(0)
                        colN3.download_button("3️⃣ 📥 Docs Equivalentes (Solo Nuevos)", data=zip_eq_nuevos_io, file_name="Docs_Equivalentes_Nuevos.zip", mime="application/zip", use_container_width=True)
                        st.divider()
                    
                    if count_ceros > 0:
                        nombres_ceros = "\n* ".join([f"👤 {d['nombre_prestador']} (C.C: {d['cedula_prestador']})" for d in ceros_detectados])
                        st.warning(f"⚠️ **SE DETECTARON {count_ceros} SALDOS EN CERO O NEGATIVOS**\n* {nombres_ceros}")
                        colC1, colC2 = st.columns(2)
                        colC1.download_button("1️⃣ 📥 Cuentas de Cobro (Saldos Cero)", data=get_pdf_bytes(pdf_ct_ceros), file_name="Cuentas_Cobro_Ceros.pdf", mime="application/pdf", use_container_width=True)
                        zip_eq_ceros_io = io.BytesIO()
                        with zipfile.ZipFile(zip_eq_ceros_io, "w", zipfile.ZIP_DEFLATED) as zipf:
                            zipf.writestr("Docs_Equivalentes_Ceros.pdf", get_pdf_bytes(pdf_ct_ceros)); excel_io = io.BytesIO(); wb_eq_ceros.save(excel_io); excel_io.seek(0)
                            zipf.writestr("Docs_Equivalentes_Ceros_Excel.xlsx", excel_io.read())
                        zip_eq_ceros_io.seek(0)
                        colC2.download_button("2️⃣ 📥 Docs Equivalentes (Saldos Cero)", data=zip_eq_ceros_io, file_name="Docs_Equivalentes_Ceros.zip", mime="application/zip", use_container_width=True)
                        st.divider()

                    st.markdown(f"### 📥 Soportes Contables Aprobados (Los {count_banco} del Banco)")
                    colD1, colD2, colD3 = st.columns(3)
                    colD1.download_button(label="1️⃣ Soportes: Cuentas de Cobro (.ZIP)", data=zip_cuentas_banco_io, file_name=f"Cuentas_Cobro_Aprobadas_{corte_seleccionado.replace(' ', '_').replace('/', '-')}.zip", mime="application/zip", use_container_width=True, disabled=(count_banco == 0))
                    colD2.download_button(label="2️⃣ Soportes: Docs. Equivalentes (.ZIP)", data=zip_eq_banco_io, file_name=f"Docs_Equivalentes_Aprobados_{corte_seleccionado.replace(' ', '_').replace('/', '-')}.zip", mime="application/zip", use_container_width=True, disabled=(count_banco == 0))
                    colD3.download_button(label="3️⃣ Archivo Excel PAB Banco (.XLSX)", data=archivo_pab_bytes, file_name=f"FORMATOPAB_{corte_seleccionado.replace(' ', '_').replace('/', '-')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, disabled=(count_banco == 0))

                except Exception as e:
                    mensaje_carga.empty()
                    st.error(f"Error en el proceso: {e}")

# ==============================================================================
# PESTAÑA 2: PANEL DE INFORMES GERENCIALES CON GRÁFICOS Y TABLAS 
# ==============================================================================
with tab_informes:
    st.markdown(f"### 📊 Informe Gerencial - Corte: {corte_seleccionado}")
    
    if cortes_disponibles:
        df_informe = df_pagos_completo[df_pagos_completo['CORTE'] == corte_seleccionado].copy()
        
        col_total_pagar = obtener_nombre_columna(df_informe, ['TOTAL A PAGAR', 'TOTAL_A_PAGAR'])
        df_informe['Valor Numérico'] = df_informe[col_total_pagar].apply(limpiar_dinero)
        
        col_horas_inf = obtener_nombre_columna(df_informe, ['NÚMERO DE HORAS', 'NUMERO DE HORAS', 'HORAS'])
        if col_horas_inf:
            df_informe[col_horas_inf] = pd.to_numeric(df_informe[col_horas_inf], errors='coerce').fillna(0)
        
        df_informe['_ced_prestador_clean'] = df_informe[col_cedula_prestador].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.strip()
        df_informe['_ced_banco_clean'] = df_informe[col_cedula_banco].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.strip()
        
        valid_records = df_informe[df_informe['_ced_prestador_clean'].isin(['nan', '', 'None']) == False]
        total_cuentas = len(valid_records.drop_duplicates(subset=['_ced_prestador_clean', '_ced_banco_clean']))
        
        cuentas_cero = len(df_informe[df_informe['Valor Numérico'] <= 0])
        total_consignar = df_informe['Valor Numérico'].sum()
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Documentos Emitidos", total_cuentas)
        c2.metric("Cuentas con Pago de $0", cuentas_cero)
        c3.metric("Cuentas Efectivas por Pagar", total_cuentas - cuentas_cero)
        c4.metric("Valor Total Quincena", f"${total_consignar:,.0f}")
        
        st.divider()
        st.markdown("### 📈 Análisis Financiero Detallado")
        
        t_cli, t_ban, t_con = st.tabs(["🏢 Consolidado por Cliente", "🏦 Distribución Bancaria", "🛵 Análisis de Conductores"])
        
        with t_cli:
            st.markdown("#### 📊 Inversión y Facturación por Cliente")
            col_cliente = obtener_nombre_columna(df_informe, ['CLIENTE', 'EMPRESA'])
            if col_cliente:
                df_cli = df_informe.groupby(col_cliente)['Valor Numérico'].sum().reset_index()
                df_cli = df_cli.sort_values('Valor Numérico', ascending=True) 
                
                cc1, cc2 = st.columns([3, 2])
                with cc1:
                    fig_cliente = px.bar(
                        df_cli, x='Valor Numérico', y=col_cliente, orientation='h',
                        title="Distribución Total a Pagar por Cliente", text_auto='$.3s', color='Valor Numérico', color_continuous_scale='Reds'
                    )
                    fig_cliente.update_layout(showlegend=False, xaxis_title="Total a Pagar ($)", yaxis_title="")
                    st.plotly_chart(fig_cliente, use_container_width=True)
                
                with cc2:
                    st.markdown("**📋 Tabla Detallada Exacta**")
                    df_cli_table = df_cli.sort_values('Valor Numérico', ascending=False).rename(columns={col_cliente: "Cliente", "Valor Numérico": "Total a Pagar"})
                    df_cli_table["Total a Pagar"] = pd.to_numeric(df_cli_table["Total a Pagar"], errors='coerce').fillna(0)
                    st.dataframe(df_cli_table.style.format({"Total a Pagar": "${:,.0f}"}), hide_index=True, use_container_width=True)
        
        with t_ban:
            st.markdown("#### 🏦 Distribución de Fondos por Entidad Bancaria")
            col_banco_graf = obtener_nombre_columna(df_informe, ['BANCO', 'BANCO DESTINO'])
            if col_banco_graf:
                df_ban = df_informe.groupby(col_banco_graf).agg(
                    Cuentas=(col_banco_graf, 'count'), Total=('Valor Numérico', 'sum')
                ).reset_index().sort_values('Total', ascending=False)
                
                cb1, cb2 = st.columns([3, 2])
                with cb1:
                    fig_banco = px.pie(
                        df_ban, names=col_banco_graf, values='Total', title="Porcentaje de Dinero Concentrado por Banco",
                        hole=0.45, color_discrete_sequence=px.colors.sequential.Reds_r
                    )
                    fig_banco.update_traces(textposition='inside', textinfo='percent+label')
                    fig_banco.update_layout(showlegend=False)
                    st.plotly_chart(fig_banco, use_container_width=True)
                    
                with cb2:
                    st.markdown("**📋 Directorio de Transferencias Bancarias**")
                    df_ban_table = df_ban.rename(columns={col_banco_graf: "Entidad Bancaria", "Cuentas": "Cant. Transferencias", "Total": "Dinero a Girar"})
                    df_ban_table["Dinero a Girar"] = pd.to_numeric(df_ban_table["Dinero a Girar"], errors='coerce').fillna(0)
                    st.dataframe(df_ban_table.style.format({"Dinero a Girar": "${:,.0f}"}), hide_index=True, use_container_width=True)

        with t_con:
            st.markdown("#### 🛵 Top 15 de Conductores con Mayor Volumen de Pago")
            col_cond = obtener_nombre_columna(df_informe, ['CONDUCTOR', 'NOMBRES'])
            if col_cond:
                agg_dict_cond = {'Valor Numérico': 'sum'}
                if col_horas_inf: agg_dict_cond[col_horas_inf] = 'sum'
                df_cond = df_informe.groupby(col_cond).agg(agg_dict_cond).reset_index()
                df_cond_top = df_cond.sort_values('Valor Numérico', ascending=False).head(15).sort_values('Valor Numérico', ascending=True)
                
                c_c1, c_c2 = st.columns([3, 2])
                with c_c1:
                    fig_cond = px.bar(
                        df_cond_top, x='Valor Numérico', y=col_cond, orientation='h', title="Top 15 - Liquidación por Conductor",
                        text_auto='$.3s', color='Valor Numérico', color_continuous_scale='Reds'
                    )
                    fig_cond.update_layout(showlegend=False, xaxis_title="Total a Pagar ($)", yaxis_title="")
                    st.plotly_chart(fig_cond, use_container_width=True)
                
                with c_c2:
                    st.markdown(f"**📋 Directorio Completo (Todos los {len(df_cond)} Conductores)**")
                    df_cond_table = df_cond.sort_values('Valor Numérico', ascending=False)
                    rename_dict = {col_cond: "Conductor", "Valor Numérico": "Total a Pagar"}
                    if col_horas_inf: rename_dict[col_horas_inf] = "Horas Facturadas"
                    df_cond_table = df_cond_table.rename(columns=rename_dict)
                    df_cond_table["Total a Pagar"] = pd.to_numeric(df_cond_table["Total a Pagar"], errors='coerce').fillna(0)
                    
                    format_dict = {"Total a Pagar": "${:,.0f}"}
                    if "Horas Facturadas" in df_cond_table.columns: 
                        df_cond_table["Horas Facturadas"] = pd.to_numeric(df_cond_table["Horas Facturadas"], errors='coerce').fillna(0)
                        format_dict["Horas Facturadas"] = "{:,.1f}"
                    
                    st.dataframe(df_cond_table.style.format(format_dict), hide_index=True, use_container_width=True, height=400)

# ==============================================================================
# PESTAÑA 3: MÓDULO DE RENTABILIDAD OPERATIVA (CARGA Y VISUALIZACIÓN EN VIVO)
# ==============================================================================
with tab_rentabilidad:
    
    # 1. Leer los periodos EXISTENTES EXCLUSIVAMENTE de Rentabilidad LTSA y Pollos
    periodos_ltsa = df_rent_hist['PERIODO'].dropna().astype(str).str.strip().str.upper().tolist() if not df_rent_hist.empty and 'PERIODO' in df_rent_hist.columns else []
    periodos_pollos = df_rent_pollos['PERIODO'].dropna().astype(str).str.strip().str.upper().tolist() if not df_rent_pollos.empty and 'PERIODO' in df_rent_pollos.columns else []
    
    todos_periodos_rent = list(set(periodos_ltsa + periodos_pollos))
    todos_periodos_rent = [p for p in todos_periodos_rent if p not in ["", "NAN", "NONE"]]
    todos_periodos_rent.sort(reverse=True)
    
    opciones_rentabilidad = ["🌐 TODOS LOS CORTES (GLOBAL)"] + todos_periodos_rent
    
    st.markdown("<div class='caja-destacada'>", unsafe_allow_html=True)
    corte_a_evaluar = st.selectbox("📌 **SELECCIONA EL CORTE A ANALIZAR EN RENTABILIDAD:**", opciones_rentabilidad, help="Selecciona el periodo que deseas consultar en los gráficos inferiores. 'GLOBAL' consolidará todos los cortes cargados.")
    st.markdown("</div>", unsafe_allow_html=True)
    
    if corte_a_evaluar == "🌐 TODOS LOS CORTES (GLOBAL)":
        corte_evaluado_str = "GLOBAL"
        df_pagos_base = df_pagos_completo.copy()
    else:
        corte_evaluado_str = corte_a_evaluar
        # Para LTSA y Directo, cruzamos contra la base general filtrada por el mismo nombre
        df_pagos_base = df_pagos_completo[df_pagos_completo['CORTE'] == corte_a_evaluar].copy()
        
        if df_pagos_base.empty:
            st.warning(f"⚠️ **Aviso Importante:** En la pestaña principal de 'Pagos Personal' no existe ningún corte llamado exactamente '{corte_a_evaluar}'. Por esta razón, el sistema asumirá que el Costo de Nómina es $0. Si deseas ver los costos, asegúrate de que el nombre del corte coincida.")

    # Recalculamos los PAGOS REALES GLOBALES (Para LTSA y Personal Directo)
    col_total_pagar = obtener_nombre_columna(df_pagos_base, ['TOTAL A PAGAR', 'TOTAL_A_PAGAR', 'NETO'])
    col_ced_conductor = obtener_nombre_columna(df_pagos_base, ['CÉDULA', 'CEDULA', 'C.C.', 'CC'])
    col_nombre_conductor = obtener_nombre_columna(df_pagos_base, ['CONDUCTOR', 'NOMBRES', 'NOMBRE'])
    
    if col_total_pagar and col_ced_conductor and not df_pagos_base.empty:
        df_pagos_base['_valor_pagar_num'] = df_pagos_base[col_total_pagar].apply(limpiar_dinero)
        df_pagos_base['_cedula_clean'] = df_pagos_base[col_ced_conductor].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.strip()
        
        df_pagos_agrupados = df_pagos_base.groupby('_cedula_clean').agg(
            NOMBRE_EMPLEADO=(col_nombre_conductor if col_nombre_conductor else col_prestador, 'first'),
            VALOR_PAGADO_NETO=('_valor_pagar_num', 'sum')
        ).reset_index()
        
        df_pagos_reales_rent = df_pagos_agrupados[df_pagos_agrupados['_cedula_clean'].isin(['nan', '', 'None']) == False]
        total_nomina_bd_rent = df_pagos_reales_rent['VALOR_PAGADO_NETO'].sum()
    else:
        df_pagos_reales_rent = pd.DataFrame(columns=['_cedula_clean', 'VALOR_PAGADO_NETO', 'NOMBRE_EMPLEADO'])
        total_nomina_bd_rent = 0

    sub_ltsa, sub_pollos, sub_directo = st.tabs(["🚚 Rentabilidad LTSA", "🍗 Pollos y Panadería", "👷 Personal Directo"])
    
    # ------------------ LTSA ------------------
    with sub_ltsa:
        st.markdown("#### 📤 1. Cargar Validador a la Base de Datos")
        c1, c2 = st.columns([1, 2])
        # Solo pre-llenamos si no está en GLOBAL
        per_ltsa = c1.text_input("Digita el PERIODO a inyectar:", value=corte_seleccionado if corte_a_evaluar == "🌐 TODOS LOS CORTES (GLOBAL)" else corte_evaluado_str, key="per_ltsa", help="Sugerimos dejar este nombre para que coincida exactamente con la hoja de Pagos.")
        file_ltsa = c2.file_uploader("Archivo Validador Excel (LTSA)", type=["xlsx", "xls"], key="up_ltsa")
        
        if file_ltsa and per_ltsa:
            per_ltsa_clean = per_ltsa.strip().upper()
            reemplazar_ltsa = False
            
            if per_ltsa_clean in periodos_ltsa:
                st.warning(f"⚠️ ¡Atención! Ya existen datos registrados para el corte **{per_ltsa_clean}**.")
                reemplazar_ltsa = st.checkbox(f"Sobrescribir datos (Se borrarán los registros anteriores de {per_ltsa_clean} para evitar duplicados)", value=True)
            
            btn_disabled = (per_ltsa_clean in periodos_ltsa) and not reemplazar_ltsa
            
            if st.button("🚀 Enviar a Google Sheets (LTSA)", use_container_width=True, type="primary", disabled=btn_disabled):
                with st.spinner("Preparando archivo y subiendo a la nube (Puede tardar hasta 1 minuto)..."):
                    df_raw_ltsa = extraer_df_desde_excel(file_ltsa, 'REPORTE')
                    cols_ltsa_gsheets = ['ORIGEN', 'PUNTO DE VENTA', 'OPERACIÓN', 'CÓDIGO', 'CIUDAD ORIGEN', 'CIUDAD DESTINO', 'CEDULA', 'NOMBRE', 'TIPO DE VEHICULO', 'PLACA', 'ESTADO', 'CONCEPTO TARIFA', 'TARIFA', 'Días/Horas', 'TOTAL', 'OBSERVACION OP', 'OBSERVACIONES ÉXITO', 'HORAS RVS', 'DIFERENCIA', 'TOTAL FACTURAR', 'OBSERVACIÓN', 'OPERADOR', 'PERIODO']
                    
                    bulk_data_ltsa = preparar_df_para_sheets(df_raw_ltsa, cols_ltsa_gsheets, per_ltsa)
                    param_replace = per_ltsa_clean if reemplazar_ltsa else None
                    res_ltsa = subir_bulk_a_sheets(GAS_URL, "RENTABILIDAD_HISTORICA", bulk_data_ltsa, clear_first=False, replace_period=param_replace)
                    
                    if res_ltsa.get('status') == 'success':
                        st.success("✅ ¡Datos inyectados exitosamente! Refrescando tablero automáticamente...")
                        cargar_datos.clear()
                        st.cache_data.clear()
                        st.session_state.pop('up_ltsa', None)
                        st.rerun()
                    else:
                        st.error(f"Error al subir: {res_ltsa.get('message')}")

        st.divider()
        if corte_evaluado_str == "GLOBAL":
            st.markdown(f"#### 📊 2. Tablero de Resultados (HISTÓRICO GLOBAL ACUMULADO)")
        else:
            st.markdown(f"#### 📊 2. Tablero de Resultados (Corte: {corte_evaluado_str})")
            
        if not df_rent_hist.empty:
            procesar_rentabilidad_db(df_rent_hist, "LTSA", df_pagos_reales_rent, corte_evaluado_str, total_nomina_bd_rent)
        else:
            st.info("No hay datos en la base de datos para mostrar. Sube el validador en el paso 1.")

    # ------------------ POLLOS Y PANADERÍA (CON DOBLE PESTAÑA) ------------------
    with sub_pollos:
        st.markdown("#### 📤 1. Cargar Validador a la Base de Datos")
        c3, c4 = st.columns([1, 2])
        per_pollos = c3.text_input("Digita el PERIODO a inyectar:", value=corte_seleccionado if corte_a_evaluar == "🌐 TODOS LOS CORTES (GLOBAL)" else corte_evaluado_str, key="per_pollos")
        file_pollos = c4.file_uploader("Archivo Validador Excel (POLLOS)", type=["xlsx", "xls"], key="up_pollos")
        
        if file_pollos and per_pollos:
            per_pollos_clean = per_pollos.strip().upper()
            reemplazar_pollos = False
            
            if per_pollos_clean in periodos_pollos:
                st.warning(f"⚠️ ¡Atención! Ya existen datos registrados para el corte **{per_pollos_clean}**.")
                reemplazar_pollos = st.checkbox(f"Sobrescribir datos (Se borrarán los registros anteriores de {per_pollos_clean} para evitar duplicados)", value=True, key="chk_pollos")
            
            btn_disabled_p = (per_pollos_clean in periodos_pollos) and not reemplazar_pollos

            if st.button("🚀 Enviar a Google Sheets (POLLOS)", use_container_width=True, type="primary", disabled=btn_disabled_p):
                with st.spinner("Procesando pestañas COBRO y PAGO. Subiendo a la nube (Puede tardar hasta 1 minuto)..."):
                    df_raw_pollos = extraer_df_desde_excel(file_pollos, 'COBRO')
                    cols_pollos_gsheets = ['ORIGEN', 'PUNTO DE VENTA', 'OPERACIÓN', 'CÓDIGO', 'CIUDAD ORIGEN', 'CIUDAD DESTINO', 'CEDULA', 'NOMBRE', 'TIPO DE VEHICULO', 'PLACA', 'ESTADO', 'CONCEPTO TARIFA', 'TARIFA', 'Días/Horas', 'TOTAL', 'OBSERVACION OP', 'OBSERVACIONES ÉXITO', 'HORAS RVS', 'DIFERENCIA', 'TOTAL FACTURAR', 'OBSERVACIÓN', 'OPERADOR', 'PERIODO']
                    
                    bulk_data_pollos = preparar_df_para_sheets(df_raw_pollos, cols_pollos_gsheets, per_pollos)
                    param_replace_p = per_pollos_clean if reemplazar_pollos else None
                    res_pollos = subir_bulk_a_sheets(GAS_URL, "RENTABILIDAD_POLLOS_PANADERIA", bulk_data_pollos, clear_first=False, replace_period=param_replace_p)
                    
                    try:
                        df_raw_pago = extraer_df_desde_excel(file_pollos, 'PAGO')
                        cols_pago_gsheets = ['OPERACIÓN', 'CEDULA', 'EMPLEADO', 'PLACA', 'VALOR PAGO', 'PERIODO']
                        bulk_data_pago = preparar_df_para_sheets(df_raw_pago, cols_pago_gsheets, per_pollos)
                        res_pago = subir_bulk_a_sheets(GAS_URL, "RENTABILIDAD_POLLOS_PAGO", bulk_data_pago, clear_first=False, replace_period=param_replace_p)
                    except Exception as e:
                        res_pago = {'status': 'error', 'message': str(e)}

                    if res_pollos.get('status') == 'success' and res_pago.get('status') == 'success':
                        st.success("✅ ¡Datos de Facturación y Pago inyectados exitosamente! Refrescando tablero automáticamente...")
                        cargar_datos.clear()
                        st.cache_data.clear()
                        st.session_state.pop('up_pollos', None)
                        st.rerun()
                    else:
                        st.error(f"Error al subir COBRO: {res_pollos.get('message')} | PAGO: {res_pago.get('message')}")

        st.divider()
        if corte_evaluado_str == "GLOBAL":
            st.markdown(f"#### 📊 2. Tablero de Resultados (HISTÓRICO GLOBAL ACUMULADO)")
        else:
            st.markdown(f"#### 📊 2. Tablero de Resultados (Corte: {corte_evaluado_str})")
            
        if not df_rent_pollos.empty:
            if not df_rent_pollos_pago.empty:
                if corte_evaluado_str != "GLOBAL":
                    col_per_pago = obtener_nombre_columna(df_rent_pollos_pago, ['PERIODO', 'CORTE'])
                    if col_per_pago:
                        df_pago_filtered = df_rent_pollos_pago[df_rent_pollos_pago[col_per_pago].astype(str).str.strip().str.upper() == str(corte_evaluado_str).strip().upper()]
                    else:
                        df_pago_filtered = df_rent_pollos_pago
                else:
                    df_pago_filtered = df_rent_pollos_pago

                col_ced_p = obtener_nombre_columna(df_pago_filtered, ['CÉDULA', 'CEDULA', 'CC', 'IDENTIFICACION'])
                col_val_p = obtener_nombre_columna(df_pago_filtered, ['VALOR PAGO', 'PAGO', 'VALOR'])
                col_nom_p = obtener_nombre_columna(df_pago_filtered, ['EMPLEADO', 'NOMBRE', 'NOMBRES'])
                
                if col_ced_p and col_val_p:
                    df_pago_filtered['_cedula_clean'] = df_pago_filtered[col_ced_p].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.strip()
                    df_pago_filtered['_valor_pagar_num'] = pd.to_numeric(df_pago_filtered[col_val_p], errors='coerce').fillna(0)
                    
                    df_pagos_agrupados_pollos = df_pago_filtered.groupby('_cedula_clean').agg(
                        NOMBRE_EMPLEADO=(col_nom_p if col_nom_p else col_ced_p, 'first'),
                        VALOR_PAGADO_NETO=('_valor_pagar_num', 'sum')
                    ).reset_index()
                    df_pagos_reales_pollos = df_pagos_agrupados_pollos[df_pagos_agrupados_pollos['_cedula_clean'] != '']
                    total_nomina_pollos = df_pagos_reales_pollos['VALOR_PAGADO_NETO'].sum()
                else:
                    df_pagos_reales_pollos = pd.DataFrame(columns=['_cedula_clean', 'VALOR_PAGADO_NETO', 'NOMBRE_EMPLEADO'])
                    total_nomina_pollos = 0
            else:
                df_pagos_reales_pollos = pd.DataFrame(columns=['_cedula_clean', 'VALOR_PAGADO_NETO', 'NOMBRE_EMPLEADO'])
                total_nomina_pollos = 0
                
            procesar_rentabilidad_db(df_rent_pollos, "Pollos y Panadería", df_pagos_reales_pollos, corte_evaluado_str, total_nomina_pollos, costo_label="Costo Nómina Real (Pestaña PAGO Pollos)")
        else:
            st.info("No hay datos en la base de datos para mostrar. Sube el validador en el paso 1.")

    # ------------------ PERSONAL DIRECTO ------------------
    with sub_directo:
        st.markdown("#### 📤 1. Cargar Lista de Personal Directo a la Base de Datos")
        c5, c6 = st.columns([1, 2])
        limpiar_directo = c5.checkbox("Reemplazar lista completa (Borrar anteriores)", value=False)
        file_directo = c6.file_uploader("Archivo Excel (Personal Directo)", type=["xlsx", "xls"], key="up_directo")
        
        if file_directo:
            if st.button("🚀 Enviar a Google Sheets (PERSONAL DIRECTO)", use_container_width=True, type="primary"):
                with st.spinner("Preparando archivo y subiendo a la nube (Puede tardar hasta 1 minuto)..."):
                    df_raw_directo = extraer_df_desde_excel(file_directo)
                    cols_directo_gsheets = ['NOMBRE COMPLETO', 'CÓDIGO INGRESO', 'IDENTIFICACIÓN', 'ACTIVO', 'FECHA DE INGRESO', 'NÓMINA', 'CENTRO DE COSTOS', 'NOMBRE CENTRO COSTO', 'FECHA DE RETIRO']
                    
                    bulk_data_directo = preparar_df_para_sheets(df_raw_directo, cols_directo_gsheets, periodo="")
                    res_directo = subir_bulk_a_sheets(GAS_URL, "RENTABILIDAD_PERSONAL_DIRECTO", bulk_data_directo, clear_first=limpiar_directo)
                    
                    if res_directo.get('status') == 'success':
                        st.success("✅ ¡Lista inyectada exitosamente! Refrescando tablero automáticamente...")
                        cargar_datos.clear()
                        st.cache_data.clear()
                        st.session_state.pop('up_directo', None)
                        st.rerun()
                    else:
                        st.error(f"Error al subir: {res_directo.get('message')}")

        st.divider()
        if corte_evaluado_str == "GLOBAL":
            st.markdown(f"#### 📊 2. Tablero Operativo: Directos vs Terceros (HISTÓRICO GLOBAL)")
        else:
            st.markdown(f"#### 📊 2. Tablero Operativo: Directos vs Terceros (Corte: {corte_evaluado_str})")
            
        if not df_rent_directo.empty:
            col_id = obtener_nombre_columna(df_rent_directo, ['IDENTIFICACIÓN', 'IDENTIFICACION', 'CÉDULA', 'CEDULA', 'CC', 'DOCUMENTO'])
            
            if col_id:
                cedulas_planta = df_rent_directo[col_id].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.strip().tolist()
                st.success(f"✅ Hay **{len(cedulas_planta)}** empleados marcados como Personal Directo en la base de datos.")
                
                if 'df_raw_LTSA' in st.session_state:
                    df_ltsa = st.session_state['df_raw_LTSA'].copy()
                    col_ced_ltsa = obtener_nombre_columna(df_ltsa, ['CÉDULA', 'CEDULA', 'CC', 'IDENTIFICACION'])
                    col_tot_ltsa = obtener_nombre_columna(df_ltsa, ['TOTAL', 'VALOR TOTAL', 'TOTAL FACTURAR'])
                    
                    if col_ced_ltsa and col_tot_ltsa:
                        df_ltsa['_cedula_clean'] = df_ltsa[col_ced_ltsa].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\D', '', regex=True).str.strip()
                        df_ltsa = df_ltsa[df_ltsa['_cedula_clean'] != '']
                        df_ltsa['TOTAL_COBRADO'] = pd.to_numeric(df_ltsa[col_tot_ltsa], errors='coerce').fillna(0)
                        
                        ltsa_grouped = df_ltsa.groupby('_cedula_clean')['TOTAL_COBRADO'].sum().reset_index()
                        
                        df_cruce_directo = pd.merge(ltsa_grouped, df_pagos_reales_rent, on='_cedula_clean', how='outer')
                        df_cruce_directo['TOTAL_COBRADO'] = df_cruce_directo['TOTAL_COBRADO'].fillna(0)
                        df_cruce_directo['VALOR_PAGADO_NETO'] = df_cruce_directo['VALOR_PAGADO_NETO'].fillna(0)
                        
                        df_cruce_directo['Tipo_Contratacion'] = df_cruce_directo['_cedula_clean'].isin(cedulas_planta).map({True: 'Directo (Planta)', False: 'Tercero'})
                        df_cruce_directo['UTILIDAD_REAL'] = df_cruce_directo['TOTAL_COBRADO'] - df_cruce_directo['VALOR_PAGADO_NETO']
                        
                        t_fact = float(df_cruce_directo['TOTAL_COBRADO'].sum())
                        t_cost = float(df_cruce_directo['VALOR_PAGADO_NETO'].sum())
                        t_util = float(df_cruce_directo['UTILIDAD_REAL'].sum())
                        
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Facturación Total (LTSA)", f"${t_fact:,.0f}")
                        c2.metric("Costo Nómina Real (Pagos SERGEM)", f"${t_cost:,.0f}")
                        c3.metric("Utilidad Real Total", f"${t_util:,.0f}")
                        
                        df_agrupado = df_cruce_directo.groupby('Tipo_Contratacion').agg({
                            'TOTAL_COBRADO': 'sum', 'VALOR_PAGADO_NETO': 'sum', 'UTILIDAD_REAL': 'sum'
                        }).reset_index()
                        df_agrupado['Margen_Real'] = (df_agrupado['UTILIDAD_REAL'] / df_agrupado['TOTAL_COBRADO'].replace(0, 1)).fillna(0)
                        
                        col_graf1, col_graf2 = st.columns(2)
                        with col_graf1:
                            fig1 = px.pie(df_agrupado, names='Tipo_Contratacion', values='TOTAL_COBRADO', title="Distribución de Facturación", hole=0.4, color='Tipo_Contratacion', color_discrete_map={'Directo (Planta)':'#15803d', 'Tercero':'#E3000F'})
                            st.plotly_chart(fig1, use_container_width=True)
                            
                        with col_graf2:
                            st.markdown("<br><br>", unsafe_allow_html=True)
                            st.dataframe(df_agrupado.style.format({'TOTAL_COBRADO': '${:,.0f}', 'VALOR_PAGADO_NETO': '${:,.0f}', 'UTILIDAD_REAL': '${:,.0f}', 'Margen_Real': '{:.1%}'}), hide_index=True, use_container_width=True)
                    else:
                        st.warning("⚠️ No se lograron cruzar los datos porque faltan las columnas de Cédula o Total Facturar en LTSA.")
                else:
                    st.info("💡 **Recordatorio:** Para ver el cruce Directo vs Terceros, asegúrate de haber procesado primero el reporte de LTSA en la pestaña anterior.")
            else:
                st.error("No se encontró la columna de IDENTIFICACIÓN o CÉDULA en la hoja de Personal Directo.")
        else:
            st.info("No hay datos cargados de Personal Directo. Sube el archivo en el paso 1.")
