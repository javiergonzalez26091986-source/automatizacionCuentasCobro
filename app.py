import streamlit as st
import pandas as pd
import requests
import io
import zipfile
import os
from datetime import datetime, timezone, timedelta
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage

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
    
    .instrucciones {
        background-color: #f8fafc;
        border-left: 4px solid #E3000F;
        padding: 1rem;
        border-radius: 4px;
        margin-bottom: 2rem;
    }
    .metric-box {
        background-color: #f1f5f9; padding: 15px; border-radius: 8px; text-align: center; border: 1px solid #e2e8f0;
    }
    </style>
    """, unsafe_allow_html=True)

GAS_URL = "https://script.google.com/macros/s/AKfycbyqJtrmVdNT1rxTobg6q_WoJCwMpp40hdIzJeEm4dKNLBgDVxwEY95T0EIoBu_qo8FB/exec"

def obtener_fecha_actual():
    meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    zona_colombia = timezone(timedelta(hours=-5))
    hoy = datetime.now(zona_colombia)
    return f"{hoy.day} DE {meses[hoy.month - 1]} DE {hoy.year}"

@st.cache_data(ttl=600) 
def cargar_datos(url):
    try:
        response = requests.get(url)
        return response.json()
    except Exception as e:
        return None

def limpiar_dinero(val):
    if pd.isna(val) or val == "": return 0.0
    s = str(val).upper().replace('$', '').replace(',', '').replace('.', '').replace(' ', '')
    try: return float(s)
    except: return 0.0

# ==============================================================================
# GENERACIÓN DE ARCHIVO PLANO BANCARIO (ESTRUCTURA EXACTA TXT)
# ==============================================================================
def generar_txt_banco(df_banco):
    lineas = []
    zona_colombia = timezone(timedelta(hours=-5))
    hoy = datetime.now(zona_colombia)
    fecha_ymd = hoy.strftime("%Y%m%d")
    fecha_dmy = hoy.strftime("%d/%m/%Y").ljust(13)
    
    num_registros = len(df_banco)
    suma_total = int(df_banco['VALOR_NETO_A_PAGAR'].sum() * 100)
    
    # 1. Encabezado (Tipo 1)
    header = (
        f"1000000900561833I               225NOMINA    "
        f"{fecha_ymd}AA{fecha_ymd}"
        f"{num_registros:06d}"
        f"00000000000000000"
        f"{suma_total:017d}"
        f"81016173001D"
    )
    lineas.append(header)
    
    # 2. Detalles (Tipo 6)
    for _, row in df_banco.iterrows():
        cedula = str(row['NIT_BENEFICIARIO']).replace('.', '').replace(' ', '').replace(',', '')
        cedula = cedula.zfill(15)[:15]
        
        nombre = str(row['NOMBRE_BENEFICIARIO']).upper().ljust(30)[:30]
        
        banco = "005600078" # Código de enrutamiento bancario base
        
        cuenta = str(row['NUMERO_CUENTA']).replace("'", "").strip().ljust(18)[:18]
        
        tipo_cta = "37" if "AHORRO" in str(row['TIPO_CUENTA']).upper() else "27"
        
        valor = int(row['VALOR_NETO_A_PAGAR'] * 100)
        valor_str = f"{valor:017d}"
        
        filler = "00000                                                                                               000000000000000                           "
        
        linea = (
            f"6{cedula}{nombre}{banco}{cuenta}{tipo_cta}{valor_str}"
            f"{fecha_ymd}NOMINA   {fecha_dmy}{filler}"
        )
        lineas.append(linea)
        
    return "\n".join(lineas)

# ==============================================================================
# LÓGICA DE PDFS 
# ==============================================================================
def agregar_pagina_pdf_cuenta_cobro(pdf, datos):
    pdf.add_page()
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("helvetica", "", 11)
    
    pdf.cell(0, 6, f"{datos['ciudad']} {datos['fecha_emision']}".upper(), 0, 1)
    pdf.ln(12)
    
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 6, "SERGEM MENSAJERIA S.A.S.", 0, 1, 'C')
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "NIT: 900.561.833-1", 0, 1, 'C')
    pdf.ln(8)
    
    pdf.cell(0, 6, "DEBE A:", 0, 1, 'C')
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 6, str(datos['nombre_titular']).upper(), 0, 1, 'C')
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"C.C {datos['cedula_titular']}", 0, 1, 'C')
    pdf.ln(12)

    pdf.set_font("helvetica", "B", 11)
    pdf.cell(80, 6, "VALOR BASE:", 0, 0)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"$ {datos['ingreso_base']:,.0f}", 0, 1)

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
    pdf.cell(80, 8, "VALOR TOTAL NETO A PAGAR:", 0, 0)
    pdf.cell(0, 8, f"$ {datos['neto_pagar']:,.0f}", 0, 1)
    pdf.ln(8)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(25, 6, "CONCEPTO:", 0, 0)
    pdf.set_font("helvetica", "", 11)
    concepto = f"SERVICIO PRESTADO EN EL CORTE DE {datos['corte_fechas']}, POR EL CONDUCTOR {datos['nombre_conductor']} CÉDULA {datos['cedula_conductor']}."
    pdf.multi_cell(0, 6, concepto)
    pdf.ln(10)
    
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "Autorizo me sea consignado en:", 0, 1)
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(0, 6, f"CUENTA # {datos['num_cuenta']}", 0, 1)
    pdf.cell(0, 6, f"{datos['tipo_cuenta'].upper()}", 0, 1)
    pdf.cell(0, 6, f"NOMBRE DEL BANCO: {datos['banco'].upper()}", 0, 1)
    pdf.cell(0, 6, f"NOMBRE TITULAR CUENTA: {datos['nombre_titular']}", 0, 1)
    pdf.cell(0, 6, f"CÉDULA TITULAR CUENTA: {datos['cedula_titular']}", 0, 1)
    pdf.ln(15)
    
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "Atentamente,", 0, 1)
    pdf.ln(12)
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(80, 5, str(datos['nombre_titular']).upper(), "T", 1, "L")
    pdf.set_font("helvetica", "", 11)
    pdf.cell(80, 5, f"C.C {datos['cedula_titular']}", 0, 1, "L")


def agregar_pagina_pdf_doc_equivalente(pdf, datos):
    pdf.add_page()
    try:
        if os.path.exists('sergemLogo_2.png'): pdf.image('sergemLogo_2.png', 10, 8, w=45)
        elif os.path.exists('sergemLogo.png'): pdf.image('sergemLogo.png', 10, 8, w=45)
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
    pdf.cell(100, 6, datos['nombre_titular'][:45], 1)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(20, 6, "C.C / NIT:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(0, 6, datos['cedula_titular'], 1, 1)
    
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(25, 6, "Ciudad:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(100, 6, datos['ciudad'], 1)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(20, 6, "Conductor:", 1)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(0, 6, datos['nombre_conductor'][:20], 1, 1)
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
    pdf.cell(10, 6, "1", 1, 0, 'C')
    pdf.cell(90, 6, f"Servicio de mensajería - Corte: {datos['corte_fechas']}", 1, 0, 'L')
    pdf.cell(20, 6, "1", 1, 0, 'C')
    pdf.cell(35, 6, f"$ {datos['ingreso_base']:,.0f}", 1, 0, 'R')
    pdf.cell(0, 6, f"$ {datos['ingreso_base']:,.0f}", 1, 1, 'R')

    if datos.get('fuera_perimetro', 0) > 0:
        pdf.cell(10, 6, "2", 1, 0, 'C')
        pdf.cell(90, 6, "Servicios Fuera de Perímetro", 1, 0, 'L')
        pdf.cell(20, 6, "1", 1, 0, 'C')
        pdf.cell(35, 6, f"$ {datos['fuera_perimetro']:,.0f}", 1, 0, 'R')
        pdf.cell(0, 6, f"$ {datos['fuera_perimetro']:,.0f}", 1, 1, 'R')

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
    pdf.set_text_color(227, 0, 15)
    pdf.cell(35, 8, "NETO A PAGAR:", 1, 0, 'R', fill=True)
    pdf.set_font('helvetica', 'B', 11)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 8, f"$ {datos['neto_pagar']:,.0f}", 1, 1, 'R', fill=True)

    pdf.ln(15)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(80, 5, "________________________________________________", 0, 1)
    pdf.cell(80, 5, "FIRMA PRESTADOR DEL SERVICIO", 0, 1)
    pdf.cell(80, 5, f"C.C. / NIT: {datos['cedula_titular']}", 0, 1)
    pdf.cell(80, 5, f"NOMBRE: {datos['nombre_titular']}", 0, 1)

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
        if os.path.exists('sergemLogo_2.png'):
            img = XLImage('sergemLogo_2.png')
            img.width = 150; img.height = 60
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

    ws['B14'] = "Nombre:"; ws['B14'].font = bold_font; ws['C14'] = datos['nombre_titular']; ws.merge_cells('C14:E14')
    ws['G14'] = "C.C / NIT:"; ws['G14'].font = bold_font; ws['H14'] = datos['cedula_titular']
    ws['G16'] = "Conductor:"; ws['G16'].font = bold_font; ws['H16'] = datos['nombre_conductor']

    fila = 18
    for i, h in enumerate(["Ítem", "Concepto", "Cantidad", "V. Unitario", "V. Total"]):
        c = ['B', 'C', 'F', 'G', 'H'][i] + str(fila)
        ws[c] = h; ws[c].font = header_font; ws[c].fill = header_fill; ws[c].alignment = center_align; ws[c].border = border_thin
    ws.merge_cells(f'C{fila}:E{fila}')
    
    fila += 1
    ws[f'B{fila}'] = 1; ws[f'C{fila}'] = f"Servicio de mensajería - Corte: {datos['corte_fechas']}"; ws.merge_cells(f'C{fila}:E{fila}')
    ws[f'F{fila}'] = 1; ws[f'G{fila}'] = datos['ingreso_base']; ws[f'H{fila}'] = datos['ingreso_base']
    for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']: ws[f'{c}{fila}'].border = border_thin
    ws[f'G{fila}'].number_format = '"$"#,##0'; ws[f'H{fila}'].number_format = '"$"#,##0'

    if datos.get('fuera_perimetro', 0) > 0:
        fila += 1
        ws[f'B{fila}'] = 2; ws[f'C{fila}'] = "Servicios Fuera de Perímetro"; ws.merge_cells(f'C{fila}:E{fila}')
        ws[f'F{fila}'] = 1; ws[f'G{fila}'] = datos['fuera_perimetro']; ws[f'H{fila}'] = datos['fuera_perimetro']
        for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']: ws[f'{c}{fila}'].border = border_thin
        ws[f'G{fila}'].number_format = '"$"#,##0'; ws[f'H{fila}'].number_format = '"$"#,##0'
        
    fila += 2
    totales = [("SUBTOTAL:", datos['ingreso_bruto_total']), ("IVA (19%):", ""), ("RETEIVA:", ""), 
               ("RTE FTE (1%):", -datos['retefuente'] if datos['retefuente']>0 else 0),
               ("RETEICA (1%):", -datos['ica'] if datos['ica']>0 else 0), ("NETO A PAGAR:", datos['neto_pagar'])]
    
    fila_firma = fila + 1
    for label, valor in totales:
        ws[f'G{fila}'] = label; ws[f'H{fila}'] = valor
        ws[f'G{fila}'].font = bold_font; ws[f'G{fila}'].alignment = right_align
        ws[f'G{fila}'].border = border_thin; ws[f'H{fila}'].border = border_thin
        if valor != "": ws[f'H{fila}'].number_format = '"$"#,##0'
        if label == "NETO A PAGAR:":
            ws[f'G{fila}'].font = Font(bold=True, color="E3000F"); ws[f'H{fila}'].font = Font(bold=True, size=12)
            ws[f'H{fila}'].fill = PatternFill(start_color="F4F6F9", end_color="F4F6F9", fill_type="solid")
        fila += 1

    ws[f'B{fila_firma}'] = "________________________________________________"
    ws[f'B{fila_firma+1}'] = "FIRMA PRESTADOR DEL SERVICIO"
    ws[f'B{fila_firma+1}'].font = bold_font
    ws[f'B{fila_firma+2}'] = f"C.C. / NIT: {datos['cedula_titular']}"
    ws[f'B{fila_firma+3}'] = f"NOMBRE: {datos['nombre_titular']}"

    ws.column_dimensions['B'].width = 16; ws.column_dimensions['C'].width = 12; ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12; ws.column_dimensions['F'].width = 10; ws.column_dimensions['G'].width = 22; ws.column_dimensions['H'].width = 22

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT; ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.5; ws.page_margins.right = 0.5; ws.page_margins.top = 0.5; ws.page_margins.bottom = 0.5

# ==============================================================================
# PROCESO MATEMÁTICO PRINCIPAL 
# ==============================================================================
def calcular_valores_conductor(row, df_fuera, corte_seleccionado):
    ingreso_neto_esperado = limpiar_dinero(row.get('TOTAL A PAGAR', 0))
    if ingreso_neto_esperado <= 0: return None
    
    nombre_conductor = str(row.get('CONDUCTOR', '')).upper().strip()
    ciudad = str(row.get('CIUDAD', '')).upper().strip()
    
    fuera_perimetro_neto = 0.0
    if not df_fuera.empty and 'TOTAL' in df_fuera.columns and 'CONDUCTOR' in df_fuera.columns:
        match = df_fuera[(df_fuera['CORTE'].astype(str).str.strip() == corte_seleccionado) & (df_fuera['CONDUCTOR'].astype(str).str.upper().str.strip() == nombre_conductor)]
        fuera_perimetro_neto = sum(match['TOTAL'].apply(limpiar_dinero))
    
    total_neto_esperado = ingreso_neto_esperado + fuera_perimetro_neto
    
    porcentaje_retefuente = 0.01
    porcentaje_ica = 0.01 if ciudad == 'CALI' else 0.0
    tasa_total_impuestos = porcentaje_retefuente + porcentaje_ica
    
    ingreso_bruto_total = round(total_neto_esperado / (1 - tasa_total_impuestos))
    retefuente = round(ingreso_bruto_total * porcentaje_retefuente)
    ica = round(ingreso_bruto_total * porcentaje_ica)
    
    fuera_perimetro_bruto = round(fuera_perimetro_neto / (1 - tasa_total_impuestos)) if fuera_perimetro_neto > 0 else 0.0
        
    ingreso_base_bruta = ingreso_bruto_total - fuera_perimetro_bruto
    neto_a_pagar = ingreso_bruto_total - retefuente - ica
    
    return {
        'nombre_conductor': nombre_conductor,
        'ciudad': ciudad,
        'cedula_conductor': str(row.get('CEDULA', '')).strip(),
        'nombre_titular': str(row.get('A NOMBRE DE QUIEN HACE CUENTA DE COBRO', row.get('NOMBRE TITULAR CUENTA BANCARIA', 'S/N'))).strip(),
        'cedula_titular': str(row.get('CÉDULA DE CUENTA DE COBRO', row.get('CÉDULA TITULAR', ''))).strip(),
        'banco': str(row.get('BANCO', '')).strip(),
        'tipo_cuenta': str(row.get('TIPO CUENTA', '')).strip(),
        'num_cuenta': str(row.get('NO. CUENTA', '')).strip(),
        'ingreso_base': ingreso_base_bruta,
        'fuera_perimetro': fuera_perimetro_bruto,
        'ingreso_bruto_total': ingreso_bruto_total,
        'retefuente': retefuente,
        'ica': ica,
        'neto_pagar': neto_a_pagar
    }

# ==============================================================================
# INTERFAZ DE USUARIO
# ==============================================================================
col1, col2 = st.columns([1.5, 3.5])
with col1:
    try: st.image("sergemLogo_2.png", width=250)
    except: pass
with col2:
    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
    st.title("Generador Automático de Documentos")
    st.markdown("**SERGEM Mensajería S.A.S.**")

if st.button("🔄 Sincronizar Base de Datos", key="btn_sync", type="secondary"):
    st.cache_data.clear()
    st.rerun()

data_cruda = cargar_datos(GAS_URL)
if not data_cruda:
    st.error("Error conectando a Google Sheets. Revise la URL.")
    st.stop()

df_pagos_completo = pd.DataFrame(data_cruda.get('pagos', []))
df_fuera = pd.DataFrame(data_cruda.get('fueras_perimetro', []))

if df_pagos_completo.empty:
    st.warning("No se encontraron datos en Google Sheets.")
    st.stop()

df_pagos_completo.columns = df_pagos_completo.columns.str.strip().str.upper()
if not df_fuera.empty: df_fuera.columns = df_fuera.columns.str.strip().str.upper()

cortes_disponibles = [c for c in df_pagos_completo['CORTE'].unique() if str(c).strip() != "" and str(c).lower() != "nan"]
corte_seleccionado = st.selectbox("📅 Seleccione el Corte a procesar:", cortes_disponibles)

st.divider()
modo_trabajo = st.radio("⚙️ Modo de trabajo:", 
                        ["🗂️ Generación Masiva (Todos los conductores del corte)", 
                         "👤 Vista Previa Individual (Revisar un conductor específico)"], horizontal=True)

df_pagos_corte = df_pagos_completo[df_pagos_completo['CORTE'] == corte_seleccionado]

# ==============================================================================
# MODO INDIVIDUAL
# ==============================================================================
if "Individual" in modo_trabajo:
    lista_conductores = sorted(df_pagos_corte['CONDUCTOR'].dropna().unique().tolist())
    cond_seleccionado = st.selectbox("Busque o seleccione el conductor:", lista_conductores)
    
    if st.button("🔍 Calcular y Previsualizar"):
        row_conductor = df_pagos_corte[df_pagos_corte['CONDUCTOR'] == cond_seleccionado].iloc[0]
        calculos = calcular_valores_conductor(row_conductor, df_fuera, corte_seleccionado)
        
        if not calculos:
            st.warning("Este conductor tiene saldo neto en $0 para este corte.")
        else:
            st.markdown(f"### Resumen Financiero: {calculos['nombre_titular']}")
            cA, cB, cC, cD = st.columns(4)
            cA.markdown(f"<div class='metric-box'><b>BASE BRUTA</b><br>${calculos['ingreso_base']:,.0f}</div>", unsafe_allow_html=True)
            cB.markdown(f"<div class='metric-box'><b>RETEFUENTE (1%)</b><br>${calculos['retefuente']:,.0f}</div>", unsafe_allow_html=True)
            cC.markdown(f"<div class='metric-box'><b>RETEICA</b><br>${calculos['ica']:,.0f}</div>", unsafe_allow_html=True)
            cD.markdown(f"<div class='metric-box' style='background-color:#E3000F; color:white;'><b>NETO EXACTO</b><br>${calculos['neto_pagar']:,.0f}</div>", unsafe_allow_html=True)
            
            st.info("💡 El sistema ha inflado la base bruta automáticamente para que al descontar los impuestos, el resultado sea exactamente el que usted parametrizó en el Excel de origen.")
            
            datos_doc = calculos.copy()
            datos_doc.update({'id': "PREVIEW", 'fecha_emision': obtener_fecha_actual(), 'corte_fechas': corte_seleccionado})
            
            pdf_ct = FPDF(); agregar_pagina_pdf_cuenta_cobro(pdf_ct, datos_doc)
            pdf_eq = FPDF(); agregar_pagina_pdf_doc_equivalente(pdf_eq, datos_doc)
            
            out_ct = pdf_ct.output()
            pdf_ct_bytes = out_ct.encode('latin-1') if isinstance(out_ct, str) else bytes(out_ct)
            
            out_eq = pdf_eq.output()
            pdf_eq_bytes = out_eq.encode('latin-1') if isinstance(out_eq, str) else bytes(out_eq)
            
            colBtn1, colBtn2 = st.columns(2)
            colBtn1.download_button("📥 Descargar Cuenta de Cobro (PDF)", data=pdf_ct_bytes, file_name=f"Cuenta_{cond_seleccionado}.pdf", mime="application/pdf", use_container_width=True)
            colBtn2.download_button("📥 Descargar Doc. Equivalente (PDF)", data=pdf_eq_bytes, file_name=f"DocEq_{cond_seleccionado}.pdf", mime="application/pdf", use_container_width=True)

# ==============================================================================
# MODO MASIVO (LOTE)
# ==============================================================================
else:
    if st.button("🚀 Procesar Lote y Generar Paquete ZIP", use_container_width=True, type="primary"):
        mensaje_carga = st.info(f"📥 Procesando la información y unificando archivos para impresión...")
        
        try:
            pagos_procesados_banco = []
            ignorados = 0
            zip_buffer = io.BytesIO()
            fecha_actual = obtener_fecha_actual() 
            
            pdf_maestro_cuentas = FPDF()
            pdf_maestro_equivalentes = FPDF()
            wb_maestro_equivalentes = openpyxl.Workbook()
            wb_maestro_equivalentes.remove(wb_maestro_equivalentes.active)
            
            contador = 1
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for index, row in df_pagos_corte.iterrows():
                    calculos = calcular_valores_conductor(row, df_fuera, corte_seleccionado)
                    
                    if not calculos:
                        ignorados += 1
                        continue
                    
                    datos_doc = calculos.copy()
                    datos_doc.update({'id': str(contador).zfill(3), 'fecha_emision': fecha_actual, 'corte_fechas': corte_seleccionado})

                    agregar_pagina_pdf_cuenta_cobro(pdf_maestro_cuentas, datos_doc)
                    agregar_pagina_pdf_doc_equivalente(pdf_maestro_equivalentes, datos_doc)

                    nombre_pestana = f"{contador}_{datos_doc['nombre_titular'][:20]}".replace(":", "").replace("/", "-")
                    ws_nuevo = wb_maestro_equivalentes.create_sheet(title=nombre_pestana)
                    construir_hoja_documento_equivalente_excel(ws_nuevo, datos_doc)

                    pagos_procesados_banco.append({
                        'NIT_BENEFICIARIO': datos_doc['cedula_titular'],
                        'NOMBRE_BENEFICIARIO': datos_doc['nombre_titular'],
                        'BANCO_DESTINO': datos_doc['banco'],
                        'TIPO_CUENTA': datos_doc['tipo_cuenta'],
                        'NUMERO_CUENTA': datos_doc['num_cuenta'],
                        'VALOR_NETO_A_PAGAR': datos_doc['neto_pagar'],
                        'FECHA_PAGO': datetime.now(timezone(timedelta(hours=-5))).strftime("%Y/%m/%d"),
                        'CONCEPTO': 'NOMINA'
                    })
                    contador += 1
                
                out_maestro_cuentas = pdf_maestro_cuentas.output()
                pdf_maestro_ct_bytes = out_maestro_cuentas.encode('latin-1') if isinstance(out_maestro_cuentas, str) else bytes(out_maestro_cuentas)
                zip_file.writestr("1_IMPRESION_MASIVA_Cuentas_de_Cobro.pdf", pdf_maestro_ct_bytes)
                
                out_maestro_eq = pdf_maestro_equivalentes.output()
                pdf_maestro_eq_bytes = out_maestro_eq.encode('latin-1') if isinstance(out_maestro_eq, str) else bytes(out_maestro_eq)
                zip_file.writestr("2_IMPRESION_MASIVA_Documentos_Equivalentes.pdf", pdf_maestro_eq_bytes)
                
                excel_maestro_io = io.BytesIO()
                wb_maestro_equivalentes.save(excel_maestro_io)
                excel_maestro_io.seek(0)
                zip_file.writestr("3_ARCHIVO_Documentos_Equivalentes_Pestañas.xlsx", excel_maestro_io.read())
                
                # --- NUEVA INTEGRACIÓN: ARCHIVO PLANO EN TXT Y FORMATO EN EXCEL ---
                df_banco = pd.DataFrame(pagos_procesados_banco)
                
                # 1. Archivo Plano (.txt) estructurado para subir al portal
                texto_banco = generar_txt_banco(df_banco)
                zip_file.writestr(f"4_PLANO_BANCARIO_{corte_seleccionado.replace(' ', '_').replace('/', '-')}.txt", texto_banco)

                # 2. Archivo en Excel para revisión visual de Yesenia
                excel_banco_io = io.BytesIO()
                df_banco.to_excel(excel_banco_io, index=False, sheet_name="PLANO_BANCO")
                excel_banco_io.seek(0)
                zip_file.writestr(f"5_REVISION_PLANO_BANCO_{corte_seleccionado.replace(' ', '_').replace('/', '-')}.xlsx", excel_banco_io.read())
            
            zip_buffer.seek(0)
            mensaje_carga.empty() 
            st.success(f"✅ ¡Éxito! Archivos listos para imprimir con 1 solo clic. {len(df_banco)} conductores procesados. ({ignorados} omitidos por saldo $0).")
            
            if len(df_banco) > 0:
                st.download_button(
                    label=f"📥 DESCARGAR PAQUETE GERENCIAL (.ZIP)",
                    data=zip_buffer,
                    file_name=f"SERGEM_Paquete_{corte_seleccionado.replace(' ', '_').replace('/', '-')}.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )
                st.divider()
                st.subheader("Vista Previa - Datos del Banco")
                st.text_area("Previsualización Archivo Plano (.txt)", texto_banco, height=250)

        except Exception as e:
            mensaje_carga.empty()
            st.error(f"Error en el proceso: {e}")
