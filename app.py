import streamlit as st
import pandas as pd
import requests
import io
import zipfile
import os
from datetime import datetime
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage

# 1. Configuración de la página
st.set_page_config(page_title="SERGEM - Nómina y Cuentas", page_icon="sergemLogo.ico", layout="wide")

# 2. Inyección de CSS corporativo
st.markdown("""
<style>
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
    .stButton>button {
        background-color: #E3000F !important;
        color: white !important;
        border-radius: 6px;
        border: none;
        padding: 0.75rem 1.5rem;
        font-weight: bold;
        font-size: 1.1rem;
        transition: all 0.3s ease;
    }
    .stButton>button:hover { background-color: #B3000C !important; box-shadow: 0px 4px 10px rgba(227, 0, 15, 0.3); }
    .instrucciones {
        background-color: #f8fafc;
        border-left: 4px solid #E3000F;
        padding: 1rem;
        border-radius: 4px;
        margin-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# --- URL INTEGRADA ---
GAS_URL = "https://script.google.com/macros/s/AKfycbyqJtrmVdNT1rxTobg6q_WoJCwMpp40hdIzJeEm4dKNLBgDVxwEY95T0EIoBu_qo8FB/exec"

# --- OBTENER FECHA ACTUAL EN FORMATO COLOMBIANO ---
def obtener_fecha_actual():
    meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    hoy = datetime.now()
    return f"{hoy.day} DE {meses[hoy.month - 1]} DE {hoy.year}"

# --- CLASE PARA PDF CON EL FORMATO DE LA CLIENTA ---
class PDFCuentaCobro(FPDF):
    def header(self):
        try:
            if os.path.exists('sergemLogo.png'):
                self.image('sergemLogo.png', 10, 8, 35)
        except:
            pass
        self.set_font('helvetica', 'B', 12)
        self.set_text_color(51, 51, 51)
        self.cell(0, 6, 'SERGEM MENSAJERIA S.A.S.', 0, 1, 'R')
        self.set_font('helvetica', '', 10)
        self.cell(0, 5, 'NIT. 900.561.833-1', 0, 1, 'R')
        self.ln(15)

def generar_pdf_cuenta_cobro(datos):
    pdf = PDFCuentaCobro()
    pdf.add_page()
    
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("helvetica", "B", 11)
    
    # Fecha de emisión dinámica
    pdf.cell(0, 6, f"{datos['ciudad']} {datos['fecha_emision']}".upper(), 0, 1)
    pdf.ln(8)
    
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "Debe a:", 0, 1)
    
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 6, str(datos['nombre_titular']).upper(), 0, 1)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"C.C {datos['cedula_titular']}", 0, 1)
    pdf.ln(5)

    # Detalle de pago (Se amplió el ancho a 80 para evitar superposición)
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(80, 6, "VALOR BASE:", 0, 0)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"$ {datos['ingreso_base']:,.0f}", 0, 1)

    if datos.get('fuera_perimetro', 0) > 0:
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(80, 6, "FUERA DE PERÍMETRO:", 0, 0)
        pdf.set_font("helvetica", "", 11)
        pdf.cell(0, 6, f"$ {datos['fuera_perimetro']:,.0f}", 0, 1)

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
    
    return pdf.output()

def generar_excel_documento_equivalente(datos):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Documento Equivalente"

    # Estilos profesionales
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="E3000F", end_color="E3000F", fill_type="solid")
    bold_font = Font(bold=True)
    border_thin = Border(left=Side(style='thin', color='BFBFBF'), 
                         right=Side(style='thin', color='BFBFBF'), 
                         top=Side(style='thin', color='BFBFBF'), 
                         bottom=Side(style='thin', color='BFBFBF'))
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    try:
        if os.path.exists('sergemLogo.png'):
            img = XLImage('sergemLogo.png')
            img.width = 150
            img.height = 60
            ws.add_image(img, 'B2')
    except:
        pass

    ws.merge_cells('D2:I3')
    ws['D2'] = "DOCUMENTO SOPORTE EN ADQUISICIONES\nEFECTUADAS A NO OBLIGADOS A FACTURAR"
    ws['D2'].font = Font(bold=True, size=13, color="1E293B")
    ws['D2'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Info de la empresa (Caja superior)
    ws['B5'] = "Empresa:"
    ws['C5'] = "SERGEM MENSAJERIA S.A.S."
    ws['C5'].font = bold_font
    ws['B6'] = "NIT:"
    ws['C6'] = "900.561.833-1"

    ws['G5'] = "Documento No:"
    ws['H5'] = datos['id']
    ws['H5'].font = Font(bold=True, size=12, color="E3000F")
    ws['H5'].alignment = right_align
    ws['G6'] = "Fecha Emisión:"
    ws['H6'] = datos['fecha_emision']
    ws['H6'].alignment = right_align

    # Info del Proveedor (Caja central con bordes)
    ws['B8'] = " DATOS DEL BENEFICIARIO / PROVEEDOR"
    ws['B8'].font = Font(bold=True, color="FFFFFF")
    ws['B8'].fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
    ws.merge_cells('B8:H8')

    ws['B9'] = "Nombre:"
    ws['B9'].font = bold_font
    ws['C9'] = datos['nombre_titular']
    ws.merge_cells('C9:E9')
    ws['G9'] = "C.C / NIT:"
    ws['G9'].font = bold_font
    ws['H9'] = datos['cedula_titular']

    ws['B10'] = "Ciudad:"
    ws['B10'].font = bold_font
    ws['C10'] = datos['ciudad']
    ws.merge_cells('C10:E10')
    ws['G10'] = "Conductor:"
    ws['G10'].font = bold_font
    ws['H10'] = datos['nombre_conductor']

    # Aplicar bordes suaves a la caja del proveedor
    for r in range(9, 11):
        for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws[f'{c}{r}'].border = border_thin

    # Tabla de Conceptos
    for i, h in enumerate(["Ítem", "Concepto", "Cantidad", "V. Unitario", "V. Total"]):
        c = ['B', 'C', 'F', 'G', 'H'][i] + '12'
        ws[c] = h
        ws[c].font = header_font
        ws[c].fill = header_fill
        ws[c].alignment = center_align
        ws[c].border = border_thin
    ws.merge_cells('C12:E12')

    ws['B13'] = 1
    ws['C13'] = f"Servicio mensajería ({datos['corte_fechas']})"
    ws.merge_cells('C13:E13')
    ws['F13'] = 1
    ws['G13'] = datos['ingreso_base']
    ws['H13'] = datos['ingreso_base']

    fila_subtotal = 15
    if datos.get('fuera_perimetro', 0) > 0:
        ws['B14'] = 2
        ws['C14'] = "Servicios Fuera de Perímetro"
        ws.merge_cells('C14:E14')
        ws['F14'] = 1
        ws['G14'] = datos['fuera_perimetro']
        ws['H14'] = datos['fuera_perimetro']
        for c in ['B14', 'C14', 'D14', 'E14', 'F14', 'G14', 'H14']:
            ws[c].border = border_thin
        ws['G14'].number_format = '"$"#,##0'
        ws['H14'].number_format = '"$"#,##0'
        ws['B14'].alignment = center_align
        ws['F14'].alignment = center_align
        ws['C14'].alignment = left_align
        fila_subtotal = 16
    
    for c in ['B13', 'C13', 'D13', 'E13', 'F13', 'G13', 'H13']:
        ws[c].border = border_thin
    ws['G13'].number_format = '"$"#,##0'
    ws['H13'].number_format = '"$"#,##0'
    ws['B13'].alignment = center_align
    ws['F13'].alignment = center_align
    ws['C13'].alignment = left_align

    # Totales (Más organizados)
    ws[f'G{fila_subtotal}'] = "SUBTOTAL:"
    ws[f'H{fila_subtotal}'] = datos['ingreso_bruto_total']
    ws[f'G{fila_subtotal+1}'] = "Retefuente (1%):"
    ws[f'H{fila_subtotal+1}'] = -datos['retefuente']
    ws[f'G{fila_subtotal+2}'] = "ICA (1%):"
    ws[f'H{fila_subtotal+2}'] = -datos['ica']
    ws[f'G{fila_subtotal+3}'] = "TOTAL A PAGAR:"
    ws[f'H{fila_subtotal+3}'] = datos['neto_pagar']

    for r in range(fila_subtotal, fila_subtotal+4):
        ws[f'G{r}'].font = bold_font
        ws[f'G{r}'].alignment = right_align
        ws[f'H{r}'].number_format = '"$"#,##0'
        ws[f'G{r}'].border = border_thin
        ws[f'H{r}'].border = border_thin
        
    ws[f'G{fila_subtotal+3}'].font = Font(bold=True, color="E3000F")
    ws[f'H{fila_subtotal+3}'].font = Font(bold=True, size=12)
    ws[f'H{fila_subtotal+3}'].fill = PatternFill(start_color="F4F6F9", end_color="F4F6F9", fill_type="solid")

    # Anchos de columna optimizados
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['G'].width = 20
    ws.column_dimensions['H'].width = 20

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io.read()

def limpiar_dinero(val):
    if pd.isna(val) or val == "":
        return 0.0
    s = str(val).upper().replace('$', '').replace(',', '').replace('.', '').replace(' ', '')
    try:
        return float(s)
    except:
        return 0.0

# --- INTERFAZ DE USUARIO ---
col1, col2 = st.columns([1, 4])
with col1:
    try:
        st.image("sergemLogo.png", width=140)
    except:
        pass
with col2:
    st.title("Generador de Pagos y Documentos")
    st.markdown("**SERGEM Mensajería S.A.S.**")

st.markdown("""
<div class="instrucciones">
    <strong>💡 Instrucciones:</strong><br>
    Asegúrese de haber ingresado las horas trabajadas en la hoja <i>PAGOS PERSONAL POR SERVICIOS</i>. 
    Los conductores que tengan un "TOTAL A PAGAR" de $0 serán omitidos automáticamente.
</div>
""", unsafe_allow_html=True)

if st.button("🚀 Procesar Nómina y Generar ZIP", use_container_width=True):
    
    # Eliminamos el st.status expandible y mostramos los mensajes en vivo en la pantalla principal
    mensaje_carga = st.info("📥 Conectando con Google Sheets y procesando la información. Por favor espere...")
    
    try:
        response = requests.get(GAS_URL)
        data = response.json()
        
        df_pagos = pd.DataFrame(data.get('pagos', []))
        df_fuera = pd.DataFrame(data.get('fueras_perimetro', []))
        
        if df_pagos.empty:
            st.warning("No se encontraron datos en Google Sheets.")
            st.stop()
        
        df_pagos.columns = df_pagos.columns.str.strip().str.upper()
        if not df_fuera.empty:
            df_fuera.columns = df_fuera.columns.str.strip().str.upper()

        pagos_procesados = []
        ignorados = 0
        zip_buffer = io.BytesIO()
        fecha_actual = obtener_fecha_actual() # Se captura la fecha dinámicamente
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            contador = 1
            for index, row in df_pagos.iterrows():
                
                ingreso_base = limpiar_dinero(row.get('TOTAL A PAGAR', 0))

                if ingreso_base <= 0:
                    ignorados += 1
                    continue

                nombre_conductor = str(row.get('CONDUCTOR', '')).upper().strip()
                ciudad = str(row.get('CIUDAD', '')).upper().strip()
                
                fuera_perimetro = 0.0
                if "MILTON" in nombre_conductor and not df_fuera.empty:
                    if 'TOTAL' in df_fuera.columns:
                        fuera_perimetro = sum(df_fuera['TOTAL'].apply(limpiar_dinero))

                ingreso_bruto_total = ingreso_base + fuera_perimetro
                
                retefuente = ingreso_bruto_total * 0.01
                ica = ingreso_bruto_total * 0.01 if ciudad == 'CALI' else 0.0
                neto_a_pagar = ingreso_bruto_total - retefuente - ica
                
                nombre_titular = str(row.get('A NOMBRE DE QUIEN HACE CUENTA DE COBRO', row.get('NOMBRE TITULAR CUENTA BANCARIA', 'S/N'))).strip()
                cedula_titular = str(row.get('CÉDULA DE CUENTA DE COBRO', row.get('CÉDULA TITULAR', ''))).strip()
                banco = str(row.get('BANCO', '')).strip()
                tipo_cuenta = str(row.get('TIPO CUENTA', '')).strip()
                num_cuenta = str(row.get('NO. CUENTA', '')).strip()
                
                pagos_procesados.append({
                    'CÉDULA': cedula_titular,
                    'NOMBRE': nombre_titular,
                    'BANCO': banco,
                    'TIPO CUENTA': tipo_cuenta,
                    'NÚMERO CUENTA': f"'{num_cuenta}", 
                    'VALOR NETO A PAGAR': round(neto_a_pagar, 0)
                })
                
                datos_doc = {
                    'id': str(contador).zfill(3),
                    'fecha_emision': fecha_actual, # Usando la fecha de Colombia
                    'nombre_titular': nombre_titular,
                    'cedula_titular': cedula_titular,
                    'nombre_conductor': nombre_conductor,
                    'cedula_conductor': str(row.get('CEDULA', '')).strip(),
                    'ciudad': ciudad,
                    'corte_fechas': str(row.get('CORTE', '')).strip(),
                    'ingreso_base': ingreso_base,
                    'fuera_perimetro': fuera_perimetro,
                    'ingreso_bruto_total': ingreso_bruto_total,
                    'retefuente': retefuente,
                    'ica': ica,
                    'neto_pagar': neto_a_pagar,
                    'banco': banco,
                    'tipo_cuenta': tipo_cuenta,
                    'num_cuenta': num_cuenta
                }

                zip_file.writestr(f"Cuentas_Cobro/{contador}_{nombre_titular.replace(' ', '_')}_Cuenta.pdf", generar_pdf_cuenta_cobro(datos_doc))
                zip_file.writestr(f"Docs_Equivalentes/{contador}_{nombre_titular.replace(' ', '_')}_DocEq.xlsx", generar_excel_documento_equivalente(datos_doc))
                contador += 1
            
            df_resultado = pd.DataFrame(pagos_procesados)
            zip_file.writestr("Archivo_Plano_Bancario_Don_Jose.csv", df_resultado.to_csv(index=False, sep=';', encoding='utf-8-sig'))
        
        zip_buffer.seek(0)
        
        # Ocultamos el mensaje de carga y mostramos el éxito directamente
        mensaje_carga.empty() 
        st.success(f"✅ ¡Proceso finalizado con éxito! Se procesaron {len(df_resultado)} pagos. (Se omitieron {ignorados} registros con saldo $0).")
        
        if len(df_resultado) > 0:
            st.download_button(
                label="📥 DESCARGAR ARCHIVOS DE LA QUINCENA (.ZIP)",
                data=zip_buffer,
                file_name="SERGEM_Documentos_Quincena.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

            st.divider()
            st.subheader("Vista Previa - Consolidado Bancario para Don José")
            st.dataframe(df_resultado, use_container_width=True)

    except Exception as e:
        mensaje_carga.empty()
        st.error(f"Error en el proceso: {e}")
