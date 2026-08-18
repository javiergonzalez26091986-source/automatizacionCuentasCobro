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
st.set_page_config(page_title="SERGEM - Generador de documentos", page_icon="sergemLogo.ico", layout="wide")

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

# --- FUNCIÓN PARA CARGAR DATOS CON CACHÉ ---
# Esto evita consultar Google Sheets repetitivamente mientras el usuario interactúa
@st.cache_data(ttl=300) # Se actualiza cada 5 minutos
def cargar_datos(url):
    try:
        response = requests.get(url)
        return response.json()
    except Exception as e:
        return None

# --- CLASE PARA PDF ---
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
    
    pdf.cell(0, 6, f"{datos['ciudad']} {datos['fecha_emision']}".upper(), 0, 1)
    pdf.ln(8)
    
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "Debe a:", 0, 1)
    
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 6, str(datos['nombre_titular']).upper(), 0, 1)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"C.C {datos['cedula_titular']}", 0, 1)
    pdf.ln(5)

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

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="E3000F", end_color="E3000F", fill_type="solid")
    dark_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
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

    ws.merge_cells('D2:H4')
    ws['D2'] = "DOCUMENTO EQUIVALENTE A LA FACTURA DE VENTA\n(DECRETO 522 DE 2003)\nDOCUMENTO SOPORTE EN ADQUISICIONES A NO OBLIGADOS A FACTURAR"
    ws['D2'].font = Font(bold=True, size=11, color="1E293B")
    ws['D2'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    hoy = datetime.now()
    ws['B6'] = "Fecha de Expedición:"
    ws['B6'].font = bold_font
    ws['C6'] = "Año:"
    ws['D6'] = hoy.year
    ws['E6'] = "Mes:"
    ws['F6'] = f"{hoy.month:02d}"
    ws['G6'] = "Día:"
    ws['H6'] = f"{hoy.day:02d}"
    
    ws['B7'] = "Fecha de Radicación:"
    ws['B7'].font = bold_font
    ws['C7'] = "Año:"
    ws['D7'] = hoy.year
    ws['E7'] = "Mes:"
    ws['F7'] = f"{hoy.month:02d}"
    ws['G7'] = "Día:"
    ws['H7'] = f"{hoy.day:02d}"

    ws['G5'] = "CONSECUTIVO NO:"
    ws['G5'].font = Font(bold=True, size=11, color="E3000F")
    ws['G5'].alignment = right_align
    ws['H5'] = datos['id']
    ws['H5'].font = Font(bold=True, size=12, color="E3000F")
    ws['H5'].alignment = center_align

    ws['B9'] = " INFORMACIÓN DE LA EMPRESA (COMPRADOR)"
    ws['B9'].font = header_font
    ws['B9'].fill = dark_fill
    ws.merge_cells('B9:H9')

    ws['B10'] = "Razón Social:"
    ws['B10'].font = bold_font
    ws['C10'] = "SERGEM MENSAJERIA S.A.S."
    ws.merge_cells('C10:E10')
    ws['G10'] = "NIT:"
    ws['G10'].font = bold_font
    ws['H10'] = "900.561.833-1"

    ws['B11'] = "Dirección:"
    ws['B11'].font = bold_font
    ws['C11'] = "CRA 62 9 235"
    ws.merge_cells('C11:D11')
    ws['E11'] = "Teléfono:"
    ws['E11'].font = bold_font
    ws['F11'] = "3994620"
    ws['G11'] = "Ciudad:"
    ws['G11'].font = bold_font
    ws['H11'] = "CALI"

    ws['B13'] = " DATOS DEL BENEFICIARIO / PROVEEDOR (VENDEDOR)"
    ws['B13'].font = header_font
    ws['B13'].fill = dark_fill
    ws.merge_cells('B13:H13')

    ws['B14'] = "Nombre:"
    ws['B14'].font = bold_font
    ws['C14'] = datos['nombre_titular']
    ws.merge_cells('C14:E14')
    ws['G14'] = "C.C / NIT:"
    ws['G14'].font = bold_font
    ws['H14'] = datos['cedula_titular']

    ws['B15'] = "Dirección:"
    ws['B15'].font = bold_font
    ws['C15'] = "" 
    ws.merge_cells('C15:D15')
    ws['E15'] = "Teléfono:"
    ws['E15'].font = bold_font
    ws['F15'] = "" 
    ws['G15'] = "Ciudad:"
    ws['G15'].font = bold_font
    ws['H15'] = datos['ciudad']

    ws['B16'] = "Email:"
    ws['B16'].font = bold_font
    ws.merge_cells('C16:E16')
    ws['G16'] = "Conductor:"
    ws['G16'].font = bold_font
    ws['H16'] = datos['nombre_conductor']

    for r in range(10, 12):
        for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws[f'{c}{r}'].border = border_thin
    for r in range(14, 17):
        for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws[f'{c}{r}'].border = border_thin

    fila = 18
    for i, h in enumerate(["Ítem", "Concepto", "Cantidad", "V. Unitario", "V. Total"]):
        c = ['B', 'C', 'F', 'G', 'H'][i] + str(fila)
        ws[c] = h
        ws[c].font = header_font
        ws[c].fill = header_fill
        ws[c].alignment = center_align
        ws[c].border = border_thin
    ws.merge_cells(f'C{fila}:E{fila}')
    
    fila += 1
    ws[f'B{fila}'] = 1
    ws[f'C{fila}'] = f"Servicio de mensajería - Corte: {datos['corte_fechas']}"
    ws.merge_cells(f'C{fila}:E{fila}')
    ws[f'F{fila}'] = 1
    ws[f'G{fila}'] = datos['ingreso_base']
    ws[f'H{fila}'] = datos['ingreso_base']
    
    for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{c}{fila}'].border = border_thin
    ws[f'B{fila}'].alignment = center_align
    ws[f'C{fila}'].alignment = left_align
    ws[f'F{fila}'].alignment = center_align
    ws[f'G{fila}'].number_format = '"$"#,##0'
    ws[f'H{fila}'].number_format = '"$"#,##0'

    if datos.get('fuera_perimetro', 0) > 0:
        fila += 1
        ws[f'B{fila}'] = 2
        ws[f'C{fila}'] = "Servicios Fuera de Perímetro"
        ws.merge_cells(f'C{fila}:E{fila}')
        ws[f'F{fila}'] = 1
        ws[f'G{fila}'] = datos['fuera_perimetro']
        ws[f'H{fila}'] = datos['fuera_perimetro']
        
        for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws[f'{c}{fila}'].border = border_thin
        ws[f'B{fila}'].alignment = center_align
        ws[f'C{fila}'].alignment = left_align
        ws[f'F{fila}'].alignment = center_align
        ws[f'G{fila}'].number_format = '"$"#,##0'
        ws[f'H{fila}'].number_format = '"$"#,##0'
        
    fila += 2
    totales = [
        ("SUBTOTAL:", datos['ingreso_bruto_total']),
        ("IVA (19%):", ""), 
        ("RETEIVA:", ""), 
        ("RTE FTE (1%):", -datos['retefuente']),
        ("RETEICA (1%):", -datos['ica']),
        ("NETO A PAGAR:", datos['neto_pagar'])
    ]
    
    fila_firma = fila + 1

    for label, valor in totales:
        ws[f'G{fila}'] = label
        ws[f'H{fila}'] = valor
        ws[f'G{fila}'].font = bold_font
        ws[f'G{fila}'].alignment = right_align
        ws[f'G{fila}'].border = border_thin
        ws[f'H{fila}'].border = border_thin
        
        if valor != "":
            ws[f'H{fila}'].number_format = '"$"#,##0'
        
        if label == "NETO A PAGAR:":
            ws[f'G{fila}'].font = Font(bold=True, color="E3000F")
            ws[f'H{fila}'].font = Font(bold=True, size=12)
            ws[f'H{fila}'].fill = PatternFill(start_color="F4F6F9", end_color="F4F6F9", fill_type="solid")
            
        fila += 1

    ws[f'B{fila_firma}'] = "________________________________________________"
    ws[f'B{fila_firma+1}'] = "FIRMA PRESTADOR DEL SERVICIO"
    ws[f'B{fila_firma+1}'].font = bold_font
    ws[f'B{fila_firma+2}'] = f"C.C. / NIT: {datos['cedula_titular']}"
    ws[f'B{fila_firma+3}'] = f"NOMBRE: {datos['nombre_titular']}"

    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 22
    ws.column_dimensions['H'].width = 22

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
    Seleccione el <b>Corte</b> que desea procesar del menú desplegable. 
    Los conductores que tengan un "TOTAL A PAGAR" de $0 serán omitidos automáticamente.
</div>
""", unsafe_allow_html=True)

# 1. CARGAMOS LOS DATOS PRIMERO
data_cruda = cargar_datos(GAS_URL)

if not data_cruda:
    st.error("Error conectando a la base de datos de Google Sheets. Revise la URL.")
    st.stop()

df_pagos_completo = pd.DataFrame(data_cruda.get('pagos', []))
df_fuera = pd.DataFrame(data_cruda.get('fueras_perimetro', []))

if df_pagos_completo.empty:
    st.warning("No se encontraron datos en Google Sheets.")
    st.stop()

# Limpiamos columnas
df_pagos_completo.columns = df_pagos_completo.columns.str.strip().str.upper()
if not df_fuera.empty:
    df_fuera.columns = df_fuera.columns.str.strip().str.upper()

# 2. CREAMOS EL MENÚ DESPLEGABLE CON LOS CORTES
if 'CORTE' in df_pagos_completo.columns:
    # Filtramos valores nulos o vacíos para el menú
    cortes_disponibles = [c for c in df_pagos_completo['CORTE'].unique() if str(c).strip() != "" and str(c).lower() != "nan"]
    if cortes_disponibles:
        corte_seleccionado = st.selectbox("📅 Seleccione el Corte a procesar:", cortes_disponibles)
    else:
        st.error("La columna CORTE en Google Sheets no tiene datos válidos.")
        st.stop()
else:
    st.error("No se encontró la columna 'CORTE' en la hoja de Google Sheets.")
    st.stop()

# 3. BOTÓN DE PROCESAMIENTO
if st.button("🚀 Procesar Nómina y Generar ZIP", use_container_width=True):
    
    mensaje_carga = st.info(f"📥 Procesando la información para el corte: {corte_seleccionado}...")
    
    try:
        # Filtramos la tabla solo con el corte seleccionado
        df_pagos = df_pagos_completo[df_pagos_completo['CORTE'] == corte_seleccionado]
        
        pagos_procesados = []
        ignorados = 0
        zip_buffer = io.BytesIO()
        fecha_actual = obtener_fecha_actual() 
        
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
                    'fecha_emision': fecha_actual, 
                    'nombre_titular': nombre_titular,
                    'cedula_titular': cedula_titular,
                    'nombre_conductor': nombre_conductor,
                    'cedula_conductor': str(row.get('CEDULA', '')).strip(),
                    'ciudad': ciudad,
                    'corte_fechas': corte_seleccionado, # Usamos directamente el corte seleccionado
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
            zip_file.writestr(f"Archivo_Plano_Bancario_{corte_seleccionado.replace(' ', '_').replace('/', '-')}.csv", df_resultado.to_csv(index=False, sep=';', encoding='utf-8-sig'))
        
        zip_buffer.seek(0)
        
        mensaje_carga.empty() 
        st.success(f"✅ ¡Proceso finalizado con éxito! Se generaron los documentos del corte **{corte_seleccionado}**. Se procesaron {len(df_resultado)} pagos y se omitieron {ignorados} registros con saldo $0.")
        
        if len(df_resultado) > 0:
            st.download_button(
                label=f"📥 DESCARGAR ARCHIVOS (.ZIP) - {corte_seleccionado}",
                data=zip_buffer,
                file_name=f"SERGEM_Docs_{corte_seleccionado.replace(' ', '_').replace('/', '-')}.zip",
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
