import streamlit as st
import pandas as pd
import requests
import io
import zipfile
import os
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

# --- URL INTEGRADA (Oculta al usuario) ---
GAS_URL = "https://script.google.com/macros/s/AKfycbyqJtrmVdNT1rxTobg6q_WoJCwMpp40hdIzJeEm4dKNLBgDVxwEY95T0EIoBu_qo8FB/exec"

# --- CLASE PARA PDF CON FPDF2 ---
class PDFCuentaCobro(FPDF):
    def header(self):
        try:
            if os.path.exists('sergemLogo.png'):
                self.image('sergemLogo.png', 10, 8, 35)
        except:
            pass
        self.set_font('helvetica', 'B', 14)
        self.set_text_color(51, 51, 51)
        self.cell(0, 8, 'SERGEM MENSAJERIA S.A.S.', 0, 1, 'R')
        self.set_font('helvetica', '', 9)
        self.set_text_color(119, 119, 119)
        self.cell(0, 5, 'NIT. 900.561.833-1', 0, 1, 'R')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

def generar_pdf_cuenta_cobro(datos):
    pdf = PDFCuentaCobro()
    pdf.add_page()
    
    # Título
    pdf.set_font("helvetica", "B", 16)
    pdf.set_text_color(227, 0, 15)
    pdf.cell(0, 8, "CUENTA DE COBRO", 0, 1, "C")
    
    pdf.set_font("helvetica", "", 10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 6, f"Documento No. {datos['id']} | Fecha: {datos['fecha_actual']}", 0, 1, "C")
    pdf.ln(8)
    
    # Debe a
    pdf.set_text_color(51, 51, 51)
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(0, 5, "DEBE A:", 0, 1)
    
    pdf.set_font("helvetica", "B", 13)
    pdf.cell(0, 6, datos['nombre_titular'], 0, 1)
    pdf.set_font("helvetica", "", 10)
    pdf.cell(0, 6, f"C.C. {datos['cedula_titular']}", 0, 1)
    pdf.ln(5)
    
    # Caja de detalles
    start_y = pdf.get_y()
    pdf.set_fill_color(253, 251, 247)
    pdf.set_draw_color(221, 221, 221)
    pdf.rect(10, start_y, 190, 75, style='FD')
    
    pdf.set_xy(15, start_y + 5)
    pdf.set_font("helvetica", "", 10)
    pdf.cell(0, 6, f"Concepto: Servicio de mensajería prestado en el corte del {datos['corte_fechas']}.", 0, 1)
    pdf.set_x(15)
    pdf.cell(0, 6, f"Conductor: {datos['nombre_conductor']} (C.C. {datos['cedula_conductor']})", 0, 1)
    pdf.set_x(15)
    pdf.cell(0, 6, f"Ciudad de Operación: {datos['ciudad']}", 0, 1)
    pdf.ln(4)
    
    pdf.set_x(15)
    pdf.cell(90, 6, "Valor Base Negociado:", 0, 0)
    pdf.cell(85, 6, f"$ {datos['ingreso_bruto']:,.0f}", 0, 1, "R")
    
    pdf.set_x(15)
    pdf.set_text_color(217, 83, 79)
    pdf.cell(90, 6, "Retención en la Fuente (1%):", 0, 0)
    pdf.cell(85, 6, f"- $ {datos['retefuente']:,.0f}", 0, 1, "R")
    
    pdf.set_x(15)
    pdf.cell(90, 6, "Descuento ICA (1%):", 0, 0)
    pdf.cell(85, 6, f"- $ {datos['ica']:,.0f}", 0, 1, "R")
    
    pdf.set_text_color(51, 51, 51)
    pdf.set_x(15)
    pdf.set_font("helvetica", "B", 11)
    pdf.set_text_color(227, 0, 15)
    pdf.cell(90, 8, "NETO A PAGAR:", "T", 0)
    pdf.cell(85, 8, f"$ {datos['neto_pagar']:,.0f}", "T", 1, "R")
    
    # Info Bancaria
    pdf.set_xy(10, start_y + 80)
    pdf.set_text_color(51, 51, 51)
    pdf.set_fill_color(244, 244, 244)
    pdf.rect(10, pdf.get_y(), 190, 32, style='FD')
    
    pdf.set_xy(15, pdf.get_y() + 3)
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(0, 5, "Por favor consignar en la siguiente cuenta bancaria:", 0, 1)
    pdf.set_font("helvetica", "", 10)
    pdf.set_x(15)
    pdf.cell(0, 5, f"Banco: {datos['banco']}  |  Tipo: {datos['tipo_cuenta']}  |  Número: {datos['num_cuenta']}", 0, 1)
    pdf.set_x(15)
    pdf.cell(0, 5, f"Titular: {datos['nombre_titular']}", 0, 1)
    
    # Firma
    pdf.ln(20)
    pdf.set_x(15)
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(80, 5, datos['nombre_titular'], "T", 1, "C")
    pdf.set_x(15)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(80, 5, f"C.C. {datos['cedula_titular']}", 0, 1, "C")
    
    return pdf.output()

def generar_excel_documento_equivalente(datos):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Documento Equivalente"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="E3000F", end_color="E3000F", fill_type="solid")
    bold_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

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
    ws['D2'].font = Font(bold=True, size=12)
    ws['D2'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws['B5'] = "Empresa:"
    ws['C5'] = "SERGEM MENSAJERIA S.A.S."
    ws['C5'].font = bold_font
    ws['B6'] = "NIT:"
    ws['C6'] = "900.561.833-1"

    ws['G5'] = "Documento No:"
    ws['H5'] = datos['id']
    ws['H5'].font = bold_font
    ws['G6'] = "Fecha:"
    ws['H6'] = datos['fecha_actual']

    ws['B8'] = "DATOS DEL BENEFICIARIO / PROVEEDOR"
    ws['B8'].font = Font(bold=True, color="E3000F")
    ws.merge_cells('B8:E8')

    ws['B9'] = "Nombre:"
    ws['C9'] = datos['nombre_titular']
    ws['G9'] = "C.C / NIT:"
    ws['H9'] = datos['cedula_titular']

    ws['B10'] = "Ciudad:"
    ws['C10'] = datos['ciudad']
    ws['G10'] = "Conductor:"
    ws['H10'] = datos['nombre_conductor']

    for i, h in enumerate(["Ítem", "Concepto", "Cantidad", "V. Unitario", "V. Total"]):
        c = ['B', 'C', 'F', 'G', 'H'][i] + '12'
        ws[c] = h
        ws[c].font = header_font
        ws[c].fill = header_fill
        ws[c].alignment = center_align
        ws[c].border = border
    ws.merge_cells('C12:E12')

    ws['B13'] = 1
    ws['C13'] = f"Servicio mensajería ({datos['corte_fechas']})"
    ws.merge_cells('C13:E13')
    ws['F13'] = 1
    ws['G13'] = datos['ingreso_bruto']
    ws['H13'] = datos['ingreso_bruto']

    for c in ['B13', 'C13', 'F13', 'G13', 'H13']:
        ws[c].border = border
    ws['G13'].number_format = '"$"#,##0'
    ws['H13'].number_format = '"$"#,##0'
    ws['B13'].alignment = center_align
    ws['F13'].alignment = center_align

    ws['G15'] = "SUBTOTAL:"
    ws['H15'] = datos['ingreso_bruto']
    ws['G16'] = "Retefuente (1%):"
    ws['H16'] = -datos['retefuente']
    ws['G17'] = "ICA (1%):"
    ws['H17'] = -datos['ica']
    ws['G18'] = "TOTAL A PAGAR:"
    ws['H18'] = datos['neto_pagar']

    for r in range(15, 19):
        ws[f'G{r}'].font = bold_font
        ws[f'G{r}'].alignment = right_align
        ws[f'H{r}'].number_format = '"$"#,##0'
        ws[f'H{r}'].border = border
    ws['H18'].font = bold_font
    ws['H18'].fill = PatternFill(start_color="e6f2ff", end_color="e6f2ff", fill_type="solid")

    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['G'].width = 18
    ws.column_dimensions['H'].width = 18

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io.read()

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
    Asegúrese de haber actualizado las horas en el archivo de Google Sheets (pestañas <i>CUENTAS</i> y <i>PAGOS PERSONAL</i>). 
    Cuando esté listo, presione el botón rojo para generar automáticamente todos los documentos de la quincena.
</div>
""", unsafe_allow_html=True)

if st.button("🚀 Procesar Nómina y Generar ZIP", use_container_width=True):
    with st.status("Procesando la información. Por favor espere...", expanded=True) as status:
        try:
            st.write("📥 Leyendo datos desde Google Sheets...")
            response = requests.get(GAS_URL)
            
            try:
                data = response.json()
            except Exception as e:
                st.error("Error conectando a la base de datos. Verifique que el script de Google esté publicado correctamente.")
                st.stop()
            
            st.write("⚙️ Calculando Retefuente e ICA por ciudad...")
            df_cuentas = pd.DataFrame(data.get('cuentas', []))
            df_pagos = pd.DataFrame(data.get('pagos', []))
            
            if df_pagos.empty or df_cuentas.empty:
                st.warning("No se encontraron datos en las pestañas de Sheets.")
                st.stop()
            
            df_pagos['CIUDAD'] = df_pagos['CIUDAD'].astype(str).str.upper().str.strip()
            df_completo = pd.merge(
                df_pagos, 
                df_cuentas[['NOMBRE DEL CONDUCTOR (PLANILLA)', 'NOMBRE DEL TITULAR DE LA CUENTA', 'DOCUMENTO DEL TITULAR DE LA CUENTA', 'NOMBRE DE LA ENTIDAD FINANCIERA - BANCO', 'TIPO DE CUENTA (AHORROS-CORRIENTE)', 'NUMERO DE CUENTA']], 
                left_on='CONDUCTOR', 
                right_on='NOMBRE DEL CONDUCTOR (PLANILLA)', 
                how='left'
            )

            st.write("📄 Creando Cuentas de Cobro y Documentos Equivalentes...")
            pagos_procesados = []
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                contador = 1
                for index, row in df_completo.iterrows():
                    if pd.notna(row.get('TOTAL A PAGAR')) and row.get('TOTAL A PAGAR', 0) > 0:
                        ingreso_bruto = float(row['TOTAL A PAGAR'])
                        retefuente = ingreso_bruto * 0.01
                        ica = ingreso_bruto * 0.01 if row['CIUDAD'] == 'CALI' else 0.0
                        neto_a_pagar = ingreso_bruto - retefuente - ica
                        
                        nombre_titular = str(row.get('NOMBRE DEL TITULAR DE LA CUENTA', 'S/N'))
                        
                        pagos_procesados.append({
                            'NIT_TITULAR': str(row.get('DOCUMENTO DEL TITULAR DE LA CUENTA', '')),
                            'NOMBRE_TITULAR': nombre_titular,
                            'BANCO': str(row.get('NOMBRE DE LA ENTIDAD FINANCIERA - BANCO', '')),
                            'TIPO_CUENTA': str(row.get('TIPO DE CUENTA (AHORROS-CORRIENTE)', '')),
                            'NUM_CUENTA': str(row.get('NUMERO DE CUENTA', '')),
                            'NETO_A_PAGAR': round(neto_a_pagar, 0)
                        })
                        
                        datos_doc = {
                            'id': str(contador).zfill(3),
                            'fecha_actual': '16/08/2026',
                            'nombre_titular': nombre_titular,
                            'cedula_titular': str(row.get('DOCUMENTO DEL TITULAR DE LA CUENTA', '')),
                            'nombre_conductor': str(row.get('CONDUCTOR', '')),
                            'cedula_conductor': str(row.get('CÉDULA', '')),
                            'ciudad': row['CIUDAD'],
                            'corte_fechas': str(row.get('CORTE', '1 AL 15 AGOSTO')),
                            'ingreso_bruto': ingreso_bruto,
                            'retefuente': retefuente,
                            'ica': ica,
                            'neto_pagar': neto_a_pagar,
                            'banco': str(row.get('NOMBRE DE LA ENTIDAD FINANCIERA - BANCO', '')),
                            'tipo_cuenta': str(row.get('TIPO DE CUENTA (AHORROS-CORRIENTE)', '')),
                            'num_cuenta': str(row.get('NUMERO DE CUENTA', ''))
                        }

                        zip_file.writestr(f"Cuentas_Cobro/{contador}_{nombre_titular.replace(' ', '_')}_Cuenta.pdf", generar_pdf_cuenta_cobro(datos_doc))
                        zip_file.writestr(f"Docs_Equivalentes/{contador}_{nombre_titular.replace(' ', '_')}_DocEq.xlsx", generar_excel_documento_equivalente(datos_doc))
                        contador += 1
            
            df_resultado = pd.DataFrame(pagos_procesados)
            zip_file.writestr("Archivo_Plano_Bancario.csv", df_resultado.to_csv(index=False, sep=';').encode('utf-8'))
            zip_buffer.seek(0)
            
            status.update(label="¡Proceso finalizado con éxito!", state="complete", expanded=False)
            
            st.success(f"✅ ¡Todo listo! Se procesaron {len(df_resultado)} pagos.")
            
            st.download_button(
                label="📥 DESCARGAR ARCHIVOS DE LA QUINCENA (.ZIP)",
                data=zip_buffer,
                file_name="SERGEM_Documentos_Quincena.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

            st.divider()
            st.subheader("Vista Previa - Consolidado Bancario")
            st.dataframe(df_resultado, use_container_width=True)

        except Exception as e:
            status.update(label="Ocurrió un error", state="error")
            st.error(f"Error en el proceso: {e}")
