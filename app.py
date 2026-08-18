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

# --- URL INTEGRADA ---
GAS_URL = "https://script.google.com/macros/s/AKfycbyqJtrmVdNT1rxTobg6q_WoJCwMpp40hdIzJeEm4dKNLBgDVxwEY95T0EIoBu_qo8FB/exec"

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
    
    # Ciudad y Fecha
    pdf.cell(0, 6, f"{datos['ciudad']} {datos['corte_fechas']}".upper(), 0, 1)
    pdf.ln(8)
    
    # Debe a
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "Debe a:", 0, 1)
    
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 6, str(datos['nombre_titular']).upper(), 0, 1)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"C.C {datos['cedula_titular']}", 0, 1)
    pdf.ln(5)

    # Detalle de pago (Valores NETOS según audio 3)
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(60, 6, "VALOR BASE:", 0, 0)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, f"$ {datos['ingreso_base']:,.0f}", 0, 1)

    if datos.get('fuera_perimetro', 0) > 0:
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(60, 6, "FUERA DE PERÍMETRO:", 0, 0)
        pdf.set_font("helvetica", "", 11)
        pdf.cell(0, 6, f"$ {datos['fuera_perimetro']:,.0f}", 0, 1)

    pdf.set_font("helvetica", "B", 12)
    pdf.cell(60, 8, "VALOR TOTAL NETO A PAGAR:", 0, 0)
    pdf.cell(0, 8, f"$ {datos['neto_pagar']:,.0f}", 0, 1)
    pdf.ln(8)
    
    # Concepto
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(25, 6, "CONCEPTO:", 0, 0)
    pdf.set_font("helvetica", "", 11)
    concepto = f"SERVICIO PRESTADO EN EL CORTE DE {datos['corte_fechas']}, POR EL CONDUCTOR {datos['nombre_conductor']} CÉDULA {datos['cedula_conductor']}."
    pdf.multi_cell(0, 6, concepto)
    pdf.ln(10)
    
    # Autorización Bancaria
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 6, "Autorizo me sea consignado en:", 0, 1)
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(0, 6, f"CUENTA # {datos['num_cuenta']}", 0, 1)
    pdf.cell(0, 6, f"{datos['tipo_cuenta'].upper()}", 0, 1)
    pdf.cell(0, 6, f"NOMBRE DEL BANCO: {datos['banco'].upper()}", 0, 1)
    pdf.cell(0, 6, f"NOMBRE TITULAR CUENTA: {datos['nombre_titular']}", 0, 1)
    pdf.cell(0, 6, f"CÉDULA TITULAR CUENTA: {datos['cedula_titular']}", 0, 1)
    pdf.ln(15)
    
    # Firma
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
    ws['G6'] = "Corte:"
    ws['H6'] = datos['corte_fechas']

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
    ws['G13'] = datos['ingreso_base']
    ws['H13'] = datos['ingreso_base']

    # Fila para Fuera de Perímetro si aplica
    fila_subtotal = 15
    if datos.get('fuera_perimetro', 0) > 0:
        ws['B14'] = 2
        ws['C14'] = "Servicios Fuera de Perímetro"
        ws.merge_cells('C14:E14')
        ws['F14'] = 1
        ws['G14'] = datos['fuera_perimetro']
        ws['H14'] = datos['fuera_perimetro']
        for c in ['B14', 'C14', 'F14', 'G14', 'H14']:
            ws[c].border = border
        ws['G14'].number_format = '"$"#,##0'
        ws['H14'].number_format = '"$"#,##0'
        ws['B14'].alignment = center_align
        ws['F14'].alignment = center_align
        fila_subtotal = 16
    
    for c in ['B13', 'C13', 'F13', 'G13', 'H13']:
        ws[c].border = border
    ws['G13'].number_format = '"$"#,##0'
    ws['H13'].number_format = '"$"#,##0'
    ws['B13'].alignment = center_align
    ws['F13'].alignment = center_align

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
        ws[f'H{r}'].border = border
    ws[f'H{fila_subtotal+3}'].font = bold_font
    ws[f'H{fila_subtotal+3}'].fill = PatternFill(start_color="e6f2ff", end_color="e6f2ff", fill_type="solid")

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
    Asegúrese de haber actualizado las horas en el archivo de Google Sheets. 
    Cuando esté listo, presione el botón rojo para generar automáticamente todos los documentos de la quincena.
</div>
""", unsafe_allow_html=True)

if st.button("🚀 Procesar Nómina y Generar ZIP", use_container_width=True):
    with st.status("Procesando la información. Por favor espere...", expanded=True) as status:
        try:
            st.write("📥 Leyendo datos desde Google Sheets...")
            response = requests.get(GAS_URL)
            data = response.json()
            
            df_pagos = pd.DataFrame(data.get('pagos', []))
            df_fuera = pd.DataFrame(data.get('fuera_perimetro', [])) # Por si existe la hoja 3 que mencionó en el audio
            
            if df_pagos.empty:
                st.warning("No se encontraron datos en Google Sheets.")
                st.stop()
            
            # CRÍTICO: Limpiar los nombres de las columnas quitando espacios al inicio y final
            df_pagos.columns = df_pagos.columns.str.strip()
            if not df_fuera.empty:
                df_fuera.columns = df_fuera.columns.str.strip()

            st.write("📄 Creando Cuentas de Cobro y Documentos Equivalentes...")
            pagos_procesados = []
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                contador = 1
                for index, row in df_pagos.iterrows():
                    
                    # Extraer el valor base
                    valor_crudo = str(row.get('TOTAL A PAGAR', '0')).replace('$', '').replace(',', '').replace('.', '').strip()
                    try:
                        ingreso_base = float(valor_crudo)
                    except ValueError:
                        ingreso_base = 0.0

                    nombre_conductor = str(row.get('CONDUCTOR', '')).upper().strip()
                    ciudad = str(row.get('CIUDAD', '')).upper().strip()
                    
                    # Lógica para "Fuera de Perímetro" de Milton Javier Cortes
                    fuera_perimetro = 0.0
                    if "MILTON" in nombre_conductor:
                        # Si la hoja adicional de fuera de perímetro existe en la API, la buscamos
                        if not df_fuera.empty and 'CONDUCTOR' in df_fuera.columns:
                            match = df_fuera[df_fuera['CONDUCTOR'].str.upper().str.contains("MILTON", na=False)]
                            if not match.empty:
                                val_fp = str(match.iloc[0].get('VALOR', '0')).replace('$', '').replace(',', '').replace('.', '').strip()
                                try:
                                    fuera_perimetro = float(val_fp)
                                except:
                                    pass

                    ingreso_bruto_total = ingreso_base + fuera_perimetro

                    # Solo procesa si hay dinero a pagar
                    if ingreso_bruto_total > 0:
                        retefuente = ingreso_bruto_total * 0.01
                        ica = ingreso_bruto_total * 0.01 if ciudad == 'CALI' else 0.0
                        neto_a_pagar = ingreso_bruto_total - retefuente - ica
                        
                        nombre_titular = str(row.get('A NOMBRE DE QUIEN HACE CUENTA DE COBRO', row.get('NOMBRE TITULAR CUENTA BANCARIA', 'S/N')))
                        cedula_titular = str(row.get('CÉDULA DE CUENTA DE COBRO', row.get('CÉDULA TITULAR', '')))
                        banco = str(row.get('BANCO', ''))
                        tipo_cuenta = str(row.get('TIPO CUENTA', ''))
                        num_cuenta = str(row.get('NO. CUENTA', ''))
                        
                        # Archivo plano para Don José
                        pagos_procesados.append({
                            'CÉDULA': cedula_titular,
                            'NOMBRE': nombre_titular,
                            'BANCO': banco,
                            'TIPO CUENTA': tipo_cuenta,
                            'NÚMERO CUENTA': f"'{num_cuenta}", # La comilla evita que Excel convierta el número en formato científico
                            'VALOR A PAGAR': round(neto_a_pagar, 0)
                        })
                        
                        datos_doc = {
                            'id': str(contador).zfill(3),
                            'nombre_titular': nombre_titular,
                            'cedula_titular': cedula_titular,
                            'nombre_conductor': nombre_conductor,
                            'cedula_conductor': str(row.get('CEDULA', '')),
                            'ciudad': ciudad,
                            'corte_fechas': str(row.get('CORTE', '')),
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
                zip_file.writestr("Archivo_Plano_Bancario.csv", df_resultado.to_csv(index=False, sep=';', encoding='utf-8-sig'))
            
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
            st.subheader("Vista Previa - Consolidado Bancario para Don José")
            st.dataframe(df_resultado, use_container_width=True)

        except Exception as e:
            status.update(label="Ocurrió un error", state="error")
            st.error(f"Error en el proceso: {e}")
