import streamlit as st
import pandas as pd
import requests
import io
import zipfile
import base64
from weasyprint import HTML
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
import os

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
# Se extrae de los mensajes anteriores la URL proporcionada por el usuario
GAS_URL = "https://script.google.com/macros/s/AKfycbyqJtrmVdNT1rxTobg6q_WoJCwMpp40hdIzJeEm4dKNLBgDVxwEY95T0EIoBu_qo8FB/exec"

# --- FUNCIONES DE GENERACIÓN ---
@st.cache_data
def get_base64_logo():
    try:
        with open('sergemLogo.png', 'rb') as f:
            return base64.b64encode(f.read()).decode('utf-8')
    except Exception:
        return ""

LOGO_B64 = get_base64_logo()

def generar_pdf_cuenta_cobro(datos):
    html_content = f'''
    <html>
    <head>
    <style>
        @page {{ size: Letter; margin: 20mm; background-color: #ffffff; }}
        body {{ font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #333; }}
        .header {{ display: flex; align-items: center; border-bottom: 2px solid #E3000F; padding-bottom: 10px; margin-bottom: 20px; }}
        .header img {{ width: 180px; }}
        .header-text {{ margin-left: 20px; text-align: right; width: 100%; }}
        .header-text h1 {{ margin: 0; font-size: 16pt; color: #333; }}
        .header-text p {{ margin: 0; font-size: 10pt; color: #777; }}
        .title {{ font-size: 18pt; font-weight: bold; color: #E3000F; text-align: center; margin-top: 10px; margin-bottom: 5px; }}
        .subtitle {{ text-align: center; font-size: 11pt; color: #555; margin-bottom: 30px; }}
        .content {{ font-size: 11pt; line-height: 1.6; }}
        .box {{ border: 1px solid #ddd; padding: 20px; border-radius: 8px; margin-top: 20px; background-color: #fdfbf7; }}
        .box p {{ margin: 5px 0; }}
        .total {{ font-size: 16pt; font-weight: bold; color: #E3000F; text-align: right; margin-top: 15px; border-top: 1px solid #ddd; padding-top: 10px; }}
        .bank-info {{ margin-top: 20px; background-color: #f4f4f4; padding: 15px; border-left: 4px solid #E3000F; }}
        .signature {{ margin-top: 80px; border-top: 1px solid #333; width: 250px; padding-top: 5px; text-align: center; font-weight: bold; }}
    </style>
    </head>
    <body>
        <div class="header">
            <img src="data:image/png;base64,{LOGO_B64}" />
            <div class="header-text">
                <h1>SERGEM MENSAJERIA S.A.S.</h1>
                <p>NIT. 900.561.833-1</p>
            </div>
        </div>
        <div class="title">CUENTA DE COBRO</div>
        <div class="subtitle">Documento No. {datos['id']} | Fecha: {datos['fecha_actual']}</div>
        <div class="content">
            <p><strong>DEBE A:</strong></p>
            <p style="font-size: 14pt; font-weight: bold; margin: 5px 0;">{datos['nombre_titular']}</p>
            <p style="margin: 0;">C.C. {datos['cedula_titular']}</p>
            
            <div class="box">
                <p><strong>Concepto:</strong> Servicio de mensajería prestado en el corte del {datos['corte_fechas']}.</p>
                <p><strong>Conductor:</strong> {datos['nombre_conductor']} (C.C. {datos['cedula_conductor']})</p>
                <p><strong>Ciudad de Operación:</strong> {datos['ciudad']}</p>
                <br>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 5px 0;">Valor Base Negociado:</td><td style="text-align: right;">$ {datos['ingreso_bruto']:,.0f}</td></tr>
                    <tr><td style="padding: 5px 0; color: #d9534f;">Retención en la Fuente (1%):</td><td style="text-align: right; color: #d9534f;">- $ {datos['retefuente']:,.0f}</td></tr>
                    <tr><td style="padding: 5px 0; color: #d9534f;">Descuento ICA (1%):</td><td style="text-align: right; color: #d9534f;">- $ {datos['ica']:,.0f}</td></tr>
                </table>
                <div class="total">NETO A PAGAR: $ {datos['neto_pagar']:,.0f}</div>
            </div>

            <div class="bank-info">
                <p style="margin:0 0 10px 0;"><strong>Por favor consignar en la siguiente cuenta bancaria:</strong></p>
                <p style="margin: 2px 0;"><strong>Banco:</strong> {datos['banco']}</p>
                <p style="margin: 2px 0;"><strong>Tipo de Cuenta:</strong> {datos['tipo_cuenta']}</p>
                <p style="margin: 2px 0;"><strong>Número:</strong> {datos['num_cuenta']}</p>
                <p style="margin: 2px 0;"><strong>Titular:</strong> {datos['nombre_titular']}</p>
            </div>

            <div class="signature">
                {datos['nombre_titular']}<br>C.C. {datos['cedula_titular']}
            </div>
        </div>
    </body>
    </html>
    '''
    return HTML(string=html_content).write_pdf()

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

# Botón Principal
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
            
            # Botón de descarga destacado
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
