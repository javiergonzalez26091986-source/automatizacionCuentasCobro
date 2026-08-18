import streamlit as st
import pandas as pd
import requests
import io

# Configuración de la página
st.set_page_config(page_title="SERGEM - Gestión de Pagos", page_icon="sergemLogo.ico", layout="wide")

# Cabecera
col1, col2 = st.columns([1, 4])
with col1:
    try:
        st.image("sergemLogo.png", width=150)
    except:
        pass
with col2:
    st.title("Generador Automático de Pagos y Formatos")
    st.markdown("**SERGEM Mensajería S.A.S.**")

st.divider()

# Input para la conexión con Sheets
GAS_URL = st.text_input(
    "URL de conexión (Google Apps Script):", 
    placeholder="Pega aquí la URL que termina en /exec..."
)

if st.button("Sincronizar Datos y Calcular Nómina", type="primary"):
    if GAS_URL:
        with st.spinner("Descargando datos de Google Sheets..."):
            try:
                # 1. Consumir la API de Google Sheets
                response = requests.get(GAS_URL)
                data = response.json()
                
                df_cuentas = pd.DataFrame(data['cuentas'])
                df_pagos = pd.DataFrame(data['pagos'])
                
                # 2. Lógica de negocio y cruce de datos
                # Aseguramos que la columna CIUDAD esté en mayúsculas para evitar errores
                df_pagos['CIUDAD'] = df_pagos['CIUDAD'].astype(str).str.upper().str.strip()
                
                # Hacemos un merge (BUSCARV) para traer los datos bancarios al df de pagos
                df_completo = pd.merge(
                    df_pagos, 
                    df_cuentas[['NOMBRE DEL CONDUCTOR (PLANILLA)', 'NOMBRE DEL TITULAR DE LA CUENTA', 'DOCUMENTO DEL TITULAR DE LA CUENTA', 'NOMBRE DE LA ENTIDAD FINANCIERA - BANCO', 'TIPO DE CUENTA (AHORROS-CORRIENTE)', 'NUMERO DE CUENTA']], 
                    left_on='CONDUCTOR', 
                    right_on='NOMBRE DEL CONDUCTOR (PLANILLA)', 
                    how='left'
                )

                # 3. Cálculos Matemáticos
                pagos_procesados = []
                for index, row in df_completo.iterrows():
                    # Solo procesamos si hay un valor a pagar
                    if pd.notna(row.get('TOTAL A PAGAR')) and row.get('TOTAL A PAGAR', 0) > 0:
                        
                        ingreso_bruto = float(row['TOTAL A PAGAR']) # Este valor ya viene del Excel base
                        
                        # Retefuente: 1% para todos
                        retefuente = ingreso_bruto * 0.01
                        
                        # ICA: 1% SOLO para Cali
                        ica = ingreso_bruto * 0.01 if row['CIUDAD'] == 'CALI' else 0.0
                        
                        neto_a_pagar = ingreso_bruto - retefuente - ica
                        
                        pagos_procesados.append({
                            'NIT_TITULAR': row['DOCUMENTO DEL TITULAR DE LA CUENTA'],
                            'NOMBRE_TITULAR': row['NOMBRE DEL TITULAR DE LA CUENTA'],
                            'BANCO': row['NOMBRE DE LA ENTIDAD FINANCIERA - BANCO'],
                            'TIPO_CUENTA': row['TIPO DE CUENTA (AHORROS-CORRIENTE)'],
                            'NUM_CUENTA': row['NUMERO DE CUENTA'],
                            'NETO_A_PAGAR': round(neto_a_pagar, 0)
                        })
                
                df_resultado = pd.DataFrame(pagos_procesados)
                
                st.success(f"¡Cálculos finalizados! Se procesaron {len(df_resultado)} pagos.")
                
                # 4. Mostrar vista previa y botón de descarga del Plano Bancario
                st.subheader("Vista Previa - Archivo Plano para el Banco")
                st.dataframe(df_resultado)
                
                # Convertir a CSV para descarga
                csv = df_resultado.to_csv(index=False, sep=';').encode('utf-8')
                st.download_button(
                    label="Descargar Archivo Plano Banco (CSV)",
                    data=csv,
                    file_name='Archivo_Plano_Banco_Quincena.csv',
                    mime='text/csv',
                )
                
                st.info("💡 En la siguiente fase, conectaremos las librerías `weasyprint` y `openpyxl` aquí mismo para descargar también el archivo ZIP con todos los PDFs de cobro y Excel equivalentes generados en bloque.")
                
            except Exception as e:
                st.error(f"Error procesando los datos. Verifica la URL o la estructura del Excel. Detalles: {e}")
    else:
        st.warning("Por favor ingresa la URL de conexión de Apps Script.")
