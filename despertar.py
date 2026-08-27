from playwright.sync_api import sync_playwright

def despertar_portal():
    print("🚀 Iniciando navegador virtual en la nube (modo silencioso)...")
    with sync_playwright() as p:
        # Lanzamos un navegador Chromium invisible (headless)
        browser = p.chromium.launch(headless=True)
        
        # Creamos una nueva pestaña en el navegador
        page = browser.new_page()
        
        # URL de tu app Streamlit
        url = "https://automatizacioncuentascobro-mc522jamiklvxlltpjhb8g.streamlit.app/" 
        
        print(f"🔗 Visitando el portal: {url}")
        
        try:
            # Ir a la URL y esperar a que la red esté inactiva (lo que confirma carga básica)
            page.goto(url, wait_until="networkidle")
            
            # --- LA CLAVE: BUSCAR Y PULSAR EL BOTÓN DE DESPERTAR ---
            # Buscamos el botón específico que aparece en la pantalla de "Zzzz"
            # Utilizamos un selector de texto exacto para ser precisos.
            wake_up_button = page.locator('text="Yes, get this app back up!"')
            
            if wake_up_button.count() > 0:
                print("⚠️  ¡Portal dormido! Haciendo clic en el botón para despertarlo...")
                wake_up_button.click()
                
                # Después de pulsar el botón, el servidor tardará un poco en arrancar.
                # En lugar de time.sleep(20), vamos a esperar a que cargue un elemento real de tu app.
                print("⏳ Esperando a que el servidor de Streamlit arranque la aplicación...")
                # Por ejemplo, podemos esperar a que aparezca tu título principal 'Portal de Pagos' (h1)
                page.locator('h1:has-text("Portal de Pagos")').wait_for()
                print(f"🎉 ¡Portal despertado con éxito!")
            else:
                # Si el botón no aparece, es que el portal ya estaba despierto.
                print("✅ El portal ya estaba despierto y operativo.")
                # Confirmamos buscando el título de tu app.
                page.locator('h1:has-text("Portal de Pagos")').wait_for()
            
            # Para mayor seguridad, imprimimos el título de la página
            print(f"Título de la app verificado: '{page.title()}'")
            
        except Exception as e:
            print(f"❌ Ocurrió un error al intentar despertar el portal: {e}")
        
        finally:
            # Cerramos el navegador para liberar recursos
            browser.close()
            print("🛑 Navegador virtual cerrado.")

if __name__ == "__main__":
    despertar_portal()
