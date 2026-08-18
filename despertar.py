import time
from playwright.sync_api import sync_playwright

def despertar_portal():
    print("Iniciando navegador virtual en la nube...")
    with sync_playwright() as p:
        # Lanzamos un navegador Chromium invisible (headless)
        browser = p.chromium.launch(headless=True)
        page = browser.new_workbook() if hasattr(browser, 'new_workbook') else browser.new_page()
        
        # URL
        url = "https://automatizacioncuentascobro-mc522jamiklvxlltpjhb8g.streamlit.app/" 
        
        print(f"Visitando el portal: {url}")
        page.goto(url)
        
        # Esperamos 20 segundos para darle tiempo al servidor de Streamlit de despertar por completo
        print("Esperando a que cargue la interfaz completa...")
        time.sleep(20)
        
        # Tomamos el título de la página para confirmar que cargó con éxito
        print(f"Portal despierto. Título de la app: '{page.title()}'")
        browser.close()

if __name__ == "__main__":
    despertar_portal()
