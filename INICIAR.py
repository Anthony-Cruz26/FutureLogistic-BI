#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ROBOT SAT - VERSIÓN FINAL
Nombre: INICIAR.py
"""

import os
import sys
import time
import shutil
import subprocess
import ctypes
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# ============================================
# CONFIGURACIÓN
# ============================================
class Config:
    USUARIO = "supervisor"
    CONTRASENA = "distribucion2024"
    RUTA_HTML = "C:/Users/yaps2/OneDrive - Ormesby Primary/Escritorio/SistemaSAT_Prueba/sap_futurelogistic.html"
    URL_SAT = f"file:///{RUTA_HTML.replace('\\', '/')}"
    CARPETA_BASE = "C:/Users/yaps2/OneDrive - Ormesby Primary/Escritorio/SistemaSAT_Prueba/REPORTES"
    CARPETA_DESCARGAS = os.path.join(os.path.expanduser("~"), "Downloads")
    
    ESPERA_CARGA = 3
    ESPERA_LOGIN = 3
    ESPERA_DESCARGA = 20
    ESPERA_BOTON = 10
    
    SELECTORES = {
        'usuario': "input#loginUser",
        'contrasena': "input#loginPass",
        'boton_login': "button.btn-login",
        'boton_descarga': "button.toolbar-btn[title='Exportar Excel']",
        'panel_principal': "#mainApp"
    }

# ============================================
# FUNCIÓN PARA GUARDAR LOGS (OPCIONAL)
# ============================================
def guardar_log(paso, mensaje):
    try:
        with open("C:\\temp\\robot_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')} - {paso}: {mensaje}\n")
    except:
        pass

# ============================================
# FUNCIONES DE MENSAJES (VENTANAS EMERGENTES)
# ============================================
def mostrar_mensaje_exito():
    """Muestra ventana emergente con mensaje de éxito"""
    try:
        ctypes.windll.user32.MessageBoxW(0, 
            "PROCESO COMPLETADO CON EXITO\n\nLos datos han sido descargados y cargados correctamente en la base de datos.", 
            "FutureLogistic S.A. - Robot SAT", 0x40 | 0x0)
    except:
        pass

def mostrar_mensaje_error():
    """Muestra ventana emergente con mensaje de error personalizado"""
    try:
        ctypes.windll.user32.MessageBoxW(0, 
            "ERROR EN EL PROCESO\n\nLos datos no han podido ser cargados correctamente en la base de datos.", 
            "FutureLogistic S.A. - Robot SAT", 0x10 | 0x0)
    except:
        pass

# ============================================
# FUNCIONES PRINCIPALES
# ============================================
def crear_estructura_carpetas():
    carpeta_reportes = Config.CARPETA_BASE
    if not os.path.exists(carpeta_reportes):
        os.makedirs(carpeta_reportes, exist_ok=True)
    return True

def iniciar_navegador():
    try:
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)
        
        prefs = {
            "download.default_directory": Config.CARPETA_DESCARGAS,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False,
            "safebrowsing.disable_download_protection": True,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_settings.popups": 0,
            "profile.content_settings.exceptions.automatic_downloads.*.setting": 1
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
    except:
        return None

def verificar_archivo_html():
    return os.path.exists(Config.RUTA_HTML)

def navegar_al_sat(driver):
    try:
        driver.get(Config.URL_SAT)
        time.sleep(Config.ESPERA_CARGA)
        return True
    except:
        return False

def realizar_login(driver):
    try:
        time.sleep(2)
        usuario_field = WebDriverWait(driver, Config.ESPERA_BOTON).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, Config.SELECTORES['usuario']))
        )
        usuario_field.clear()
        usuario_field.send_keys(Config.USUARIO)
        
        contrasena_field = driver.find_element(By.CSS_SELECTOR, Config.SELECTORES['contrasena'])
        contrasena_field.clear()
        contrasena_field.send_keys(Config.CONTRASENA)
        
        boton_login = driver.find_element(By.CSS_SELECTOR, Config.SELECTORES['boton_login'])
        boton_login.click()
        time.sleep(Config.ESPERA_LOGIN)
        return True
    except:
        return False

def establecer_fechas(driver):
    try:
        fecha_ayer = "2024-10-15"
        fecha_hoy = "2024-10-16"
        driver.execute_script(f"""
            document.getElementById('fechaDesde').value = '{fecha_ayer}';
            document.getElementById('fechaHasta').value = '{fecha_hoy}';
        """)
        time.sleep(1)
        return True
    except:
        return False

def descargar_reporte_excel(driver):
    try:
        time.sleep(3)
        boton_descarga = WebDriverWait(driver, Config.ESPERA_BOTON).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, Config.SELECTORES['boton_descarga']))
        )
        
        archivos_antes = set()
        if os.path.exists(Config.CARPETA_DESCARGAS):
            for f in os.listdir(Config.CARPETA_DESCARGAS):
                if f.lower().endswith((".xlsx", ".xls")):
                    archivos_antes.add(f)
        
        boton_descarga.click()
        time.sleep(8)
        
        archivo_descargado = monitorear_descarga(archivos_antes)
        if archivo_descargado:
            time.sleep(2)
            return procesar_archivo_descargado(archivo_descargado) is not None
        return False
    except:
        return False

def monitorear_descarga(archivos_previos):
    try:
        tiempo_inicio = time.time()
        while (time.time() - tiempo_inicio) < Config.ESPERA_DESCARGA:
            time.sleep(2)
            archivos_actuales = []
            for archivo in os.listdir(Config.CARPETA_DESCARGAS):
                if archivo.lower().endswith(('.xlsx', '.xls')):
                    archivos_actuales.append(archivo)
            
            nuevos_archivos = [a for a in archivos_actuales if a not in archivos_previos]
            for archivo in nuevos_archivos:
                ruta_completa = os.path.join(Config.CARPETA_DESCARGAS, archivo)
                try:
                    tamano1 = os.path.getsize(ruta_completa)
                    time.sleep(1)
                    tamano2 = os.path.getsize(ruta_completa)
                    if tamano1 == tamano2 and tamano1 > 0:
                        return archivo
                except:
                    continue
        return None
    except:
        return None

def procesar_archivo_descargado(nombre_archivo):
    try:
        ruta_origen = os.path.join(Config.CARPETA_DESCARGAS, nombre_archivo)
        if not os.path.exists(ruta_origen):
            return None
        
        carpeta_destino = Config.CARPETA_BASE
        if not os.path.exists(carpeta_destino):
            os.makedirs(carpeta_destino, exist_ok=True)
        
        fecha_actual = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_base, extension = os.path.splitext(nombre_archivo)
        nombre_final = f"SAT_{nombre_base}_{fecha_actual}{extension}"
        ruta_destino = os.path.join(carpeta_destino, nombre_final)
        
        contador = 1
        while os.path.exists(ruta_destino):
            nombre_final = f"SAT_{nombre_base}_{fecha_actual}_{contador}{extension}"
            ruta_destino = os.path.join(carpeta_destino, nombre_final)
            contador += 1
        
        shutil.move(ruta_origen, ruta_destino)
        return ruta_destino
    except:
        return None

def ejecutar_etl():
    """Ejecuta el ETL, guarda la salida y retorna True si fue exitoso"""
    try:
        ruta_etl = os.path.join(os.path.dirname(__file__), "etl_cargar_excel.py")
        
        if not os.path.exists(ruta_etl):
            guardar_log("ETL", f"Archivo no encontrado: {ruta_etl}")
            return False
        
        # Ejecutar ETL y capturar salida
        resultado = subprocess.run(
            [sys.executable, ruta_etl, "--silent"],
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        
        # Guardar la salida del ETL para depuración
        with open("C:\\temp\\etl_salida.txt", "w", encoding="utf-8") as f:
            f.write(f"=== CÓDIGO DE RETORNO ===\n{resultado.returncode}\n\n")
            f.write(f"=== STDOUT ===\n{resultado.stdout}\n\n")
            f.write(f"=== STDERR ===\n{resultado.stderr}\n")
        
        guardar_log("ETL", f"Código retorno: {resultado.returncode}")
        
        # El ETL retorna 0 si fue exitoso, 1 si hubo errores
        return resultado.returncode == 0
        
    except Exception as e:
        guardar_log("ETL", f"Excepción: {str(e)}")
        return False

def ejecutar_descarga_completa():
    driver = None
    try:
        if not verificar_archivo_html():
            return False
        
        driver = iniciar_navegador()
        if not driver:
            return False
        
        if not navegar_al_sat(driver):
            driver.quit()
            return False
        
        if not realizar_login(driver):
            driver.quit()
            return False
        
        time.sleep(3)
        
        try:
            driver.execute_script("showScreen('logistics')")
            time.sleep(2)
            driver.execute_script("selectReport('ZFL_REP01')")
            time.sleep(2)
        except:
            pass
        
        establecer_fechas(driver)
        time.sleep(2)
        
        try:
            driver.execute_script("ejecutarReporte()")
            time.sleep(5)
        except:
            pass
        
        if not descargar_reporte_excel(driver):
            driver.quit()
            return False
        
        driver.quit()
        
        # ✅ Verificamos el resultado REAL del ETL
        if not ejecutar_etl():
            return False
        
        return True
        
    except:
        if driver:
            try:
                driver.quit()
            except:
                pass
        return False

# ============================================
# EJECUCIÓN PRINCIPAL
# ============================================
if __name__ == "__main__":
    try:
        if not os.path.exists("C:\\temp"):
            os.makedirs("C:\\temp")
    except:
        pass
    
    if ejecutar_descarga_completa():
        mostrar_mensaje_exito()
    else:
        mostrar_mensaje_error()