#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
============================================================
                    ROBOT SAT - FUTURELOGISTIC
============================================================
Autor: Anthony Cruz
Fecha: Febrero 2026
Descripción:
    Este script automatiza la extracción, transformación y carga
    (ETL) de los reportes de despacho generados por el sistema
    SAT de FutureLogistic S.A. Realiza las siguientes tareas:

    1. Inicia sesión automáticamente en el sistema SAT.
    2. Navega al módulo de logística y ejecuta el reporte
       de despachos con un rango de fechas predefinido.
    3. Descarga el archivo Excel generado.
    4. Ejecuta el proceso ETL que valida, transforma y carga
       los datos en la base de datos SQL Server.
    5. Muestra una ventana emergente indicando éxito o error.

    El script está diseñado para ejecutarse en un bucle continuo
    con un intervalo de 1 minuto, permitiendo una actualización
    constante de los datos. La ventana de consola se minimiza
    automáticamente al inicio para no interferir con el usuario.
============================================================
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


# ============================================================
# MINIMIZAR VENTANA DE CONSOLA (al inicio)
# ============================================================
if sys.platform == "win32":
    try:
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)  # 6 = SW_MINIMIZE
    except:
        pass


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================
class Config:
    USUARIO = "supervisor"
    CONTRASENA = "distribucion2024"
    RUTA_HTML = "C:/Users/yaps2/OneDrive - Ormesby Primary/Escritorio/SistemaSAT_Prueba/sap_futurelogistic.html"
    URL_SAT = f"file:///{RUTA_HTML.replace('\\', '/')}"
    CARPETA_REPORTES = "C:/Users/yaps2/OneDrive - Ormesby Primary/Escritorio/SistemaSAT_Prueba/REPORTES"
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


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================
def guardar_log(paso, mensaje):
    try:
        ruta_log = "C:\\temp\\robot_log.txt"
        os.makedirs(os.path.dirname(ruta_log), exist_ok=True)
        with open(ruta_log, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')} - {paso}: {mensaje}\n")
    except:
        pass


def mostrar_mensaje_exito():
    try:
        ctypes.windll.user32.MessageBoxW(
            0,
            "PROCESO COMPLETADO CON EXITO\n\nLos datos han sido descargados y cargados correctamente en la base de datos.",
            "FutureLogistic S.A. - Robot SAT",
            0x40 | 0x0
        )
    except:
        pass


def mostrar_mensaje_error():
    try:
        ctypes.windll.user32.MessageBoxW(
            0,
            "ERROR EN EL PROCESO\n\nLos datos no han podido ser cargados correctamente en la base de datos.",
            "FutureLogistic S.A. - Robot SAT",
            0x10 | 0x0
        )
    except:
        pass


# ============================================================
# FUNCIONES DE AUTOMATIZACIÓN
# ============================================================
def crear_estructura_carpetas():
    if not os.path.exists(Config.CARPETA_REPORTES):
        os.makedirs(Config.CARPETA_REPORTES, exist_ok=True)
    return True


def iniciar_navegador():
    try:
        opciones = Options()
        opciones.add_argument("--start-maximized")
        opciones.add_argument("--disable-notifications")
        opciones.add_argument("--disable-popup-blocking")
        opciones.add_argument("--disable-blink-features=AutomationControlled")
        opciones.add_experimental_option("excludeSwitches", ["enable-automation"])
        opciones.add_experimental_option("useAutomationExtension", False)
        preferencias = {
            "download.default_directory": Config.CARPETA_DESCARGAS,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False,
            "safebrowsing.disable_download_protection": True,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_settings.popups": 0,
            "profile.content_settings.exceptions.automatic_downloads.*.setting": 1
        }
        opciones.add_experimental_option("prefs", preferencias)
        servicio = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=servicio, options=opciones)
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
        campo_usuario = WebDriverWait(driver, Config.ESPERA_BOTON).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, Config.SELECTORES['usuario']))
        )
        campo_usuario.clear()
        campo_usuario.send_keys(Config.USUARIO)

        campo_contrasena = driver.find_element(By.CSS_SELECTOR, Config.SELECTORES['contrasena'])
        campo_contrasena.clear()
        campo_contrasena.send_keys(Config.CONTRASENA)

        boton = driver.find_element(By.CSS_SELECTOR, Config.SELECTORES['boton_login'])
        boton.click()
        time.sleep(Config.ESPERA_LOGIN)
        return True
    except:
        return False


def establecer_fechas(driver):
    try:
        fecha_desde = "2024-10-15"
        fecha_hasta = "2024-10-16"
        driver.execute_script(f"""
            document.getElementById('fechaDesde').value = '{fecha_desde}';
            document.getElementById('fechaHasta').value = '{fecha_hasta}';
        """)
        time.sleep(1)
        return True
    except:
        return False


def descargar_reporte_excel(driver):
    try:
        time.sleep(3)
        boton = WebDriverWait(driver, Config.ESPERA_BOTON).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, Config.SELECTORES['boton_descarga']))
        )
        archivos_antes = set()
        if os.path.exists(Config.CARPETA_DESCARGAS):
            for f in os.listdir(Config.CARPETA_DESCARGAS):
                if f.lower().endswith((".xlsx", ".xls")):
                    archivos_antes.add(f)
        boton.click()
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
        inicio = time.time()
        while (time.time() - inicio) < Config.ESPERA_DESCARGA:
            time.sleep(2)
            archivos_actuales = []
            for archivo in os.listdir(Config.CARPETA_DESCARGAS):
                if archivo.lower().endswith(('.xlsx', '.xls')):
                    archivos_actuales.append(archivo)
            nuevos = [a for a in archivos_actuales if a not in archivos_previos]
            for archivo in nuevos:
                ruta = os.path.join(Config.CARPETA_DESCARGAS, archivo)
                try:
                    tam1 = os.path.getsize(ruta)
                    time.sleep(1)
                    tam2 = os.path.getsize(ruta)
                    if tam1 == tam2 and tam1 > 0:
                        return archivo
                except:
                    continue
        return None
    except:
        return None


def procesar_archivo_descargado(nombre_archivo):
    try:
        origen = os.path.join(Config.CARPETA_DESCARGAS, nombre_archivo)
        if not os.path.exists(origen):
            return None
        if not os.path.exists(Config.CARPETA_REPORTES):
            os.makedirs(Config.CARPETA_REPORTES, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_base, extension = os.path.splitext(nombre_archivo)
        nombre_final = f"SAT_{nombre_base}_{timestamp}{extension}"
        destino = os.path.join(Config.CARPETA_REPORTES, nombre_final)
        contador = 1
        while os.path.exists(destino):
            nombre_final = f"SAT_{nombre_base}_{timestamp}_{contador}{extension}"
            destino = os.path.join(Config.CARPETA_REPORTES, nombre_final)
            contador += 1
        shutil.move(origen, destino)
        return destino
    except:
        return None


def ejecutar_etl():
    try:
        ruta_etl = os.path.join(os.path.dirname(__file__), "etl_cargar_excel.py")
        if not os.path.exists(ruta_etl):
            return False
        resultado = subprocess.run(
            [sys.executable, ruta_etl, "--silent"],
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        with open("C:\\temp\\etl_salida.txt", "w", encoding="utf-8") as f:
            f.write(f"=== CÓDIGO ===\n{resultado.returncode}\n\n")
            f.write(f"=== STDOUT ===\n{resultado.stdout}\n\n")
            f.write(f"=== STDERR ===\n{resultado.stderr}\n")
        return resultado.returncode == 0
    except:
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
        if not establecer_fechas(driver):
            driver.quit()
            return False
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


def tarea_cada_minuto():
    """Ejecuta el proceso completo de descarga y ETL."""
    if ejecutar_descarga_completa():
        mostrar_mensaje_exito()
    else:
        mostrar_mensaje_error()
    time.sleep(60)


# ============================================================
# PUNTO DE ENTRADA PRINCIPAL
# ============================================================
if __name__ == "__main__":
    # Crear carpeta de logs si no existe
    try:
        os.makedirs("C:\\temp", exist_ok=True)
    except:
        pass

    print("=" * 60)
    print("🤖 ROBOT SAT INICIADO")
    print("=" * 60)
    print("Modo automático: cada 1 minuto")
    print("La consola está minimizada.")
    print("Para detener, cierra esta ventana o presiona Ctrl+C")
    print("=" * 60)

    try:
        while True:
            tarea_cada_minuto()
    except KeyboardInterrupt:
        print("\n🛑 Robot detenido por el usuario")
        sys.exit(0)