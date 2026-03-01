#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
============================================================
                    ROBOT SAT - FUTURELOGISTIC
============================================================
Autor: Anthony Cruz
Fecha: Febrero 2026
Descripción:
    Este script automatiza la descarga de reportes de despacho
    desde el sistema SAT de FutureLogistic S.A. Inicia sesión,
    navega al módulo de logística, ejecuta el reporte con fechas
    específicas y descarga el archivo Excel. Luego ejecuta el ETL
    para cargar los datos en SQL Server.
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
# CONFIGURACIÓN GENERAL (ajustar según el entorno)
# ============================================================
class Config:
    # Credenciales de acceso al SAT (simuladas)
    USUARIO = "supervisor"
    CONTRASENA = "distribucion2024"

    # Ruta del archivo HTML que simula el SAT (local)
    RUTA_HTML = "C:/Users/yaps2/OneDrive - Ormesby Primary/Escritorio/SistemaSAT_Prueba/sap_futurelogistic.html"
    URL_SAT = f"file:///{RUTA_HTML.replace('\\', '/')}"

    # Carpetas de trabajo
    CARPETA_REPORTES = "C:/Users/yaps2/OneDrive - Ormesby Primary/Escritorio/SistemaSAT_Prueba/REPORTES"
    CARPETA_DESCARGAS = os.path.join(os.path.expanduser("~"), "Downloads")

    # Tiempos de espera (ajustados empíricamente)
    ESPERA_CARGA = 3
    ESPERA_LOGIN = 3
    ESPERA_DESCARGA = 20
    ESPERA_BOTON = 10

    # Selectores CSS (obtenidos del HTML del SAT)
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
    """
    Guarda mensajes en un archivo de log para depuración.
    Útil cuando el robot falla y no hay ventanas emergentes.
    """
    try:
        ruta_log = "C:\\temp\\robot_log.txt"
        os.makedirs(os.path.dirname(ruta_log), exist_ok=True)
        with open(ruta_log, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')} - {paso}: {mensaje}\n")
    except:
        pass  # Si no se puede escribir, no detenemos el proceso


def mostrar_mensaje_exito():
    """Ventana emergente cuando todo sale bien."""
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
    """Ventana emergente cuando algo falla."""
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
# FUNCIONES DE AUTOMATIZACIÓN CON SELENIUM
# ============================================================
def crear_estructura_carpetas():
    """Crea la carpeta REPORTES si no existe."""
    if not os.path.exists(Config.CARPETA_REPORTES):
        os.makedirs(Config.CARPETA_REPORTES, exist_ok=True)
    return True


def iniciar_navegador():
    """
    Inicia Chrome con las opciones necesarias para evitar
    detecciones de automatización y configurar descargas.
    """
    try:
        opciones = Options()
        opciones.add_argument("--start-maximized")
        opciones.add_argument("--disable-notifications")
        opciones.add_argument("--disable-popup-blocking")
        opciones.add_argument("--disable-blink-features=AutomationControlled")
        opciones.add_experimental_option("excludeSwitches", ["enable-automation"])
        opciones.add_experimental_option("useAutomationExtension", False)

        # Configurar carpeta de descargas
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
        driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        return driver
    except Exception as e:
        guardar_log("ERROR", f"No se pudo iniciar el navegador: {str(e)}")
        return None


def verificar_archivo_html():
    """Verifica que el archivo HTML del SAT exista."""
    return os.path.exists(Config.RUTA_HTML)


def navegar_al_sat(driver):
    """Navega a la URL del SAT (local o remota)."""
    try:
        driver.get(Config.URL_SAT)
        time.sleep(Config.ESPERA_CARGA)
        return True
    except:
        return False


def realizar_login(driver):
    """Llena el formulario de login y hace clic en el botón."""
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
    """
    Establece el rango de fechas para el reporte.
    Por ahora son fijas, pero se puede modificar para que sean dinámicas.
    """
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
    """
    Busca el botón de exportar Excel, hace clic y espera a que
    el archivo aparezca en la carpeta de descargas.
    """
    try:
        time.sleep(3)
        boton = WebDriverWait(driver, Config.ESPERA_BOTON).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, Config.SELECTORES['boton_descarga']))
        )

        # Tomar lista de archivos antes de la descarga
        archivos_antes = set()
        if os.path.exists(Config.CARPETA_DESCARGAS):
            for f in os.listdir(Config.CARPETA_DESCARGAS):
                if f.lower().endswith((".xlsx", ".xls")):
                    archivos_antes.add(f)

        boton.click()
        time.sleep(8)  # Espera generosa para que termine la descarga

        archivo_descargado = monitorear_descarga(archivos_antes)
        if archivo_descargado:
            time.sleep(2)
            return procesar_archivo_descargado(archivo_descargado) is not None
        return False
    except:
        return False


def monitorear_descarga(archivos_previos):
    """
    Monitorea la carpeta de descargas hasta que aparezca un archivo nuevo.
    """
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
    """
    Mueve el archivo descargado de Downloads a REPORTES,
    renombrándolo con un timestamp para evitar colisiones.
    """
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

        # Evitar sobrescritura (por si acaso)
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
    """
    Ejecuta el script ETL en modo silencioso y captura su salida.
    Retorna True si el ETL terminó con código 0 (éxito).
    """
    try:
        ruta_etl = os.path.join(os.path.dirname(__file__), "etl_cargar_excel.py")
        if not os.path.exists(ruta_etl):
            guardar_log("ETL", "Archivo no encontrado")
            return False

        resultado = subprocess.run(
            [sys.executable, ruta_etl, "--silent"],
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        # Guardar salida por si hay que depurar
        with open("C:\\temp\\etl_salida.txt", "w", encoding="utf-8") as f:
            f.write(f"=== CÓDIGO ===\n{resultado.returncode}\n\n")
            f.write(f"=== STDOUT ===\n{resultado.stdout}\n\n")
            f.write(f"=== STDERR ===\n{resultado.stderr}\n")

        guardar_log("ETL", f"Código retorno: {resultado.returncode}")
        return resultado.returncode == 0

    except Exception as e:
        guardar_log("ETL", f"Excepción: {str(e)}")
        return False


def ejecutar_descarga_completa():
    """
    Orquesta todo el proceso de descarga:
    - Verifica HTML
    - Inicia navegador
    - Login
    - Navegación
    - Fechas
    - Ejecución de reporte
    - Descarga
    - ETL
    """
    driver = None
    try:
        if not verificar_archivo_html():
            guardar_log("ERROR", "HTML no encontrado")
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

        # Ir a la pantalla de logística y seleccionar reporte
        try:
            driver.execute_script("showScreen('logistics')")
            time.sleep(2)
            driver.execute_script("selectReport('ZFL_REP01')")
            time.sleep(2)
        except:
            pass  # Si falla, continuamos (a veces ya está seleccionado)

        if not establecer_fechas(driver):
            driver.quit()
            return False

        time.sleep(2)

        # Ejecutar reporte
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


# ============================================================
# PUNTO DE ENTRADA PRINCIPAL
# ============================================================
if __name__ == "__main__":
    # Crear carpeta de logs si no existe
    try:
        os.makedirs("C:\\temp", exist_ok=True)
    except:
        pass

    if ejecutar_descarga_completa():
        mostrar_mensaje_exito()
    else:
        mostrar_mensaje_error()