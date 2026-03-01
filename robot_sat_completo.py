#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
============================================================
        ROBOT SAT COMPLETO - VERSIÓN CON CONSOLA
============================================================
Autor: Anthony Cruz
Fecha: Febrero 2026
Descripción:
    Este script es la versión con interfaz de consola del robot.
    Permite ejecutar el proceso de descarga de forma manual,
    automática (programada) o en modo prueba. Es útil para
    depuración y para ver los mensajes en tiempo real.
============================================================
"""

import os
import sys
import time
import shutil
import subprocess
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

import schedule
import traceback


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================
class Config:
    # Credenciales
    USUARIO = "supervisor"
    CONTRASENA = "distribucion2024"

    # Ruta del SAT simulado
    RUTA_HTML = "C:/Users/yaps2/OneDrive - Ormesby Primary/Escritorio/SistemaSAT_Prueba/sap_futurelogistic.html"
    URL_SAT = f"file:///{RUTA_HTML.replace('\\', '/')}"

    # Carpetas
    CARPETA_REPORTES = "C:/Users/yaps2/OneDrive - Ormesby Primary/Escritorio/SistemaSAT_Prueba/REPORTES"
    CARPETA_DESCARGAS = os.path.join(os.path.expanduser("~"), "Downloads")

    # Tiempos de espera
    ESPERA_CARGA = 3
    ESPERA_LOGIN = 3
    ESPERA_DESCARGA = 20
    ESPERA_BOTON = 10

    # Programación (modo automático)
    HORA_INICIO = 7      # 7:00 AM
    HORA_FIN = 20        # 8:00 PM
    SOLO_LABORABLES = True

    # Selectores CSS (del HTML)
    SELECTORES = {
        'usuario': "input#loginUser",
        'contrasena': "input#loginPass",
        'boton_login': "button.btn-login",
        'boton_descarga': "button.toolbar-btn[title='Exportar Excel']",
        'panel_principal': "#mainApp"
    }


# ============================================================
# FUNCIONES DE CONFIGURACIÓN Y LOG
# ============================================================
def crear_estructura_carpetas():
    """Crea la carpeta REPORTES si no existe."""
    if not os.path.exists(Config.CARPETA_REPORTES):
        os.makedirs(Config.CARPETA_REPORTES, exist_ok=True)
        print(f"📁 Carpeta creada: {Config.CARPETA_REPORTES}")
    return True


def registrar_log(mensaje, nivel="INFO"):
    """
    Muestra mensajes en consola con un formato de tiempo.
    Los niveles determinan el ícono que se muestra.
    """
    iconos = {
        "INFO": "📝",
        "SUCCESS": "✅",
        "ERROR": "❌",
        "WARNING": "⚠️",
        "DEBUG": "🔧"
    }
    icono = iconos.get(nivel, "📝")
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {icono} {mensaje}")
    return True


# ============================================================
# FUNCIONES DE SELENIUM
# ============================================================
def iniciar_navegador():
    """Inicia Chrome con la configuración necesaria para automatización."""
    try:
        registrar_log("Iniciando navegador Chrome...", "INFO")

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

        # Ocultar el hecho de que estamos automatizando
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

        registrar_log("Navegador iniciado correctamente", "SUCCESS")
        return driver

    except Exception as e:
        registrar_log(f"Error al iniciar navegador: {str(e)}", "ERROR")
        return None


def verificar_archivo_html():
    """Verifica que el archivo HTML del SAT exista."""
    if not os.path.exists(Config.RUTA_HTML):
        registrar_log(f"ARCHIVO NO ENCONTRADO: {Config.RUTA_HTML}", "ERROR")
        return False
    tamano = os.path.getsize(Config.RUTA_HTML) / 1024
    registrar_log(f"SAT encontrado: {Config.RUTA_HTML} ({tamano:.1f} KB)", "SUCCESS")
    return True


def navegar_al_sat(driver):
    """Navega a la URL del SAT."""
    try:
        registrar_log(f"Navegando a: {Config.URL_SAT}", "INFO")
        driver.get(Config.URL_SAT)
        time.sleep(Config.ESPERA_CARGA)

        if "SISTEMA SAT" in driver.title or "SAP Logon" in driver.title:
            registrar_log("SAT cargado correctamente", "SUCCESS")
        else:
            registrar_log("SAT cargado pero título no reconocido", "WARNING")
        return True

    except Exception as e:
        registrar_log(f"Error navegando al SAT: {str(e)}", "ERROR")
        return False


def realizar_login(driver):
    """Llena el formulario de login y hace clic en el botón."""
    try:
        registrar_log("Realizando login automático...", "INFO")
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

        # Verificar si el login fue exitoso
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, Config.SELECTORES['panel_principal']))
            )
            registrar_log("Login exitoso", "SUCCESS")
        except:
            page_text = driver.page_source
            if "Sesión iniciada" in page_text or "mainApp" in page_text:
                registrar_log("Login exitoso", "SUCCESS")
            else:
                registrar_log("Login completado pero panel no detectado", "WARNING")

        return True

    except Exception as e:
        registrar_log(f"Error durante login: {str(e)}", "ERROR")
        return False


def establecer_fechas_ayer_hoy(driver):
    """
    Establece las fechas del reporte como ayer y hoy.
    Se puede modificar para usar fechas fijas si se prefiere.
    """
    try:
        hoy = datetime.now()
        ayer = hoy - timedelta(days=1)

        fecha_ayer = ayer.strftime("%Y-%m-%d")
        fecha_hoy = hoy.strftime("%Y-%m-%d")

        registrar_log(f"Estableciendo fechas: {fecha_ayer} → {fecha_hoy}", "INFO")

        driver.execute_script(f"""
            document.getElementById('fechaDesde').value = '{fecha_ayer}';
            document.getElementById('fechaHasta').value = '{fecha_hoy}';
        """)

        time.sleep(1)
        registrar_log("Fechas actualizadas correctamente", "SUCCESS")
        return True

    except Exception as e:
        registrar_log(f"Error estableciendo fechas: {str(e)}", "ERROR")
        return False


def descargar_reporte_excel(driver):
    """Hace clic en el botón de exportar y espera la descarga."""
    try:
        registrar_log("Buscando botón de descarga Excel...", "INFO")
        time.sleep(3)

        boton = WebDriverWait(driver, Config.ESPERA_BOTON).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, Config.SELECTORES['boton_descarga']))
        )

        # Lista de archivos antes de la descarga
        archivos_antes = set()
        if os.path.exists(Config.CARPETA_DESCARGAS):
            for f in os.listdir(Config.CARPETA_DESCARGAS):
                if f.lower().endswith((".xlsx", ".xls")):
                    archivos_antes.add(f)

        boton.click()
        registrar_log("Clic en botón de descarga realizado", "SUCCESS")
        time.sleep(8)

        archivo = monitorear_descarga(archivos_antes)

        if archivo:
            registrar_log(f"Descarga completada: {archivo}", "SUCCESS")
            time.sleep(2)
            return procesar_archivo_descargado(archivo) is not None
        else:
            registrar_log("Descarga iniciada pero archivo no detectado", "WARNING")
            return procesar_archivo_descargado() is not None

    except Exception as e:
        registrar_log(f"Error al descargar Excel: {str(e)}", "ERROR")
        return False


def monitorear_descarga(archivos_previos=None):
    """
    Monitorea la carpeta de descargas hasta que aparezca un archivo nuevo.
    """
    try:
        registrar_log("Monitoreando descargas...", "INFO")

        if not os.path.exists(Config.CARPETA_DESCARGAS):
            return None

        if archivos_previos is None:
            iniciales = set()
            for f in os.listdir(Config.CARPETA_DESCARGAS):
                if f.lower().endswith(('.xlsx', '.xls')):
                    iniciales.add(f)
        else:
            iniciales = archivos_previos

        inicio = time.time()
        while (time.time() - inicio) < Config.ESPERA_DESCARGA:
            time.sleep(2)
            actuales = []
            for f in os.listdir(Config.CARPETA_DESCARGAS):
                if f.lower().endswith(('.xlsx', '.xls')):
                    actuales.append(f)

            nuevos = [f for f in actuales if f not in iniciales]
            for f in nuevos:
                ruta = os.path.join(Config.CARPETA_DESCARGAS, f)
                try:
                    tam1 = os.path.getsize(ruta)
                    time.sleep(1)
                    tam2 = os.path.getsize(ruta)
                    if tam1 == tam2 and tam1 > 0:
                        return f
                except:
                    continue
        return None

    except Exception as e:
        registrar_log(f"Error monitoreando descarga: {str(e)}", "ERROR")
        return None


def procesar_archivo_descargado(nombre_archivo=None):
    """
    Mueve el archivo descargado de Downloads a REPORTES,
    renombrándolo con un timestamp.
    """
    try:
        archivo_origen = None
        nombre_original = ""

        if nombre_archivo:
            ruta = os.path.join(Config.CARPETA_DESCARGAS, nombre_archivo)
            if os.path.exists(ruta):
                archivo_origen = ruta
                nombre_original = nombre_archivo

        if not archivo_origen:
            # Buscar el Excel más reciente (últimos 10 minutos)
            exceles = []
            if os.path.exists(Config.CARPETA_DESCARGAS):
                for f in os.listdir(Config.CARPETA_DESCARGAS):
                    if f.lower().endswith(('.xlsx', '.xls')):
                        ruta = os.path.join(Config.CARPETA_DESCARGAS, f)
                        modif = os.path.getmtime(ruta)
                        if time.time() - modif < 600:
                            exceles.append((ruta, modif, f))
            if exceles:
                exceles.sort(key=lambda x: x[1], reverse=True)
                archivo_origen = exceles[0][0]
                nombre_original = exceles[0][2]

        if not archivo_origen or not os.path.exists(archivo_origen):
            registrar_log("No se encontró ningún archivo Excel para procesar", "ERROR")
            return None

        os.makedirs(Config.CARPETA_REPORTES, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d")
        base, ext = os.path.splitext(nombre_original)
        if not base.upper().startswith("SAT_"):
            base = f"SAT_{base}"
        nombre_final = f"{base}_{timestamp}{ext}"
        destino = os.path.join(Config.CARPETA_REPORTES, nombre_final)

        contador = 1
        while os.path.exists(destino):
            nombre_final = f"{base}_{timestamp}_{contador}{ext}"
            destino = os.path.join(Config.CARPETA_REPORTES, nombre_final)
            contador += 1

        for intento in range(3):
            try:
                time.sleep(1)
                shutil.copy2(archivo_origen, destino)
                if os.path.exists(destino):
                    tam = os.path.getsize(destino) / 1024
                    try:
                        os.remove(archivo_origen)
                    except:
                        pass
                    registrar_log(f"Archivo guardado: {nombre_final} ({tam:.1f} KB)", "SUCCESS")
                    return destino
            except:
                if intento < 2:
                    time.sleep(2)
                else:
                    raise
        return None

    except Exception as e:
        registrar_log(f"ERROR procesando archivo: {str(e)}", "ERROR")
        return None


# ============================================================
# FUNCIÓN PARA EJECUTAR EL ETL
# ============================================================
def ejecutar_etl():
    """Ejecuta el script ETL en modo silencioso."""
    try:
        registrar_log("Ejecutando ETL para cargar datos...", "INFO")
        ruta_etl = os.path.join(os.path.dirname(__file__), "etl_cargar_excel.py")

        if os.path.exists(ruta_etl):
            resultado = subprocess.run(
                [sys.executable, ruta_etl, "--silent"],
                capture_output=True,
                text=True
            )
            if resultado.returncode == 0:
                registrar_log("ETL completado exitosamente", "SUCCESS")
                for linea in resultado.stdout.split('\n')[-5:]:
                    if linea.strip():
                        print(f"   📊 {linea}")
            else:
                registrar_log(f"ETL falló: {resultado.stderr}", "ERROR")
        else:
            registrar_log(f"Archivo ETL no encontrado: {ruta_etl}", "WARNING")
    except Exception as e:
        registrar_log(f"Error ejecutando ETL: {str(e)}", "ERROR")


# ============================================================
# PROCESO COMPLETO DE DESCARGA
# ============================================================
def ejecutar_descarga_completa():
    """Orquesta todo el flujo de descarga y ETL."""
    registrar_log("INICIANDO PROCESO DE DESCARGA", "INFO")

    driver = None
    exito = False

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

        # Seleccionar reporte
        registrar_log("Seleccionando reporte de despachos...", "INFO")
        try:
            driver.execute_script("showScreen('logistics')")
            time.sleep(2)
            driver.execute_script("selectReport('ZFL_REP01')")
            time.sleep(2)
            registrar_log("Reporte seleccionado correctamente", "SUCCESS")
        except Exception as e:
            registrar_log(f"Error seleccionando reporte: {str(e)}", "WARNING")

        # Establecer fechas
        registrar_log("Configurando fechas del reporte...", "INFO")
        if not establecer_fechas_ayer_hoy(driver):
            registrar_log("Usando fechas por defecto del sistema", "WARNING")
            driver.execute_script("setFechasDefault()")
        time.sleep(2)

        # Ejecutar reporte
        registrar_log("Ejecutando reporte...", "INFO")
        try:
            driver.execute_script("ejecutarReporte()")
            time.sleep(5)
            registrar_log("Reporte ejecutado", "SUCCESS")
        except Exception as e:
            registrar_log(f"Error ejecutando reporte: {str(e)}", "WARNING")

        # Descargar Excel
        if not descargar_reporte_excel(driver):
            driver.quit()
            return False

        driver.quit()

        registrar_log("PROCESO COMPLETADO", "SUCCESS")
        exito = True

        # Ejecutar ETL
        ejecutar_etl()

    except Exception as e:
        registrar_log(f"ERROR CRÍTICO: {str(e)}", "ERROR")
        if driver:
            try:
                driver.quit()
            except:
                pass

    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass

    return exito


# ============================================================
# SISTEMA DE PROGRAMACIÓN (MODO AUTOMÁTICO)
# ============================================================
def deberia_ejecutar():
    """Determina si debe ejecutarse según la hora y día."""
    ahora = datetime.now()

    if Config.SOLO_LABORABLES and ahora.weekday() >= 5:
        return False

    if ahora.hour < Config.HORA_INICIO or ahora.hour >= Config.HORA_FIN:
        return False

    return True


def tarea_programada():
    """Tarea que se ejecuta según la programación."""
    print(f"\n⏰ EJECUCIÓN: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    if deberia_ejecutar():
        ejecutar_descarga_completa()
    else:
        print("⏸️ Ejecución omitida (fuera de horario o fin de semana)")


# ============================================================
# MODOS DE EJECUCIÓN (desde el menú)
# ============================================================
def modo_manual():
    """Ejecuta el proceso una sola vez."""
    crear_estructura_carpetas()
    registrar_log("MODO MANUAL - EJECUCIÓN ÚNICA", "INFO")
    ejecutar_descarga_completa()


def modo_automatico():
    """Inicia el modo programado (cada 1 minuto para pruebas)."""
    crear_estructura_carpetas()

    print("\n" + "=" * 70)
    print("🔄 MODO AUTOMÁTICO - PROGRAMACIÓN ACTIVADA")
    print("=" * 70)
    registrar_log("Iniciando modo automático", "SUCCESS")
    print(f"\n📋 CONFIGURACIÓN:")
    print(f"   📁 Carpeta destino: {Config.CARPETA_REPORTES}")
    print(f"   🎯 URL SAT: {Config.URL_SAT}")
    print(f"   👤 Usuario: {Config.USUARIO}")
    print(f"   ⏱️  Horario: {Config.HORA_INICIO}:00 a {Config.HORA_FIN}:00")
    print("\n" + "=" * 70)
    print("🔄 El sistema se ejecutará CADA 1 MINUTO (para pruebas)")
    print("   (Para cambiarlo a 2 horas, modificar la línea schedule)")
    print("   Presiona Ctrl+C para detener")
    print("=" * 70 + "\n")

    schedule.every(1).minutes.do(tarea_programada)
    tarea_programada()

    try:
        while True:
            schedule.run_pending()
            time.sleep(60)
    except KeyboardInterrupt:
        print("\n🛑 Sistema detenido")


def modo_prueba():
    """Solo verifica conexión y elementos principales."""
    crear_estructura_carpetas()

    print("\n🔧 MODO PRUEBA")

    if not verificar_archivo_html():
        print("\n⏎ Prueba completada. Volviendo al menú...")
        time.sleep(2)
        return

    driver = iniciar_navegador()
    if not driver:
        print("\n⏎ Prueba completada. Volviendo al menú...")
        time.sleep(2)
        return

    if navegar_al_sat(driver):
        registrar_log("Navegación al SAT exitosa", "SUCCESS")
        print(f"\n📄 INFORMACIÓN:")
        print(f"   Título: {driver.title}")

        try:
            driver.find_element(By.CSS_SELECTOR, Config.SELECTORES['usuario'])
            print(f"   ✅ Campo usuario encontrado")
        except:
            print(f"   ❌ Campo usuario NO encontrado")

        try:
            driver.find_element(By.CSS_SELECTOR, Config.SELECTORES['boton_login'])
            print(f"   ✅ Botón login encontrado")
        except:
            print(f"   ❌ Botón login NO encontrado")

        time.sleep(2)

    driver.quit()
    print("\n⏎ Prueba completada. Volviendo al menú principal...")
    time.sleep(2)


def buscar_archivos_manual():
    """Busca archivos Excel en Downloads y pregunta si moverlos."""
    print("\n🔍 BUSQUEDA MANUAL")

    if not os.path.exists(Config.CARPETA_DESCARGAS):
        print("❌ Carpeta Downloads no encontrada")
        return []

    exceles = []
    for f in os.listdir(Config.CARPETA_DESCARGAS):
        if f.lower().endswith(('.xlsx', '.xls')):
            ruta = os.path.join(Config.CARPETA_DESCARGAS, f)
            tam = os.path.getsize(ruta) / 1024
            exceles.append((f, tam))

    if exceles:
        print(f"\n✅ Se encontraron {len(exceles)} archivos Excel:")
        for i, (arch, tam) in enumerate(exceles, 1):
            print(f"{i:2}. {arch} ({tam:.1f} KB)")

        resp = input("\n¿Deseas mover estos archivos a REPORTES? (s/n): ").strip().lower()
        if resp == 's':
            for arch, _ in exceles:
                print(f"Procesando: {arch}")
                procesar_archivo_descargado(arch)
    else:
        print("\n📭 No se encontraron archivos Excel")

    return exceles


# ============================================================
# MENÚ PRINCIPAL
# ============================================================
def mostrar_menu():
    """Muestra las opciones disponibles."""
    print("\n" + "=" * 50)
    print("🎯 ROBOT SAT - SISTEMA AUTOMATIZADO")
    print("=" * 50)
    print("🤖 ROBOT SAT - MENÚ PRINCIPAL")
    print("=" * 50)
    print("\n📋 OPCIONES:")
    print("1. 📥 Ejecutar UNA VEZ")
    print("2. 🔄 Modo AUTOMÁTICO (programado)")
    print("3. 🔧 Modo PRUEBA (verificar conexión)")
    print("4. 📂 Abrir carpeta REPORTES")
    print("5. 🔍 Buscar archivos manualmente")
    print("6. 🚪 Salir")
    print("\n" + "=" * 50)


def main():
    """Punto de entrada principal."""
    if len(sys.argv) > 1 and sys.argv[1] == "--silent":
        crear_estructura_carpetas()
        ejecutar_descarga_completa()
        return

    crear_estructura_carpetas()
    os.system('cls' if os.name == 'nt' else 'clear')

    while True:
        mostrar_menu()
        print()

        try:
            opcion = input("📝 Seleccione opción (1-6): ").strip()

            if opcion == "1":
                modo_manual()
                print("\n✅ Proceso completado. Volviendo al menú...")
                time.sleep(2)
                os.system('cls' if os.name == 'nt' else 'clear')

            elif opcion == "2":
                modo_automatico()
                # Cuando salga del automático, limpia pantalla
                os.system('cls' if os.name == 'nt' else 'clear')

            elif opcion == "3":
                modo_prueba()
                os.system('cls' if os.name == 'nt' else 'clear')

            elif opcion == "4":
                if os.path.exists(Config.CARPETA_REPORTES):
                    os.startfile(Config.CARPETA_REPORTES)
                    print("📂 Carpeta REPORTES abierta")
                else:
                    print("❌ La carpeta no existe")
                time.sleep(2)
                os.system('cls' if os.name == 'nt' else 'clear')

            elif opcion == "5":
                buscar_archivos_manual()
                print("\nVolviendo al menú...")
                time.sleep(2)
                os.system('cls' if os.name == 'nt' else 'clear')

            elif opcion == "6":
                print("\n👋 ¡Hasta luego!")
                break

            else:
                print("\n❌ Opción no válida")
                time.sleep(2)
                os.system('cls' if os.name == 'nt' else 'clear')

        except KeyboardInterrupt:
            print("\n🛑 Operación cancelada")
            break
        except Exception as e:
            print(f"\n❌ Error: {str(e)}")
            time.sleep(2)
            os.system('cls' if os.name == 'nt' else 'clear')


# ============================================================
# EJECUCIÓN PRINCIPAL
# ============================================================
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        input("\nPresiona ENTER para salir...")