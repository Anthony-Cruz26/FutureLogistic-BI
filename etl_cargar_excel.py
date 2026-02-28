#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETL - CARGA DE EXCEL A SQL SERVER (VERSIÓN FINAL CORREGIDA)
Empresa: FutureLogistic S.A. - Durán, Ecuador
Versión: 6.0 - CON RETORNO DE CÓDIGO DE SALIDA
"""

import os
import sys
import time
import pyodbc
import pandas as pd
from datetime import datetime
import glob
import shutil

# ============================================================
# CONFIGURACIÓN DE MODO SILENCIOSO
# ============================================================
MODO_SILENCIOSO = len(sys.argv) > 1 and sys.argv[1] == "--silent"

def print_silent(*args, **kwargs):
    """Print solo si no está en modo silencioso"""
    if not MODO_SILENCIOSO:
        print(*args, **kwargs)

# ============================================================
# CONFIGURACIÓN
# ============================================================
class Config:
    CARPETA_REPORTES = "C:/Users/yaps2/OneDrive - Ormesby Primary/Escritorio/SistemaSAT_Prueba/REPORTES"
    CARPETA_PROCESADOS = os.path.join(CARPETA_REPORTES, "PROCESADOS")
    SERVIDOR = "localhost\\SQLEXPRESS"
    BASE_DATOS = "FutureLogistic_BI"
    CONN_STR = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVIDOR};DATABASE={BASE_DATOS};Trusted_Connection=yes;'

# ============================================================
# FUNCIONES DE AYUDA
# ============================================================
def conectar_bd():
    try:
        conn = pyodbc.connect(Config.CONN_STR)
        print_silent(f"✅ Conectado a {Config.BASE_DATOS} en {Config.SERVIDOR}")
        return conn
    except Exception as e:
        print_silent(f"❌ Error conectando a BD: {str(e)}")
        return None

def registrar_log(conn, archivo, leidos, insertados, duplicados, errores, duracion, estado, mensaje):
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO log_etl 
            (fecha_proceso, archivo_origen, registros_leidos, registros_insertados, 
             registros_duplicados, registros_error, duracion_seg, estado, mensaje)
            VALUES (GETDATE(), ?, ?, ?, ?, ?, ?, ?, ?)
        """, (archivo, leidos, insertados, duplicados, errores, duracion, estado, mensaje))
        conn.commit()
    except Exception as e:
        print_silent(f"⚠️ Error registrando log: {str(e)}")

def obtener_id_operador(conn, nombre_operador):
    cursor = conn.cursor()
    cursor.execute("SELECT id_operador FROM dim_operador WHERE nombre = ?", (nombre_operador,))
    row = cursor.fetchone()
    return row[0] if row else None

def obtener_id_plataforma(conn, codigo_plataforma):
    cursor = conn.cursor()
    cursor.execute("SELECT id_plataforma FROM dim_plataforma WHERE codigo = ?", (codigo_plataforma,))
    row = cursor.fetchone()
    return row[0] if row else None

def obtener_id_bodega(conn, codigo_bodega):
    cursor = conn.cursor()
    cursor.execute("SELECT id_bodega FROM dim_bodega WHERE codigo = ?", (codigo_bodega,))
    row = cursor.fetchone()
    return row[0] if row else None

def obtener_id_cliente(conn, cadena, sucursal):
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id_cliente FROM dim_cliente 
        WHERE cadena = ? AND sucursal = ?
    """, (cadena, sucursal))
    row = cursor.fetchone()
    return row[0] if row else None

def obtener_id_producto(conn, tipo_azucar, presentacion):
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id_producto FROM dim_producto 
        WHERE tipo_azucar = ? AND presentacion = ?
    """, (tipo_azucar, presentacion))
    row = cursor.fetchone()
    return row[0] if row else None

def guia_existe(conn, guia):
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM hechos_despachos WHERE guia_remision = ?", (guia,))
    return cursor.fetchone()[0] > 0

# ============================================================
# PROCESAR ARCHIVO EXCEL
# ============================================================
def procesar_archivo(ruta_archivo):
    nombre_archivo = os.path.basename(ruta_archivo)
    print_silent(f"\n📄 Procesando: {nombre_archivo}")
    
    inicio = time.time()
    conn = None
    exito = False
    
    try:
        conn = conectar_bd()
        if not conn:
            return False
        
        # Leer Excel (hoja 1)
        df = pd.read_excel(ruta_archivo, sheet_name=0)
        registros_leidos = len(df)
        print_silent(f"📊 Registros leídos: {registros_leidos}")
        
        if registros_leidos == 0:
            registrar_log(conn, nombre_archivo, 0, 0, 0, 0, 
                         time.time() - inicio, "ADVERTENCIA", "Archivo vacío")
            return True
        
        insertados = 0
        duplicados = 0
        errores = 0
        
        for idx, row in df.iterrows():
            try:
                guia = row['GUIA_REMISION']
                
                if guia_existe(conn, guia):
                    duplicados += 1
                    continue
                
                # Obtener IDs
                id_operador = obtener_id_operador(conn, row['OPERADOR'])
                if not id_operador:
                    print_silent(f"⚠️ Operador no encontrado: {row['OPERADOR']}")
                    errores += 1
                    continue
                
                id_plataforma = obtener_id_plataforma(conn, row['PLATAFORMA'])
                if not id_plataforma:
                    print_silent(f"⚠️ Plataforma no encontrada: {row['PLATAFORMA']}")
                    errores += 1
                    continue
                
                id_bodega = obtener_id_bodega(conn, row['BODEGA_ORIGEN'])
                if not id_bodega:
                    print_silent(f"⚠️ Bodega no encontrada: {row['BODEGA_ORIGEN']}")
                    errores += 1
                    continue
                
                id_cliente = obtener_id_cliente(conn, row['CLIENTE'], row['SUCURSAL'])
                if not id_cliente:
                    print_silent(f"⚠️ Cliente no encontrado: {row['CLIENTE']} - {row['SUCURSAL']}")
                    errores += 1
                    continue
                
                id_producto = obtener_id_producto(conn, row['TIPO_AZUCAR'], row['PRESENTACION'])
                if not id_producto:
                    print_silent(f"⚠️ Producto no encontrado: {row['TIPO_AZUCAR']} - {row['PRESENTACION']}")
                    errores += 1
                    continue
                
                # ====================================================
                # LÓGICA DE NEGOCIO PARA HORAS Y TIEMPOS
                # ====================================================
                if row['ESTADO'] == 'CANCELADO':
                    # Si está cancelado, forzar hora NULL y tiempo 0
                    hora_llegada = None
                    tiempo_ruta = 0
                else:
                    # Para otros estados, tomar valores del Excel (o NULL si vacío)
                    hora_llegada = None if pd.isna(row['HORA_LLEGADA']) else row['HORA_LLEGADA']
                    tiempo_ruta = 0 if pd.isna(row['TIEMPO_RUTA_MIN']) else row['TIEMPO_RUTA_MIN']
                
                # Insertar en hechos_despachos
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO hechos_despachos (
                        fecha, guia_remision, id_operador, id_plataforma, id_bodega,
                        id_cliente, id_producto, num_pallets, num_fardos, kilos_totales,
                        hora_salida, hora_llegada, tiempo_ruta_min, incidencia, estado, observaciones,
                        fuente
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    row['FECHA'], guia, id_operador, id_plataforma, id_bodega,
                    id_cliente, id_producto, row['NUM_PALLETS'], row['NUM_FARDOS'], row['KILOS_TOTALES'],
                    row['HORA_SALIDA'], hora_llegada, tiempo_ruta,
                    row['INCIDENCIA'], row['ESTADO'], row['OBSERVACIONES'],
                    'EXCEL_REAL'
                ))
                conn.commit()
                insertados += 1
                
            except Exception as e:
                print_silent(f"❌ Error fila {idx}: {str(e)}")
                errores += 1
        
        duracion = time.time() - inicio
        
        estado_log = "EXITO" if errores == 0 else "ERROR" if errores == registros_leidos else "ADVERTENCIA"
        mensaje_log = f"Insertados: {insertados}, Duplicados: {duplicados}, Errores: {errores}"
        
        registrar_log(conn, nombre_archivo, registros_leidos, insertados, 
                     duplicados, errores, duracion, estado_log, mensaje_log)
        
        print_silent(f"✅ Insertados: {insertados}")
        print_silent(f"🔄 Duplicados: {duplicados}")
        print_silent(f"❌ Errores: {errores}")
        print_silent(f"⏱️ Duración: {duracion:.2f} seg")
        
        exito = (errores == 0)  # Éxito si no hubo errores
        return exito
        
    except Exception as e:
        print_silent(f"❌ Error crítico: {str(e)}")
        if conn:
            duracion = time.time() - inicio
            registrar_log(conn, nombre_archivo, 0, 0, 0, 0, 
                         duracion, "ERROR", str(e))
        return False
    
    finally:
        if conn:
            conn.close()

# ============================================================
# MOVER ARCHIVO PROCESADO
# ============================================================
def mover_a_procesados(ruta_archivo):
    try:
        if not os.path.exists(Config.CARPETA_PROCESADOS):
            os.makedirs(Config.CARPETA_PROCESADOS)
        
        nombre = os.path.basename(ruta_archivo)
        fecha_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_nuevo = f"{fecha_hora}_{nombre}"
        destino = os.path.join(Config.CARPETA_PROCESADOS, nombre_nuevo)
        
        shutil.move(ruta_archivo, destino)
        print_silent(f"📦 Archivo movido a: {destino}")
        return True
    except Exception as e:
        print_silent(f"⚠️ Error moviendo archivo: {str(e)}")
        return False

# ============================================================
# PROCESAR TODOS LOS ARCHIVOS
# ============================================================
def procesar_todos():
    print_silent("\n" + "=" * 60)
    print_silent("🚀 ETL - CARGA DE DATOS FUTURELOGISTIC")
    print_silent("=" * 60)
    
    if not os.path.exists(Config.CARPETA_REPORTES):
        print_silent(f"❌ Carpeta no encontrada: {Config.CARPETA_REPORTES}")
        return False
    
    archivos = glob.glob(os.path.join(Config.CARPETA_REPORTES, "*.xlsx"))
    archivos += glob.glob(os.path.join(Config.CARPETA_REPORTES, "*.xls"))
    
    if not archivos:
        print_silent("📭 No hay archivos Excel para procesar")
        return True  # No es error, solo no hay archivos
    
    print_silent(f"📁 Encontrados: {len(archivos)} archivo(s)\n")
    
    todos_exitosos = True
    
    for archivo in archivos:
        if "PROCESADOS" in archivo:
            continue
        
        if procesar_archivo(archivo):
            mover_a_procesados(archivo)
        else:
            todos_exitosos = False
        
        print_silent("-" * 40)
    
    print_silent("\n✅ ETL COMPLETADO")
    return todos_exitosos

# ============================================================
# FUNCIÓN PARA MODO SILENCIOSO
# ============================================================
def ejecutar_etl_silencioso():
    resultado = procesar_todos()
    return resultado

# ============================================================
# MENÚ PRINCIPAL
# ============================================================
def main():
    if MODO_SILENCIOSO:
        exito = procesar_todos()
        sys.exit(0 if exito else 1)
        return
    
    while True:
        print("\n" + "=" * 60)
        print("📊 ETL FUTURELOGISTIC - MENÚ PRINCIPAL")
        print("=" * 60)
        print("1. 🚀 Ejecutar ETL")
        print("2. 📁 Ver carpeta REPORTES")
        print("3. 📂 Ver carpeta PROCESADOS")
        print("4. 🔍 Ver últimos logs")
        print("5. 🚪 Salir")
        print("=" * 60)
        
        opcion = input("\n📝 Seleccione opción (1-5): ").strip()
        
        if opcion == "1":
            procesar_todos()
            input("\n✅ Presione ENTER para continuar...")
        elif opcion == "2":
            if os.path.exists(Config.CARPETA_REPORTES):
                os.startfile(Config.CARPETA_REPORTES)
        elif opcion == "3":
            if os.path.exists(Config.CARPETA_PROCESADOS):
                os.startfile(Config.CARPETA_PROCESADOS)
        elif opcion == "4":
            try:
                conn = conectar_bd()
                if conn:
                    query = "SELECT TOP 10 * FROM log_etl ORDER BY fecha_proceso DESC"
                    df = pd.read_sql(query, conn)
                    print("\n📋 ÚLTIMOS 10 LOGS:")
                    print(df.to_string(index=False))
                    conn.close()
            except Exception as e:
                print(f"❌ Error: {str(e)}")
            input("\n✅ Presione ENTER para continuar...")
        elif opcion == "5":
            print("\n👋 ¡Hasta luego!")
            break
        else:
            print("\n❌ Opción no válida")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n🛑 Proceso cancelado")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        if not MODO_SILENCIOSO:
            input("Presione ENTER para salir...")
        sys.exit(1)