# FutureLogistic BI

Sistema automatizado de inteligencia de negocios para la empresa FutureLogistic S.A., dedicada a la distribución de azúcar San Carlos en Durán, Ecuador.

---

## 📂 Archivos del Proyecto

### 1. `INICIAR.py` – Robot Silencioso (Versión Final)

**Uso recomendado:** Para ejecución diaria con doble clic desde el escritorio.

- ✅ Automatiza la descarga del reporte desde el sistema SAT simulado.
- ✅ Ejecuta el proceso ETL automáticamente al finalizar.
- ✅ **No muestra ventana de consola**, solo ventanas emergentes de éxito o error.
- ✅ Ideal para usuarios finales o para dejar ejecutándose de forma continua.
- ✅ La ventana de consola se minimiza automáticamente al inicio.
- ✅ Se ejecuta en un bucle infinito con intervalo de 1 minuto (ajustable).  

### 2. `robot_sat_completo.py` - Robot con Consola (Versión de Depuración)
**Uso recomendado:** Para pruebas, depuración o cuando quieras ver los mensajes en tiempo real.

- ✅ Misma funcionalidad que `INICIAR.py`, pero con **menú interactivo** en consola
- ✅ Muestra todos los pasos del proceso con iconos y tiempos
- ✅ Incluye **modo automático programado** (cada 1 minuto para pruebas)
- ✅ Ideal para entender el flujo o detectar errores

### 3. `etl_cargar_excel.py` - Procesador ETL
**Uso:** Común para ambos robots. Procesa los archivos Excel y los carga en SQL Server.

- ✅ Lee los Excel, busca IDs en dimensiones y carga en `hechos_despachos`
- ✅ Control de duplicados por `guia_remision`
- ✅ Registra cada ejecución en la tabla `log_etl`
- ✅ Puede ejecutarse en **modo silencioso** (para los robots) o en **modo interactivo** (para depuración)

---

## 🚀 Comparación Rápida

| Archivo | Consola | Menú | Automático | Uso recomendado |
|---------|---------|------|------------|------------------|
| `INICIAR.py` | ❌ No | ❌ No | ❌ No (solo manual) | Usuario final, doble clic |
| `robot_sat_completo.py` | ✅ Sí | ✅ Sí | ✅ Sí (cada 1 min) | Depuración, pruebas |
| `etl_cargar_excel.py` | Depende | ✅ Sí | ❌ No | Procesamiento de datos |

---

## ⚙️ Instalación y Configuración

### 1. Clonar el repositorio
```bash
git clone https://github.com/Anthony-Cruz26/FutureLogistic-BI.git
cd FutureLogistic-BI
2. Instalar dependencias
Ejecutar el siguiente comando para instalar todas las librerías necesarias:

bash
pip install selenium webdriver-manager pandas pyodbc openpyxl
3. Configurar rutas y credenciales
Abrir los archivos INICIAR.py y robot_sat_completo.py y ajustar las siguientes variables según tu entorno:

python
# Credenciales de acceso al SAT
USUARIO = "supervisor"
CONTRASENA = "distribucion2024"

# Ruta del archivo HTML que simula el SAT
RUTA_HTML = "C:/Users/tu_usuario/ruta/al/sap_futurelogistic.html"

# Carpeta donde se guardarán los reportes descargados
CARPETA_REPORTES = "C:/Users/tu_usuario/ruta/a/REPORTES"
Nota: Asegúrate de que la carpeta REPORTES exista o tenga permisos de escritura.

4. Configurar la base de datos SQL Server
Crear una base de datos llamada FutureLogistic_BI

Ejecutar el script de creación de tablas incluido en la documentación (o crearlas manualmente según el modelo)

Verificar que la cadena de conexión en etl_cargar_excel.py apunte a tu servidor:

python
SERVIDOR = "localhost\\SQLEXPRESS"
BASE_DATOS = "FutureLogistic_BI"
▶️ Cómo usar cada archivo
Opción 1: Uso diario (sin consola)
Haz doble clic en INICIAR.py (o en su acceso directo en el escritorio).
Al finalizar, aparecerá una ventana emergente indicando si el proceso fue exitoso o si hubo errores.

Opción 2: Depuración o pruebas
Ejecutar en terminal:

bash
python robot_sat_completo.py
Luego elegir la opción deseada en el menú:

1 → Ejecutar una vez

2 → Modo automático (cada 1 minuto)

3 → Modo prueba (verificar conexión)

4 → Abrir carpeta REPORTES

5 → Buscar archivos manualmente

Opción 3: Procesar Excel manualmente
bash
python etl_cargar_excel.py
Luego elegir la opción 1 para procesar todos los archivos en REPORTES.

📊 Estructura de la Base de Datos
dim_operador: Operadores de transporte

dim_cliente: Clientes y sucursales

dim_producto: Tipos de azúcar y presentaciones

dim_plataforma: Vehículos y plataformas

dim_bodega: Bodegas de origen

hechos_despachos: Tabla de hechos con todas las transacciones

log_etl: Registro de ejecuciones del ETL
