# FutureLogistic BI

Sistema automatizado de inteligencia de negocios para la empresa FutureLogistic S.A., dedicada a la distribución de azúcar San Carlos en Durán, Ecuador.

## 📋 Descripción del Proyecto

Este proyecto consiste en un robot automatizado que descarga reportes de despacho desde un sistema SAT simulado, procesa los archivos Excel mediante un ETL (Extract, Transform, Load) y carga los datos en SQL Server para su posterior visualización en Power BI.

## 🚀 Funcionalidades

- **Robot automatizado**: Utiliza Selenium para navegar, hacer login, seleccionar reportes y descargar archivos Excel.
- **ETL robusto**: Lee los Excel, busca IDs en tablas de dimensiones, inserta en hechos con control de duplicados.
- **Logs de auditoría**: Cada ejecución queda registrada en la tabla `log_etl`.
- **Mensajes emergentes**: Ventanas de éxito/error al finalizar el proceso.
- **Integración con Power BI**: Datos actualizados automáticamente en la nube cada 2 horas.

## 🛠️ Tecnologías Utilizadas

- **Python 3.14** - Lenguaje principal
- **Selenium** - Automatización del navegador
- **Pandas** - Procesamiento de archivos Excel
- **PyODBC** - Conexión a SQL Server
- **SQL Server** - Base de datos
- **Power BI** - Visualización de datos

## ⚙️ Instalación

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/Anthony-Cruz26/FutureLogistic-BI.git
