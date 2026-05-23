# Financial Excel Simulations with Python

## Descripción General
Este repositorio contiene un conjunto de scripts en Python diseñados para automatizar el procesamiento, análisis y consolidación de datos financieros relacionados con predios inmobiliarios. 

## El Problema que Resuelve
El procesamiento manual de grandes volúmenes de datos financieros en Excel es propenso a errores y consume una cantidad significativa de tiempo. Este proyecto soluciona estos inconvenientes de la siguiente manera:
1. Automatizando la ingesta de datos desde un archivo plano (CSV) para insertarlos directamente en plantillas financieras de Excel complejas.
2. Generando masivamente modelos financieros de manera individual por cada registro o predio.
3. Consolidando métricas financieras clave (como TIR, VAN, ROI, entre otras) extraídas de múltiples archivos Excel resultantes, agrupándolas en un único reporte de resumen listo para el análisis.

## Tecnologías Utilizadas
- **Python 3.x**: Lenguaje principal de desarrollo.
- **openpyxl**: Manipulación y lectura/escritura de archivos Excel (`.xlsx`).
- **xlsxwriter**: Creación y formateo avanzado de archivos Excel.
- **pywin32 (win32com.client)**: Interfaz con la aplicación de escritorio de Excel para forzar el recálculo de fórmulas complejas antes de la extracción de datos.
- **tqdm**: Visualización de barras de progreso en la consola para monitorear el avance.
- **csv / os**: Librerías estándar de Python para lectura de datos y manejo del sistema de archivos y rutas.

## Estructura y Funcionamiento del Proyecto

El proyecto está diseñado para ejecutarse en dos fases o flujos principales:

### 1. Procesamiento Financiero Individual
Los scripts principales (como `B_NRM_REHAB.py` o `B_NRM_REHAB_SCOM.py`) leen un archivo CSV con un listado de predios. Aplican filtros de validación (por ejemplo, verificando el estado de conservación) y distribuyen los datos numéricos en hojas específicas de una plantilla base de Excel (hojas como `CalendarioInv`, `PresupuestoCost`, `PresupuestoIng` y `Flujo`). Al finalizar, el código guarda un archivo de Excel independiente para cada predio procesado.

### 2. Generador de Resumen Consolidado (`summary_loadbar.py`)
Este script escanea un directorio de salida en busca de los modelos financieros generados en la fase anterior. Utiliza la interfaz de Windows COM para abrir cada archivo Excel, recalcular de forma segura todas sus fórmulas internas, y extraer indicadores financieros clave. Todo el conjunto de datos se vuelca en un archivo final consolidado.

## Instalación y Requisitos Previos

1. Clonar este repositorio en su entorno local.
2. Asegurar que Python 3 está instalado en el sistema.
3. Instalar las dependencias necesarias mediante `pip`:
   ```bash
   pip install openpyxl tqdm XlsxWriter pywin32
   ```
4. Contar con Microsoft Excel instalado en la máquina (requisito indispensable para el script de resumen).

## Guía de Uso

1. **Configurar Rutas**: Abra los scripts y configure las rutas de los archivos (`input_file_a`, `template_file_b`, `output_folder`, `summary_file`) para apuntar a sus directorios locales.
2. **Ajustar Variables**: Modifique variables de entorno o parámetros según la necesidad de la operación (por ejemplo, `filtro_estatus` o `apreciacion_values`).
3. **Ejecutar Fase 1**: Ejecute el script principal de procesamiento desde su terminal.
   ```bash
   python B_NRM_REHAB.py
   ```
4. **Ejecutar Fase 2 (Consolidación)**: Ejecute el script de resumen para extraer los resultados.
   ```bash
   python summary_loadbar.py
   ```

*Nota Importante: Para la ejecución del script de resumen (`summary_loadbar.py`), es estrictamente necesario correr el código en un entorno Windows y asegurarse de que todos los archivos Excel objetivo se encuentren cerrados. El script controla la aplicación de Excel en segundo plano, por lo que archivos abiertos pueden ocasionar bloqueos de lectura/escritura.*
