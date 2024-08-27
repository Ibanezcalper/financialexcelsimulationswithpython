# financialexcelsimulationswithpython
This task automates financial data processing from a CSV file, inserting calculated values into an Excel template, and generating output files. The script streamlines bulk data handling, ensuring accurate financial data in templates. The guide covers file path setup, variable modification, and output customization for specific needs.


#----------------------------------#
Autor: Mario Ignacio Ibañez Castro
Github: https://github.com/MarioIb
#----------------------------------#

#---------------------------------------------------------------------------------------------------------------------------------------#
ESPAÑOL
#---------------------------------------------------------------------------------------------------------------------------------------#

Guía de Uso del Código: Automatización de Procesamiento Financiero en Excel

Descripción General
Este script está diseñado para automatizar el procesamiento y análisis de datos financieros a partir de un archivo CSV. 
El programa carga los datos, realiza cálculos específicos y luego los inserta en una plantilla de Excel, generando archivos 
de salida personalizados para cada registro en el archivo CSV. Este documento explica cómo utilizar el script, modificar 
las variables y rutas, y personalizar la funcionalidad según sea necesario.

Configuración de Archivos y Directorios
Archivo de Entrada "A" (CSV):

Ruta: input_file_a
Propósito: Contiene los datos financieros de varios predios. Este archivo se carga y se procesa en el script.
Ejemplo de ruta: C:\06_Py\excel\DATA FINANCIERA 3.csv
Archivo de Plantilla "B" (Excel):

Ruta: template_file_b
Propósito: Es un archivo de Excel que actúa como plantilla, en el cual se insertarán los datos procesados del archivo CSV. 
Este archivo contiene varias hojas donde se distribuyen los datos.
Ejemplo de ruta: C:\06_Py\excel\Financiera inmobiliaria (plantilla con operacion) v3 %.xlsx
Carpeta de Salida:

Ruta: output_folder
Propósito: Carpeta donde se guardarán los archivos Excel generados, personalizados para cada predio.
Ejemplo de ruta: C:\06_Py\05_TRAD03_DOR06\01_salidas\01_COP
Variables de Configuración
Filtros de Estatus:

Variable: filtro_estatus
Propósito: Especifica los estados de los predios que serán procesados. Solo se procesarán aquellos predios cuyo estado esté en esta lista.
Valores posibles: ['BUENO', 'NECESITA REPARACIONES MENORES']
Modificación: Puedes añadir o quitar estatus según lo que necesites filtrar.
Valores de Apreciación:

Variable: apreciacion_values
Propósito: Define el porcentaje de apreciación inmobiliaria basado en la zona del predio (tradicional o dorada).
Valores:
tradicional: 4.5%
dorada: 9%
Modificación: Se puede ajustar el valor de apreciación cambiando los valores en el diccionario.
Porcentaje del Primer Año:

Variable: porcentajeAnho1
Propósito: Define un porcentaje específico respecto a la inversión que se usará en cálculos dentro de la hoja de Excel.
Valor predeterminado: 0.7 (70% de la inversión se realiza en el año 1)
Modificación: Este valor puede ser ajustado según los requisitos de los cálculos.


Funcionamiento del Código
Lectura del Archivo CSV:

El archivo CSV se lee y se convierte en una lista de registros, que luego se procesan uno por uno.
Se utiliza tqdm para mostrar una barra de progreso, lo que facilita el seguimiento del procesamiento.
Procesamiento de Datos:

Cada fila en el archivo CSV representa un predio y contiene múltiples campos como predio, estado_uno, rehabilitacion, entre otros.
Los valores se convierten a números mediante la función to_number para asegurar que se manejen correctamente en los cálculos.
Aplicación de Filtros:

Solo se procesan los predios cuyo estado esté incluido en la lista filtro_estatus.
Inserción de Datos en la Plantilla Excel:

Los datos de cada predio se insertan en la plantilla Excel en las celdas correspondientes.
Se actualizan múltiples hojas de la plantilla (CalendarioInv, PresupuestoCost, PresupuestoIng, Flujo), dependiendo de los datos 
disponibles en el CSV.
Generación y Guardado de Archivos:

Para cada predio procesado, se genera un archivo de Excel personalizado con un nombre basado en el predio y el porcentaje de 
apreciación inmobiliaria.
Los archivos se guardan en la carpeta especificada en output_folder.
Modificaciones y Personalización
Cambiar Rutas: Para cambiar los archivos de entrada, plantilla o la carpeta de salida, simplemente modifica las variables input_file_a, 
template_file_b, y output_folder con las rutas nuevas.
Añadir Nuevos Filtros: Si necesitas filtrar por otros estados, simplemente agrega esos estados a la lista filtro_estatus.
Modificar Fórmulas o Celdas: Si la plantilla de Excel cambia o necesitas actualizar otras celdas, ajusta las líneas correspondientes 
en el bloque de código donde se actualizan las hojas.
Ejecución del Programa
Una vez que las variables están configuradas correctamente, puedes ejecutar el script. El programa procesará el archivo CSV, generará 
los archivos Excel y los guardará en la carpeta de salida especificada.

Nota: Asegúrate de que las rutas de los archivos sean correctas y que las plantillas de Excel estén disponibles antes de ejecutar el 
script para evitar errores.


#---------------------------------------------------------------------------------------------------------------------------------------#


Guía de Uso del Código: Generación Automática de Resumen de Resultados Financieros
Descripción General
Este script está diseñado para procesar múltiples archivos de Excel, extraer valores específicos y generar un resumen consolidado en un 
archivo Excel separado. El proceso incluye la recalculación de cada archivo, la extracción de datos de celdas clave y la organización de 
estos datos en un nuevo archivo de resumen. Esta guía te ayudará a configurar, ejecutar y modificar el script según tus necesidades.

Configuración de Archivos y Directorios
Carpeta de Salida:

Ruta: output_folder
Propósito: Contiene todos los archivos de Excel que serán procesados. El script recorrerá esta carpeta para encontrar los archivos pertinentes.
Ejemplo de ruta: C:\06_Py\05_TRAD03_DOR06\01_salidas\02_SOP
Archivo de Resumen:

Ruta: summary_file
Propósito: Es el archivo Excel donde se consolidarán los datos extraídos de cada archivo en la carpeta de salida.
Ejemplo de ruta: C:\06_Py\05_TRAD03_DOR06\01_salidas\02_SOP\02_Resumen_SOP_220824.xlsx
Variables de Configuración
Encabezados del Resumen:

Variable: headers
Propósito: Define los nombres de las columnas en el archivo de resumen. Estos encabezados corresponden a los datos que se extraerán de 
cada archivo Excel.
Valores:
['Escenario', 'Inversion', 'Tir', 'Tasa de descuento', 'VA Ingresos', 'VA Egresos', 'Costo + Inversion', 'B/C', 'VAN', 'ROI', 'TIIE', 
'FC VAN AC 1', 'FC VAN AC 2', 'FC VAN AC 3', 'FC VAN AC 4', 'FC VAN AC 5', 'FC VAN AC 6']
Modificación: Puedes agregar, quitar o renombrar los encabezados según los datos que necesites consolidar.
Porcentaje de Dinero en el Primer Año:

Variable: dinero_en_anho_1
Propósito: Define un porcentaje que puede ser utilizado en los cálculos dentro de los archivos de Excel procesados. 
Este valor es fijo dentro del código, pero podría ser usado para comparaciones o cálculos adicionales.
Funcionamiento del Código
Recorrer la Carpeta de Salida:

El script busca todos los archivos Excel en la carpeta de salida (output_folder), excluyendo cualquier archivo de resumen ya existente.
Utiliza tqdm para mostrar una barra de progreso, lo que facilita el seguimiento del estado del procesamiento.
Recalcular y Guardar Archivos de Excel:

Cada archivo Excel encontrado se recalcula utilizando la librería win32com.client, asegurando que todas las fórmulas y referencias 
se actualicen antes de extraer los datos.
Función: recalculate_excel(filepath)
Extracción de Datos:

El script abre cada archivo Excel y extrae los valores de celdas específicas (por ejemplo, E47, E48, etc.) de la hoja nombrada Flujo.
Los valores se almacenan en una lista y luego se insertan en la fila correspondiente del archivo de resumen.
Generación del Archivo de Resumen:

Los datos extraídos de cada archivo se consolidan en un nuevo archivo Excel con el nombre y ruta especificados en summary_file.
Encabezados: Los encabezados de las columnas se insertan en la primera fila del archivo de resumen, y los valores de cada archivo 
procesado se añaden a las filas subsecuentes.
Modificaciones y Personalización
Cambiar la Ruta de la Carpeta de Salida: Si deseas procesar archivos de otra carpeta, ajusta la variable output_folder con la nueva ruta.
Modificar las Celdas Extraídas: Si necesitas extraer datos de celdas diferentes, cambia las referencias de celdas en el bucle que recorre 
['E47','E48', 'E49', 'E50', ...].
Actualizar el Nombre de la Hoja: Si la hoja de trabajo en los archivos Excel tiene un nombre diferente al de Flujo, cambia el nombre en la 
línea sheet = wb['Flujo'].
Ejecución del Programa
Asegúrate de que las rutas de los archivos y carpetas estén correctamente configuradas.
Ejecuta el script para procesar todos los archivos Excel en la carpeta especificada y generar un resumen consolidado.
El archivo de resumen se guardará en la ubicación especificada por summary_file.



Nota: Asegúrate de que Excel esté instalado en tu sistema, ya que el script utiliza win32com.client para la recalculación de los archivos Excel.
También es importante cerrar y guardar los archivos excel antes de la ejecución de este código ya que forza la ejecución de excel, por lo tanto cierra
y abre la aplicación n veces hasta completar la tarea, y si está abierto antes de la ejecucion, no guarda los archivos y/o sus modificaciones-


#---------------------------------------------------------------------------------------------------------------------------------------#
NOTAS GENERALES


1. La carpeta "COP" hace referencia a los escenarios con operación

2. La carpeta "SOP" hace referencia a los escenarios sin operación

3. Para la ejecución del código se necesitan las siguientes librerias: (pegar el código de instalación en la terminal)
        openpyxl
            Instalación: pip install openpyxl
            
        csv
            Instalación: Incluida en la biblioteca estándar de Python (no requiere instalación).
            
        os
            Instalación: Incluida en la biblioteca estándar de Python (no requiere instalación).
            
        tqdm
            Instalación: pip install tqdm
            
        xlsxwriter
            Instalación: pip install XlsxWriter
            
        win32com.client
            Instalación: pip install pywin32

4. Los archivos excel base se encuentran en la ruta "simulaciones/excel"
5. Se tiene un resumen de la información en un archivo PowerBI (Se debe instalar PowerBI previamente)
6. Se adjunta un archivo ".kmz" con las ubicaciones de los predios.
7. Los archivos Python se encuentran en la ruta "simulaciones/00_py"
8. Los archivos de salida resultantes se encuentran en la ruta "simulaciones/01_salidas"
9. Los archivos de salida del resumen se encuentran en la ruta "simulaciones/02_resumenes"
