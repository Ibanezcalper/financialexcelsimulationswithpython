#----------------------------------#
#Autor: Mario Ignacio Ibañez Castro
#Github: https://github.com/MarioIb
#----------------------------------#

#---------------------------------------------------------------------------------------------------------------------------------------#

import openpyxl
import os
import csv
from tqdm import tqdm  # Importar tqdm

# Ruta del archivo de entrada "A" (CSV)
input_file_a = r'C:\06_Py\excel\DATA FINANCIERA 3.csv'
# Ruta del archivo de plantilla "B" (Excel)
template_file_b = r'C:\06_Py\excel\Financiera inmobiliaria (plantilla con operacion) v3 %.xlsx'
# Ruta de la carpeta de salida
output_folder = r'C:\06_Py\05_TRAD03_DOR06\01_salidas\02_SOP'

# Filtros específicos para los estatus
filtro_estatus = ['BUENO', 'NECESITA REPARACIONES MENORES']  # Cambiar según sea necesario

def to_number(value):
    try:
        return float(value)
    except ValueError:
        return  0.0 

# Variables declaradas 
porcentajeAnho1 = 0.7
apreciacion_values = {'tradicional': 0.045, 'dorada': 0.09}  # Diccionario de valores de apreciación según la zona

# Leer el archivo CSV
with open(input_file_a, newline='', encoding='utf-8-sig') as csvfile:
    reader = csv.DictReader(csvfile)
    
    rows = list(reader)  # Convertir a lista para usar con tqdm
    total_rows = len(rows)  # Total de filas para la barra de progreso

    for row in tqdm(rows, desc="Procesando registros", total=total_rows):  # Añadir tqdm para la barra de progreso
        predio = row['predio']
        id_ = row['id']
        estado_uno = row['estado_uno']
        estatus_agosto = row['estatus_agosto']
        idea_final = row['idea_final']
        nombre = row['nombre']
        demolicion = to_number(row['demolicion'])
        rehabilitacion = to_number(row['rehabilitacion'])
        cn_cdh = to_number(row['cn_cdh'])
        cn_rest = to_number(row['cn_rest'])
        costo_alberca = to_number(row['costo_alberca'])
        costo_av = to_number(row['costo_av'])
        costo_estacionamiento = to_number(row['costo_estacionamiento'])
        costo_terraza = to_number(row['costo_terraza'])
        costo_terreno = to_number(row['costo_terreno'])
        aumento_amenidades = to_number(row['aumento_amenidades'])
        aumento_rehab = to_number(row['aumento_rehab'])
        aumento_cdh = to_number(row['aumento_cdh'])
        aumento_rest = to_number(row['aumento_rest'])
        aumento_comer = to_number(row['aumento_comer'])

        # Obtener el valor de la columna "zona"
        zona = row.get('zona', '').lower()  # Obtener la zona y convertirla a minúsculas
        apreciacionInmobiliaria = apreciacion_values.get(zona, 0.03)  # Asignar la apreciación según la zona

        # Verificar si el estado del predio está en los filtros de estatus
        if estado_uno in filtro_estatus:
            # Abrir el archivo de plantilla B
            wb_b = openpyxl.load_workbook(template_file_b)
            sheet_b = wb_b['CalendarioInv']  # Cambiar "CalendarioInv" al nombre real de la hoja
            sheet_c = wb_b['Flujo']

            # Llenar solo la celda de rehabilitación en B con los datos de A
            # Modificar si es necesario dependiendo del estatus
           
            sheet_b['F12'].value = rehabilitacion / 1000
            sheet_b['F12'].number_format = '0.00' if isinstance(rehabilitacion, float) else '@'

            sheet_b['F7'].value = costo_terreno  / 1000
            sheet_b['F7'].number_format = '0.00' if isinstance(costo_terreno, float) else '@'
            sheet_b['F14'].value = costo_av  / 1000
            sheet_b['F14'].number_format = '0.00' if isinstance(costo_av, float) else '@'
            sheet_b['F15'].value = costo_alberca  / 1000
            sheet_b['F15'].number_format = '0.00' if isinstance(costo_alberca, float) else '@'
            sheet_b['F16'].value = costo_estacionamiento  / 1000
            sheet_b['F16'].number_format = '0.00' if isinstance(costo_estacionamiento, float) else '@'
            sheet_b['F17'].value = costo_terraza  / 1000
            sheet_b['F17'].number_format = '0.00' if isinstance(costo_terraza, float) else '@'
                      
            # sheet_c
            sheet_c['K1'].value = apreciacionInmobiliaria
            sheet_c['K1'].number_format = '0.00' if isinstance(apreciacionInmobiliaria, float) else '@'
            sheet_c['E3'].value = porcentajeAnho1
            sheet_c['E3'].number_format = '0.00' if isinstance(porcentajeAnho1, float) else '@'   

            # Guardar el nuevo archivo
            output_file_name = f'{predio}_SOP.APRECIACION.{apreciacionInmobiliaria:.2f}_B.NRM.REHAB.xlsx'
            output_path = os.path.join(output_folder, output_file_name)
            wb_b.save(output_path)

print('Proceso completado.')
