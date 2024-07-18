import os
import configparser
import shutil
from datetime import datetime, timedelta
import openpyxl
from copy import copy
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.utils import get_column_letter, column_index_from_string

# ===========================
# = Manejo de Filas Excel =
# ===========================

def insertar_fila(sheet, now):
    try:
        sheet.insert_rows(8)
        for col in range(1, sheet.max_column + 1):
            column_letter = get_column_letter(col)
            
            celda_vieja = sheet[column_letter + '9']
            celda_nueva = sheet[column_letter + '8']

            celda_nueva.font = copy(celda_vieja.font)
            celda_nueva.fill = copy(celda_vieja.fill)
            celda_nueva.alignment = copy(celda_vieja.alignment)

        now = datetime.strftime(now, "%d/%m/%Y")
        sheet.cell(row=8, column=1).value = now

    except InvalidFileException as e:
        print(f"Error al trabajar con el archivo de Excel: {e}")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

def copiar_fila(origen_row, origen_col, detino_row, detino_col, sh):
    celda_origen = sh.cell(row=origen_row, column=origen_col)
    celda_destino = sh.cell(row=detino_row, column=detino_col)

    celda_destino.font = copy(celda_origen.font)
    celda_destino.fill = copy(celda_origen.fill)
    celda_destino.alignment = copy(celda_origen.alignment)

def crear_archivo_excel_si_no_existe(carpeta_excel, ruta_excel_base, now):
    try:
        archivo_excel_year = os.path.join(carpeta_excel, f"Bitacora_de_respaldos_BD_{now.year}.xlsx")

        if not os.path.exists(archivo_excel_year):
            print(f"Creando archivo excel en la ruta: {carpeta_excel}")
            shutil.copy2(ruta_excel_base, archivo_excel_year)
        else:
            print(f"El archivo excel ya existe en la ruta: {archivo_excel_year}")
    except Exception as e:
        print(f"Error inesperado de la excepcion (129): {e}")
    return archivo_excel_year

# ================================
# = Llenado de bitacora en excel =
# ================================

def comprobar_bakcup_realizado_sial(now, esquemas, ruta_base, letras, sheet):
    for i, esquema in enumerate(esquemas): 
        ruta_folder = os.path.join(ruta_base, esquema)
        if os.path.exists(ruta_folder):
            archivo = f"{esquema}-respaldo-{now.strftime('%Y%m%d')}.tar.gz"
            ruta_archivo = os.path.join(ruta_folder, archivo)
            if os.path.exists(ruta_archivo):
                sheet.cell(row=8, column=column_index_from_string(letras[i])).value = "X"
                print(f"El backup del esquema {esquema} se realizó correctamente")

def if_statement_in_backup_notSial(ruta_folder, idx, esquema, now, sheet, letras):
    try:
        lista_mongo = ['SIAL_HDE','SialCFDI','CAMPOBDB']

        if esquema == 'SEGUIDORES' or esquema == 'REPCIU_AYTO_ZAMORA':
            archivo = f"{esquema}-respaldo-{now.strftime('%Y%m%d')}.tar.gz"
            ruta_archivo = os.path.join(ruta_folder, archivo)
            if os.path.exists(ruta_archivo):
                sheet.cell(row=8, column=column_index_from_string(letras[idx])).value = "X"
                print(f"El backup del esquema {esquema} se realizó correctamente")
    
        if esquema == 'ZAM-SV-MORPHO2':
            archivo = f"MorphoManager-{now.strftime('%Y%m%d')}.7z"
            ruta_archivo = os.path.join(ruta_folder, archivo)
            if os.path.exists(ruta_archivo):
                sheet.cell(row=8, column=column_index_from_string(letras[idx])).value = "X"
                print(f"El backup del esquema {esquema} se realizó correctamente")

        if esquema in lista_mongo:
            archivo = f"{now.strftime('%Y%m%d')}-{esquema}.7z"
            ruta_archivo = os.path.join(ruta_folder, archivo)
            if os.path.exists(ruta_archivo):
                sheet.cell(row=8, column=column_index_from_string(letras[idx])).value = "X"
                print(f"El backup del esquema {esquema} se realizó correctamente")

    except Exception as e:
        print(f"Ocurrió un error inesperado al comprobar el backup del esquema {esquema}: {e}")

def comprobar_backup_realizado_not_sial(now, esquemas, ruta_base, letras, sheet):
    for i, esquema in enumerate(esquemas):
        ruta_folder = os.path.join(ruta_base, esquema)
        ruta_folder_mg = os.path.join(ruta_base, 'Mongo')

        if esquema in {'SIAL_HDE', 'SialCFDI', 'CAMPOBDB'} and os.path.exists(ruta_folder_mg):
            if_statement_in_backup_notSial(ruta_folder_mg, i, esquema, now, sheet, letras)
        if os.path.exists(ruta_folder):
            if_statement_in_backup_notSial(ruta_folder, i, esquema, now, sheet, letras)

# ===========================
# = Manejo del archivo .ini =
# ===========================

def escribir_configuracion(ruta_archivo, esquemas, letras):
    config = configparser.ConfigParser()
    config['ESQUEMAS'] = {
        'esquema': ','.join(esquemas),
        'letras': ','.join(letras)
    }
    
    with open(ruta_archivo, 'w') as configfile:
        config.write(configfile)

def leer_configuracion(ruta_archivo):
    config = configparser.ConfigParser()
    config.read(ruta_archivo)
    
    esquemas = config['ESQUEMAS']['esquema'].split(',')
    letras = config['ESQUEMAS']['letras'].split(',')
    
    return esquemas, letras

def actualizar_esquemas_y_letras(ruta_base, esquemas, letras, lista_mongo):
    esquema_set = set(esquemas) # conjunto de esquemas para hacer busquedas más eficientes
    pos_final_num = column_index_from_string(letras[-1]) # número de la última letra en la lista de letras

    nombres_nuevos, letras_nuevas = [], [] # listas para almacenar los nuevos esquemas y letras
    ctr = 1

    exclusiones = {'SIAL_PRUEBA', 'LOBORH', 'FREXPORT', 'Mongo'}

    try:
        dirs_to_process = [d for d in os.listdir(ruta_base) 
                           if os.path.isdir(os.path.join(ruta_base, d)) and 
                           d not in esquema_set and 
                           d not in exclusiones]
        dirs_to_process.extend(subdir for subdir in lista_mongo if subdir not in esquema_set)

        for nombre_dir in dirs_to_process:
            nombres_nuevos.append(nombre_dir)
            letras_nuevas.append(get_column_letter(pos_final_num + ctr))
            ctr += 1

    except PermissionError as e:
        print(f"No se pudo acceder a la ruta {ruta_base}: {e}")
        
    return nombres_nuevos, letras_nuevas

def actualizar_archivo_de_config(ruta_base, esquemas, letras):
    lista_mongo = ['SIAL_HDE','SialCFDI','CAMPOBDB']

    nombres_nuevos, letras_nuevas = actualizar_esquemas_y_letras(ruta_base, esquemas, letras, lista_mongo)

    if nombres_nuevos or letras_nuevas:
        esquemas.extend(nombres_nuevos)
        letras.extend(letras_nuevas)

        config = configparser.ConfigParser()
        config.read('config.ini')

        esquema_str = ','.join(esquemas)
        letra_str = ','.join(letras)

        config.set('ESQUEMAS', 'esquema', esquema_str)
        config.set('ESQUEMAS', 'letras', letra_str)

        with open('config.ini', 'w') as configfile:
            config.write(configfile)

        print(f"Se añadieron los siguientes esquemas: {nombres_nuevos}")
        print(f"Se añadieron las siguientes letras: {letras_nuevas}")

    return esquemas, letras

# ===========================
# = Llamado a las funciones =
# ===========================

def main_function(sheet, now, esquemas, ruta_base, letras):
    insertar_fila(sheet, now)

    comprobar_bakcup_realizado_sial(now, esquemas, ruta_base, letras, sheet)
    comprobar_backup_realizado_not_sial(now, esquemas, ruta_base, letras, sheet)
    
    # se añaden los nombres de las carpetas a la bitácora
    # esto deberia ir primero y despues checar los backups
    for i, letra in enumerate(letras):
        if esquemas[i] not in ['Mongo', 'LOBORH', 'FREXPORT', 'SIAL_PRUEBA']:            
            copiar_fila(7, column_index_from_string(letra) - 1, 7, column_index_from_string(letra), sheet)
            sheet.cell(row=7, column=column_index_from_string(letra)).value = esquemas[i]

# =========================

# ruta actual: /home/julianq@local.lobos.com.mx/Documents/Python/Bitacora_BD_v2
ruta_actual = os.path.dirname(__file__)
nombre_excel_base = "Bitacora_de_respaldos_BD"
carpeta_excel = os.path.join(ruta_actual, nombre_excel_base)
ruta_excel_base = os.path.join(ruta_actual, nombre_excel_base + ".xlsx")

#ruta_base = '/mnt/DMP/'
ruta_base = r'\\192.0.0.175\respaldos'

ruta_archivo_config = os.path.join(ruta_actual, "config.ini")

esquemas, letras = leer_configuracion(ruta_archivo_config)
esquemas, letras = actualizar_archivo_de_config(ruta_base, esquemas, letras)

if len(letras) != len(esquemas):
    print('Las listas no tienen la misma cantidad de elementos')

# now = datetime.now().date()
now = datetime(2024, 6, 18).date() # Fecha de prueba
fecha_modificacion = datetime(2024, 6, 1).date()

os.makedirs(carpeta_excel, exist_ok=True)

archivo_excel_year = crear_archivo_excel_si_no_existe(carpeta_excel, ruta_excel_base, fecha_modificacion)

workbook = openpyxl.load_workbook(archivo_excel_year)
sheet = workbook.active

llenar_bitacora = False # Variable para controlar si se debe llenar la bitácora o no

# ______________________________________________________________________________________________________________________
if llenar_bitacora:
    while fecha_modificacion < now :
        main_function(sheet, fecha_modificacion, esquemas, ruta_base, letras)

        fecha_modificacion += timedelta(days=1)
else:
    main_function(sheet, now, esquemas, ruta_base, letras)
# ______________________________________________________________________________________________________________________

workbook.save(archivo_excel_year)