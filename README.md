# Proyecto de Bitácora de Respaldos BD

Este proyecto está diseñado para manejar y registrar los respaldos de esquemas de bases de datos en un archivo de Excel. Utiliza la biblioteca `openpyxl` para la manipulación del archivo de Excel y `configparser` para manejar la configuración del proyecto.

## Estructura del Código

### Manejo de Filas en Excel

- **insertar_fila(sheet, now):** Inserta una nueva fila en la hoja de Excel y copia los estilos de la fila anterior. Establece la fecha actual en la primera celda de la nueva fila.
- **copiar_fila(origen_row, origen_col, detino_row, detino_col, sh):** Copia el contenido y los estilos de una celda a otra.
- **crear_archivo_excel_si_no_existe(carpeta_excel, ruta_excel_base, now):** Crea un nuevo archivo de Excel si no existe para el año actual.

### Llenado de Bitácora en Excel

- **comprobar_bakcup_realizado_sial(now, esquemas, ruta_base, letras, sheet):** Verifica si los respaldos de ciertos esquemas se han realizado y marca una "X" en la hoja de Excel.
- **if_statement_in_backup_notSial(ruta_folder, idx, esquema, now, sheet, letras):** Verifica si los respaldos de otros esquemas se han realizado y marca una "X" en la hoja de Excel.
- **comprobar_backup_realizado_not_sial(now, esquemas, ruta_base, letras, sheet):** Controla la lógica para verificar los respaldos de esquemas no incluidos en `SIAL`.

### Manejo del Archivo de Configuración

- **escribir_configuracion(ruta_archivo, esquemas, letras):** Escribe la configuración de esquemas y letras en un archivo `.ini`.
- **leer_configuracion(ruta_archivo):** Lee la configuración de esquemas y letras desde un archivo `.ini`.
- **actualizar_esquemas_y_letras(ruta_base, esquemas, letras, lista_mongo):** Actualiza las listas de esquemas y letras con nuevos esquemas encontrados en la ruta base.
- **actualizar_archivo_de_config(ruta_base, esquemas, letras):** Actualiza el archivo de configuración con los nuevos esquemas y letras.

### Función Principal

- **main_function(sheet, now, esquemas, ruta_base, letras):** Llama a las funciones necesarias para insertar filas, verificar respaldos y actualizar la hoja de Excel.

## Uso

1. **Configurar Rutas:**

   Asegúrate de que las rutas `ruta_actual`, `nombre_excel_base`, `carpeta_excel`, `ruta_excel_base` y `ruta_base` estén configuradas correctamente en el código.

2. **Archivo de Configuración:**

   El archivo de configuración `config.ini` debe estar en el mismo directorio que el script principal y debe contener las secciones `ESQUEMAS` con las claves `esquema` y `letras`.

3. **Ejecución del Script:**

   Para ejecutar el script, simplemente ejecuta el archivo principal. El script verificará si los respaldos se han realizado y actualizará la bitácora en el archivo de Excel correspondiente.

   ```python
   python nombre_del_script.py
