# Automatización Bitácora

## Breve descripción 
El script realiza el llenado de una bitácora de la siguiente forma:

1. **Carga de datos de configuración**: Se cargan los datos necesarios desde los archivos de configuración.

2. **Actualización de datos**: Si es necesario, se actualizan estos datos. Es posible que se hayan añadido nuevos esquemas y se deba añadir letras correspondientes.

3. **Ejecución del archivo Java**: Se ejecuta el archivo Java para procesar la información necesaria.

4. **Inserción en el archivo Excel**: Se inserta una fila en el archivo Excel y se agrega la fecha correspondiente.

5. **Comprobación de backups**: Se verifica la existencia de backups para la fecha en varios directorios mediante una función específica.

6. **Marcado de casillas**: En caso de encontrar los backups, se marca la casilla correspondiente con una "X". Este proceso se repite para cada una de las empresas en la lista de esquemas y algunas más adicionales.

  

## Funcionamiento de las Rutas y cómo Modificarlas

Las rutas en el código están definidas principalmente al inicio del script y se pueden modificar directamente editando las variables ruta_actual, ruta_base, ruta_excel_base, file_path_conf, file_path_lobo, ruta_archivo_config, y ruta_java. Estas rutas son utilizadas para ubicar archivos de configuración, archivos de Excel, ejecutables JAR, y directorios de datos necesarios para el funcionamiento del sistema.


## Permisos Necesarios

Para ejecutar correctamente este código, se necesitan permisos de lectura y escritura en los directorios y archivos especificados por las rutas definidas. Especialmente, asegúrate de tener permisos para modificar archivos de configuración (.conf, .lobo, .ini) y crear archivos Excel en la carpeta designada.

1. ### modificar_valcontenidoor_conf(ruta_archivo, clave, nuevo_valor)

   **Qué**: Modifica un valor específico dentro de un archivo de configuración.  
   **Cómo**: Abre el archivo de configuración especificado, busca la clave proporcionada y actualiza su valor con nuevo_valor.  
   **Dónde**: Utilizado para ajustar configuraciones en archivos .conf.  
   **Por Qué**: Permite la personalización dinámica de configuraciones sin necesidad de editar manualmente los archivos.  

   #### Parámetros:

   **ruta_archivo (str)**: Ruta completa al archivo de configuración (.conf).  
   **clave (str)**: Clave del parámetro que se desea modificar dentro del archivo.  
   **nuevo_valor (str)**: Nuevo valor que se desea asignar a la clave especificada.  

2. ### limpiar_archivo(ruta_archivo)

   **Qué**: Limpia el contenido de un archivo, eliminando saltos de línea y normalizando el formato.  
   **Cómo**: Lee el archivo, elimina los caracteres de salto de línea y sobrescribe el archivo con su contenido limpio.  
   **Dónde**: Utilizado para preparar archivos de configuración antes de su lectura o escritura.  
   **Por Qué**: Asegura que los archivos de configuración estén en un formato coherente para evitar errores de lectura.  

   #### Parámetros:

   **ruta_archivo (str)**: Ruta completa al archivo que se desea limpiar y normalizar.

3. ### leer_conf(ruta_completa)

   **Qué**: Lee y analiza un archivo de configuración .conf o .lobo para obtener información específica.  
   **Cómo**: Utiliza configparser para parsear y extraer datos según el tipo de archivo especificado.  
   **Dónde**: Esencial para cargar datos de configuración inicial.  
   **Por Qué**: Facilita la carga de parámetros críticos como asuntos y fechas desde archivos de configuración.  

   #### Parámetros:

   **ruta_completa (str)**: Ruta completa al archivo de configuración (.conf o .lobo) que se desea leer.

4. ### ejecucionJava(ruta_actual, ruta_java)

   **Qué**: Ejecuta un archivo JAR de Java con parámetros específicos y gestiona la salida y errores del proceso.  
   **Cómo**: Utiliza subprocess.Popen para iniciar el proceso Java con el archivo JAR y capturar la salida estándar y errores.  
   **Dónde**: Se aplica para ejecutar un componente Java que realiza tareas específicas de marcado (en este caso, marcarAltex.jar).  
   **Por Qué**: Es crucial para la integración con sistemas Java y para automatizar tareas complejas que requieren procesamiento externo.  

   #### Parámetros:

   **ruta_actual (str)**: Ruta del directorio actual desde donde se ejecuta el proceso.  
   **ruta_java (str)**: Ruta completa al archivo JAR de Java que se desea ejecutar.

5. ### insertar_fila(sheet, now)

   **Qué**: Inserta una nueva fila en una hoja de Excel en la posición 8, copiando el estilo de la fila siguiente y añadiendo la fecha actual en la primera celda.  
   **Cómo**: Utiliza métodos de openpyxl para insertar y manipular filas y celdas en la hoja especificada.  
   **Dónde**: Utilizado en la bitácora de respaldos Excel para registrar nuevas entradas diarias.  
   **Por Qué**: Permite mantener una bitácora organizada y estructurada de los respaldos realizados.  

   #### Parámetros:

   **sheet (openpyxl.Worksheet)**: Hoja de Excel donde se desea insertar la nueva fila.  
   **now (datetime.datetime)**: Objeto de fecha y hora actual que se insertará en la primera celda de la nueva fila.

6. ### copiar_fila(origen_row, origen_col, destino_row, destino_col, sh)

   **Qué**: Copia el estilo de una celda de origen a una celda de destino en una hoja de Excel.  
   **Cómo**: Copia propiedades como fuente, relleno y alineación de la celda de origen a la de destino.  
   **Dónde**: Utilizado para mantener consistencia visual al insertar nuevas filas en la bitácora Excel.  
   **Por Qué**: Asegura que todas las filas en la bitácora mantengan el formato correcto al insertar nuevas entradas.  

   #### Parámetros:

   **origen_row (int)**: Índice de la fila de origen desde donde se copiará el estilo.  
   **origen_col (int)**: Índice de la columna de origen desde donde se copiará el estilo.  
   **destino_row (int)**: Índice de la fila de destino donde se aplicará el estilo copiado.  
   **destino_col (int)**: Índice de la columna de destino donde se aplicará el estilo copiado.  
   **sh (openpyxl.Worksheet)**: Hoja de Excel donde se encuentran las celdas de origen y destino.

7. ### crear_archivo_excel_si_no_existe(carpeta_excel, ruta_excel_base, now)

   **Qué**: Crea un archivo Excel de bitácora si no existe en la carpeta especificada, basándose en una plantilla Excel base.  
   **Cómo**: Verifica la existencia del archivo en carpeta_excel, lo crea copiando la plantilla ruta_excel_base si no está presente.  
   **Dónde**: Aplicado al inicio del programa para garantizar que la bitácora Excel esté lista para registrar los respaldos diarios.  
   **Por Qué**: Facilita la gestión y registro ordenado de respaldos en un formato accesible y estructurado.  

   #### Parámetros:

   **carpeta_excel (str)**: Ruta de la carpeta donde se buscará o creará el archivo Excel.  
   **ruta_excel_base (str)**: Ruta completa a la plantilla Excel base que se utilizará para crear el archivo de bitácora.  
   **now (datetime.datetime)**: Objeto de fecha y hora actual, utilizado para identificar el archivo de bitácora correspondiente al día.

8. ### comprobar_backup(now, esquemas, ruta_base, letras, sheet, marcado_lista)

   **Qué**: Comprueba la existencia de respaldos para una lista de esquemas dados y marca la bitácora Excel en consecuencia.  
   **Cómo**: Itera sobre los esquemas proporcionados, verifica la existencia de archivos de respaldo y actualiza la celda correspondiente en la hoja de Excel sheet.  
   **Dónde**: Aplicado para actualizar diariamente la bitácora Excel con el estado de los respaldos.  
   **Por Qué**: Proporciona visibilidad inmediata sobre el estado de los respaldos para múltiples esquemas, facilitando la gestión y supervisión.  

   #### Parámetros:

   **now (datetime.datetime)**: Objeto de fecha y hora actual, utilizado para identificar el estado de los respaldos correspondientes al día.  
   **esquemas (list)**: Lista de nombres de esquemas cuyos respaldos se verificarán.  
   **ruta_base (str)**: Ruta base donde se buscarán los directorios de respaldo de cada esquema.  
   **letras (list)**: Lista de letras que representan las columnas en la hoja de Excel donde se realizará el marcado.  
   **sheet (openpyxl.Worksheet)**: Hoja de Excel donde se actualizará el estado de los respaldos.  
   **marcado_lista (list)**: Lista utilizada para registrar el estado de marcado de los respaldos.

9. ### marcado_de_esquemas(ruta_folder, idx, esquema, now, sheet, letras, marcado_lista)

   **Qué**: Marca la existencia de respaldos de un esquema específico en la hoja de Excel según la presencia de archivos de respaldo en un directorio.  
   **Cómo**: Verifica la existencia de archivos de respaldo en ruta_folder/esquema, actualiza la celda correspondiente en la hoja de Excel sheet y registra el estado en marcado_lista.  
   **Dónde**: Utilizado dentro de comprobar_backup para cada esquema individualmente.  
   **Por Qué**: Permite una representación visual clara y precisa del estado de los respaldos para cada esquema en la bitácora Excel.  

   #### Parámetros:

   **ruta_folder (str)**: Ruta base del directorio donde se encuentran los respaldos de los esquemas.  
   **idx (int)**: Índice utilizado para realizar un seguimiento del estado de marcado en marcado_lista.  
   **esquema (str)**: Nombre del esquema del cual se verificará la existencia de respaldos.  
   **now (datetime.datetime)**: Objeto de fecha y hora actual, utilizado para identificar el estado de los respaldos correspondientes al día.  
   **sheet (openpyxl.Worksheet)**: Hoja de Excel donde se actualizará el estado de los respaldos.  
   **letras (list)**: Lista de letras que representan las columnas en la hoja de Excel donde se realizará el marcado.  
   **marcado_lista (list)**: Lista utilizada para registrar el estado de marcado de los respaldos.

10. ### leer_configuracion(ruta_archivo)

    **Qué**: Lee un archivo de configuración .ini para obtener listas de esquemas y letras asociadas.  
    **Cómo**: Utiliza configparser para parsear el archivo y extraer los datos necesarios.  
    **Dónde**: Utilizado al inicio del programa para cargar configuraciones iniciales de esquemas y columnas de la bitácora Excel.  
    **Por Qué**: Facilita la personalización y configuración inicial del programa mediante la lectura de un archivo de configuración estructurado.  

    #### Parámetros:

    **ruta_archivo (str)**: Ruta completa al archivo de configuración .ini que se desea leer para cargar las configuraciones.

11. ### actualizar_archivo_de_config(ruta_base, esquemas, letras, lista_mongo, asunto)

    **Qué**: Actualiza dinámicamente un archivo de configuración .ini con nuevos esquemas y letras según la existencia de directorios y criterios específicos.  
    **Cómo**: Añade nuevos esquemas y letras al archivo de configuración si no están presentes, basándose en condiciones predefinidas.  
    **Dónde**: Utilizado para mantener actualizado el archivo de configuración con los esquemas y columnas actuales de la bitácora Excel.  
    **Por Qué**: Automatiza la gestión de configuraciones, adaptándolas dinámicamente a cambios en la estructura de directorios y necesidades del programa.  

    #### Parámetros:

    **ruta_base (str)**: Ruta base donde se encuentran los directorios de los esquemas.  
    **esquemas (list)**: Lista de nombres de esquemas cuyos directorios se deben verificar y añadir al archivo de configuración.  
    **letras (list)**: Lista de letras asociadas a los esquemas que se añadirán al archivo de configuración.  
    **lista_mongo (list)**: Lista de datos adicionales que pueden ser requeridos para la actualización del archivo.  
    **asunto (str)**: Información adicional o contexto que puede ser relevante para la actualización del archivo.

## Dependencias
Para ejecutar este script, es necesario tener Python 3 instalado.

### Instalación de Python

**Para Linux:**

1. Para instalar Python 3:
    ```bash
    sudo apt update
    sudo apt install python3
    ```

2. Para actualizar Python 3 a la última versión:
    ```bash
    sudo apt update
    sudo apt upgrade python3
    ```

**Para macOS:**

1. Python 2 viene preinstalado en macOS. Para instalar Python 3 utilizando Homebrew:
    ```bash
    brew install python@3
    ```

2. Para actualizar Python 3 a la última versión:
    ```bash
    brew upgrade python@3
    ```

**Para Windows:**

1. Descarga el instalador de Python desde el sitio web oficial [Python.org](https://www.python.org/downloads/) y ejecútalo. Asegúrate de marcar la casilla "Add Python to PATH" durante la instalación para poder acceder a Python desde la línea de comandos.

2. Para actualizar Python en Windows, descarga el instalador de la versión más reciente desde el sitio web oficial y ejecútalo, seleccionando "Modificar" o "Actualizar" para actualizar tu instalación existente.

### Creación del entorno virtual

**Para Linux y macOS:**

1. Crear un nuevo entorno virtual en el directorio 'myenv':
    ```bash
    python3 -m venv myenv
    ```

2. Activar el entorno virtual:
    ```bash
    source myenv/bin/activate
    ```

   Ahora estás dentro del entorno virtual 'myenv'.

3. Para desactivar el entorno virtual, simplemente ejecuta:
    ```bash
    deactivate
    ```

**Para Windows:**

1. Crear un nuevo entorno virtual en el directorio 'myenv':
    ```bash
    python -m venv myenv
    ```

2. Activar el entorno virtual:
    ```bash
    myenv\Scripts\activate
    ```

   Ahora estás dentro del entorno virtual 'myenv'.

3. Para desactivar el entorno virtual en Windows, ejecuta el mismo comando que para Linux:
    ```bash
    deactivate
    ```

Una vez creado el entorno, puedes proceder a instalar las librerías necesarias. Este programa requiere la instalación de varios módulos:
- **openpyxl**
- **configparser**
- **et-xmlfile**

Para instalar estos módulos, ejecuta el siguiente comando en la terminal desde el directorio actual del proyecto:  

```bash
pip install -r requirements.txt
```

Para la ejecución del archivo .jar se necesitan ciertas librerias, las cuales estan la carpeta lib  

## Archivos de configuración 

### config.ini
Dentro de este archivo, encontrarás dos listas: `esquemas` y `letras`.

1. La lista `esquemas` contiene los nombres de las empresas.
2. La lista `letras` está relacionada con `esquemas`, ya que cada elemento de `letras` representa una letra que simboliza una columna en una hoja de cálculo de Excel.

Específicamente:
- El primer elemento de la lista `letras` corresponde a la letra de la columna asociada con el primer elemento de la lista `esquemas`.
- De manera similar, el segundo elemento de `letras` se relaciona con el segundo elemento de `esquemas`, y así sucesivamente.

Esta estructura permite mapear de manera directa cada nombre de empresa en `esquemas` a una columna específica en Excel mediante las letras en `letras`.

### configMarcado.conf
Estas variables son utilizadas por el archivo JAR:

- **`asunto`**: Representa los nombres de las empresas.
- **`fechaToday`**: Puede tener dos valores posibles: "Sí" o "No". 
  - Si `fechaToday` es afirmativo ("Sí"), se utilizará la fecha actual.
  - Si `fechaToday` es negativo ("No"), se utilizará el valor de la variable `fecha`.

La variable `fecha` puede ser modificada para consultar el estado de los backups en una fecha específica.

### Marcado_altex.lobo
Este archivo no se modifica. Su propósito es proporcionar al script los valores de retorno del archivo .jar para que pueda realizar operaciones con ellos.  

## Funcionameinto del archivo JAVA
El script `MarcarAltex` es una aplicación Java que se conecta a una cuenta de correo electrónico utilizando el protocolo IMAP para buscar mensajes recibidos en una fecha específica. Luego, verifica si los correos contienen ciertos asuntos y palabras clave, y marca en un archivo de configuración si los respaldos correspondientes se han realizado.  

### Funcionamiento

1. **Carga de configuración**:
    - Se cargan las configuraciones necesarias desde el archivo `configMarcado.conf`.

2. **Conexión al servidor de correo**:
    - Utiliza las propiedades configuradas para conectarse al servidor de correo mediante IMAP.

3. **Búsqueda de mensajes**:
    - Busca los mensajes en la bandeja de entrada (`INBOX`) que fueron recibidos en una fecha específica.

4. **Verificación del contenido de los mensajes**:
    - Revisa si el asunto y el contenido de los mensajes contienen las palabras clave especificadas en la configuración.

5. **Actualización de estado**:
    - Si se encuentran las palabras clave correspondientes, se actualiza el estado en el archivo `Marcado_altex.lobo`.

## Estructura del proyecto
BITACORA_BD_V2/
* src/
  * mainv2.py              # Script principal para ejecutar el llenado de la bitácora
  * config.ini             # Archivo de configuración principal
  * configMarcado.conf     # Archivo de variables adicionales
  * marcarAltex.jar        # Archivo Java que se ejecuta durante el proceso
  * Marcado_altex.lobo     # Archivo de retorno del ejecutable jar
    
* Bitacora_de_respaldos_BD/
  * Bitacora_APP.xlsx      # Archivo Excel base de la bitácora
  * Bitacora_APP_2024.xlsx # Archivo Excel de la bitácora para el año 2024
* logs/
  * archivo_log.txt        # Archivo de logs para registrar errores y eventos
* README.md                # Documentación del proyecto
