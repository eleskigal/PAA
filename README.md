# Descripción general
El código es un script de Python que se encarga de descargar la versión actualizada del plan anual de adquisiciones, manipular sus datos, crear una tabla y guardar el archivo modificado en una nueva ubicación. Además, utiliza la librería schedule para ejecutar la tarea repetidamente cada 50 segundos.

# Funciones
El código utiliza varias funciones para realizar tareas específicas:

download_file(url, save_path, file_name): Descarga un archivo desde una URL y lo guarda en la ubicación especificada.
open_workbook(file_path): Abre un archivo Excel a partir de su ruta de acceso y devuelve el objeto del libro.
save_workbook(wb, file_path): Guarda el objeto de libro de Excel modificado en la ruta de acceso especificada y devuelve la ruta de acceso modificada.
copy_rows_to_new_workbook(wb): Copia las filas de la hoja de trabajo activa del libro de Excel especificado en un nuevo libro de trabajo y devuelve el objeto de libro de trabajo nuevo.
create_table(ws): Crea un objeto de tabla de Excel en la hoja de trabajo especificada y devuelve la hoja de trabajo modificada.
scrape_and_save(): Descarga un archivo de Excel desde una URL, manipula los datos, guarda los datos modificados en un nuevo archivo y elimina el archivo original.
main(): Función principal que ejecuta la tarea de raspado repetidamente a intervalos regulares de tiempo.


# Bibliotecas
El código utiliza varias bibliotecas de Python, que incluyen:

requests: para hacer solicitudes HTTP.
os: para realizar operaciones en el sistema de archivos.
urllib3: para desactivar la verificación SSL.
openpyxl: para manipular archivos de Excel.
io: para trabajar con objetos de bytes.
datetime: para trabajar con fechas.
pyautogui: para la automatización de GUI.
time: para trabajar con el tiempo.
schedule: para programar la tarea de raspado en intervalos regulares.

# Constantes

El código utiliza varias constantes, que incluyen:

SAVE_PATH: la ruta de acceso donde se guarda el archivo Excel descargado.
FILE_NAME: el nombre del archivo Excel descargado.
URL: la URL desde la cual se descarga el archivo Excel.

# Mantenimiento del código
Para mantener este código a lo largo del tiempo, es importante tener en cuenta lo siguiente:

Asegurarse de que todas las bibliotecas de Python utilizadas estén actualizadas a las versiones más recientes.
Realizar pruebas regulares en el código para asegurarse de que sigue funcionando correctamente.
Utilizar un sistema de control de versiones como Git para controlar los cambios realizados en el código a lo largo del tiempo.
Documentar el código de manera adecuada y clara para facilitar su lectura y mantenimiento en el futuro.
Implementar la funcionalidad de registro para monitorear cualquier error o problema que pueda surgir.
