# Obtener información de las tablas
Este algoritmo permite obtener la información de las figuras:
- Título de la tabla
- Tabla
- Nota de la tabla

## Instalación

### Prerequisitos
- Instalar python
- Instalar un entorno virtual
- Instalar las dependencias dentro del entorno: cairosvg, pypandoc, python-docx, beautifulSoup, canvasapi
## Requisitos
- Especificar los cursos a analizar en el archivo "courses.txt"
- Especificar el token de extración dentro del archivo "get_images.py" en la variable "API_KEY"
- Crear las carpetas "imagenes", "svg_images", "table_results".
## Proceso
- Ejecutar el archivo ""get_tables.py" con el comando "python get_tables.py" o "python3 get_tables.py"
## Resultados
- Se genera un documento ".docx" con el código SIS del curso analizado.
