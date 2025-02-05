# Obtener información de las tablas
Este algoritmo permite obtener la información de las figuras: título de la tabla, tabla, nota de la tabla.

## Instalación

### Prerequisitos
- Descargar e instalar python de la página oficial.
- Crear y ejecutar un un entorno virtual.
- Instalar las dependencias dentro del entorno con el comando
```
pip install -r .\requirements.txt
```
- Instalar "gtk3-runtime" de la página oficial (https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases)
## Requisitos
- Especificar los cursos a analizar en el archivo "courses.txt"
- Especificar el token de extración dentro del archivo "get_images.py" en la variable "API_KEY"
- Crear las carpetas "imagenes", "svg_images", "table_results".
## Proceso
- Ejecutar el archivo ""get_tables.py" con el comando "python get_tables.py" o "python3 get_tables.py"
## Resultados
- Se genera un documento ".docx" con el código SIS del curso analizado.
