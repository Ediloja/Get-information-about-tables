import os
import re
import html
import requests

from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from canvasapi import Canvas
from bs4 import BeautifulSoup, Tag

import cairosvg

import pypandoc

API_V1 = "https://utpl.instructure.com/api/v1"
API_URL = "https://utpl.instructure.com"
API_KEY = ""

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json",  # Opcional, dependiendo de la API
}


def get_connection():
    """
    Establece la conexión con la plataforma Canvas

    Parámetros:
    API_URL (str): es el dominio de Canvas
    API_KEY (str): es el token con el cual se va a trabajar

    Retorna:
    canvas (obj): Instancia a Canvas para la conexión respectiva
    """
    canvas = Canvas(API_URL, API_KEY)

    return canvas


def get_number(text):
    """
    Obtene un número de una cadena de  texto

    Parámetros:
    text (str): corresponde a la URL del curso

    Retorna:
    int:
        - id_curso -> si encontró un número
        - 0 -> si no encontro números
    """

    expression = re.findall(r"\d+", text)

    if len(expression) != 0:
        return expression[0]
    else:
        return 0


def get_url_pages(course):
    """
    Obtiene las URLs de las páginas ordenadas

    Parámetros:
    course (obj): curso que se desea obtener la información

    Retorna:
    (list): lista de IDs de las páginas
    """

    modules = course.get_modules()

    list_pages = []

    for m in modules:
        items = m.get_module_items()

        for i in items:
            if i.type == "Page" and "semana" in i.title.lower():
                list_pages.append(i.page_url)
            elif i.type == "Page" and "week" in i.title.lower():
                list_pages.append(i.page_url)
    return list_pages


def delete_tags(html):
    """
    Elimina etiquetas del HTML "link" y "script"

    Parámetros:
    html (str): html que se desea limpiar

    Retorna:
    (str): html limpio
    """
    soup = BeautifulSoup(html, "html.parser")

    # Se elimina las etiquetas link y script
    for tag in soup(["link", "script"]):
        tag.decompose()

    html = str(soup)

    return html


def identify_class(course, html):
    """
    Identifica si existe botón de continuar y se agrega el html de las páginas adicionales

    Parámetros:
    course (obj): curso que se debe analizar
    html (str): html donde se debe identificar las clases

    Retorna:
    (str): html con la información adicional (si fuera el caso)
    """

    soup = BeautifulSoup(html, "html.parser")

    # Se obtiene todas las etiquetas <a> siempre y cuando tengan el atributo "data-api-returntype = True" y no digan semana o week
    btn_continue = [
        link
        for link in soup.find_all("a")
        if link.get("data-api-returntype", "").lower() == "page"
        and "semana" not in link.text.strip().lower()
        and "week" not in link.text.strip().lower()  # Excluir si contiene "semana"
    ]

    for b in btn_continue:
        url = b["data-api-endpoint"]
        response = requests.get(url, headers=HEADERS)

        data = response.json()
        url_page = data.get("url")
        page = course.get_page(url_page)

        # Se incluye el html de la página externa dentro de la página
        tag_p = b.parent
        # Se limpia el html de la página externa
        html_cleaning = delete_tags(page.body)
        soup_external = BeautifulSoup(html_cleaning, "html.parser")

        tag_p.replace_with(soup_external)

    return str(soup)


def get_tables(course, html):
    """
    Identifica título de tablas, contenido y nota

    Parámetros:
    course (obj): curso que se desea analizar
    html (str): html donde se debe identificar las tablas

    Retorna:
    (str): html con la información en html con la información respectiva
    """

    soup_result = BeautifulSoup("", "html.parser")

    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")

    for table in tables:
        search_limit = 5  # Límite de búsqueda de hermanos
        pre = table.find_next_sibling()
        soup_result.append(table)

        # Buscar hasta el límite de búsqueda
        for _ in range(search_limit):
            if pre is None:
                break  # Evita el error si no hay más elementos
            if pre.name == "table":
                break
            if isinstance(pre, Tag) and pre.name == "pre":
                # Elimina estilos de la etiqueta pre
                del pre["style"]

                # Reemplazamos la etiqueta por <p>
                str_pre = str(pre).replace("<pre>", "<p>")
                str_pre = str_pre.replace("</pre>", "</p>")
                soup_result.append(str_pre)

                break  # Se encontró <pre>, detener búsqueda

            pre = (
                pre.find_next_sibling()
            )  # Continuar con el siguiente hermano si no es <pre>

    return str(soup_result)


def agregar_hipervinculo(parrafo, texto, url, document):
    """
    Agrega un hipervínculo a un párrafo en un documento Word.

    :param parrafo: Párrafo donde se insertará el enlace.
    :param texto: Texto visible del enlace.
    :param url: URL del hipervínculo.
    :param document: Documento Word donde se inserta la relación del enlace.
    """
    # Crear una relación de hipervínculo en el documento
    r_id = document.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )

    # Crear el elemento <w:hyperlink>
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)  # Asignar ID de la relación

    # Crear el "run" (<w:r>) que contendrá el texto del enlace
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    # Aplicar color azul
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0000FF")

    # Aplicar subrayado
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")

    rPr.append(color)
    rPr.append(underline)
    run.append(rPr)

    # Agregar el texto al "run"
    text = OxmlElement("w:t")
    text.text = texto
    run.append(text)

    # Agregar el "run" al hipervínculo y el hipervínculo al párrafo
    hyperlink.append(run)
    parrafo._element.append(hyperlink)


def write_file(filename, html):
    """
    Escribe el archivo

    Parámetros:
    filename (str): ruta, nombre y extensión del archivo (ej. 'archivo.html')
    html (str): html que se desea escribir en el archivo

    Retorna:
    (file): archivo guardado
    """
    with open(filename, "w", encoding="utf-8") as file:
        file.write(html)

    print(f"Documento guardado como: {filename}")


def decoding_html(html_text):
    """
    Limpia el HTML decodificando caracteres especiales.

    Parámetros:
    html (str): html que se desea limpiar

    Retorna:
    (str): html decodificado
    """
    return html.unescape(html_text)


def replace_br(html):
    """
    Reemplaza los saltos de línea por un punto y un espacio (título de la imagen)

    Parámetros:
    html (str): html que se desea reemplazar

    Retorna:
    (str): html sin etiquetas <br>
    """
    soup = BeautifulSoup(html, "html.parser")

    for br in soup.find_all("br"):
        br.replace_with(" ")

    return str(soup)

def process_tables(html):
    """
    Procesa las tablas y las imágenes en el HTML.

    Parámetros:
    html (str): HTML que se desea procesar

    Retorna:
    (str): HTML procesado
    """
    image_folder = "imagenes"
    os.makedirs(image_folder, exist_ok=True)
    soup = BeautifulSoup(html, "html.parser")

    # Agregar estilo de bordes a todas las tablas
    for table in soup.find_all("table"):
        table["style"] = "border: 1px solid black; border-collapse: collapse;"
        table["border"] = "1"  # Compatibilidad

    # Asegurar que todo dentro de <thead> esté en negrita
    for thead in soup.find_all("thead"):
        for element in thead.find_all():
            element.wrap(soup.new_tag("strong"))

    # Procesar imágenes
    for i, img_tag in enumerate(soup.find_all("img")):
        img_url = img_tag.get("src")

        if img_url and img_url.startswith("http"):
            img_filename = os.path.join(image_folder, f"imagen_externa_{i}")
            response = requests.get(img_url)

            if response.status_code == 200:
                content_type = response.headers.get("Content-Type", "")

                if "image/svg+xml" in content_type or b"<svg" in response.content[:500]:
                    svg_filename = img_filename + ".svg"
                    png_filename = img_filename + ".png"

                    with open(svg_filename, "wb") as svg_file:
                        svg_file.write(response.content)

                    cairosvg.svg2png(url=svg_filename, write_to=png_filename)
                    img_tag["src"] = png_filename
                else:
                    img_extension = content_type.split("/")[-1]
                    img_filename = f"{img_filename}.{img_extension}"

                    with open(img_filename, "wb") as img_file:
                        img_file.write(response.content)

                    img_tag["src"] = img_filename

    for thead in soup.find_all("thead"):
        thead["style"] = "background-color: blue; color: white;"
    return str(soup)


def html_to_word(output_filename):
    """
    Convierte un archivo HTML a DOCX.

    Parámetros:
    output_filename (str): nombre del archivo de salida
    """
    html_file = "tablas.html"

    pypandoc.convert_file(
        html_file,
        format="html",
        to="docx",
        outputfile=f"table_results/{output_filename}",
        extra_args=["-RTS"],

    )

    print(
        "Conversión completada. El archivo DOCX ha sido generado con las imágenes en PNG o su formato original."
    )


def main():
    """
    Función principal
    """

    # Se instancia la conexión a Canvas
    canvas = get_connection()

    i = 0  # Contador para el número de cursos

    # Se abre el archivo donde se encuentra el listado de cursos
    with open("courses.txt", encoding="utf8") as f:
        for line in f:
            i += 1
            url = line.strip()
            courseId = get_number(url)
            course = canvas.get_course(courseId)

            print("\n%s) %s\n" % (i, course.name))

            html_course = ""

            # Listado de páginas a partir de "Módulos"
            list_pages = get_url_pages(course=course)

            for p in list_pages:
                page = course.get_page(p)  # Se instancia a la página
                html = delete_tags(page.body)  # Se eliminan etiquetas basura
                html = identify_class(
                    course, html
                )  # Se suma el html de una página en particular a la página de semana
                html_course += html  # se almacena el html de todas las páginas

            html_tables = get_tables(
                course, html_course
            )  # Se crea un html con: titulo de la tabla, tabla y nota
            html_dec = decoding_html(
                html_tables
            )  # se decodifica el html en las notas <pre>

            html_parsed = process_tables(html_dec)
            write_file("tablas.html", html_parsed)
            html_to_word(f"{course.sis_course_id}_tablas.docx")


if __name__ == "__main__":
    main()
