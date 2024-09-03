# Importamos los módulos necesarios para el manejo de archivos, hojas de cálculo, PDF y expresiones regulares
from os import makedirs  # Para crear directorios
from os.path import abspath, dirname, join  # Para manejar rutas de archivos y directorios
import xlwings as xw  # Para manipular archivos de Excel
import pandas as pd  # Para trabajar con DataFrames, que son estructuras de datos tabulares
import pdfplumber  # Para extraer texto e imágenes de archivos PDF
import re  # Para trabajar con expresiones regulares, útiles para buscar patrones de texto

# Configurar las rutas de los archivos
python_ubicacion = abspath(dirname(__file__))  # Obtiene la ruta absoluta del directorio del script actual
ruta_archivo = join(python_ubicacion, "maquillaje.pdf")  # Ruta del archivo PDF que se procesará
ruta_carpeta_imagen = join(python_ubicacion, "imagen")  # Ruta del directorio donde se guardarán las imágenes extraídas
ruta_respuesta = join(python_ubicacion, "respuesta.xlsx")  # Ruta del archivo Excel de salida

# Crear carpeta para guardar imágenes si no existe
makedirs(ruta_carpeta_imagen, exist_ok=True)  # Crea el directorio para imágenes si aún no existe

def extraer_datos(page):
    """Extrae datos de las tablas en una página PDF"""
    records = []  # Lista para almacenar los datos extraídos de la página
    # Expresión regular para capturar los productos y sus precios (anterior y actual)
    pattern = re.compile(r'(.*?)\s+(?:S/\.\s*([\d,]+\.\d+)?\s+)?S/\.\s*([\d,]+\.\d+|0\.00)', re.DOTALL)
    
    text_lines = page.extract_text_lines()[1: -1]  # Extrae líneas de texto, excluyendo la primera y última para evitar encabezados/pies de página
    dato = " ".join([i["text"] for i in text_lines])  # Une las líneas de texto en una sola cadena
    matches = pattern.findall(dato)  # Busca todas las coincidencias del patrón en el texto

    for match in matches:  # Itera sobre cada coincidencia encontrada
        producto = match[0].replace('\n', ' ').strip()  # Limpia el texto del producto
        precio_anterior = f'{match[1].strip()}'  # Obtiene el precio anterior, si está disponible
        precio_actual = f'{match[2].strip()}'  # Obtiene el precio actual
        # Añade un diccionario con los datos del producto a la lista de registros
        records.append({
            "Producto": producto,
            "Precio Anterior": precio_anterior,
            "Precio Actual": precio_actual,
        })    
    
    return records  # Retorna la lista de registros extraídos

def extraer_imagenes(page, page_index):
    """Extrae y guarda imágenes de una página PDF"""
    images = []  # Lista para almacenar las rutas de las imágenes extraídas
    for image_index, image in enumerate(page.images):  # Itera sobre cada imagen encontrada en la página
        x0, y0, x1, y1 = image["x0"], image["y0"], image["x1"], image["y1"]  # Coordenadas de la imagen dentro de la página
        cropped_image = page.within_bbox((x0, y0, x1, y1)).to_image()  # Recorta la imagen utilizando las coordenadas
        img = cropped_image.original.convert("RGB")  # Convierte la imagen recortada a formato RGB
        
        # Ruta donde se guardará la imagen extraída
        ruta_imagen = join(ruta_carpeta_imagen, f"image_{page_index}_{image_index}.png")
        img.save(ruta_imagen)  # Guarda la imagen en la ruta especificada
        images.append(ruta_imagen)  # Añade la ruta de la imagen a la lista
    return images  # Retorna la lista de rutas de imágenes extraídas

def guardar_excel(df, ruta_excel):
    """Guarda el DataFrame en un archivo Excel formateado"""
    
    df["Producto"] = df["Producto"].str.replace("(cid:1)(cid:2)", "")  # Limpia el texto de la columna "Producto"
    df["Precio Anterior"] = pd.to_numeric(df["Precio Anterior"], errors='coerce')  # Convierte los precios a números
    df["Precio Actual"] = pd.to_numeric(df["Precio Actual"], errors='coerce')  # Convierte los precios a números
    df["imagen"] = ''  # Añade una columna para las imágenes

    with pd.ExcelWriter(ruta_excel, engine='xlsxwriter') as writer:  # Crea un objeto para escribir en Excel
        df.to_excel(writer, index=False)  # Guarda el DataFrame en el archivo Excel

        # Acceso al workbook y worksheet
        worksheet = writer.sheets['Sheet1']
       
        # Ajusta el ancho de las columnas para mejor visualización
        worksheet.set_column('A:A', 60)  # Columna A para productos
        worksheet.set_column('B:B', 15)  # Columna B para precios anteriores
        worksheet.set_column('C:C', 15)  # Columna C para precios actuales
        worksheet.set_column('D:D', 50)  # Columna D para rutas de imágenes
        worksheet.set_column('E:E', 15)  # Columna E donde se insertarán las imágenes
        
        worksheet.freeze_panes(1, 0)  # Congela la primera fila para facilitar la navegación en Excel
        
        for i, _ in enumerate(df.index, start=1):  # Itera sobre las filas del DataFrame para ajustar la altura
            worksheet.set_row(i, 50)  # Ajusta la altura de cada fila

def insertar_imagen_excel(excel_path, start_cell='E2', image_width=60, image_height=40):
    """Inserta imágenes en un archivo Excel centradas en la celda."""
    wb = xw.Book(excel_path)  # Abre el archivo Excel utilizando xlwings
    sheet = wb.sheets.active  # Obtiene la hoja activa (la primera por defecto)
    image_values = sheet.range('D2').expand('down').value  # Obtiene las rutas de las imágenes desde la columna D
    
    for image_index, image_path in enumerate(image_values):  # Itera sobre las rutas de las imágenes
        current_cell = sheet.range(start_cell).offset(row_offset=image_index, column_offset=0)  # Calcula la celda destino para cada imagen
        
        # Calcular las posiciones centrales
        cell_width = current_cell.width  # Ancho de la celda
        cell_height = current_cell.height  # Altura de la celda
        left = current_cell.left + (cell_width - image_width) / 2  # Calcula la posición izquierda para centrar la imagen
        top = current_cell.top + (cell_height - image_height) / 2  # Calcula la posición superior para centrar la imagen
        
        # Insertar la imagen centrada en la celda
        picture = sheet.pictures.add(
            image_path,  # Ruta de la imagen
            left=left,  # Posición izquierda calculada
            top=top,  # Posición superior calculada
            width=image_width,  # Ancho de la imagen
            height=image_height,  # Altura de la imagen
            scale=True,  # Mantiene la escala de la imagen
        )
        picture.api.Placement = 1  # Fija la imagen a la celda
    
    sheet.range('D:D').api.EntireColumn.Hidden = True  # Oculta la columna D que contiene las rutas de las imágenes
        
    wb.save()  # Guarda los cambios en el archivo Excel
    wb.close()  # Cierra el archivo Excel

def procesar_pagina(page_index, page_content):
    """Procesa una sola página del PDF."""
    records = extraer_datos(page_content)  # Extrae datos de la página
    images = extraer_imagenes(page_content, page_index)  # Extrae imágenes de la página

    images.reverse()  # Invierte la lista de imágenes para alinearlas correctamente con los registros de productos
    for index, ruta_imagen in enumerate(images):  # Itera sobre las imágenes extraídas
        records[index]["ruta_imagen"] = ruta_imagen  # Asigna la ruta de la imagen correspondiente a cada registro
    
    print(f"Página {page_index + 1} procesada correctamente.")  # Mensaje de confirmación para el usuario
    return records  # Retorna los registros procesados de la página

def main():
    """Función principal para extraer datos e imágenes del PDF usando paralelización limitada a 4 cores."""
    records = []  # Lista para almacenar todos los registros extraídos de todas las páginas

    with pdfplumber.open(ruta_archivo) as pdf:  # Abre el archivo PDF
        # Leer todas las páginas antes de cerrar el archivo
        for index, page in enumerate(pdf.pages):  # Itera sobre cada página del PDF
            if index == 20:  # Limita el procesamiento a las primeras 20 páginas
                break
            
            result_records = procesar_pagina(index, page)  # Procesa cada página
            records.extend(result_records)  # Añade los registros de la página a la lista total

    df = pd.DataFrame(records)  # Crea un DataFrame de pandas a partir de la lista de registros
    
    guardar_excel(df, ruta_respuesta)  # Guarda el DataFrame en un archivo Excel
    insertar_imagen_excel(ruta_respuesta)  # Inserta imágenes en el archivo Excel
    
    print("Proceso de extracción de datos e imágenes completado.")  # Mensaje de confirmación para el usuario

if __name__ == "__main__":
    main()  # Ejecuta la función principal si el script se ejecuta directamente
