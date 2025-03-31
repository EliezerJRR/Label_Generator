import os
import pandas as pd
from reportlab.pdfgen import canvas
from PyPDF2 import PdfMerger
from reportlab.graphics.barcode import code128
from reportlab.lib.units import inch
from win32 import win32print

def limpiar_carpeta_pdf(carpeta_pdf):
    """
    Limpia la carpeta de PDFs antes de generar nuevos archivos.
    """
    if os.path.exists(carpeta_pdf):
        for archivo in os.listdir(carpeta_pdf):
            ruta_archivo = os.path.join(carpeta_pdf, archivo)
            if os.path.isfile(ruta_archivo):
                os.remove(ruta_archivo)
    else:
        os.makedirs(carpeta_pdf)
    print("Carpeta PDF limpiada y lista.")

def generar_hoja_personalizada(pdf_file, numero_pedido, conteo_actual, total_etiquetas, modalidad, contenido_fila_2, contenido_fila_3_nombre):
    """
    Generar un PDF con elementos centrados en las filas y columnas.
    """
    ancho_hoja = 6 * 72
    alto_hoja = 3 * 72
    c = canvas.Canvas(pdf_file, pagesize=(ancho_hoja, alto_hoja))

    # Fila 1: Pedido y contador
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(ancho_hoja * 0.4, alto_hoja * 0.88, str(numero_pedido))
    c.setFont("Helvetica-Bold", 25)
    c.drawCentredString(ancho_hoja * 0.9, alto_hoja * 0.88, f"({conteo_actual}/{total_etiquetas})")

    # Código de barras
    barcode = code128.Code128(str(numero_pedido), barHeight=alto_hoja / 3.8, barWidth=0.038 * inch)
    barcode.drawOn(c, ancho_hoja * 0.05, alto_hoja * 0.58)

    # Fila 2: SKU y modalidad
    c.setFont("Helvetica-Bold", 35)
    c.drawCentredString(ancho_hoja * 0.4, alto_hoja * 0.4, contenido_fila_2[0])
    c.drawCentredString(ancho_hoja * 0.9, alto_hoja * 0.4, modalidad)

    # Fila 3: Nombre del cliente
    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(ancho_hoja / 2, alto_hoja * 0.1, contenido_fila_3_nombre)

    c.save()
    print(f"Etiqueta generada: {pdf_file}")

def unir_pdfs(lista_archivos, archivo_salida):
    """
    Une múltiples archivos PDF en un solo archivo.
    """
    if not lista_archivos:
        print("No hay archivos para unir.")
        return

    try:
        merger = PdfMerger()
        for pdf in lista_archivos:
            merger.append(pdf)
        merger.write(archivo_salida)
        merger.close()
        print(f"Archivo PDF unificado generado: {archivo_salida}")
    except Exception as e:
        print(f"Error al generar el archivo PDF unificado: {e}")

def procesar_pedidos(datos, carpeta_pdf):
    """
    Procesa los pedidos del DataFrame y genera PDFs individuales.
    """
    archivos_pdf = []
    pedidos_unicos = datos['ORDEN'].unique()

    for pedido in pedidos_unicos:
        productos = datos[datos['ORDEN'] == pedido]
        total_etiquetas = productos['CANTIDAD'].sum()
        etiqueta_actual = 1

        for _, fila in productos.iterrows():
            for i in range(fila['CANTIDAD']):
                archivo_pdf = os.path.join(carpeta_pdf, f"{pedido}_{etiqueta_actual}.pdf")
                archivos_pdf.append(archivo_pdf)

                generar_hoja_personalizada(
                    pdf_file=archivo_pdf,
                    numero_pedido=pedido,
                    conteo_actual=etiqueta_actual,
                    total_etiquetas=total_etiquetas,
                    modalidad=fila['MODALIDAD'],
                    contenido_fila_2=[fila['SKU'], fila['MODALIDAD']],
                    contenido_fila_3_nombre=fila['NOMBRE CLIENTE']
                )
                etiqueta_actual += 1

    return archivos_pdf

def label_generator():
    """
    Genera etiquetas en PDF y unifica en un archivo final.
    """
    ruta_excel = os.path.join(os.path.dirname(__file__), "data", "Tabla.xlsx")
    carpeta_pdf = os.path.join(os.path.dirname(__file__), "pdf")
    carpeta_output = os.path.join(os.path.dirname(__file__), "output")

    limpiar_carpeta_pdf(carpeta_pdf)
    if not os.path.exists(carpeta_output):
        os.makedirs(carpeta_output)

    try:
        datos = pd.read_excel(ruta_excel)
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo Excel en {ruta_excel}.")
        return
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return

    columnas_requeridas = {'ORDEN', 'CANTIDAD', 'SKU', 'MODALIDAD', 'NOMBRE CLIENTE'}
    if not columnas_requeridas.issubset(datos.columns):
        print(f"Error: El archivo Excel no contiene las columnas necesarias: {columnas_requeridas}")
        return

    try:
        archivos_pdf = procesar_pedidos(datos, carpeta_pdf)
        if archivos_pdf:
            archivo_pdf_unido = os.path.join(carpeta_output, "orden_completa.pdf")
            unir_pdfs(archivos_pdf, archivo_pdf_unido)
            print(f"Proceso completado. Archivo generado en: {archivo_pdf_unido}")
        else:
            print("No se generaron archivos PDF.")
    except Exception as e:
        print(f"Error durante la generación de PDFs: {e}")

def imprimir_en_modo_raw(nombre_impresora, ruta_archivo):
    """
    Envía un archivo a la impresora en formato RAW.

    Parámetros:
        nombre_impresora (str): Nombre de la impresora donde se enviará el archivo.
        ruta_archivo (str): Ruta completa del archivo que se desea imprimir.
    """
    try:
        # Abrir la impresora especificada
        hprinter = win32print.OpenPrinter(nombre_impresora)

        # Configurar el trabajo de impresión
        hprinter_job = win32print.StartDocPrinter(hprinter, 1, ("Trabajo Python", None, "RAW"))
        win32print.StartPagePrinter(hprinter)

        # Leer los datos del archivo y enviarlos a la impresora
        with open(ruta_archivo, "rb") as archivo:
            datos = archivo.read()
            win32print.WritePrinter(hprinter, datos)

        # Finalizar el trabajo de impresión
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
        win32print.ClosePrinter(hprinter)

        print(f"El archivo '{ruta_archivo}' se imprimió correctamente en '{nombre_impresora}' en modo RAW.")
    except Exception as e:
        print(f"Ocurrió un error al intentar imprimir en RAW: {e}")

#nombre_impresora = win32print.GetDefaultPrinter()  # Obtiene el nombre de la impresora predeterminada
#ruta_pdf = r"output/orden_completa.pdf"
#(nombre_impresora, ruta_pdf)

if __name__ == "__main__":
    label_generator()