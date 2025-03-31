import pandas as pd
import tkinter as tk
from tkinter import ttk  # Para la tabla (Treeview)
import os
import re  # Para buscar patrones en nombres de archivos
from GeneradorPDF import label_generator, limpiar_carpeta_pdf
from win32 import win32print

# Función para generar datos en la tabla y ejecutar GeneradorPDF.py
def generar_datos():
    # Inhabilitar el botón de imprimir al iniciar la generación
    btn_imprimir.config(state=tk.DISABLED)
    progreso["value"] = 0  # Reiniciar barra de progreso

    # Ruta del archivo Excel para contar los registros
    ruta_excel = os.path.join(os.path.dirname(__file__), "Tabla.xlsx")
    datos = pd.read_excel(ruta_excel)
    total_registros = len(datos)

    # Configurar el máximo de la barra de progreso para representar el 100%
    progreso["maximum"] = total_registros + (0.25 * total_registros)

    # Actualizar el progreso inicial
    estado_label.config(text=f"Progreso: 0/{total_registros} registros procesados.")

    # Ejecutar el archivo GeneradorPDF.py
    #subprocess.run(["python", ruta_etiquetas], check=True)
    label_generator()

    # Ruta de la carpeta donde están los archivos PDF generados
    carpeta_pdf = os.path.join(os.path.dirname(__file__), "pdf")

    # Procesar los registros del Excel (80% de la barra)
    for index, fila in datos.iterrows():
        # Actualizar barra de progreso y etiqueta
        progreso["value"] = (index + 1)
        estado_label.config(text=f"Progreso: {index + 1}/{total_registros} registros procesados.")
        root.update_idletasks()  # Actualizar la interfaz en tiempo real

    # Verificar que se ha generado el archivo unificado
    archivo_pdf_unificado = os.path.join(os.path.dirname(__file__), "output", "orden_completa.pdf")
    if not os.path.exists(archivo_pdf_unificado):
        print("Error: No se encontró el archivo PDF unificado.")
        estado_label.config(text="Error: No se generó el archivo PDF unificado.")
        return

    # Completar el 20% restante de la barra
    progreso["value"] = progreso["maximum"]
    estado_label.config(text="Generando archivo PDF unificado...")

    #print(f"Archivo PDF unificado generado: {archivo_pdf_unificado}")

    # Obtener todos los archivos PDF en la carpeta
    archivos = [archivo for archivo in os.listdir(carpeta_pdf) if archivo.endswith(".pdf")]

    # Agrupar los archivos por el número antes del "_"
    grupos = {}
    for archivo in archivos:
        match = re.match(r"(\d+)_\d+\.pdf", archivo)  # Buscar el número antes del "_"
        if match:
            numero_grupo = match.group(1)
            if numero_grupo not in grupos:
                grupos[numero_grupo] = []
            grupos[numero_grupo].append(archivo)

    # Obtener el archivo con el número más alto para cada grupo
    datos_tabla = []
    for numero_grupo, archivos in grupos.items():
        numeros = [int(re.search(r"_(\d+)\.pdf", archivo).group(1)) for archivo in archivos]
        max_numero = max(numeros)  # Obtener el número más alto
        datos_tabla.append((numero_grupo, max_numero))  # Guardar resultado como (Número de orden, Nro de etiquetas)

    # Mostrar los datos en la tabla con colores alternativos
    tabla.delete(*tabla.get_children())  # Limpiar tabla antes de insertar nuevos datos
    for i, (numero_orden, nro_etiquetas) in enumerate(datos_tabla, start=1):
        tag = "even" if i % 2 == 0 else "odd"  # Alternar entre "even" y "odd"
        tabla.insert("", "end", values=(i, numero_orden, nro_etiquetas), tags=(tag,))

    print("Datos de la tabla generados correctamente.")
    btn_imprimir.config(state=tk.NORMAL)
    estado_label.config(text="Generación completada con éxito.")

def Limpiar():
    ruta_pdf = r"output/"
    limpiar_carpeta_pdf(ruta_pdf)
    carpeta_pdf = os.path.join(os.path.dirname(__file__), "pdf")
    limpiar_carpeta_pdf(carpeta_pdf)
    tabla.delete(*tabla.get_children())
    btn_imprimir.config(state=tk.DISABLED)
    estado_label.config(text="Generar nuevos pdf")
    #pass
def imprimir_pdf():
    # Habilitar el botón de imprimir cuando se cumplan ambos criterios
    nombre_impresora = win32print.GetDefaultPrinter()  # Obtiene el nombre de la impresora predeterminada
    ruta_pdf = r"output/orden_completa.pdf"
    imprimir_en_modo_raw(nombre_impresora, ruta_pdf)

# Crear ventana principal
root = tk.Tk()
root.title("Gestor de PDFs")
root.geometry("400x550")  # Ajustar altura para la barra de progreso
root.resizable(False, False)  # Evitar redimensionar la ventana

# Configurar una sola columna en la ventana principal
root.columnconfigure(0, weight=1)

# Crear el LabelFrame "Buscador PDF"
lf_buscador = tk.LabelFrame(root, text="Buscador PDF", padx=10, pady=10)
lf_buscador.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

# Configurar proporción para los elementos dentro del LabelFrame
lf_buscador.columnconfigure(0, weight=1)

# Crear el botón "Generar" centrado
btn_generar = tk.Button(lf_buscador, text="Generar", command=generar_datos)
btn_generar.grid(row=0, column=0, padx=0, pady=0,sticky="nsew")
btn_generar = tk.Button(lf_buscador, text="Limpiar", command=Limpiar)
btn_generar.grid(row=0, column=1, padx=0, pady=0, sticky="nsew")

# Crear la sección "Archivos"
lf_archivos = tk.LabelFrame(root, text="Archivos", padx=10, pady=10)
lf_archivos.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

# Hacer que la tabla ocupe el espacio completo
lf_archivos.columnconfigure(0, weight=1)
lf_archivos.rowconfigure(0, weight=1)

# Crear tabla con Treeview
columns = ("ID", "orden", "etiquetas")
tabla = ttk.Treeview(lf_archivos, columns=columns, show="headings")
tabla.grid(row=0, column=0, sticky="nsew")

# Configurar las cabeceras de la tabla
tabla.heading("ID", text="ID")
tabla.heading("orden", text="Números de Orden")
tabla.heading("etiquetas", text="Nro de Etiquetas")

# Configurar ancho de las columnas
tabla.column("ID", width=20, anchor="center")
tabla.column("orden", width=150, anchor="center")
tabla.column("etiquetas", width=150, anchor="center")

# Agregar barra de desplazamiento a la tabla
scrollbar = ttk.Scrollbar(lf_archivos, orient="vertical", command=tabla.yview)
tabla.configure(yscrollcommand=scrollbar.set)
scrollbar.grid(row=0, column=1, sticky="ns")

# Configurar estilos de las filas para alternar colores
tabla.tag_configure("odd", background="white")  # Color para filas impares
tabla.tag_configure("even", background="#f0f0f0")  # Color para filas pares

# Crear barra de progreso
progreso = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progreso.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

# Crear etiqueta de estado
estado_label = tk.Label(root, text="Esperando acción...", anchor="w", fg="blue")
estado_label.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

# Crear botón debajo de la tabla para "Imprimir", inicialmente inhabilitado
btn_imprimir = tk.Button(root, text="Imprimir", command=imprimir_pdf, state=tk.DISABLED)
btn_imprimir.grid(row=5, column=0, padx=10, pady=10, sticky="ew")

# Ejecutar la ventana principal
root.mainloop()
