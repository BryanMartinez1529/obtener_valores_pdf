import pdfplumber
import xlwings as xw
import os
import logging

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Función para extraer valores según los índices de cada PDF
def extraer_valores_indices(ruta_pdf, indices_buscados):
    valores_encontrados = {indice: 0 for indice in indices_buscados}  # Inicializar con 0

    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            for pagina in pdf.pages:
                palabras = pagina.extract_words()

                for i in range(len(palabras)):
                    palabra_actual = palabras[i]['text']

                    if palabra_actual in valores_encontrados:
                        if i + 1 < len(palabras):
                            valor_extraido = palabras[i + 1]['text'].replace('.', ',')  # Reemplazar punto por coma
                            
                            # Convertir a número si es posible
                            try:
                                valor_extraido = float(valor_extraido.replace(',', '.'))  # Convertir a número
                            except ValueError:
                                pass  

                            valores_encontrados[palabra_actual] = valor_extraido
    except Exception as e:
        logging.error(f"Error al procesar el archivo PDF {ruta_pdf}: {e}")

    return valores_encontrados

# Función para convertir el número de mes a la letra de columna correspondiente
def mes_a_columna(mes):
    columnas = ["C", "F", "I", "L", "O", "R", "U", "X", "AA", "AD", "AG", "AJ"]
    return columnas[mes - 1]

# Función para escribir datos en la plantilla Excel
# Función para escribir datos en la plantilla Excel
def escribir_en_plantilla(datos, mes, ruta_plantilla, ruta_salida, ubicaciones):
    try:
        # Verificar si el archivo de salida ya existe
        if os.path.exists(ruta_salida):
            # Abrir el archivo de salida existente
            app = xw.App(visible=False)
            wb = xw.Book(ruta_salida)
        else:
            # Si no existe, usar la plantilla como base
            app = xw.App(visible=False)
            wb = xw.Book(ruta_plantilla)
        
        for indice, valor in datos.items():
            if indice in ubicaciones:
                hoja_nombre, celda_base = ubicaciones[indice]
                fila_base = int(celda_base[1:])  # Extraer el número de fila base (ej. 10)
                columna_mes = mes_a_columna(mes)  # Obtener la columna correspondiente al mes

                if hoja_nombre in [sheet.name for sheet in wb.sheets]:
                    hoja = wb.sheets[hoja_nombre]
                    celda_destino = f"{columna_mes}{fila_base}"  # Construir la celda destino
                    hoja.range(celda_destino).value = valor  # Insertar el valor
                else:
                    logging.warning(f"La hoja '{hoja_nombre}' no existe en la plantilla.")

        wb.save(ruta_salida)
        wb.close()
        app.quit()
        logging.info(f"Datos del mes {mes} guardados en '{ruta_salida}' correctamente.")
    except Exception as e:
        logging.error(f"Error al escribir en la plantilla Excel: {e}")

# Configuración de archivos y datos
ruta_plantilla = "./pdf/plantilla.xlsx"
ruta_excel_salida = "datos_anuales.xlsx"

indices_a_buscar = [
    "303", "3030", "304", "304B", "307", "308", "309", "310", "311", "312", "312A", "3121",
    "314", "319", "320", "322", "323", "324", "325", "326", "327", "328", "332", "332G",
    "336", "337", "343", "344", "3440", "345", "346", "421"
]  # Lista de índices a buscar

# Ubicaciones de los valores en la plantilla Excel
ubicaciones_celdas = {
    "303": ("103 VS ATS", "C10"),  # Columna B, fila 2 (Enero), fila 3 (Febrero), etc.
    "3030": ("103 VS ATS", "C11"),
    "304": ("103 VS ATS", "C12"),
    "304B": ("103 VS ATS", "C13"),
    "307": ("103 VS ATS", "C14"),
    "308": ("103 VS ATS", "C15"),
    "309": ("103 VS ATS", "C16"),
    "310": ("103 VS ATS", "C17"),
    "311": ("103 VS ATS", "C18"),
    "312": ("103 VS ATS", "C19"),
    "312A": ("103 VS ATS", "C20"),
    "3121": ("103 VS ATS", "C21"),
    "314": ("103 VS ATS", "C22"),
    "319": ("103 VS ATS", "C23"),
    "320": ("103 VS ATS", "C24"),
    "322": ("103 VS ATS", "C25"),
    "323": ("103 VS ATS", "C26"),
    "324": ("103 VS ATS", "C27"),
    "325": ("103 VS ATS", "C28"),
    "326": ("103 VS ATS", "C29"),
    "327": ("103 VS ATS", "C30"),
    "328": ("103 VS ATS", "C31"),
    "332": ("103 VS ATS", "C32"),
    "332G": ("103 VS ATS", "C33"),
    "336": ("103 VS ATS", "C34"),
    "337": ("103 VS ATS", "C35"),
    "343": ("103 VS ATS", "C36"),
    "344": ("103 VS ATS", "C37"),
    "3440": ("103 VS ATS", "C38"),
    "345": ("103 VS ATS", "C39"),
    "346": ("103 VS ATS", "C40"),
    "421": ("103 VS ATS", "C41")
}

# Procesar los 12 archivos PDF (uno por mes)
for mes in range(1, 13):
    ruta_pdf = os.path.join(".", "pdf", f"{mes}.pdf")

    if os.path.exists(ruta_pdf):
        valores_extraidos = extraer_valores_indices(ruta_pdf, indices_a_buscar)
        escribir_en_plantilla(valores_extraidos, mes, ruta_plantilla, ruta_excel_salida, ubicaciones_celdas)
    else:
        logging.warning(f"No se encontró el archivo: {ruta_pdf}")