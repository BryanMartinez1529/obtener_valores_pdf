# Description: Script para extraer datos de archivos PDF y escribirlos en una plantilla de Excel
import pdfplumber
import xlwings as xw
import os
import logging
import re

indices_a_buscar_103 = [
    "302","303", "3030", "304", "304B", "307", "308", "309", "310", "311", "312", "312A", "3121",
    "314", "319", "320", "322", "323", "324", "325", "326", "327", "328", "332", "332G",
    "336", "337", "343", "344", "3440", "345", "346", "421"
]  
indices_a_buscar_104 = [
    "500", "501", "502", "503", "505", "506", "507", "508", "510", "511", "512", "513", "515", "516", "517", "518",
    "531", "532", "535", "540", "550", "721", "723", "725", "729","731"
]


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

# Función para extraer códigos de retención y sus valores
def extraer_codigos_retencion(pdf_path):
    codigos_retencion = {}

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if "RESUMEN DE RETENCIONES - AGENTE DE RETENCION" in texto:
                seccion = texto.split("RESUMEN DE RETENCIONES - AGENTE DE RETENCION")[1]
                lineas = seccion.strip().split("\n")

                for linea in lineas:
                    match = re.match(r"^(\d{3}[A-Z]?)\s+.*?(\d+(?:,\d{3})*(?:\.\d+)?)\s+(\d+(?:,\d{3})*(?:\.\d+)?)$", linea.strip())
                    if match:
                        codigo = match.group(1)
                        base = float(match.group(2).replace(",", ""))
                        codigos_retencion[codigo] = base

    return codigos_retencion

# Función para extraer totales de compras
def extraer_totales_compras(pdf_path):
    totales = {}

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if "COMPRAS" in texto:
                seccion = texto.split("COMPRAS")[1]
                lineas = seccion.strip().split("\n")

                for linea in lineas:
                    if "TOTAL:" in linea:
                        numeros = re.findall(r"(\d{1,3}(?:,\d{3})*(?:\.\d+)|\d+\.\d+)", linea)
                        if len(numeros) >= 4:
                            totales = {
                                "BI tarifa 0%": float(numeros[0].replace(",", "")),
                                "BI tarifa diferente 0%": float(numeros[1].replace(",", "")),
                                "BI No Objeto IVA": float(numeros[2].replace(",", ""))
                            }
                        break
    return totales

# Función para convertir el número de mes a la letra de columna correspondiente
def mes_a_columna(mes):
    columnas = ["C", "F", "I", "L", "O", "R", "U", "X", "AA", "AD", "AG", "AJ"]
    return columnas[mes - 1]
# Función para escribir datos en una hoja específica de la plantilla Excel (para filas por mes y columnas por índice)
# Función para escribir datos en una hoja específica de la plantilla Excel
def escribir_en_hoja(datos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja):
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
        
        # Verificar si la hoja existe, si no, crearla
        if nombre_hoja in [sheet.name for sheet in wb.sheets]:
            hoja = wb.sheets[nombre_hoja]
        else:
            hoja = wb.sheets.add(name=nombre_hoja)

        for indice, valor in datos.items():
            if indice in ubicaciones:
                celda_base = ubicaciones[indice][1]
                fila_base = int(celda_base[1:])  # Extraer el número de fila base (ej. 10)
                columna_mes = mes_a_columna(mes)  # Obtener la columna correspondiente al mes

                celda_destino = f"{columna_mes}{fila_base}"  # Construir la celda destino
                hoja.range(celda_destino).value = valor  # Insertar el valor

        wb.save(ruta_salida)
        wb.close()
        app.quit()
        logging.info(f"Datos del mes {mes} guardados en la hoja '{nombre_hoja}' en '{ruta_salida}' correctamente.")
    except Exception as e:
        logging.error(f"Error al escribir en la hoja '{nombre_hoja}' de la plantilla Excel: {e}")

# Procesar múltiples conjuntos de datos
def procesar_datos_por_hoja(ruta_pdf_base, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja, indices_a_buscar):
    for mes in range(1, 13):
        ruta_pdf = os.path.join(ruta_pdf_base, f"{mes}.pdf")

        if os.path.exists(ruta_pdf):
            valores_extraidos = extraer_valores_indices(ruta_pdf, indices_a_buscar)
            escribir_en_hoja(valores_extraidos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja)
        else:
            logging.warning(f"No se encontró el archivo: {ruta_pdf}")

def escribir_en_hoja_por_filas(datos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja):
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
        
        # Verificar si la hoja existe, si no, crearla
        if nombre_hoja in [sheet.name for sheet in wb.sheets]:
            hoja = wb.sheets[nombre_hoja]
        else:
            hoja = wb.sheets.add(name=nombre_hoja)

        for indice, valor in datos.items():
            if indice in ubicaciones:
                fila_mes = 11 + (mes - 1)  # Ajustar la fila según el mes (Enero = Fila 11, Febrero = Fila 12, etc.)
                columna_base = ubicaciones[indice][1][0]  # Extraer la letra de la columna base (ej. "F")

                celda_destino = f"{columna_base}{fila_mes}"  # Construir la celda destino
                hoja.range(celda_destino).value = valor  # Insertar el valor

        wb.save(ruta_salida)
        wb.close()
        app.quit()
        logging.info(f"Datos del mes {mes} guardados en la hoja '{nombre_hoja}' en '{ruta_salida}' correctamente.")
    except Exception as e:
        logging.error(f"Error al escribir en la hoja '{nombre_hoja}' de la plantilla Excel: {e}")


# Configuración de archivos y datos

def escribir_en_hoja_por_ubicaciones(datos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja):
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
        
        # Verificar si la hoja existe, si no, crearla
        if nombre_hoja in [sheet.name for sheet in wb.sheets]:
            hoja = wb.sheets[nombre_hoja]
        else:
            hoja = wb.sheets.add(name=nombre_hoja)

        for indice, valor in datos.items():
            if indice in ubicaciones:
                # Obtener la columna base y ajustar la fila según el mes
                columna_base = ubicaciones[indice][1][0]  # Extraer la letra de la columna (ej. "I")
                fila_base = int(ubicaciones[indice][1][1:])  # Extraer la fila base (ej. 40)
                fila_mes = fila_base + (mes - 1)  # Ajustar la fila según el mes (Enero = Fila base, Febrero = Fila base + 1, etc.)

                celda_destino = f"{columna_base}{fila_mes}"  # Construir la celda destino
                hoja.range(celda_destino).value = valor  # Insertar el valor

        wb.save(ruta_salida)
        wb.close()
        app.quit()
        logging.info(f"Datos del mes {mes} guardados en la hoja '{nombre_hoja}' en '{ruta_salida}' correctamente.")
    except Exception as e:
        logging.error(f"Error al escribir en la hoja '{nombre_hoja}' de la plantilla Excel: {e}")


# Procesar datos para la hoja "A4"
def procesar_datos_por_ubicaciones(ruta_pdf_base, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja, indices_a_buscar):
    for mes in range(1, 13):
        ruta_pdf = os.path.join(ruta_pdf_base, f"{mes}.pdf")

        if os.path.exists(ruta_pdf):
            valores_extraidos = extraer_valores_indices(ruta_pdf, indices_a_buscar)
            escribir_en_hoja_por_ubicaciones(valores_extraidos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja)
        else:
            logging.warning(f"No se encontró el archivo: {ruta_pdf}")


# Procesar el segundo conjunto de datos y escribir en la hoja "103 VS 104"
def procesar_datos_por_filas(ruta_pdf_base, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja, indices_a_buscar):
    for mes in range(1, 13):
        ruta_pdf = os.path.join(ruta_pdf_base, f"{mes}.pdf")

        if os.path.exists(ruta_pdf):
            valores_extraidos = extraer_valores_indices(ruta_pdf, indices_a_buscar)
            escribir_en_hoja_por_filas(valores_extraidos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja)
        else:
            logging.warning(f"No se encontró el archivo: {ruta_pdf}")

# Procesar datos de tablas
def procesar_datos_tablas(ruta_pdf_base, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja, extractor_func):
    for mes in range(1, 13):
        ruta_pdf = os.path.join(ruta_pdf_base, f"{mes}.pdf")

        if os.path.exists(ruta_pdf):
            datos_extraidos = extractor_func(ruta_pdf)
            escribir_en_hoja(datos_extraidos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja)
        else:
            logging.warning(f"No se encontró el archivo: {ruta_pdf}")


ruta_plantilla = "./pdf/plantilla.xlsx"
ruta_excel_salida = "datos_anuales.xlsx"

# Ubicaciones para el primer conjunto de datos
ubicaciones_celdas_hoja1 = {
    "302": ("103 VS ATS", "C44"),
    "303": ("103 VS ATS", "C10"),
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

ubicaciones_celdas_hoja_ats = {
    "302": ("103 VS ATS", "B44"),
    "303": ("103 VS ATS", "B10"),
    "303A": ("103 VS ATS", "B11"),
    "304": ("103 VS ATS", "B12"),
    "304B": ("103 VS ATS", "B13"),
    "307": ("103 VS ATS", "B14"),
    "308": ("103 VS ATS", "B15"),
    "309": ("103 VS ATS", "B16"),
    "310": ("103 VS ATS", "B17"),
    "311": ("103 VS ATS", "B18"),
    "312": ("103 VS ATS", "B19"),
    "312A": ("103 VS ATS", "B20"),
    "3121": ("103 VS ATS", "B21"),
    "314": ("103 VS ATS", "B22"),
    "319": ("103 VS ATS", "B23"),
    "320": ("103 VS ATS", "B24"),
    "322": ("103 VS ATS", "B25"),
    "323": ("103 VS ATS", "B26"),
    "324": ("103 VS ATS", "B27"),
    "325": ("103 VS ATS", "B28"),
    "326": ("103 VS ATS", "B29"),
    "327": ("103 VS ATS", "B30"),
    "328": ("103 VS ATS", "B31"),
    "332": ("103 VS ATS", "B32"),
    "332G": ("103 VS ATS", "B33"),
    "336": ("103 VS ATS", "B34"),
    "337": ("103 VS ATS", "B35"),
    "343": ("103 VS ATS", "B36"),
    "344": ("103 VS ATS", "B37"),
    "3440": ("103 VS ATS", "B38"),
    "345": ("103 VS ATS", "B39"),
    "346": ("103 VS ATS", "B40"),
    "421": ("103 VS ATS", "B41"),
    "501": ("103 VS 104", "B42"),
}

# Ubicaciones para el segundo conjunto de datos
ubicaciones_celdas_hoja2 = {
    "500": ("103 VS 104", "F11"),
    "501": ("103 VS 104", "G11"),
    "502": ("103 VS 104", "H11"),
    "503": ("103 VS 104", "I11"),
    "540": ("103 VS 104", "J11"),
    "505": ("103 VS 104", "K11"),
    "506": ("103 VS 104", "L11"),
    "507": ("103 VS 104", "M11"),
    "508": ("103 VS 104", "N11"),
    "531": ("103 VS 104", "O11"),
    "532": ("103 VS 104", "P11"),
    "535": ("103 VS 104", "Q11"),
}

# Ubicaciones para el tercer conjunto de datos
ubicaciones_celdas_hoja3 = {
    "510": ("104 VS ATS", "F11"),
    "511": ("104 VS ATS", "G11"),
    "512": ("104 VS ATS", "H11"),
    "513": ("104 VS ATS", "I11"),
    "550": ("104 VS ATS", "J11"),
    "515": ("104 VS ATS", "K11"),
    "516": ("104 VS ATS", "L11"),
    "517": ("104 VS ATS", "M11"),
    "518": ("104 VS ATS", "N11")
}


ubicaciones_celdas_hoja4 = {
    "721": ("A4", "I40"),
    "723": ("A4", "J40"),
    "725": ("A4", "K40"),
    "727": ("A4", "L40"),
    "729": ("A4", "M40"),
    "731": ("A4", "N40")
}




# procesar_datos_tablas(
#     "./pdf/ats",
#     ruta_plantilla,
#     ruta_excel_salida,
#     ubicaciones_celdas_hoja_ats,
#     "103 VS ATS",
#     extraer_codigos_retencion
# )

procesar_datos_por_hoja(
    "./impuestos/103",
    ruta_plantilla,
    ruta_excel_salida,
    ubicaciones_celdas_hoja1,
    "103 VS ATS",
    indices_a_buscar_103
)

procesar_datos_por_filas(
    "./impuestos/104",
    ruta_plantilla,
    ruta_excel_salida,
    ubicaciones_celdas_hoja2,
    "103 VS 104",
    indices_a_buscar_104
)

procesar_datos_por_filas(
    "./impuestos/104",
    ruta_plantilla,
    ruta_excel_salida,
    ubicaciones_celdas_hoja3,
    "104 VS ATS",
    indices_a_buscar_104
)


procesar_datos_por_ubicaciones(
    "./impuestos/104",
    ruta_plantilla,
    ruta_excel_salida,
    ubicaciones_celdas_hoja4,
    "A4",
    indices_a_buscar_104
)