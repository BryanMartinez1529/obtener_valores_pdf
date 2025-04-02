# Description: Script para extraer datos de archivos PDF y escribirlos en una plantilla de Excel
import pdfplumber
import xlwings as xw
import os
import logging
import re

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
                        valor = float(match.group(3).replace(",", ""))
                        codigos_retencion[codigo] = {"base": base, "valor": valor}

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

# Función para escribir datos en una hoja específica de la plantilla Excel
def escribir_en_hoja(datos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja):
    try:
        # Verificar si el archivo de salida ya existe
        if os.path.exists(ruta_salida):
            app = xw.App(visible=False)
            wb = xw.Book(ruta_salida)
        else:
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

# Función para escribir datos en filas por mes y columnas por índice
def escribir_en_hoja_por_filas(datos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja):
    try:
        if os.path.exists(ruta_salida):
            app = xw.App(visible=False)
            wb = xw.Book(ruta_salida)
        else:
            app = xw.App(visible=False)
            wb = xw.Book(ruta_plantilla)

        if nombre_hoja in [sheet.name for sheet in wb.sheets]:
            hoja = wb.sheets[nombre_hoja]
        else:
            hoja = wb.sheets.add(name=nombre_hoja)

        for indice, valor in datos.items():
            if indice in ubicaciones:
                fila_mes = 11 + (mes - 1)  # Ajustar la fila según el mes
                columna_base = ubicaciones[indice][1][0]  # Extraer la letra de la columna base

                celda_destino = f"{columna_base}{fila_mes}"  # Construir la celda destino
                hoja.range(celda_destino).value = valor  # Insertar el valor

        wb.save(ruta_salida)
        wb.close()
        app.quit()
        logging.info(f"Datos del mes {mes} guardados en la hoja '{nombre_hoja}' en '{ruta_salida}' correctamente.")
    except Exception as e:
        logging.error(f"Error al escribir en la hoja '{nombre_hoja}' de la plantilla Excel: {e}")

# Función para escribir datos en ubicaciones específicas
def escribir_en_hoja_por_ubicaciones(datos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja):
    try:
        if os.path.exists(ruta_salida):
            app = xw.App(visible=False)
            wb = xw.Book(ruta_salida)
        else:
            app = xw.App(visible=False)
            wb = xw.Book(ruta_plantilla)

        if nombre_hoja in [sheet.name for sheet in wb.sheets]:
            hoja = wb.sheets[nombre_hoja]
        else:
            hoja = wb.sheets.add(name=nombre_hoja)

        for indice, valor in datos.items():
            if indice in ubicaciones:
                columna_base = ubicaciones[indice][1][0]
                fila_base = int(ubicaciones[indice][1][1:])
                fila_mes = fila_base + (mes - 1)

                celda_destino = f"{columna_base}{fila_mes}"
                hoja.range(celda_destino).value = valor

        wb.save(ruta_salida)
        wb.close()
        app.quit()
        logging.info(f"Datos del mes {mes} guardados en la hoja '{nombre_hoja}' en '{ruta_salida}' correctamente.")
    except Exception as e:
        logging.error(f"Error al escribir en la hoja '{nombre_hoja}' de la plantilla Excel: {e}")

# Procesar datos por hoja
def procesar_datos_por_hoja(ruta_pdf_base, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja, indices_a_buscar):
    for mes in range(1, 13):
        ruta_pdf = os.path.join(ruta_pdf_base, f"{mes}.pdf")

        if os.path.exists(ruta_pdf):
            valores_extraidos = extraer_valores_indices(ruta_pdf, indices_a_buscar)
            escribir_en_hoja(valores_extraidos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja)
        else:
            logging.warning(f"No se encontró el archivo: {ruta_pdf}")

# Procesar datos por filas
def procesar_datos_por_filas(ruta_pdf_base, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja, indices_a_buscar):
    for mes in range(1, 13):
        ruta_pdf = os.path.join(ruta_pdf_base, f"{mes}.pdf")

        if os.path.exists(ruta_pdf):
            valores_extraidos = extraer_valores_indices(ruta_pdf, indices_a_buscar)
            escribir_en_hoja_por_filas(valores_extraidos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja)
        else:
            logging.warning(f"No se encontró el archivo: {ruta_pdf}")

# Procesar datos por ubicaciones específicas
def procesar_datos_por_ubicaciones(ruta_pdf_base, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja, indices_a_buscar):
    for mes in range(1, 13):
        ruta_pdf = os.path.join(ruta_pdf_base, f"{mes}.pdf")

        if os.path.exists(ruta_pdf):
            valores_extraidos = extraer_valores_indices(ruta_pdf, indices_a_buscar)
            escribir_en_hoja_por_ubicaciones(valores_extraidos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja)
        else:
            logging.warning(f"No se encontró el archivo: {ruta_pdf}")

# Procesar datos de tablas
def procesar_datos_tablas(ruta_pdf_base, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja, extractor_func):
    for mes in range(1, 13):
        ruta_pdf = os.path.join(ruta_pdf_base, f"{mes}.pdf")

        if os.path.exists(ruta_pdf):
            datos_extraidos = extractor_func(ruta_pdf)
            escribir_en_hoja_por_ubicaciones(datos_extraidos, mes, ruta_plantilla, ruta_salida, ubicaciones, nombre_hoja)
        else:
            logging.warning(f"No se encontró el archivo: {ruta_pdf}")

# Configuración de archivos y datos
ruta_plantilla = "./pdf/plantilla.xlsx"
ruta_excel_salida = "datos_anuales.xlsx"

# Ejemplo de uso
procesar_datos_por_hoja(
    "./pdf/103",
    ruta_plantilla,
    ruta_excel_salida,
    ubicaciones_celdas_hoja1,
    "103 VS ATS",
    indices_a_buscar_103
)

procesar_datos_por_filas(
    "./pdf/104",
    ruta_plantilla,
    ruta_excel_salida,
    ubicaciones_celdas_hoja2,
    "103 VS 104",
    indices_a_buscar_104
)

procesar_datos_por_ubicaciones(
    "./pdf/104",
    ruta_plantilla,
    ruta_excel_salida,
    ubicaciones_celdas_hoja4,
    "A4",
    indices_a_buscar_104
)

procesar_datos_tablas(
    "./pdf/compras",
    ruta_plantilla,
    ruta_excel_salida,
    ubicaciones_celdas_hoja4,
    "Tablas",
    extraer_totales_compras
)

procesar_datos_tablas(
    "./pdf/retenciones",
    ruta_plantilla,
    ruta_excel_salida,
    ubicaciones_celdas_hoja4,
    "Tablas",
    extraer_codigos_retencion
)