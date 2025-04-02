
import pdfplumber
import re

def extraer_valores_indices(ruta_pdf, indices_buscados):
    valores_encontrados = {indice: None for indice in indices_buscados}  # Diccionario inicial

    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            palabras = pagina.extract_words()  # Extraer palabras individuales

            for i in range(len(palabras)):
                palabra_actual = palabras[i]['text']

                # Si la palabra actual es un índice buscado
                if palabra_actual in valores_encontrados:
                    if i + 1 < len(palabras):  # Asegurar que hay un valor después
                        valor_extraido = palabras[i + 1]['text'].replace('.', ',')  # Convertir punto a coma
                        valores_encontrados[palabra_actual] = valor_extraido

    return valores_encontrados  # Retorna el diccionario con los valores extraídos
 # Retorna el diccionario con los valores extraídos



ruta_pdf = "./pdf/104/1.pdf"
indices_103 = [
    "303",
    "3030",
    "304",
    "304B",
    "307",
    "308",
    "309",
    "310",
    "311",
    "312",
    "312A",
    "3121",
    "314",
    "319",
    "320",
    "322",
    "323",
    "324",
    "325",
    "326",
    "327",
    "328",
    "332",
    "332G",
    "336",
    "337",
    "343",
    "344",
    "3440",
    "345",
    "346",
    "421"
]  # Lista de índices a buscar


indices_104 = [
    "500",
    "501",
    "502",
    "503",
    "540",
    "505",
    "506",
    "507",
    "508",
    "531",
    "532",
    "535",
    "510",
    "511",
    "512",
    "513",
    "550",
    "515",
    "516",
    "517",
    "518",
    "721",
    "723",
    "725",
    "729",
    "731"
]

valores_extraidos = extraer_valores_indices(ruta_pdf, indices_104)

# Imprimir el diccionario con los resultados
print("Valores extraídos:", valores_extraidos)