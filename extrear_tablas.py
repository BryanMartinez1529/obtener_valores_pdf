import pdfplumber
import re

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


# === EJECUCIÓN ===
ruta_pdf = "./pdf/3_ats.pdf"

# 1. Extraer Retenciones
retenciones = extraer_codigos_retencion(ruta_pdf)
print("🧾 Retenciones encontradas:")
for cod, datos in retenciones.items():
    print(f"  Código: {cod} | Base: {datos['base']} | Valor: {datos['valor']}")

# 2. Extraer Totales COMPRAS
compras_totales = extraer_totales_compras(ruta_pdf)
print("\n🛒 Totales de COMPRAS:")
for nombre, valor in compras_totales.items():
    print(f"  {nombre}: {valor}")
