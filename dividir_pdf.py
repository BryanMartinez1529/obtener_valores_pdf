import os
import re
import win32com.client
from PyPDF2 import PdfReader, PdfWriter

def convertir_docx_a_pdf(ruta_docx, ruta_pdf):
    print("🔄 Convirtiendo DOCX a PDF...")
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(ruta_docx)
    doc.SaveAs(ruta_pdf, FileFormat=17)  # 17 = wdFormatPDF
    doc.Close()
    word.Quit()

def extraer_nombre_funcionario(texto):
    # Buscar la línea después de "Estimado(a)"
    lineas = texto.splitlines()
    for i, linea in enumerate(lineas):
        if "Estimado/a:" in linea:
            if i + 1 < len(lineas):
                nombre = lineas[i + 1].strip()
                # Limpiar caracteres no válidos para nombres de archivo
                nombre = re.sub(r'[\\/*?:"<>|]', "", nombre)
                # Eliminar espacios adicionales
                nombre = re.sub(r'\s+', " ", nombre).strip()
                return nombre
    return "Funcionario_Desconocido"

def dividir_pdf_en_funcionarios(pdf_path, paginas_por_funcionario, salida_dir):
    print("✂️ Dividiendo el PDF...")
    reader = PdfReader(pdf_path)
    total_paginas = len(reader.pages)
    grupo = 1

    for i in range(0, total_paginas, paginas_por_funcionario):
        writer = PdfWriter()

        # Extraer nombre desde la primera página del grupo
        texto = reader.pages[i].extract_text()
        nombre_funcionario = extraer_nombre_funcionario(texto) or f"Funcionario_{grupo:03}"

        for j in range(i, min(i + paginas_por_funcionario, total_paginas)):
            writer.add_page(reader.pages[j])
        
        salida_path = os.path.join(salida_dir, f"{nombre_funcionario}.pdf")
        with open(salida_path, "wb") as salida:
            writer.write(salida)
        print(f"📄 Guardado: {salida_path}")
        grupo += 1

def main():
    ruta_docx = os.path.abspath(r"C:\Users\HP VICTUS\Desktop\BRYAN\BRYAN\PORTAFOLIO\proyecto_imp\notificaciones_generadas\convocatoria_borrador.docx")
    ruta_pdf = os.path.abspath("documento_temporal.pdf")
    carpeta_salida = os.path.abspath("convocatoria_borrador")


    os.makedirs(carpeta_salida, exist_ok=True)

    convertir_docx_a_pdf(ruta_docx, ruta_pdf)
    dividir_pdf_en_funcionarios(ruta_pdf, paginas_por_funcionario=1, salida_dir=carpeta_salida)

    print("✅ ¡Proceso completado!")

if __name__ == "__main__":
    main()
