import fitz  # PyMuPDF
import re

def analizar_informe_local(ruta_pdf):
    """
    Analiza un PDF de inspección leyendo las últimas dos páginas
    y buscando explícitamente "certifica que Cumple" o "certifica que No Cumple".
    Retorna el estado y las observaciones.
    """
    try:
        doc = fitz.open(ruta_pdf)
        texto = ""
        paginas = list(range(len(doc)))
        ultimas_dos = paginas[-2:] if len(paginas) >= 2 else paginas
        for i in ultimas_dos:
            texto += doc[i].get_text("text") + "\n"
        doc.close()
        
        if not texto.strip():
            return "ERROR LECTURA", "PDF sin texto extraíble."

        # Búsqueda usando regex case-insensitive
        if re.search(r'certifica\s+que\s+no\s+cumple', texto, re.IGNORECASE):
            return "NO CUMPLE", "Equipo NO cumple con los requerimientos."
        elif re.search(r'certifica\s+que\s+cumple', texto, re.IGNORECASE):
            return "CUMPLE", "Equipo cumple con los requerimientos."
        else:
            return "DESCONOCIDO", "No se encontró el texto esperado en el informe."

    except Exception as e:
        return "ERROR LECTURA", f"Fallo al leer PDF local: {str(e)[:40]}"
