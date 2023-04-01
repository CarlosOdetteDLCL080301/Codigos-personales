import os
import fitz  # esta es la biblioteca PyMuPDF

def algunErrorConElPdf(archivo_pdf):
    try:
    # Verificar si el archivo PDF existe
        if os.path.exists(archivo_pdf):
            # Obtener el tamaño del archivo en bytes
            tamano_archivo = os.path.getsize(archivo_pdf)
            
            # Verificar si el archivo tiene contenido
            if tamano_archivo > 0:
                return True
            else:
                os.remove(archivo_pdf)
                return False
        else:
            return False
    except ValueError as error:
        os.remove(archivo_pdf)
        return False

def pdfEncontroLaCadena(nombreArchivo,cadena):
    try:
        doc = fitz.open(nombreArchivo)
        # Itera sobre las páginas del archivo PDF
        for page in doc:
            # Obtiene el texto de la página actual
            text = page.get_text()
            
            #este if ternario, me permite identificar la materia que necesito filtrar, sin embargo habría otro problema el cual si otro pdf tiene esta cadena lo guardara en esta nueva ubicación, es un riesgo que lo vale
            return True if cadena in text else False
    except ValueError:
        return False

def moverPdf(rutaActualArchivo,rutaDestinoArchivo_tentativa):
    simbolosInvalidos = ['<', '>', ':', '"', '\\', '|', '?', '*']
    
    for simbolo in simbolosInvalidos:
        resultado = rutaDestinoArchivo_tentativa[3:].replace(simbolo, "_")
        rutaDestinoArchivo_tentativa = rutaDestinoArchivo_tentativa[:3] + resultado
    
    if not os.path.exists(rutaDestinoArchivo_tentativa):
        os.makedirs(rutaDestinoArchivo_tentativa)
    rutaDestinoArchivo = os.path.join(rutaDestinoArchivo_tentativa, os.path.basename(rutaActualArchivo))
    os.rename(rutaActualArchivo,rutaDestinoArchivo)

def filtrarPDFs(ruta_carpeta,cadena):
    # Itera sobre todos los archivos en la carpeta
    for nombreArchivo in os.listdir(ruta_carpeta):
        # Comprueba si el archivo es un archivo PDF
        if nombreArchivo.endswith('.pdf'):#Si existe un pdf, pondremos a buscar la palabra para poder cambiar su ubicación en el directorio
            filepath = os.path.join(ruta_carpeta, nombreArchivo)
            #if pdfEncontroLaCadena(filepath,cadena):
            #print(f'Archivo: {filepath} {ruta_carpeta}')
            if (pdfEncontroLaCadena(filepath,cadena) and algunErrorConElPdf(filepath)):
                moverPdf(filepath,ruta_carpeta + cadena)
    print("".center(100,"*"))
    print(f"\tSe finalizó la busqueda de la cadena \"{cadena}\"\t".center(100,"*"))
    print("".center(100,"*"))

filtrarPDFs("F:/Copia de respaldo/Pdf´s/","Grupo: 13")
