import datetime
import os
from shutil import copy2,copy
import random

def alterarIndices(strings):
    indices = list(range(len(strings)))
    random.shuffle(indices)
    return [strings[i] for i in indices]

def extraerArchivos(ruta_carpeta):
    listaCanciones = []
    ordenDeCanciones = 0
    for raiz,direccion, archivo in os.walk(ruta_carpeta):
        for nombreArchivo in archivo:
            if nombreArchivo.endswith(".mp3"):
                ordenDeCanciones += 1
                listaCanciones.append([nombreArchivo,os.path.join(raiz,nombreArchivo)])
    return alterarIndices(listaCanciones) # Devuelve una lista con los nombres de las canciones en orden aleatorio

def copiar_archivos_mp3(origen, destino):
    for root, dirs, files in os.walk(origen):
        for file in files:
            if file.endswith(".mp3"):
                origen_archivo = os.path.join(root, file)
                destino_archivo = os.path.join(destino, file)
                copy2(origen_archivo, destino_archivo)
                print(f"Archivo copiado: {origen_archivo} -> {destino_archivo}")    

def main():
    ####################################################################################################
    #============================================ VARIABLES ============================================
    ####################################################################################################

    #___________________________
    # VARIABLES DE FECHA Y HORA
    #---------------------------

    # Obtener la fecha y hora actuaL
    ahora = datetime.datetime.now()

    # Formatear la fecha y hora
    formato_fecha = "%d-%m-%Y"
    formato_hora = "%H_%M_%S"
    fecha_formateada = ahora.strftime(formato_fecha)
    hora_formateada = ahora.strftime(formato_hora)

    #___________________________
    # VARIABLES PARA CARPETAS NUEVAS Y LEER RUTAS DE LA MUSICA
    #---------------------------

    # Generar rutas de carpetas
    ruta_generica = "\\Users\\carlo\\OneDrive\\Escritorio\\"
    ruta_carpeta = f"{ruta_generica}Music"
    nuevaCarpeta = f"{ruta_generica}\\PlayList\\Fecha{fecha_formateada}Hora{hora_formateada}"

    # Crear la carpeta donde se almacenara la nueva playlist
    os.makedirs(nuevaCarpeta,exist_ok=True)

    ordenDeCanciones = 0
    archivos = extraerArchivos(ruta_carpeta)
    print(archivos)
    for cancion,ruta in zip(archivos,archivos):
        ordenDeCanciones += 1
        if not ordenDeCanciones % 2 == 0:
            destino_archivo = os.path.join(nuevaCarpeta, f"{ordenDeCanciones}__{cancion[0]}")
            copy(cancion[1], destino_archivo)
            print(f"Archivo copiado: {cancion[1]} -> {destino_archivo}")
        else:
            destino_archivo = os.path.join(nuevaCarpeta, f"{ordenDeCanciones}__farmacia san mateo.mp3")
            copy(f"{ruta_carpeta}\\farmacia san mateo.mp3", destino_archivo)
            pass
        
    print("EL PROGRAMA FINALIZO EXITOSAMENTE")


main()
