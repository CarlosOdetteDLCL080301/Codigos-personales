@echo off
REM Esto desactiva la visualización de los comandos en la consola

REM Ruta del intérprete de Python
set python_exe=python

REM Ruta del script de Python que deseas ejecutar
set script_path=C:\Users\carlo\OneDrive\Escritorio\Codigos-personales\Python\creadorPlaylistFarmacia.py

REM Comando para ejecutar el script de Python
%python_exe% %script_path%

REM Mostrar cuadro de diálogo de aviso con mensaje personalizado
msg * "El programa ha finalizado su ejecucion con exito"

REM Pausa opcional para mantener la ventana de la consola abierta después de la ejecución
exit
