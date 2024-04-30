Attribute VB_Name = "Módulo1"
Public Sub EnviarCorreos()
    Dim cell As Range


    ' ------------
    
    ' Inicializar Outlook
    ' Set OutApp = CreateObject("Outlook.Application")
    ' ===================================================================
    ' ALGORITMO PARA TRANSFORMAR LA HOJA DE "CONTACTO" EN UN DICCIONARIO
    ' TIENE LA POSIBILIDAD DE ANALIZAR EL TAMAÑO DE LA HOJA
    ' ===================================================================
    ' El siguiente codigo funciona para crear nuestro directorio (a partir de diccionarios)
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Establecer la hoja de trabajo
    Set ws = Sheets("Contactos")
    
    ' Encontrar la última fila en la columna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Verificar si la última fila en la columna A es menor o igual a 1 (hoja vacía)
    If lastRow <= 1 Then
        MsgBox "No hay datos en la hoja 'Contactos'."
        Exit Sub
    End If
    
    ' Establecer el rango dinámico
    Set rng = ws.Range("A2:B" & lastRow)
    
    Dim datos As Variant
    Dim fila As Integer
    Dim columna As Integer
    Dim filas As Integer
    Dim columnas As Integer
    ' Necesario para hacer diccionarios (Estructura de datos)
    Dim dic As Object
    ' Formato que adquirirá es dic = {"Usuario":"correo electronico"}
    Set dic = CreateObject("Scripting.Dictionary")
    datos = rng.Value
    filas = UBound(datos, 1)
    columnas = UBound(datos, 2)
    
    For fila = 1 To filas
        If Not IsEmpty(datos(fila, 1)) And Not IsEmpty(datos(fila, 2)) Then
            dic.Add datos(fila, 1), datos(fila, 2)
        End If
    Next fila
    
    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    ' Eliminar
    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    ' Recorrer el diccionario
    Dim clave As Variant
    For Each clave In dic.Keys
        Debug.Print "Clave: " & clave & ", Valor: " & dic(clave) & "TYPE: "; TypeName(dic(clave))
    Next clave
    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' FIN DEL ALGORITMO
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    ' ===================================================================
    ' ALGORITMO QUE FILTRARÁ A LOS CONTACTOS QUE SE NECESITARÁ ENVIAR
    ' EL CORREO
    ' ===================================================================
    Dim dict2 As Object
    ' Definir la hoja de Excel
    Set ws = Sheets("Reporte")
    
    ' Inicializar el diccionario para almacenar los valores únicos
    Set dict2 = CreateObject("Scripting.Dictionary")
    
    ' Encontrar la última fila en la columna A
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ' Verificar si la última fila en la columna A es menor o igual a 1 (hoja vacía)
    If lastRow <= 1 Then
        MsgBox "No hay datos en la hoja 'Reporte'."
        Exit Sub
    End If
    
    ' Obtener el rango de la columna E
    Set rng = ws.Range("E2:E" & lastRow)
    Debug.Print "E2:E" & lastRow
    ' Iterar sobre cada celda en la columna E y almacenar los valores únicos en el diccionario
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            ' Agregar el valor de la celda al diccionario si no está en él
            If Not dict2.exists(cell.Value) Then
                dict2.Add cell.Value, 1
            End If
        End If
    Next cell
    
    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    ' Eliminar
    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    ' Recorrer el diccionario
    For Each clave In dict2.Keys
        Debug.Print "_________Clave: " & clave & ", Valor: " & dict2(clave) & "TYPE: "; TypeName(dict2(clave))
    Next clave
    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    
    
    ' ===================================================================
    ' ALGORITMO QUE FILTRARÁ A LOS REPORTES CONSIDERANDO UNICAMENTE LOS
    ' USUARIOS EXISTENTES
    ' ===================================================================
    Dim fila2 As Range
    ' Definir el rango de la columna E
    Set rng = ws.Range("A2:H" & lastRow)
    Dim cuerpoFiltrado As String
    cuerpoFiltrado = ""
    For Each clave In dict2.Keys
        ' Iterar sobre cada fila en el rango
        Debug.Print "Correo para " & clave
        For Each fila2 In rng.Rows
            ' Verificar si el valor en la columna E de la fila es igual a la clave
            If fila2.Cells(1, 5).Value = clave Then
                ' Concatenar los valores de las fila coincidentes
                cuerpoFiltrado = cuerpoFiltrado & " | " & fila2.Cells(1, 1).Value & " | " & fila2.Cells(1, 2).Value & " | " & fila2.Cells(1, 3).Value & " | " & fila2.Cells(1, 4).Value & " | " & fila2.Cells(1, 5).Value & " | " & fila2.Cells(1, 6).Value & " | " & fila2.Cells(1, 7).Value & " | " & fila2.Cells(1, 8).Value & " | " & vbCrLf
            End If
        Next fila2
        ' Debug.Print cuerpoFiltrado
         ' Llamar a la subrutina para enviar correos personalizados
        
        ' XMXMXMXMXMXMXMXMXMXMXMMXMXXMXMXMXMXMXMXMXMXMXMXMMXMXXMXMXMXMXMXMXMXMXMXMXMMXMX
        ' SE APLICARÁ TODA LA FUNCIONALIDAD DEL ENVIO DE CORREOS
        ' XMXMXMXMXMXMXMXMXMXMXMMXMXXMXMXMXMXMXMXMXMXMXMXMMXMXXMXMXMXMXMXMXMXMXMXMXMMXMX
        
        Dim OutlookApp As Object
        Dim OutlookMail As Object
        
        ' Crear una nueva instancia de Outlook
        Set OutlookApp = CreateObject("Outlook.Application")
        Set OutlookMail = OutlookApp.CreateItem(0)
        
        ' Configurar el correo
        With OutlookMail
            .To = dic(clave) ' Dirección de correo del destinatario
            .CC = "carlosodettedlcl@gmail.comCOPIA; ferny.cruz0406@gmail.com"
            .Subject = "PRUEBA DEL PROGRAMA" ' Asunto del correo
            .Body = "Hola " & clave & vbCrLf & "Estos son tus tickets" & vbCrLf & cuerpoFiltrado & vbCrLf & "Saludos" ' Cuerpo del correo
            ' Enviar el correo
            .Send
        End With
        
        ' Liberar los objetos OutlookMail y OutlookApp
        Set OutlookMail = Nothing
        Set OutlookApp = Nothing

        
        ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        cuerpoFiltrado = ""
    Next clave

    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' FIN DEL ALGORITMO
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
End Sub


