Attribute VB_Name = "Módulo1"
Option Explicit

' Funcion para generar un codigo QR a partir de valores
' creados en la hoja, se puede configurar la cantidad
' de datos que se incertan, asi como un texto predefinido
' Si no funciona la generacion del QR, revise que la API
' del url siga funcionando, si no, cambie de API

Function GeneradorQR(dato1, dato2, dato3, dato4)
    Dim texto, url As String
    Dim frase As String
    Dim celda As Range
    
    frase = "Escribe cualquier frase "
    
    
    texto = frase & dato1 & " - " & dato2 & " - " & dato3 & " - " & dato4
    
    url = "https://qrcode.tec-it.com/API/QRCode?data=" & WorksheetFunction.EncodeURL(texto)
    
    Set celda = ActiveCell
    
    celda.Worksheet.Shapes.AddPicture url, False, True, celda.Left, celda.Top, 20, 20
    
    GeneradorQR = ""

End Function

