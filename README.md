# ExcelQRGenerator
Complemento de Excel para generar un QR consumiendo una API publica y usando valores de celdas

---

## Uso:
importa el complemento a Excel y llamalo como funcion en una celda usando:

GeneradorQR(_celda1, celda2, celda3, celda4_)

La función manda un String como argumento a una API generadora de QR y devuelve una
imagen del QR con la cadena de texto creada:

## Ejemplo:

Teniendo valores 
|  |B   |
|---| --|
| 3 |24 | 
| 4 |texto aleatorio|
| 5 |hola|
| 6 | 55 |

![image](https://github.com/user-attachments/assets/ed549867-7c1f-4cf6-bb97-acd54eaf7dac)


## usando la formula:
`=GeneradorQR(B3,B4,B5,B6)`

el string que se genera: "Escribe cualquier frase 24 - texto aleatorio - hola - 55"

el resultado que se genera:

![image](https://github.com/user-attachments/assets/30f4fed7-8d6a-4a0e-b551-b3c35d322327)

---

# CODIGO:

```
Function GeneradorQR(dato1, dato2, dato3, dato4)
Dim texto, url As String
Dim frase As String
```
Existen las variables texto y frase que son strings.
La variable texto almacena toda la cadena de valores que se pasan como parametros
y la variable frase es un texto predeterminado que se guarda para que siempre se genere
junto con el texto completo.


```
frase = "Escribe cualquier frase "
texto = frase & dato1 & " - " & dato2 & " - " & dato3 & " - " & dato4
```
Los valores dato1, dato2, dato3 y dato4 son los parametros que se mandan al ejecutar la función.


`url = "https://qrcode.tec-it.com/API/QRCode?data=" & WorksheetFunction.EncodeURL(texto)`
La variable url guarda la cadena de texto con la URL de la API y concatena todo el string guardado en texto.

```
Set celda = ActiveCell
celda.Worksheet.Shapes.AddPicture url, False, True, celda.Left, celda.Top, 20, 20
GeneradorQR = ""

End Function
```
En este apartado del codigo se crea la imagen que se agrega en el excel haciendo uso de la url.


