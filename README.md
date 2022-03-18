# trabajo-1
programacion de computadores

Private Sub precio_Click()
 Dim precio1 As Double
Dim precio2 As Double
Dim precio3 As Double
Dim horas As String
Dim fechas As String
Dim ruta As String
Dim x As String
Dim media As Double
Const articulos As Double = 3

MsgBox "introduzca precio1:", vbInformation, "precio1"
x = InputBox("Digite precio1", "precio1")
precio1 = Val(x)

MsgBox "introduzca precio2:", vbInformation, "precio2"
x = InputBox("Digite precio2", "precio2")
precio2 = Val(x)

MsgBox "introduzca precio3:", vbInformation, "precio1"
x = InputBox("Digite precio3", "precio3")
precio3 = Val(x)

hora = Format(Time, "hh:mm:ss")
fecha = Format(Date, "dd/mm/yyyy")
ruta = "C:\Users\Kevin\Desktop/entrada.dat"

media = (precio1 + precio2 + precio3) / articulos
If Dir(ruta) = "" Then
Open ruta For Output As #1
Print #1, media & vbCrLf & "hora de entrada: " & hora & vbCrLf & "fecha de entrada: " & fecha & vbCrLf
Close #1
Else
Open ruta For Append As #1
Print #1, media & vbCrLf & "hora de entrada: " & hora & vbCrLf & "fecha de entrada: " & fecha & vbCrLf
Close #1
End If

MsgBox "El precio medio es: " & media, vbInformation, "precio final"
End Sub

Private Sub saluida_Click()
Dim horas As String
Dim fechas As String
Dim ruta As String

hora = Format(Time, "hh:mm:ss")
fecha = Format(Date, "dd/mm/yyyy")
ruta = "C:\Users\Kevin\Desktop/salida.dat"
If Dir(ruta) = "" Then
Open ruta For Output As #1
Print #1, media & vbCrLf & "hora de salida: " & hora & vbCrLf & "fecha salida: " & fecha & vbCrLf
Close #1
Else
Open ruta For Append As #1
Print #1, media & vbCrLf & "hora de salida: " & hora & vbCrLf & "fecha salida: " & fecha & vbCrLf
Close #1
End If


MsgBox "fin de la clase", vbCritical, "salida"
End Sub

