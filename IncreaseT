Private Sub IncreaseT_Scroll()

   ' Definir la hoja de trabajo "Pressure"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pressure")

    ' Obtener los valores mínimo y máximo de las celdas G7 y H7
    Dim minValue As Double
    Dim maxValue As Double
    minValue = Range("G7").Value
    maxValue = Range("H7").Value
    
    ' Referencia a la barra de desplazamiento llamada "IncreaseT"
    With ws.Shapes("IncreaseT").ControlFormat
    IncreaseT.Min = minValue
    IncreaseT.Max = maxValue
    End With
    
    ' Obtener el valor actual de la celda C4
    Dim currentValue As Double
    currentValue = Range("C4").Value
    
    ' Determinar el valor de incremento o decremento
    Dim newValue As Double
    
'   Reemplaza "IncreaseT" por el nombre exacto de tu barra de desplazamiento
'    If ws.Shapes("IncreaseT").ControlFormat.Value > currentValue Then
'        newValue = currentValue + 1 ' Incrementar en 1
'    Else
'        newValue = currentValue - 1 ' Decrementar en 1
'    End If
    
    If Me.IncreaseT.Value > currentValue Then
        newValue = currentValue + 1 ' Incrementar en 1
    Else
        newValue = currentValue - 1 ' Decrementar en 1
    End If
    
    ' Verificar que el nuevo valor esté dentro del rango permitido
    If newValue >= minValue And newValue <= maxValue Then
        Range("C4").Value = newValue
    Else
        MsgBox ("Drag value out of range")
    End If
End Sub
