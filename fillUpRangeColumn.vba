Sub fillUpRange()
    
    'Limpiar la columna
    Range("J12:J" & Rows.Count).ClearContents
    
   ' Definir la hoja de trabajo "Pressure"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pressure")

    ' Obtener los valores mínimo y máximo de las celdas G7 y H7
    Dim minValue As Double
    Dim maxValue As Double
    minValue = Range("G7").Value
    maxValue = Range("H7").Value
    
    ' Incrementa y pociciona los valores de T
    Dim i As Integer
    Dim fila As Integer
    
    fila = 12 ' Inicia en la primera fila
    For i = Int(minValue) To Int(maxValue)
        Cells(fila, 10).Value = i
        fila = fila + 1
    Next i
    
End Sub
