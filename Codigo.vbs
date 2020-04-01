
Sub prueba()


Dim endOfTickerRow, beginTickerStocks, numeroDeStocks, numeroDeSheet As Integer
Dim precioOpenBY, precioCloseEY, c, d As Integer

Dim nombreDeStock As String


For numeroDeSheet = 0 To 2

Worksheets("2016").Activate
Sheets(ActiveSheet.Index + numeroDeSheet).Activate

'Crear titulos de las columnas
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Precio Inicial"
Cells(1, 12).Value = "precio final"
Cells(1, 13).Value = "Cambio YoY($)"
Cells(1, 14).Value = "Cambio YoY(%)"
Cells(1, 15).Value = "# de row del cambio"
Cells(1, 16).Value = "Volumen ($)"

'Crear un loop que me de el nombre de todas las acciones listadas y sus rendimientos

endOfTickerRow = Cells(Rows.Count, "A").End(xlUp).Row
numeroDeStocks = 2
'precioCloseEY = 03

For beginTickerStocks = 2 To endOfTickerRow

If Cells(beginTickerStocks, 1) <> Cells(beginTickerStocks - 1, 1) Then
    nombreDeStock = Cells(beginTickerStocks, 1).Value
    precioOpenBY = Cells(beginTickerStocks, 3).Value
    
    Cells(numeroDeStocks, 10).Value = nombreDeStock
    Cells(numeroDeStocks, 11).Value = precioOpenBY
    Cells(numeroDeStocks, 15).Value = beginTickerStocks
    numeroDeStocks = numeroDeStocks + 1
End If



Next beginTickerStocks

Cells(numeroDeStocks, 15).Value = endOfTickerRow + 1
'Agregar el volumen

Dim volumen, columnaFechaInicial, columnaFechaFinal As Long

For volumen = 2 To numeroDeStocks - 1
    columnaFechaInicial = Cells(volumen, 15).Value
    columnaFechaFinal = Cells(volumen + 1, 15).Value - 1

    Cells(volumen, 16).Value = Application.Sum(Range(Cells(columnaFechaInicial, 7), Cells(columnaFechaFinal, 7)))

Next volumen

'Agregar Precios finales
Dim a, b, e, h As Double

For a = 2 To numeroDeStocks - 1

b = Cells(a + 1, 15).Value
e = Cells(b - 1, 6).Value
Cells(a, 12).Value = e

Next a
 
'Generar calculos de rendimientos

For h = 2 To numeroDeStocks - 1

Cells(h, 13).Value = Cells(h, 12).Value - Cells(h, 11).Value

If Cells(h, 11).Value = 0 Then
Cells(h, 14).Value = 0
Else

Cells(h, 14).Value = (Cells(h, 12).Value / Cells(h, 11).Value) - 1
End If

If Cells(h, 13).Value >= 0 Then
Cells(h, 13).Interior.Color = RGB(0, 255, 0)
Else
Cells(h, 13).Interior.Color = RGB(255, 51, 0)
End If

Next h

'Formateando y retirando las columnas sobrantes

    Columns("N:N").Select
    Selection.Style = "Percent"
    Columns("P:P").Select
    Selection.Style = "Comma"
    Columns("K:L").Select
    Selection.Delete Shift:=xlToLeft
    Columns("M:M").Select
    Selection.Delete Shift:=xlToLeft

'Puntos extras
'--------------------------------------------------

Dim loopCounter, valorPivot, topYield, rowCounter, f As Integer

'Generando los titulos de cada casilla
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest Total Volume"

valorPivot = Cells(2, 12).Value

rowCounter = Cells(Rows.Count, "J").End(xlUp).Row

'Obteniendo el stock con mayor rendimiento
For loopCounter = 3 To rowCounter

If valorPivot < Cells(loopCounter, 12).Value Then
    valorPivot = Cells(loopCounter, 12).Value
    topYield = loopCounter
End If

Next loopCounter

Cells(2, 18).Value = valorPivot
Cells(2, 17).Value = Cells(topYield, 10).Value

valorPivot = Cells(2, 12).Value

'Obteniendo el stock con menor rendimiento
For loopCounter = 3 To rowCounter

If valorPivot > Cells(loopCounter, 12).Value Then
    valorPivot = Cells(loopCounter, 12).Value
    topYield = loopCounter
End If

Next loopCounter

Cells(3, 18).Value = valorPivot
Cells(3, 17).Value = Cells(topYield, 10).Value
valorPivot = Cells(2, 13).Value

'Obteniendo el stock con mayor volumen
For loopCounter = 3 To rowCounter

If valorPivot < Cells(loopCounter, 13).Value Then
    valorPivot = Cells(loopCounter, 13).Value
    topYield = loopCounter
End If

Next loopCounter

Cells(4, 18).Value = valorPivot
Cells(4, 17).Value = Cells(topYield, 10).Value

'Dandole formato a las tres casillas

    Range("R2:R3").Select
    Selection.Style = "Percent"
    Range("R4").Select
    Selection.Style = "Comma"
    Columns("A:R").Select
    Columns("A:R").EntireColumn.AutoFit



Next numeroDeSheet

End Sub