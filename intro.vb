Sub Test()
    MsgBox ("hello there")
End Sub

Sub Ranges()
    ThisWorkbook.Sheets("Sheet1").Range("A5:F10").Value = "hola zahid"
End Sub

Sub Ranges2()
    ThisWorkbook.Sheets("Sheet1").Range("A5:F10").Value = ThisWorkbook.Sheets("Sheet1").Range("H3").Value
End Sub

Sub Cells()
    ThisWorkbook.Sheets("Sheet1").Cells(1,1).Value = "wow"
End Sub

Sub Cells()
    ThisWorkbook.Sheets("Sheet1").Cells(1,1).Value = ThisWorkbook.Sheets("Sheet1").Cells(2,1).Value + 1
End Sub

Sub FillSelection()
    Selection.Value = "Filled"
End Sub

Sub FillSelection()
    Selection.Interior.Color = RGB(30,3,58)
    Selection.Font.Color = RGB(255,255,255)
End Sub

Sub Variables()
    Dim x As Integer, y As Integer
    x = 4
    MsgBox (x)
    MsgBox (y)
End Sub

Sub Grafica()
Charts.Add
ActiveChart.ChartType = xlXYScatterLines
ActiveChart.SetSourceData Source:=Sheets("estacion").Range("E1:E14")
ActiveChart.Location Where:=xlLocationAsObject, Name:="estacion"
End Sub