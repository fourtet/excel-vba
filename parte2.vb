Sub aGrafica2()
    Dim desplazarGrafica as Integer
    desplazarGrafica = 15
    Worksheets("ESTACIÓN 1").Activate
    For j = 1 To Worksheets("ESTACIÓN 1").Rows.Count Step 5
        Dim iMasUno As Integer
        desplazarGrafica = desplazarGrafica + 225
        If IsEmpty(Cells(j, 1).Value) Then
            Exit Sub
        Else
            Dim i As Integer
            Dim myChtObj As ChartObject

            Set myChtObj = Worksheets("GRA1").ChartObjects.Add _
                (Left:=0, Width:=650, Top:=desplazarGrafica, Height:=225)
            myChtObj.Chart.SetSourceData Source:=Sheets("ESTACIÓN 1").Range("E1:E6,G1:G6,I1:I6,K1:K6,M1:M6")
            myChtObj.Chart.ChartType = xlXYScatterLines
            myChtObj.Chart.ChartTitle.Text = "Grafica"
            For i = 1 To 5
                iMasUno = j + i
                
                'If i = 4 Then
                    'myChtObj.Chart.SeriesCollection.NewSeries
                'End If
                myChtObj.Chart.SeriesCollection(i).Values = Worksheets("ESTACIÓN 1").Range("E" & iMasUno & ",G" & iMasUno & ",I" & iMasUno & ",K" & iMasUno & ",M" & iMasUno & ",O" & iMasUno & ",Q" & iMasUno & ",S" & iMasUno & ",U" & iMasUno & ",W" & iMasUno & ",Y" & iMasUno & ",AA" & iMasUno & ",AC" & iMasUno & ",AE" & iMasUno & ",AG" & iMasUno & ",AI" & iMasUno & ",AK" & iMasUno & ",AM" & iMasUno & ",AO" & iMasUno & ",AQ" & iMasUno & ",AS" & iMasUno & ",AU" & iMasUno & ",AW" & iMasUno & ",AY" & iMasUno)
                myChtObj.Chart.SeriesCollection(i).XValues = ""
                myChtObj.Chart.SeriesCollection(i).Name = Worksheets("ESTACIÓN 1").Range("D" & iMasUno)
            Next i
        End If
    Next j
    
End Sub


Sub aQuitarSerie2()
    Dim nombreDeLaSerie As String
    Dim numeroFilaSerie As String
    Dim celdasAQuitar As String
    
    strFormula = Selection.Formula
    
    Dim strRangeFromFormula As String
    strRangeFromFormula = Mid(strFormula, _
                        InStrRev(strFormula, "!") + 1, _
                        InStrRev(strFormula, ",") - InStrRev(strFormula, "!") - 2)
    'test here
    'Debug.Print strRangeFromFormula 'result: $D$4:$D$8
    MsgBox strRangeFromFormula
    
    numeroFilaSerie = Mid(strRangeFromFormula, 5)
    MsgBox numeroFilaSerie
    
    celdasAQuitar = "E" & numeroFilaSerie & ":AZ" & numeroFilaSerie
    MsgBox celdasAQuitar

    Worksheets("ESTACIÓN 1").Range(celdasAQuitar) = ""

    
    'nombreDeLaSerie = Worksheets("ESTACIÓN 1").ChartObjects("My chart").Name
    'Selection.Delete
End Sub


Sub aPromediar2()
    Worksheets("ESTACIÓN 1").Activate
    Dim contadorFilas As Integer
    contadorFilas = 0
    Worksheets("ESTACIÓN 1").Activate
    For m = 1 To Worksheets("ESTACIÓN 1").Rows.Count
        If Not IsEmpty(Cells(m, 1).Value) Then
            contadorFilas = contadorFilas + 1
        Else
            Exit For
        End If
    Next m
    MsgBox contadorFilas
    For c = 1 To Worksheets("ESTACIÓN 1").Rows.Count Step 1
        If IsEmpty(Cells(c, 1).Value) Then
            Exit Sub 
        Else
            Dim numeroDeFilas As Integer
            numeroDeFilas = contadorFilas / 5
            
            Dim i As Integer, j As Integer, multiploCinco As Integer, contador As Integer
            ThisWorkbook.Sheets("ESTACIÓN 1").Activate
            For i = 1 To numeroDeFilas
                contador = i - 1
                multiploCinco = 5 * contador
                For j = 1 To 48
                    '2,5
                    '2,6
                    ThisWorkbook.Sheets("PROM EST 1").Cells(i + 1, j + 4).Value = _
                    Application.WorksheetFunction.Average(ActiveSheet.Range(Cells(multiploCinco + 2, j + 4), Cells(multiploCinco + 6, j + 4)))
                    '2,5 6,5
                    '2,6 6,6
                Next j
            Next i
        End If
    Next c
    

End Sub