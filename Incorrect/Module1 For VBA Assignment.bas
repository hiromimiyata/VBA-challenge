Attribute VB_Name = "Module1"
Sub stocks()
'Titles for table
     Cells(1, 9).Value = "Ticker"
     Cells(1, 10).Value = "Yearly Change"
     Cells(1, 11).Value = "Percent Change"
     Cells(1, 12).Value = "Total Stock Volume"
     Cells(1, 13).Value = "Opening Value"
     Cells(1, 14).Value = "Ending Value"
     
'Variables to go through the column are defined
    Dim arow, Column As Long
    Column = 2
    Firstopen = True
    ending = Cells(Rows.Count, 1).End(xlUp).Row

    'For loop values in ticker column
    For arow = 2 To ending
    If Firstopen = True Then
        Firstopen = False
        Cells(Column, 13).Value = Cells(arow, 3).Value
        Cells(Column, 9).Value = Cells(arow, 1).Value

    ElseIf Cells(arow, 1).Value <> Cells(arow + 1, 1).Value Then
        Cells(Column, 14).Value = Cells(arow, 6).Value
        Column = Column + 1
        Firstopen = True
        

    End If
    Next arow

End Sub

Sub stocks2()
    ending2 = Cells(Rows.Count, 13).End(xlUp).Row
    For Values = 2 To ending2
    Cells(Values, 10).Value = Cells(Values, 14).Value - Cells(Values, 13).Value
    Next Values
    Range("J:J").Style = "Currency"
End Sub
Sub stocks3()
    ending2 = Cells(Rows.Count, 13).End(xlUp).Row
    For Values = 2 To ending2
    Cells(Values, 11).Value = (Cells(Values, 14).Value - Cells(Values, 13).Value) / Cells(Values, 13).Value
    Next Values
    Range("K:K").NumberFormat = "0.00%"
End Sub

Sub stockvolume()
    Dim aroww, Columnn As Long
    ending3 = Cells(Rows.Count, 1).End(xlUp).Row
    TotalST = 0
    Columnn = 2
    For aroww = 2 To ending3
        If Cells(aroww, 1).Value = Cells(aroww + 1, 1).Value Then
            TotalST = TotalST + Cells(aroww, 7).Value
        Else
            Cells(Columnn, 12).Value = TotalST
            Columnn = Columnn + 1
        End If
    Next aroww
End Sub

Sub greatest()
    Cells(2, 17).Value = "Greatest % increase"
    Cells(3, 17).Value = "Greatest % decrease"
    Cells(4, 17).Value = "Greatest total volume"
    Cells(1, 18).Value = "Ticker"
    Cells(1, 19).Value = "Value"
    endd = Cells(Rows.Count, 11).End(xlUp).Row
    Dim best, worst, i As Integer
    best = Application.WorksheetFunction.Max(Range("K:K"))
    bestvolume = Application.WorksheetFunction.Max(Range("L:L"))
    worst = Application.WorksheetFunction.Min(Range("K:K"))
    For i = 2 To endd
        If Cells(i, 11).Value = best Then
            Cells(2, 18).Value = Cells(i, 9).Value
            Cells(2, 19).Value = Cells(i, 11).Value
        ElseIf Cells(i, 12).Value = bestvolume Then
            Cells(4, 18).Value = Cells(i, 9).Value
            Cells(4, 19).Value = bestvolume
        ElseIf Cells(i, 11).Value = worst Then
            Cells(3, 18).Value = Cells(i, 9).Value
            Cells(3, 19).Value = Cells(i, 11).Value
        End If
    Next i
    Range("S2:S3").NumberFormat = "0.00%"
End Sub
Sub redorgreen()
    ending4 = Cells(Rows.Count, 11).End(xlUp).Row
    For h = 10 To 11
    For i = 2 To ending4
        If Cells(i, h).Value > 0 Then
            Cells(i, h).Interior.Color = vbGreen
        Else
            Cells(i, h).Interior.Color = vbRed
        End If
    Next i
    Next h
End Sub

Sub Runall()
    Dim was As Worksheet
    For Each ws In Worksheets
        ws.Activate
        stocks
        stocks2
        stocks3
        stockvolume
        redorgreen
        greatest
    Next ws
End Sub

