Sub analysis():

    For Each ws in Worksheets
        Range("I1") = "Ticker" 
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"

        Dim count As Integer
        count = 2

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



        For i = 2 To LastRow
            If Cells(i,1).Value <> Cells(i-1,1).Value Then
                Cells(count,9).Value = Cells(i,1).Value
                count = count + 1
            End If
        Next i

        Worksheets(ws.Name).Column("A:L").Autofit
    Next ws

End Sub
    
