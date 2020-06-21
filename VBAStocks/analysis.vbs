Sub analysis()

    For Each ws In Worksheets
        ws.Activate
        
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"

        Dim summary_count As Integer
        Dim Ep As Double
        Dim Op As Double
        Dim YearChange As Double
        Dim Tot_Vol As Double
        
        Tot_Vol = 0
        
        summary_count = 2

        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        Op = Cells(2, 3).Value
        
        Cells(2, 9).Value = Cells(2, 1).Value
        

        Columns("K").NumberFormat = "0.00%"

        For data_count = 2 To LastRow
            
            If Cells(data_count, 1).Value <> Cells(data_count + 1, 1).Value Then
                
                Tot_Vol = Tot_Vol + Cells(data_count, 7)
                
                Cells(summary_count, 12).Value = Tot_Vol
                
                Ep = Cells(data_count, 6).Value
                
                YearChange = Ep - Op
                
                Cells(summary_count, 10).Value = YearChange
                
                If YearChange > 0 Then
                
                    Cells(summary_count, 10).Interior.Color = vbGreen
                    
                ElseIf YearChange < 0 Then
                    
                    Cells(summary_count, 10).Interior.Color = vbRed
                    
                End If
                
                If Op = 0 Then
                    
                    Cells(summary_count, 11).Value = 0
                
                Else
                    
                    Cells(summary_count, 11).Value = (Ep - Op) / Op
                
                End If
                
                summary_count = summary_count + 1
                
                Cells(summary_count, 9).Value = Cells(data_count + 1, 1).Value
                
                Op = Cells(data_count + 1, 3).Value
                
                Tot_Vol = 0
                
            Else
            
                Tot_Vol = Tot_Vol + Cells(data_count, 7)
                
            End If
            
        Next data_count

'        Worksheets(Name).Columns("A:M").AutoFit

   
        SummaryLastRow = Cells(Rows.count, 9).End(xlUp).Row

        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
            
        For j = 2 To SummaryLastRow
            If Cells(j, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & SummaryLastRow)) Then
                Cells(2, 15).Value = Cells(j, 9).Value
                Cells(2, 16).Value = Cells(j, 11).Value
                Cells(2, 16).NumberFormat = "0.00%"
                
            ElseIf Cells(j, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & SummaryLastRow)) Then
                Cells(3, 15).Value = Cells(j, 9).Value
                Cells(3, 16).Value = Cells(j, 11).Value
                Cells(3, 16).NumberFormat = "0.00%"
                
            ElseIf Cells(j, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & SummaryLastRow)) Then
                Cells(4, 15).Value = Cells(j, 9).Value
                Cells(4, 16).Value = Cells(j, 12).Value
            
            End If
        Next j

    Next ws

End Sub
