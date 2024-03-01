Sub stock()

    Dim ws As Worksheet

    For Each ws In Worksheets
        
        Dim yearopen As Double
        Dim yearclose As Double
        Dim yearchange As Double
        Dim percentchange As Double
        Dim totalvolume As Double
        Dim summaryrow As Long
        Dim ticker As String
        Dim tickerrow As Long
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_volume As Double
        Dim greatest_increase_ticker As String
        Dim greatest_decrease_ticker As String
        Dim greatest_volume_ticker As String
        
        summaryrow = 2
        ticker = ws.Cells(2, 1).Value
        tickerrow = 2
        greatest_increase = 0
        greatest_decrease = 0

        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ticker = ws.Cells(i, 1).Value
            
                ws.Range("I" & summaryrow).Value = ticker

                yearopen = ws.Cells(tickerrow, 3).Value
                yearclose = ws.Cells(i, 6).Value
                yearchange = yearclose - yearopen
                percentchange = yearchange / yearopen
            
                totalvolume = 0
                For j = tickerrow To i
                    totalvolume = totalvolume + ws.Cells(j, 7).Value
                Next j
            
                ws.Range("J" & summaryrow).Value = yearchange
                ws.Range("K" & summaryrow).Value = percentchange * 100 & "%"
                ws.Range("L" & summaryrow).Value = totalvolume
            
                If ws.Range("J" & summaryrow).Value > 0 Then
                    
                    ws.Range("J" & summaryrow).Interior.Color = RGB(0, 255, 0)
                ElseIf ws.Range("J" & summaryrow).Value < 0 Then
                    
                    ws.Range("J" & summaryrow).Interior.Color = RGB(255, 0, 0)
                End If
                

                If percentchange > greatest_increase Then
                    greatest_increase = percentchange
                    greatest_increase_ticker = ticker
                ElseIf percentchange < greatest_decrease Then
                    greatest_decrease = percentchange
                    greatest_decrease_ticker = ticker
                End If

                
                If totalvolume > greatest_volume Then
                    greatest_volume = totalvolume
                    greatest_volume_ticker = ticker
                End If

                summaryrow = summaryrow + 1
                tickerrow = i + 1
            
            End If
        Next i

        ws.Range("O2").Value = greatest_increase_ticker
        ws.Range("O3").Value = greatest_decrease_ticker
        ws.Range("O4").Value = greatest_volume_ticker
        ws.Range("P2").Value = greatest_increase * 100 & "%"
        ws.Range("P3").Value = greatest_decrease * 100 & "%"
        ws.Range("P4").Value = greatest_volume

    Next ws
End Sub