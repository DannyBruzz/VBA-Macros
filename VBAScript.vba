Sub stocks()

    Dim ws As Worksheet

    For Each ws In Worksheets

            Dim LR As Long
            LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            Dim row_Table As Integer
            row_Table = 2
    
            Dim vol As Double
            vol = 0
    
            Dim first As Double
            first = ws.Cells(2, 3).Value

            Dim last As Double
    
            Dim tickcount As Integer
            tickcount = 0
    
            Dim hightickvalue As Double
            Dim lowtickvalue As Double
            Dim highvol As Double
   

            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("I1").Value = "Ticker"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
    
    
            For i = 2 To LR
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    last = ws.Cells(i, 6).Value
                    ws.Cells(row_Table, 10).Value = last - first
                    ws.Cells(row_Table, 11).Value = ws.Cells(row_Table, 10).Value / first
                    ws.Cells(row_Table, 11).NumberFormat = "0.00%"
                    ws.Cells(row_Table, 9).Value = ws.Cells(i, 1).Value
            
                    vol = vol + Cells(i, 7).Value
                    ws.Cells(row_Table, 12).Value = vol
                    row_Table = row_Table + 1
                    first = ws.Cells(i + 1, 3).Value
                    vol = 0
                    tickcount = tickcount + 1
                Else
                    vol = vol + ws.Cells(i, 7).Value
            
                End If
            
            Next i


            For x = 2 To tickcount

                If ws.Cells(x, 10).Value >= 0 Then
                    ws.Cells(x, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(x, 10).Interior.ColorIndex = 3
                End If
    
            Next x
    
            highvol = WorksheetFunction.Max(ws.Range("L:L").Value)
            ws.Range("Q4").Value = highvol

            hightickvalue = WorksheetFunction.Max(ws.Range("J:J").Value)
            ws.Range("Q2").Value = hightickvalue

            lowtickvalue = WorksheetFunction.Min(ws.Range("J:J").Value)
            ws.Range("Q3").Value = lowtickvalue


            For x = 2 To tickcount

                If ws.Cells(x, 10).Value = hightickvalue Then
                    ws.Range("P2").Value = ws.Cells(x, 9).Value
                ElseIf ws.Cells(x, 10).Value = lowtickvalue Then
                    ws.Range("P3").Value = ws.Cells(x, 9).Value
                ElseIf ws.Cells(x, 12).Value = highvol Then
                    ws.Range("P4").Value = ws.Cells(x, 9).Value
                End If
            Next x
    
      
    Next ws
    
    End Sub
    



