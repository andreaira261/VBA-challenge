Attribute VB_Name = "Module1"
Sub stock_data()

For Each ws In Worksheets

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    Dim lastrow As Long
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim ticker_name As String
     
    Dim ticker_total As Double
    ticker_total = 0
    
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    Dim open_price_row As Long
    open_price_row = 2
    
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ticker_name = ws.Cells(i, 1).Value
            
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            
            ws.Range("I" & Summary_Table_Row).Value = ticker_name
            
            ws.Range("L" & Summary_Table_Row).Value = ticker_total
    
            ticker_total = 0
            
            close_price = ws.Cells(i, 6).Value
            
            open_price = ws.Cells(open_price_row, 3).Value
                    
            ws.Cells(Summary_Table_Row, 10).Value = close_price - open_price
                    
            ws.Cells(Summary_Table_Row, 11).Value = (close_price - open_price) / open_price
                    
            ws.Cells(Summary_Table_Row, 11).Value = FormatPercent(ws.Cells(Summary_Table_Row, 11).Value, 2)
                    
            Summary_Table_Row = Summary_Table_Row + 1
                    
            open_price_row = i + 1
        
        Else
        
            ticker_total = ticker_total + ws.Cells(i, 7).Value
    
    
        End If
    
    Next i
    
    Dim tickerlastrow As Long
    tickerlastrow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To tickerlastrow
    
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            ws.Cells(i, 11).Interior.ColorIndex = 3
        
        ElseIf ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            ws.Cells(i, 11).Interior.ColorIndex = 4
        
        End If
    
    Next i
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    Greatest_inc = ws.Cells(2, 11).Value
    Greatest_dec = ws.Cells(2, 11).Value
    Greatest_tot_vol = ws.Cells(2, 12).Value
    
    For i = 2 To tickerlastrow
    
        If ws.Cells(i, 11).Value > Greatest_inc Then
            Greatest_inc = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 17).Value = FormatPercent(ws.Cells(2, 17).Value, 2)
            
            ElseIf ws.Cells(i, 11).Value = Greatest_inc Then
            Greatest_inc = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 17).Value = FormatPercent(ws.Cells(2, 17).Value, 2)
        
        End If
        
        If ws.Cells(i, 11).Value < Greatest_dec Then
            Greatest_dec = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 17).Value = FormatPercent(ws.Cells(3, 17).Value, 2)
            
            ElseIf ws.Cells(i, 11).Value = Greatest_dec Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 17).Value = FormatPercent(ws.Cells(3, 17).Value, 2)
                
        End If
        
        If ws.Cells(i, 12).Value > Greatest_tot_vol Then
            Greatest_tot_vol = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            ElseIf ws.Cells(i, 12).Value = Greatest_tot_vol Then
            Greatest_tot_vol = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        
        End If
    
    Next i
    
    ws.Columns("A:Q").AutoFit

Next ws

End Sub
