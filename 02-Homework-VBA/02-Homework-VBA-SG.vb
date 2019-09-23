Sub VBAHw()

    'Loop through all sheets
    For Each ws In Worksheets
    
        'Create variables for volume, ticker change counter, opening, and closing prices
        Dim volume As Double
        Dim ticker As Double
        Dim opening As Double
        Dim closing As Double
        Dim pctchg As Double
    
        opening = ws.Cells(2, 3).Value
        closing = 0
        ticker = 1
        pctchg = 0
        volume = 0
    
        'Create headers for summary columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'Determine the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Iterate through rows calculate values for summary table
        For i = 2 To LastRow
        
            volume = volume + ws.Cells(i, 7).Value
            closing = ws.Cells(i, 6).Value
            pctchg = (closing - opening) / opening
            
            On Error Resume Next
        
            'Add values to summary table
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
                ticker = ticker + 1
            
                ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
            
                ws.Cells(ticker, 10).Value = volume
                
                ws.Cells(ticker, 11).Value = closing - opening
                
                ws.Cells(ticker, 12).Value = pctchg
                ws.Cells(ticker, 12).NumberFormat = "0.00%"
        
                opening = ws.Cells(i + 1, 3).Value
                closing = 0
                volume = 0
                
                ' Add conditional formatting to yearly change column
                
                ws.Cells(1, 11).Interior.ColorIndex = 0
            
                    If ws.Cells(ticker, 11).Value < 0 Then
            
                    ws.Cells(ticker, 11).Interior.ColorIndex = 3
                
                    ElseIf ws.Cells(ticker, 11).Value >= 0 Then
            
                    ws.Cells(ticker, 11).Interior.ColorIndex = 4
                    
                    End If
                
            End If
        
        Next i
        
        ' Find greatest increase, decrease and total volume
        
        maxchg = WorksheetFunction.Max(ws.Range("L:L"))
        minchg = WorksheetFunction.Min(ws.Range("L:L"))
        maxvol = WorksheetFunction.Max(ws.Range("J:J"))
        
        ws.Cells(2, 17).Value = maxchg
        ws.Cells(2, 16).Value = maxchg.Offset(0, -3).Value
        
        ws.Cells(3, 17).Value = minchg
        ws.Cells(3, 16).Value = minchg.Offset(0, -3).Value
        
        ws.Cells(4, 17).Value = maxvol
        ws.Cells(4, 16).Value = maxvol.Offset(0, -1).Value
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next ws
    
End Sub

