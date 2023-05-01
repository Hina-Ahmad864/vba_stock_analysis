Sub testing()

Dim i As Long
Dim start As Long
Dim j As Integer
Dim ticker_volume_total As Double
Dim yearly_price_change As Double
Dim percent_change As Double
Dim lastRow As Long
Dim ws As Worksheet



    For Each ws In Worksheets

    'set values for each worksheet
    j = 0
    start = 2
    ticker_volume_total = 0
    yearly_price_change = 0
    daily_price_change = 0
    
    'name columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'set variable for last row
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
        For i = 2 To lastRow
                                                               
            'find when ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'update ticker_volume_total when ticker does change to account for the last row's volume
                ticker_volume_total = ticker_volume_total + ws.Cells(i, 7).Value
                
                'condition for zero total volume
                If ticker_volume_total = 0 Then
                
                    'print
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
            
                Else
                    'find first starting price value
                    If ws.Cells(start, 3) = 0 Then
                        
                        'finds change in tickers
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    
                    
                    'calculate the yearly price change and the percent change
                    yearly_price_change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percent_change = yearly_price_change / ws.Cells(start, 3)
                    
                    'to continue to the next ticker
                    start = i + 1
                    
                    'print
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = yearly_price_change
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).Value = percent_change
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + j).Value = ticker_volume_total
                    
                    'color yearly price change
                    Select Case yearly_price_change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                    
                End If
                
                'reset variables
                ticker_volume_total = 0
                yearly_price_change = 0
                j = j + 1
            
            'ticker is the same 
            Else
                'if ticker is the same conintue to add ticker_volume
                ticker_volume_total = ticker_volume_total + ws.Cells(i, 7).Value
                
            End If
            

            
        Next i
        
        'move max an min to another part onto worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
        
        'account for header row
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
        
        
        'ticker symbol for the volume total and increase and decrese %
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        
    Next ws
    
End Sub
