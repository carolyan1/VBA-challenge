# VBA-challenge

Sub StockChallenge()

For Each ws In Worksheets

    'Dim all variables
    Dim Ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim open_price As Double
    Dim closed_price As Double
    
    'Prepare summary table
    summary_table_row = 2
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    'Find endrow
    endrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set initial value for variables
    total_volume = 0
    open_price = ws.Range("C2").Value
    
     ' Start for loop
     For i = 2 To endrow
     
    'Find Ticker
    If ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        ws.Range("I" & summary_table_row).Value = Ticker
    End If
    
    
    'Find Total Volume
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'volume
        total_volume = total_volume + ws.Cells(i, 7).Value
        ws.Range("L" & summary_table_row).Value = total_volume
        total_volume = 0
        
        'closed and open price
        closed_price = ws.Cells(i, 6).Value
        yearly_change = closed_price - open_price
        ws.Range("J" & summary_table_row).Value = yearly_change
        'calculate percent change
            If open_price = 0 Then
                percent_change = 0
            Else
            percent_change = yearly_change / open_price
            End If
        
        ws.Range("K" & summary_table_row).Value = FormatPercent(percent_change)
        
            If percent_change >= 0 Then
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
            Else
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
            End If
            
        'reset open price
        open_price = ws.Cells(i + 1, 3).Value
        'set summary_table to next row
        summary_table_row = summary_table_row + 1
    Else
        total_volume = total_volume + ws.Cells(i, 7).Value
    End If
    
    Next i

Next ws


End Sub


