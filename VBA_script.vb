Sub alphabet_draft()

'define variables
Dim ws As Worksheet
Dim lastRow As Long
Dim currentRow As Long
Dim tickerRow As Integer
Dim v_totalStock As LongLong
Dim priceChange As Double
Dim percentChange As Double
Dim v_startPrice As Double
Dim v_closePrice As Double

'iterate through all sheets
For Each ws In Sheets

    'create headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock "
    
    'row number of last row
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'initial values
    tickerRow = 2
    v_totalStock = 0
    v_startPrice = Range("C3").Value
    v_closePrice = 0
    
    'iterate through sheet
    For currentRow = 2 To lastRow
     
        'if ticker is the same
        If Cells(currentRow + 1, 1) = Cells(currentRow, 1) Then
            
            'add stock
            v_totalStock = v_totalStock + Cells(currentRow, 7).Value
            
        
        'if ticker changes next row - Cells(currentRow + 1, 1) <> Cells(currentRow, 1)
        Else
        
            'final stock count
             v_totalStock = v_totalStock + Cells(currentRow, 7).Value
        
            'store closing price
            v_closePrice = Cells(currentRow, 6).Value
            
            'calculate changes - round b/c we don't need crazy decimals
            priceChange = v_closePrice - v_startPrice
            percentChange = Round((priceChange / v_startPrice) * 100, 2)
            
            'print results to respective columns
            
            'ticker symbol
            Range("I" & tickerRow).Value = Cells(currentRow, 1).Value
            'yearly change
            Range("J" & tickerRow).Value = Round(priceChange, 2)
            'percent change
            Range("K" & tickerRow).Value = percentChange & "%"
            'final stock volume
            Range("L" & tickerRow).Value = v_totalStock
            
            'conditional formatting for colors
            If Range("J" & tickerRow).Value > 0 Then
                Range("J" & tickerRow).Interior.ColorIndex = 4
                
            ElseIf Range("J" & tickerRow).Value < 0 Then
                Range("J" & tickerRow).Interior.ColorIndex = 3
            
            Else
                Range("J" & tickerRow).Interior.ColorIndex = 0
            End If
            
            'reset variables for new ticker
            v_totalStock = 0
            'next row will contain new startPrice
            v_startPrice = Cells(currentRow + 1, 3).Value
            v_closePrice = 0
            priceChange = 0
            percentChange = 0
            'move ticker storage to next row
            tickerRow = tickerRow + 1
        
        End If
        
    
    Next currentRow

Next ws

End Sub
