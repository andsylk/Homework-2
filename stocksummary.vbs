Sub stocksummary()

'declare worksheet
Dim ws As Worksheet

'apply to all worksheets
For Each ws In Worksheets

    'declare variables
    Dim ticker As String
    Dim stock_row As Integer
    stock_row = 2

    'last row
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'declare more variables
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total As Double
    Dim current_volume As Double

    total = 0

    'loop to run through rows
    For i = 2 To lastrow
        'counting total stock volume
        current_volume = ws.Cells(i, 7).Value
        total = current_volume + total
            'get the open_price of each stock
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                open_price = ws.Cells(i, 3).Value
            End If
            'get the ticker symbols and close_price of each stock
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                close_price = ws.Cells(i, 6).Value
                ws.Cells(stock_row, 9).Value = ticker
                'calculate yearlyly change and percentage chage
                ws.Cells(stock_row, 10).Value = close_price - open_price
                 If open_price = 0 Or close_price = 0 Or close_price - open_price = 0 Then
                    ws.Cells(stock_row, 11).Value = 0
                Else
                    ws.Cells(stock_row, 11).Value = (close_price - open_price) / open_price
                End If
                'fill in total stock volume of stock
                ws.Cells(stock_row, 12).Value = total
                'formatting % and colors
                ws.Cells(stock_row, 11).NumberFormat = "0.00%"
                    If ws.Cells(stock_row, 10).Value >= 0 Then
                        ws.Cells(stock_row, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(stock_row, 10).Value < 0 Then
                        ws.Cells(stock_row, 10).Interior.ColorIndex = 3
                    End If
                stock_row = stock_row + 1
                'resetting total stock volume
                total = 0
            End If
    Next i

    'headers
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    
    'declare variables
    Dim max_inc As Double
    Dim max_dec As Double
    Dim max_vol As Double
    Dim max_inc_ticker As String
    Dim max_dec_ticker As String
    Dim max_vol_ticker As String
    max_inc = 0
    max_dec = 0
    max_vol = 0

    'getting the ticker symbol and value
    For i = 2 To lastrow
        If ws.Cells(i, 11).Value > max_inc Then
            max_inc = ws.Cells(i, 11).Value
            max_inc_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value < max_dec Then
            max_dec = ws.Cells(i, 11).Value
            max_dec_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 12).Value > max_vol Then
            max_vol = ws.Cells(i, 12).Value
            max_vol_ticker = ws.Cells(i, 9).Value
        End If
    Next i
    
    'formatting and filling in the values
    ws.Cells(2, 17).Value = max_inc_ticker
    ws.Cells(3, 17).Value = max_dec_ticker
    ws.Cells(4, 17).Value = max_vol_ticker
    ws.Cells(2, 18).Value = max_inc
    ws.Cells(2, 18).NumberFormat = "0.00%"
    ws.Cells(3, 18).Value = max_dec
    ws.Cells(3, 18).NumberFormat = "0.00%"
    ws.Cells(4, 18).Value = max_vol

Next ws

End Sub




