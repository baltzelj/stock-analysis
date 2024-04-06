Attribute VB_Name = "Module1"
Sub stock_test()
    ' Preparing variables that will function through the entire program.
    Dim ticker As String
    Dim ws As Worksheet
    Dim open_value As Double
    Dim close_value As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim stock_volume As Double
    
    ' Looping through every worksheet.
    For Each ws In Worksheets
    
     ' Setting headers for each column in all the worksheets, and preparing variables for future multi-worksheet work.
        ws.Range("I:L").EntireColumn.Insert
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("I1:L1").Font.Bold = True
        
    ' Compiling unique data for each ticker value including yearly change, percent change, and total stock volume.
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Looping to gather data.
        For i = 2 To last_row
            ' Represents a start to a new ticker.
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                open_value = ws.Cells(i, 3).Value
                stock_volume = ws.Cells(i, 7).Value
            ' Represents a continuation of the same ticker value.
            ElseIf ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value Then
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                ' Represents last ticker in a series.
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    close_value = ws.Cells(i, 6).Value
                    yearly_change = close_value - open_value
                    percent_change = yearly_change / open_value
                    
                    ' Setting cells to collected variables. Establishing conditional formatting for values.
                    last_ticker = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
                    ws.Cells(last_ticker, 9).Value = ticker
                        ws.Cells(last_ticker, 10).Value = yearly_change
                    If ws.Cells(last_ticker, 10).Value >= 0 Then
                        ws.Cells(last_ticker, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(last_ticker, 10).Value < 0 Then
                        ws.Cells(last_ticker, 10).Interior.ColorIndex = 3
                    End If
                    ws.Cells(last_ticker, 11).Value = percent_change
                    ws.Cells(last_ticker, 11).NumberFormat = "0.00%"
                    If ws.Cells(last_ticker, 11).Value >= 0 Then
                        ws.Cells(last_ticker, 11).Interior.ColorIndex = 4
                    ElseIf ws.Cells(last_ticker, 11).Value < 0 Then
                        ws.Cells(last_ticker, 11).Interior.ColorIndex = 3
                    End If
                ws.Cells(last_ticker, 12).Value = stock_volume
                End If
            End If
        Next i
    
        ' Create new columns and headers for the analysis of compiled data.
        ws.Range("N:P").EntireColumn.Insert
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Stock Volume"
            ws.Range("N2:N4").Font.Bold = True
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
            ws.Range("O1:P1").Font.Bold = True
        
        ' Finding greatest percent changes.
        last_row2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        greatest_increase = 0
        greatest_decrease = 0
        stock_vol = 0
        
        For j = 2 To last_row2
            If ws.Cells(j, 11).Value > greatest_increase Then
                ticker_inc = ws.Cells(j, 9).Value
                greatest_increase = ws.Cells(j, 11).Value
                ws.Cells(2, 15).Value = ticker_inc
                ws.Cells(2, 16).Value = greatest_increase
                ws.Cells(2, 16).NumberFormat = "0.00%"
            ElseIf ws.Cells(j, 11).Value < greatest_decrease Then
                ticker_dec = ws.Cells(j, 9).Value
                greatest_decrease = ws.Cells(j, 11).Value
                ws.Cells(3, 15).Value = ticker_dec
                ws.Cells(3, 16).Value = greatest_decrease
                    ws.Cells(3, 16).NumberFormat = "0.00%"
            End If
            If ws.Cells(j, 12) > stock_vol Then
                ticker_vol = ws.Cells(j, 9).Value
                stock_vol = ws.Cells(j, 12).Value
                ws.Cells(4, 15).Value = ticker_vol
                ws.Cells(4, 16).Value = stock_vol
            End If
        Next j
    Next
End Sub
