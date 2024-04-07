Attribute VB_Name = "Module3"
Sub experimenting()
    Dim ws As Worksheet
    ' Timing how long the code takes (just for fun).
    Dim time1 As Long, time2 As Long
    time1 = Timer
    
    ' Erasing prior attempt.
    For Each ws In Worksheets
        ws.Range("I:W").Clear
    Next

    For Each ws In Worksheets
        ' Setting up cells with their appropriate headers.
        ws.Range("I:P").EntireColumn.Insert
        ws.Range("I1:P1").Font.Bold = True
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N1:N4").Font.Bold = True
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Stock Volume"
        ws.Range("O1:P1").Font.Bold = True
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"
        
        ' Collecting data variables.
        Dim ticker As String, open_value As Double, close_value As Double, stock_value As Double
        
        final_data = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To final_data
            ' Collecting initial data set for ticker.
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                open_value = ws.Cells(i, 3).Value
                stock_value = ws.Cells(i, 7).Value
            ' Collecting all middle data sets for ticker value.
            ElseIf ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value Then
                stock_value = stock_value + ws.Cells(i, 7).Value
                ' Checking for final ticker value, and compiling data if that is the case.
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    close_value = ws.Cells(i, 6).Value
                    last_data = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
                    ws.Cells(last_data, 9).Value = ticker
                    ws.Cells(last_data, 10).Value = close_value - open_value
                        If ws.Cells(last_data, 10).Value >= 0 Then
                            ws.Cells(last_data, 10).Interior.ColorIndex = 4
                        ElseIf ws.Cells(last_data, 10).Value < 0 Then
                            ws.Cells(last_data, 10).Interior.ColorIndex = 3
                        End If
                    ws.Cells(last_data, 11).Value = (close_value - open_value) / open_value
                        ws.Cells(last_data, 11).NumberFormat = "0.00%"
                        If ws.Cells(last_data, 11).Value >= 0 Then
                            ws.Cells(last_data, 11).Interior.ColorIndex = 4
                        ElseIf ws.Cells(last_data, 11).Value < 0 Then
                            ws.Cells(last_data, 11).Interior.ColorIndex = 3
                        End If
                    ws.Cells(last_data, 12).Value = stock_value
                End If
            End If
        Next i
        ' Determining greatest percent increase and decrease, along with largest stock volume.
        last_comp_data = ws.Cells(Rows.Count, 9).End(xlUp).Row
        inc_perc = 0
        dec_perc = 0
        max_stock = 0
        For j = 2 To last_comp_data
            ' Finding the greatest increase percent.
            If ws.Cells(j, 11).Value > inc_perc Then
                inc_perc = ws.Cells(j, 11).Value
                inc_ticker = ws.Cells(j, 9).Value
                ws.Range("O2").Value = inc_ticker
                ws.Range("P2").Value = inc_perc
                    ws.Range("P2").NumberFormat = "0.00%"
            ' Finding the greatest decrease percent.
            ElseIf ws.Cells(j, 11).Value < dec_perc Then
                dec_perc = ws.Cells(j, 11).Value
                dec_ticker = ws.Cells(j, 9).Value
                ws.Range("O3").Value = dec_ticker
                ws.Range("P3").Value = dec_perc
                    ws.Range("P3").NumberFormat = "0.00%"
            End If
            ' Finding the maximum stock volume.
            If ws.Cells(j, 12).Value > max_stock Then
                max_stock = ws.Cells(j, 12).Value
                max_tick = ws.Cells(j, 9).Value
                ws.Range("O4").Value = max_tick
                ws.Range("P4").Value = max_stock
            End If
        Next j
    Next

time2 = Timer
MsgBox "This code took " & (time2 - time1) & " seconds to complete."

End Sub
