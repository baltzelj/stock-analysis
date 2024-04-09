Attribute VB_Name = "stockAnalysis"
Sub stockAnalysis()
    'Setup the worksheet so that it has appropriate headers and formatting.
    Dim ws As Worksheet, ticker As String, openValue As Double, closeValue As Double, stockVolume As LongLong
    time1 = Timer
    'Formatting headers and workspace for ease of reading.
    For Each ws In Worksheets
        ws.Range("J:Q").EntireColumn.Insert
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
            ws.Range("J1:M1,O2:O4,P1:Q1").Font.Bold = True
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    Next
    'Read data and compile variables.
    'Saving variables including ticker, opening value, closing value, and total stock volume.
    For Each ws In Worksheets
        last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To last_Row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openValue = ws.Cells(i, 3).Value
                stockVolume = ws.Cells(i, 7).Value
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
                stockVolume = stockVolume + ws.Cells(i, 7).Value
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    closeValue = ws.Cells(i, 6).Value
                    last_Input = ws.Cells(Rows.Count, 10).End(xlUp).Row + 1
                    For j = 2 To last_Input
                        ws.Cells(last_Input, 10).Value = ticker
                        ws.Cells(last_Input, 11).Value = closeValue - openValue
                        ws.Cells(last_Input, 12).Value = (closeValue - openValue) / openValue
                            ws.Cells(last_Input, 12).NumberFormat = "0.00%"
                        ws.Cells(last_Input, 13).Value = stockVolume
                    Next j
                End If
            End If
        Next i
    Next
    'Populate the new worksheet with the collected data.
    'Analysis variables.
    For Each ws In Worksheets
        last_Input = ws.Cells(Rows.Count, 10).End(xlUp).Row
        max_perc = 0
        min_perc = 0
        max_volume = 0
        For k = 2 To last_Input
            If ws.Cells(k, 12).Value > max_perc Then
                max_perc = ws.Cells(k, 12).Value
                max_ticker = ws.Cells(k, 10).Value
                ws.Range("P2").Value = max_ticker
                ws.Range("Q2").Value = max_perc
            ElseIf ws.Cells(k, 12).Value < min_perc Then
                min_perc = ws.Cells(k, 12).Value
                min_ticker = ws.Cells(k, 10).Value
                ws.Range("P3").Value = min_ticker
                ws.Range("Q3").Value = min_perc
            End If
            If ws.Cells(k, 13).Value > max_volume Then
                max_volume = ws.Cells(k, 13).Value
                max_volume_ticker = ws.Cells(k, 10).Value
                ws.Range("P4").Value = max_volume_ticker
                ws.Range("Q4").Value = max_volume
            End If
        Next k
    Next
    time2 = Timer
    MsgBox "This code took " & (time2 - time1) & " seconds to complete."
End Sub

