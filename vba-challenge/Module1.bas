Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

Dim ws As Worksheet


For Each ws In Worksheets

    Dim tickerSym As String
    Dim totalVol As Double

    Dim yearOpen As Double
    Dim yearClose As Double

    Dim summaryRow As Integer
    summaryRow = 2

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    Dim lastRow As Double
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If (ws.Cells(i, 3).Value = 0) Then
            If (ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value) Then
                tickerSym = ws.Cells(i, 1).Value
            End If
        ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
            totalVol = totalVol + ws.Cells(i, 7).Value
            If (ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value) Then
                yearOpen = ws.Cells(i, 3).Value
            End If
        Else
            tickerSym = ws.Cells(i, 1).Value
            totalVol = totalVol + ws.Cells(i, 7).Value
            yearClose = ws.Cells(i, 6).Value
            ws.Cells(summaryRow, 9).Value = tickerSym
            ws.Cells(summaryRow, 12).Value = totalVol
            If (totalVol > 0) Then
                ws.Cells(summaryRow, 10).Value = yearClose - yearOpen
                    If (ws.Cells(summaryRow, 10).Value > 0) Then
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                    End If
                ws.Cells(summaryRow, 11).Value = ws.Cells(summaryRow, 10).Value / yearOpen
            Else
                ws.Cells(summaryRow, 11).Value = 0
                ws.Cells(summaryRow, 12).Value = 0
            End If
            ws.Cells(summaryRow, 11).Style = "percent"
            totalVol = 0
            summaryRow = summaryRow + 1
        End If
    Next i

    
    Dim greatTotVol As Double

    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    greatTotVol = 0

    summaryRow = summaryRow - 2
    For i = 2 To summaryRow
        If (ws.Cells(i, 12).Value > greatTotVol) Then
            greatTotVol = ws.Cells(i, 12).Value

            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        End If
    Next i

    ws.Cells(4, 17).Value = greatTotVol

    Dim increasePerc As Double
    Dim decreasePerc As Double

   
    increasePerc = 0
    decreasePerc = 0

    For i = 2 To summaryRow
        If (ws.Cells(i, 11).Value > increasePerc) Then
            increasePerc = ws.Cells(i, 11).Value

            ws.Cells(2, 16) = ws.Cells(i, 9).Value
        ElseIf (ws.Cells(i, 11).Value < decreasePerc) Then
            decreasePerc = ws.Cells(i, 11).Value

            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        End If
    Next i

    ws.Cells(2, 17).Value = increasePerc
    ws.Cells(3, 17).Value = decreasePerc

    ws.Cells(2, 17).Style = "percent"
    ws.Cells(3, 17).Style = "percent"

    ws.Columns("J:Q").AutoFit

Next ws

End Sub

