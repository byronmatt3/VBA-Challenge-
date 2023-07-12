Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
       
        Dim ticker As String
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Double
        totalVolume = 0
        
       
        Dim summaryTableRow As Integer
        summaryTableRow = 2
        
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                ws.Cells(summaryTableRow, 9).Value = ticker
                ws.Cells(summaryTableRow, 10).Value = yearlyChange
                ws.Cells(summaryTableRow, 11).Value = percentChange
                ws.Cells(summaryTableRow, 11).NumberFormat = "0"
                ws.Cells(summaryTableRow, 12).Value = totalVolume
                
                If yearlyChange >= 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(summaryTableRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                openingPrice = ws.Cells(i + 1, 3).Value
                summaryTableRow = summaryTableRow + 1
                totalVolume = 0
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                If openingPrice = 0 Then
                    openingPrice = ws.Cells(i, 3).Value
                End If
            End If
        Next i
    Next ws
End Sub
