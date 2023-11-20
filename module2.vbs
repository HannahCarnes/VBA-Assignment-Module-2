Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerSymbol As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTable As Range
    Dim summaryTableRow As Long
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
    Set ws = ThisWorkbook.Worksheets("A")
    
    ' find the last row on the worksheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' setting up summary table
    summaryTableRow = 2
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    ' loop through and get each value of table
    For i = 2 To lastRow
        ' check if the current row has a different ticker symbol than the previous row
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
            tickerSymbol = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            totalVolume = 0
        End If
        
        ' closing price and total volume
        closingPrice = ws.Cells(i, 6).Value
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        ' check if the row is the last row for the ticker symbol
        If ws.Cells(i + 1, 1).Value <> tickerSymbol Then
            ' yearly change and percent change
            yearlyChange = closingPrice - openingPrice
            percentChange = yearlyChange / openingPrice * 100
            
            ' put the information in the summary table
            ws.Cells(summaryTableRow, 9).Value = tickerSymbol
            ws.Cells(summaryTableRow, 10).Value = yearlyChange
            ws.Cells(summaryTableRow, 11).Value = percentChange
            ws.Cells(summaryTableRow, 12).Value = totalVolume
            
            
            ' format colors for positive and negative percent changes
            If yearlyChange > 0 Then
                ws.Cells(summaryTableRow, 10).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Cells(summaryTableRow, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            ' plus one to move to next row in summary table
            summaryTableRow = summaryTableRow + 1
        End If
    Next i
  Next ws

    
End Sub