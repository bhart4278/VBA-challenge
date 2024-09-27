Sub AnalyzeTickerData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim outputRow As Long
    

    
    ' Start loop through each worksheet and row
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        outputRow = 2
        totalVolume = 0
        openingPrice = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
            
            ' check for when ticker is not the same as the one before
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ticker = ws.Cells(i, 1).Value
                ws.Cells(outputRow, 9).Value = ticker
            
                ' Set closingPrice and totalVolume per quarter count
                
                closingPrice = ws.Cells(i, 6).Value
            
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                ws.Cells(outputRow, 12).Value = totalVolume
            
                ' QuarterlyChange calculation
                quarterlyChange = closingPrice - openingPrice
                ws.Cells(outputRow, 10).Value = quarterlyChange
            
                    ' PercentChange calculation and check divisible by 0
                    If openingPrice <> 0 Then
                        percentChange = (quarterlyChange / openingPrice)
                    Else
                        percentChange = 0
                    End If
                
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                
                ' move to next row
                outputRow = outputRow + 1
                
                ' reset openingPrice to next ticker and reset volume
                openingPrice = ws.Cells(i + 1, 3).Value
                
                totalVolume = 0
            
            Else
                ' continue to sum volume under current ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
            
            End If
        Next i
        
        ' Find last row of quarterlyChange
        QcLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        ' conditional formatting for color of quarterlyChange cells
        For j = 2 To QcLastRow
            If (ws.Cells(j, 10).Value > 0) Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf (ws.Cells(j, 10).Value < 0) Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
        Next j
        
        ' Create table headers
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Outcome"
        ws.Range("N2:N4").Font.Bold = True
        
        ' Find min/max percentChange and largest totalVolume and print to table. Color & format cells
        For k = 2 To QcLastRow
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & QcLastRow)) Then
                ws.Cells(2, 15).Value = ws.Cells(k, 9).Value
                ws.Cells(2, 16).Value = ws.Cells(k, 11).Value
                ws.Cells(2, 16).Interior.ColorIndex = 4
                ws.Cells(2, 16).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & QcLastRow)) Then
                ws.Cells(3, 15).Value = ws.Cells(k, 9).Value
                ws.Cells(3, 16).Value = ws.Cells(k, 11).Value
                ws.Cells(3, 16).Interior.ColorIndex = 3
                ws.Cells(3, 16).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & QcLastRow)) Then
                ws.Cells(4, 15).Value = ws.Cells(k, 9).Value
                ws.Cells(4, 16).Value = ws.Cells(k, 12).Value
            End If
        
        Next k
        
    
    Next ws
        
            
            
    
End Sub

