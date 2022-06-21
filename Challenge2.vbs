Attribute VB_Name = "Module1"

Sub Test():

    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
    sheetname = sheet.Name
    
    Ticker = ""
    summaryTableRow = 2
    Closing = 0
    Opening = 0
    YearlyChange = 0
    PercentChange = 0
    Total = 0
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    sheet.Cells(1, 9).Value = "Ticker"
    sheet.Cells(1, 10).Value = "YearlyChange"
    sheet.Cells(1, 11).Value = "Percent Change"
    sheet.Cells(1, 12).Value = "Total Stock Volume"
    
        For Row = 2 To lastRow
        
            Total = Total + Cells(Row, 7)
            
            If Row = 2 Then
                Opening = Cells(Row, 3).Value
            
            ElseIf sheet.Cells(Row + 1, 1).Value <> sheet.Cells(Row, 1).Value Then
            
                Ticker = sheet.Cells(Row, 1).Value
                sheet.Cells(summaryTableRow, 9).Value = Ticker
                
                Closing = Cells(Row, 6).Value
                
                YearlyChange = Closing - Opening
                sheet.Cells(summaryTableRow, 10).Value = YearlyChange
                
                PercentChange = (YearlyChange) / Opening
                sheet.Cells(summaryTableRow, 11).Value = PercentChange
                sheet.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
                
                sheet.Cells(summaryTableRow, 12).Value = Total
                
                If YearlyChange < 0 Then
                    sheet.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
                Else
                    sheet.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
                End If
                
                Total = 0
                Opening = Cells(Row + 1, 3).Value
                
                summaryTableRow = summaryTableRow + 1
                
            End If
        
        Next Row
        
    Next sheet
    
End Sub









