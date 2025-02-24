Attribute VB_Name = "StockDataModule"

Sub GetUniqueValues()

Range("A2:A93001").AdvancedFilter _
Action:=xlFilterCopy, CopyToRange:=Range("I2"), Unique:=True

End Sub

Sub CalculateQuarterlyChange()
Dim ws As Worksheet
Dim lastRow As Long, i As Long
Set ws = ThisWorkbook.Worksheets("Sheet1")
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow

    Dim openingPrice As Double
    openingPrice = ws.Cells(i, 1).Value
        
    Dim closingPrice As Double
    closingPrice = ws.Cells(i + 2, 2).Value
    
    Dim quarterlyChange As Double
    quarterlyChange = (closingPrice - openingPrice) / openingPrice
    
    ws.Cells(i, 3).Value = quarterlyChange * 100
    Next i
    
End Sub
