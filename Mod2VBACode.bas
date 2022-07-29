Attribute VB_Name = "Module1"
Sub StockData()

Range("I1:L1").EntireColumn.Insert
Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")



Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim StockVolume_Total As LongLong
StockVolume_Total = 0

Dim Sumary_Table_Row As Integer
Summary_Table_Row = 2

For i = 2 To 759001

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        
        StockVolume_Total = StockVolume_Total + Cells(i, 7).Value
        
        Range("I" & Summary_Table_Row).Value = Ticker
        
        Range("L" & Summary_Table_Row).Value = StockVolume_Total
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        StockVolume_Total = 0
        
    Else
    
        StockVolume_Total = StockVolume_Total + Cells(i, 7).Value
    
    End If
Next i

End Sub
