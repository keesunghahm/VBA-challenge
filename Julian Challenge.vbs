Attribute VB_Name = "Module1"
Sub stocks()
    ' Declare Current as a worksheet object variable.
    Dim ws As Worksheet
    
    ' Loop through all of the worksheets in the active workbook.
    For Each ws In Worksheets
        ' Set an initial variable for holding the ticker
        Dim Ticker As String
        ' Set an initial variable for holding the total per stock
        Dim Total_Volume As Double
        Total_Volume = 0
        ' Keep track of the location for each credit card brand in the summary table
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
        
        Dim opening_price As Double
        opening_price = ws.Cells(2, 3).Value
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Loop through all data
        For i = 2 To RowCount
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker name
                Ticker = ws.Cells(i, 1).Value
                ' Add to the ticker Total
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                ' Print the ticker name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                yearly_change = ws.Cells(i, 6).Value - opening_price
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
                
                percent_change = yearly_change / opening_price
                ws.Range("K" & Summary_Table_Row).Value = percent_change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                ' Print the ticker Amount to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Volume
                
                'if condition for the color themes of changes column J interior. color index
                jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
                For J = 2 To jEndRow
                    If ws.Cells(J, 10) > 0 Then
                        ws.Cells(J, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(J, 10).Interior.ColorIndex = 3
                    End If
                Next J
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the Brand Total
                Total_Volume = 0
                
                opening_price = ws.Cells(i + 1, 3).Value
            ' If the cell immediately following a row is the same brand...
            Else
                ' Add to the Total
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        'worksheetfunction max, worksheetfunction match
        Dim maxPercent_change As Double, MinPercent_change As Double, maxStockVolume As LongLong
        Dim maxPercent_changeIndex, minPercent_changeIndex, maxStock_Index As Long
        
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        maxPercent_change = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
        maxPercent_changeIndex = WorksheetFunction.Match(maxPercent_change, ws.Range("K2:K" & lastRow), 0)
        ws.Range("O2").Value = ws.Range("I" & maxPercent_changeIndex + 1).Value
        ws.Range("P2").Value = maxPercent_change
        
        MinPercent_change = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
        minPercent_changeIndex = WorksheetFunction.Match(MinPercent_change, ws.Range("K2:K" & lastRow), 0)
        ws.Range("O3").Value = ws.Range("I" & minPercent_changeIndex + 1).Value
        ws.Range("P3").Value = MinPercent_change
  
maxStockVolume = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
maxstockvolumeIndex = WorksheetFunction.Match(maxStockVolume, ws.Range("L2:L" & lastRow), 0)
maxStock_Index = maxstockvolumeIndex
ws.Range("O4").Value = ws.Range("I" & maxStock_Index + 1).Value
ws.Range("P4").Value = maxStockVolume

  
  
  Next ws
  
  

End Sub

