Attribute VB_Name = "Module11"

'A Script that loops through all the stocks for 1 year and
'OUTPUTS THE FOLLOWING:
'Ticker Symbol
'Yearly Change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock
'Use conditional formatting that will highlight positive change in green and negative change in red.

Sub MultipleYearStock():

'loop through all the worksheets (2018,2019 and 2020),(included in the BONUS PART)
For Each ws In Worksheets

         
    'Add the headers to the cells
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'variable to hold the ticker
    ticker = ""
    ' variable to hold the total stock volume of the ticker
    totalStockVolume = 0
    ' variable to hold the summary table starter row
    summaryTableRow = 2
    
    
    'variables to hold Yearly Change, Percent Change, First Opening Price and Last Closing Price
    Dim yearlyChange As Double

    FirstOpenPrice = Cells(2, 3).Value
    LastClosePrice = 0
    yearlyChange = 0
    
    Dim percentChange As Double
    percentChange = 0
    
    

    'Determine the Last Row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' loop from row 2 in column A out to the last row
For Row = 2 To lastRow
        ' check to see if the brand changes
 If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                
                ' if the ticker changes, do the following
                ' first set the ticker symbol
               ticker = Cells(Row, 1).Value
               ' add the ticker to the I column in the summary table row
                ws.Cells(summaryTableRow, 9).Value = ticker
                
                
                ' add the last total stock volume from the row
                totalStockVolume = totalStockVolume + ws.Cells(Row, 7).Value
                ' add the total stock volume to the L column in the summary table row
                ws.Cells(summaryTableRow, 12).Value = totalStockVolume
              
              
              'calculate the yearly change = Last closing price -First opening Price
               LastClosePrice = ws.Cells(Row, 6).Value
               yearlyChange = LastClosePrice - FirstOpenPrice
                'add the yearly change to the J column in the summary table row
                ws.Cells(summaryTableRow, 10).Value = yearlyChange
              'Change the format of Yearly Change to currency
                ws.Cells(summaryTableRow, 10).Style = "Currency"
                
               
               
               'Calculate the Percent Change
               percentChange = yearlyChange / FirstOpenPrice
               
                'add the percent change to the K column in the summary table row
                ws.Cells(summaryTableRow, 11).Value = percentChange
                'Change the format of percent change to percentage value
                ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
                
                            
             'Use conditional formatting to color the positive and negative change
           If ws.Cells(summaryTableRow, 10).Value > 0 Then
        'highlights positive change in green
               ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
     
           ElseIf ws.Cells(summaryTableRow, 10).Value < 0 Then
        'highlights negative change in red
               ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
              
           Else
               'if there is no change, color the cell white
               ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 2
                
          End If
                ' go to the next summary table row (add 1 on to the value of the summary table row)
                summaryTableRow = summaryTableRow + 1
                ' reset the totalStockVolume to 0
                totalStockVolume = 0
                
                ' set the first opening price for each different ticker
               FirstOpenPrice = ws.Cells(Row + 1, 3).Value

        
                
    Else
                
                ' if the ticker stays the same, do the following:
                ' add on to the total stock volume from the G column
                totalStockVolume = totalStockVolume + ws.Cells(Row, 7).Value
            
            
            
 End If
                     

            
Next Row
    
      'Autofit the columns to accomodate the size of the values in the summary table
       ws.Range("I:L").Columns.AutoFit
    
 
 'THE BONUS PART:
      
      'Declare variables for the maximum and minimum values in % change and maximum value of total stock volume
       Dim maxPercent_change As Double, minPercent_change As Double, maxStockVolume As LongLong
       Dim maxPercent_changeIndex, minPercent_changeIndex, maxStock_Index As Long

        
      'Add the headers to the columns and rows of the bonus table
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"

        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        
        'find the maximum value of percent change
            maxPercent_change = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
        'find the index of the maximum value in percent change
            maxPercent_changeIndex = WorksheetFunction.Match(maxPercent_change, ws.Range("K2:K" & lastRow), 0)
         'add/print the ticker symbol of the index related to the maximum value of the percent change
            ws.Range("O2").Value = ws.Range("I" & maxPercent_changeIndex + 1).Value
         'add/print the maximum value/greatest % increase value of percent change
            ws.Range("P2").Value = maxPercent_change
            
            
        'find the minimun value of percent change
            minPercent_change = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
        'find the index of the minimum value in percent change
            minPercent_changeIndex = WorksheetFunction.Match(minPercent_change, ws.Range("K2:K" & lastRow), 0)
        'add/print the ticker symbol of the index related to the minimum value of the percent change
            ws.Range("O3").Value = ws.Range("I" & minPercent_changeIndex + 1).Value
        'add/print the minimum value/greatest % decrease value of percent change
            ws.Range("P3").Value = minPercent_change
            
        
        
        'find the maximum value of Total Stock Volume
            maxStockVolume = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
        'find the index of the maximum total stock value
            maxStock_Index = WorksheetFunction.Match(maxStockVolume, ws.Range("L2:L" & lastRow), 0)
        'add the ticker symbol of the maximum total stock value index
            ws.Range("O4").Value = ws.Range("I" & maxStock_Index + 1).Value
        'add/print the maximum value of the total stock volume
            ws.Range("P4").Value = maxStockVolume
        
        
       'Change the formats of the maximum, minimum and maximum total stock volume
            ws.Range("P2").NumberFormat = "000.00%"
            ws.Range("P3").NumberFormat = "00.00%"
            ws.Range("P4").NumberFormat = "0.00E+0"
        'Autofit the columns to accomodate the size of the values in the bonus table
            ws.Range("N:P").Columns.AutoFit
            
            
    
            
Next ws
  


End Sub






