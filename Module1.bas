Attribute VB_Name = "Module1"
Sub StockVBA()

'Create variables

Dim ticker As String
Dim next_ticker As Integer
Dim totalRow As Long
Dim openprice As Double
Dim closeprice As Double
Dim yearlychange As Double
Dim percent_change As Double
Dim total_stock_volume As Double



' Loop through all sheets
For Each ws In Worksheets

    ' Make the worksheet active.
    ws.Activate

    ' Find the last row of each worksheet
    totalRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Additional Headers for each worksheet
    ws.Cells(9, 1).Value = "Ticker"
    ws.Cells(10, 1).Value = "Yearly Change"
    ws.Cells(11, 1).Value = "Percent Change"
    ws.Cells(12, 1).Value = "Total Stock Volume"
    
    ' Reset variables for each worksheet.
    ticker = ""
    next_ticker = 0
    yearlychange = 0
    openprice = 0
    percent_change = 0
    total_stock_volume = 0
    
    ' Create For loop for ticker values
    For rowNum = 2 To totalRow

        ' Define ticker variable
        ticker = Cells(rowNum, 1).Value
        
        ' Figure out open price for ticker
        If openprice = 0 Then
            openprice = Cells(rowNum, 3).Value
        End If
        
        ' Calculate total stock volume
        total_stock_volume = total_stock_volume + Cells(rowNum, 7).Value
        
        ' Figure out when ticker value changes
        If Cells(rowNum + 1, 1).Value <> ticker Then
            next_ticker = next_ticker + 1
            Cells(next_ticker + 1, 9) = ticker
            
            ' Get the end of the year closing price for ticker
            closeprice = Cells(rowNum, 6)
            
            ' Get yearly change value
            yearlychange = closeprice - openprice
            
            ' Add yearly change value to the appropriate cell in each worksheet.
            Cells(next_ticker + 1, 10).Value = yearlychange
            
            'Color cell reference:
            'Green = 4
            'Red = 3
            'Yellow = 6
            
            ' If yearly change value is >, then green cell.
            If yearlychange > 0 Then
                Cells(next_ticker + 1, 10).Interior.ColorIndex = 4
            ' If yearly change value is < 0, then red cell.
            ElseIf yearlychange < 0 Then
                Cells(next_ticker + 1, 10).Interior.ColorIndex = 3
            ' If yearly change value = 0, then yellow cell.
            Else
                Cells(next_ticker + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Header: Percent Change Calculate percent change value
            'Establish number % change
            If openprice = 0 Then
                percent_change = 0
            'Calculation for percent change
            Else
                percent_change = (yearlychange / openprice)
            End If
        
            
            ' Convert percent_change value to a percentage
            Cells(next_ticker + 1, 11).Value = Format(percent_change, "Percent")
       
            
            ' Reset Open when ticker value changes
            openprice = 0
            
            'Header: Total Stock Volume
            'Add total stock volume value to the appropriate cell in each worksheet.
            Cells(next_ticker + 1, 12).Value = total_stock_volume
            
            ' Reset total stock value to 0 when ticker value changes
            total_stock_volume = 0
        End If
        
    Next rowNum
    
  
    
    
    
   
    
Next ws


End Sub

