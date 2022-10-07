Attribute VB_Name = "Module1"
Sub ABCStock()

    Dim vol As Double
    Dim LastRow As Double
    Dim RowDisplay As Integer
    Dim tickeropen As Double
    Dim tickerclose As Double
    Dim yearlychange As Double
    
    
    RowDisplay = 2
    
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Total Stock Volume"
    
    tickeropen = Cells(2, 3).Value
    
    
    For i = 2 To LastRow
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
            vol = vol + Cells(i, 7).Value
        
            Cells(RowDisplay, 9) = Cells(i, 1)
            Cells(RowDisplay, 10) = vol
            
            tickerclose = Cells(i, 6).Value
            
            yearlychange = tickerclose - tickeropen
            
            Cells(RowDisplay, 11).Value = tickerclose - tickeropen
        
            Cells(RowDisplay, 12).Value = yearlychange / tickeropen * 100
        
            vol = 0
            
            tickeropen = Cells(i + 1, 3).Value
            
        
            RowDisplay = RowDisplay + 1
        End If
        
            vol = vol + Cells(i, 7).Value
        
        Next i
        
    
        
    Cells(1, 11) = "Yearly Change"
    Cells(1, 12) = "Percentage Change"
    
   
    
    



End Sub

