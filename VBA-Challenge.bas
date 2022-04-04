Attribute VB_Name = "Module1"
'Write a loop that will:'
'output the ticker'
'the yearly change(Beg open to End close)'
'the percent change("ditto")'
'the total stock volume of stock'
'conditional formatting that will highlight pos/neg change'

Sub VBA_Challenge()
'Column Labels
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
'P1: Sort throught the tickers - put them in "Ticker" coulmn'

    Dim Total As Double
    Dim Change As Double
    Dim Percent As Double
    Dim Start As Long
    Dim Finish As Long
    
    
    lastrow = Cells(Rows.Count, 1).End(x1Up).Row
    Total = 0
    output = 2
    Start = 2
    
    For i = 2 To lastrow
    
'Condense Titckers to single cell'
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    Cells(output, 9).Value -Cells(i, 1).Value
    
    
    
    
    


End Sub



