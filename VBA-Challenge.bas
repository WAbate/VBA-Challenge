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
    
    
'Condense Titckers to single cell'
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    Cells(output, 9).Value = Cells(i, 1).Value
    
'Yearly change'
    Change = Cells(i, 6) - Cells(Start, 3)
    Percent = Cells(output, 10).Value / Cells(Start, 3).Value
    
'Yearly Change in Percent'
    Cells(output, 11).Value = Percent
    Cells(output, 11).NumberFormat = "00.00%"
    Start = i + 1
    
'Output Volume'
    Cells(output, 12).Value = Total
    output = output + 1
    Total = 0
Else
    Total = Total + Cells(i, 7).Value
    
End If
        Next i
        For i = 2 To 290
                If (Cells(i, 10).Value > 0) Then
                Cells(i, 10).Interior.ColorIndex = 4
                Else
                    Cells(i, 10).Interior.ColorIndex = 3
                    
        End If
        Next i
        
                
        
End Sub



