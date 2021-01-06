Sub stocks()

'Set Variables
    Dim volume_total As Double
    Dim column As Integer
    Dim ticker As Integer
    Dim total_row As Integer
    Dim opening_price As Double
    Dim closing_price As Double
    
'Variables/ticker cell
    volume_total = 0
    Range("I2").Value = Cells(2, 1).Value
    column = 1
    Row = 2
    total_row = 2
    opening_price = Range("C2").Value
    closing_price = 0
    lastrow_Ticker = Cells(Rows.Count, 9).End(xlUp).Row
    Lastrow_YearlyChg = Cells(Rows.Count, 10).End(xlUp).Row
    Lastrow_PercentageChg = Cells(Rows.Count, 11).End(xlUp).Row
    Lastrow_Totalvol = Cells(Rows.Count, 12).End(xlUp).Row
        
'Column Names
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "PercentageChg"
    Range("L1").Value = "Total Stock Vol"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Inc"
    Range("N3").Value = "Greatest % Dec"
    Range("N4").Value = "Greatest Total Volume"
    
'Loop
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
       
            
'New ticker classification
    If Cells(i, column).Value <> Cells(i + 1, column).Value Then
        
       'Plugs in ticker value into column 9
        Cells(lastrow_Ticker + 1, 9).Value = Cells(i + 1, column).Value
        
        'Resets last row in Ticker in column 9
        lastrow_Ticker = Cells(Rows.Count, 9).End(xlUp).Row
        closing_price = Cells(i, 6).Value
        
        Lastrow_YearlyChg = Cells(Rows.Count, 10).End(xlUp).Row
        
        'Plugs in yearly change into column 10
        Cells(Lastrow_YearlyChg + 1, 10).Value = Format(closing_price - opening_price, ".00")
        
        
        'Calcuate percent change in opening and closing price
        If opening_price > 0 Then
        Cells(Lastrow_PercentageChg + 1, 11).Value = (closing_price - opening_price) / opening_price
        Lastrow_PercentageChg = Cells(Rows.Count, 11).End(xlUp).Row
        
        Else
        Cells(Lastrow_PercentageChg + 1, 11).Value = 0
        Lastrow_PercentageChg = Cells(Rows.Count, 11).End(xlUp).Row
        
        End If
        
        'Sets new opening price for next ticker
        opening_price = Cells(i + 1, 3).Value
            
        
        'Plugs in Total Stock Volume in column 12
        Cells(Lastrow_Totalvol + 1, 12).Value = volume_total + Cells(i, 7).Value
        Lastrow_Totalvol = Cells(Rows.Count, 12).End(xlUp).Row
            
        volume_total = 0
        
        Else
        'Adding total stock volume within same ticker
        volume_total = Cells(i, 7).Value + volume_total
     
    
    End If

    Next i
    
'Bonus!

Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("2016")
Dim MyMax As Double, maxcell As Range

'Find the Max in Percent Chg

MyMax = Application.WorksheetFunction.Max(ws.Range("K:K"))
Set maxcell = ws.Range("K:K").Find(MyMax, Lookat:=xlWhole)

ws.Range("P2") = MyMax

Range("P2").NumberFormat = "0.00%"


'Find the Min value in Percentage Chg

Dim MyMin As Double, mincell As Range

MyMin = Application.WorksheetFunction.Min(ws.Range("K:K"))
Set mincell = ws.Range("K:K").Find(MyMin, Lookat:=xlWhole)

ws.Range("P3") = MyMin

Range("P3").NumberFormat = "0.00%"

'Find the Greatest Total Volume
MyMax = Application.WorksheetFunction.Max(ws.Range("L:L"))
Set maxcell = ws.Range("L:L").Find(MyMax, Lookat:=xlWhole)

ws.Range("P4") = MyMax
Range("P4").NumberFormat = "0"

'Return tickers for each value
increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K:K" & RowCount)), ws.Range("K:K" & RowCount), 0)
ws.Range("O2") = ws.Cells(increase_number + 1, 9)

decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K:K" & RowCount)), ws.Range("K:K" & RowCount), 0)
ws.Range("O3") = ws.Cells(decrease_number + 1, 9)

total_volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L:L" & RowCount)), ws.Range("L:L" & RowCount), 0)
ws.Range("O4") = ws.Cells(total_volume + 1, 9)

'If yearly change is <0, then format cell to turn red

For i = 2 To Lastrow_YearlyChg + 1

    If Cells(i, 10) < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
    
    '>0, then turn green
    Else
    Cells(i, 10).Interior.ColorIndex = 4

End If

Next i

'Format Percentage Chg to %

Range("K:K").NumberFormat = "0.00%"

'Format columns to expand
Range("J:L").Columns.AutoFit


End Sub
 






