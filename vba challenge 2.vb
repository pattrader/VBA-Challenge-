Sub routine_2()

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"

Dim lastrow As Variant
Dim i As Variant
Dim ticker As String
Dim j As Variant
Dim vol As Variant
Dim vol_total As Variant
Dim pct_chg As Variant
Dim open_price As Variant
Dim close_price As Variant





lastrow = Cells(Rows.Count, 3).End(xlUp).Row

vol_total = 0

For i = 2 To lastrow
    dollar_chg = Cells(i, 6) - Cells(i, 3)


    If Cells(i, 1) <> Cells(i - 1, 1) Then
    j = 1
    Count = Count + j
    Cells(1 + Count, 9) = Cells(i, 1)
    vol_total = Cells(i, 7)
    
    open_price = Cells(i, 3)
    
    
    
    Else
    vol = Cells(i, 7)
    vol_total = vol_total + vol
    Cells(1 + Count, 12) = vol_total
    
    close_price = Cells(i, 6)
    
    dollar_chg = close_price - open_price
 
    Cells(1 + Count, 10) = dollar_chg
    
    pctchg = (close_price - open_price) / open_price
    Cells(1 + Count, 11) = pctchg
    
        
    End If
    
Next i

For j = 2 To lastrow

If Cells(j, 10) > 0 Then
    Cells(j, 10).Interior.ColorIndex = 4
    
Else
    Cells(j, 10).Interior.ColorIndex = 3
    
End If

Next j

Columns("K:K").Select
    Selection.Style = "Percent"


Dim k As Variant
Dim low As Variant
Dim value As Variant



Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = " Greatest Total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"

lastrow = Cells(Rows.Count, 11).End(xlUp).Row

For k = 2 To lastrow
    If Cells(k, 11) > high Then
        high = Cells(k, 11)
        hticker = Cells(k, 9)
        Range("P2") = hticker
        Range("q2") = high
    ElseIf Cells(k, 11) < low Then
        low = Cells(k, 11)
        lticker = Cells(k, 9)
        Range("P3") = lticker
        Range("Q3") = low
    ElseIf Cells(k, 12) > vol Then
        vol = Cells(k, 12)
        vticker = Cells(k, 9)
        Range("P4") = vticker
        Range("Q4") = vol
    End If
            
   
    
Next k


Range("Q2").value = FormatPercent(Range("Q2"))
Range("Q3").value = FormatPercent(Range("Q3"))


End Sub






