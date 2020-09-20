Attribute VB_Name = "Module1"
Sub VBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call stockloop
    Next ws
End Sub

Sub stockloop()
    
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Stock Volume"

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

Dim ticker As String
Dim volume As Double
volume = 0
'initialize using the first recorded open price (update in loop)
Dim open_price As Double
open_price = Cells(2, 3).Value
Dim close_price As Double
clos_price = 0
Dim summary_row As Integer
summary_row = 2
lastrow = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To lastrow
    
    'if not equal then add to summary table
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        clos_price = Cells(i, 6).Value
        ticker = Cells(i, 1).Value
        volume = volume + Cells(i, 7).Value
        Range("I" & summary_row).Value = ticker
        'need to avoid dividing by 0
        'if open price is 0, look for the first nonzero open price if there is one
        If open_price = 0 And Cells(i, 3).Value = Cells(i + 1, 3).Value Then
            open_price = Cells(i + 1, 3).Value
        ElseIf open_price = 0 Then
            Range("K" & summary_row).Value = 0
        Else
            Range("K" & summary_row).Value = (clos_price - open_price) / open_price
        End If
        Range("J" & summary_row).Value = clos_price - open_price
        If Range("J" & summary_row).Value < 0 Then
            Range("J" & summary_row).Interior.ColorIndex = 3
        Else
            Range("J" & summary_row).Interior.ColorIndex = 4
        End If
        Range("K" & summary_row).NumberFormat = "0.00%"
        If Range("K" & summary_row).Value < 0 Then
            Range("K" & summary_row).Interior.ColorIndex = 3
        Else
            Range("K" & summary_row).Interior.ColorIndex = 4
        End If
        Range("L" & summary_row).Value = volume
        summary_row = summary_row + 1
        'update open price to reflect the next stock
        open_price = Cells(i + 1, 3).Value
        'set volume back to 0
        volume = 0
    Else
        volume = volume + Cells(i, 7).Value
    End If
Next i

'challenge
'initialize variables
Dim greatest_inc As Double
greatest_inc = Cells(2, 11).Value
Dim greatest_dec As Double
greatest_dec = Cells(2, 11).Value
Dim greatest_vol As Double
greatest_vol = Cells(2, 12).Value
Dim ticker2 As String
ticker2 = Cells(2, 9).Value
Dim ticker3 As String
ticker3 = Cells(2, 9).Value
Dim ticker4 As String
ticker4 = Cells(2, 9).Value
lastrow2 = Cells(Rows.Count, "I").End(xlUp).Row
For i = 2 To lastrow2
    If greatest_inc < Cells(i, 11).Value Then
        greatest_inc = Cells(i, 11).Value
        Range("Q2").Value = greatest_inc
        Range("Q2").NumberFormat = "0.00%"
        ticker2 = Cells(i, 9).Value
        Range("P2").Value = ticker2
    End If
    If greatest_dec > Cells(i, 11).Value Then
        greatest_dec = Cells(i, 11).Value
        Range("Q3").Value = greatest_dec
        Range("Q3").NumberFormat = "0.00%"
        ticker3 = Cells(i, 9).Value
        Range("P3").Value = ticker3
    End If
    If greatest_vol < Cells(i, 12).Value Then
        greatest_vol = Cells(i, 12).Value
        Range("Q4").Value = greatest_vol
        ticker4 = Cells(i, 9).Value
        Range("P4").Value = ticker4
    End If
Next i
Range("I:Q").Columns.AutoFit

Debug.Print ActiveSheet.Name

End Sub



