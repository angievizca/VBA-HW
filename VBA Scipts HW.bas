Attribute VB_Name = "Module1"
Sub VBAHomework()
    For Each Worksheet In ActiveWorkbook.Worksheets
    Worksheet.Activate
    
Dim TickerRow As Integer
Dim YearChange As Double
Dim PercentChange As Double
Dim Total_Stock_Volume As Double

Dim lastrow As Long
Dim TickerName As String
Dim OpenPrice As Double
Dim ClosePrice As Double


TickerRow = 2
lastrow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
Total_Stock_Volume = 0

    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
    
    
OpenPrice = Cells(2, 3).Value
For i = 2 To lastrow



If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    TickerName = Cells(i, 1).Value
    
    ClosePrice = Cells(i, 6).Value
    YearChange = ClosePrice - OpenPrice

    Range("I" & TickerRow).Value = TickerName
    
    Range("J" & TickerRow).Value = YearChange
    
    Worksheet.Cells(TickerRow, 10).NumberFormat = "$0.00"
    Worksheet.Cells(TickerRow, 11).NumberFormat = "0.00%"
    Worksheet.Cells(2, 17).NumberFormat = "0.00%"
    Worksheet.Cells(3, 17).NumberFormat = "0.00%"
    
    If (OpenPrice = 0 And ClosePrice = 0) Then
    PercentChange = 0
    
    ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
        PercentChange = 1
    Else
        PercentChange = YearChange / OpenPrice
        Cells(TickerRow, 11).Value = PercentChange
    End If
    
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    Range("L" & TickerRow).Value = Total_Stock_Volume
        
    
    TickerRow = TickerRow + 1
    OpenPrice = Cells(i + 1, 3)
    Total_Stock_Volume = 0


Else
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    


End If

Next i

    LastRow1 = Worksheet.Cells(Rows.Count, 9).End(xlUp).Row
    
    For j = 2 To LastRow1
    
    If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
        Cells(j, 10).Interior.ColorIndex = 4
    ElseIf Cells(j, 10).Value < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
        
    End If
    Next j
    
    'Set Geatest % Increase
    Cells(2, 15).Value = "Greates % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    
    For x = 2 To LastRow1
        If Cells(x, 11).Value = Application.WorksheetFunction.Max(Worksheet.Range("k2:K" & LastRow1)) Then
            Cells(2, 16).Value = Cells(x, 9).Value
            Cells(2, 17).Value = Cells(x, 11).Value
        ElseIf Cells(x, 11).Value = Application.WorksheetFunction.Min(Worksheet.Range("K2:K" & LastRow1)) Then
        Cells(3, 16).Value = Cells(x, 9).Value
        Cells(3, 17).Value = Cells(x, 11).Value
        ElseIf Cells(x, 12).Value = Application.WorksheetFunction.Max(Worksheet.Range("L2:L" & LastRow1)) Then
        Cells(4, 16).Value = Cells(x, 9).Value
        Cells(4, 17).Value = Cells(x, 12).Value
        End If
    Next x
        
Next Worksheet

End Sub











