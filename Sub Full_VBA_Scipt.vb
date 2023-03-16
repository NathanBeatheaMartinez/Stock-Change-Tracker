Sub FullFinalScript()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
 
    For i = 9 To 12
        If i = 9 Then
            ws.Cells(1, i).Value = "Ticker"
        ElseIf i = 10 Then
            ws.Cells(1, i).Value = "Yearly Change"
        ElseIf i = 11 Then
             ws.Cells(1, i).Value = "Percent Change"
        ElseIf i = 12 Then
             ws.Cells(1, i).Value = "Total Stock Volume"
        End If
    Next i
        
    Dim TickerRow As Integer
     TickerRow = 2
      
      Dim InitialValue As Double
      InitialValue = ws.Cells(2, 3).Value
      
      Dim TotalVolume As Double
      TotalVolume = 0
      
      For m = 3 To Cells(Rows.Count, 1).End(xlUp).Row
        If ws.Cells(m, 1).Value = ws.Cells(m - 1, 1).Value Then
             TotalVolume = TotalVolume + ws.Cells(m - 1, 7).Value
             ws.Cells(TickerRow, 12).Value = TotalVolume
        
        ElseIf ws.Cells(m, 1).Value <> ws.Cells(m - 1, 1).Value Then
             ws.Cells(TickerRow, 9).Value = ws.Cells(m - 1, 1).Value
             
             ws.Cells(TickerRow, 10).Value = ws.Cells(m - 1, 6) - InitialValue
             
             ws.Cells(TickerRow, 11).Value = ws.Cells(TickerRow, 10).Value / InitialValue
             
             InitialValue = ws.Cells(m, 3).Value
             TickerRow = TickerRow + 1
             TotalVolume = 0
    
         End If
         
    Next m
    
        For n = 16 To 17
            If n = 16 Then
            ws.Cells(1, n).Value = "Ticker"
        ElseIf n = 17 Then
            ws.Cells(1, n).Value = "Value"
        End If
        Next n
        
         For o = 2 To 4
            If o = 2 Then
            ws.Cells(o, 15).Value = "Greatest % Increase"
        ElseIf o = 3 Then
            ws.Cells(o, 15).Value = "Greatest % Decrease"
        ElseIf o = 4 Then
            ws.Cells(o, 15).Value = "Greatest Total Volume"
        End If
   Next o
        
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = ws.Cells(2, 12).Value
    
    Dim GreatestIncrease As Double
    GreatestIncrease = ws.Cells(2, 11).Value
    
    Dim GreatestDecrease As Double
    GreatestDecrease = ws.Cells(2, 11).Value
    
    
    For p = 3 To 5005
    
        If GreatestTotalVolume > ws.Cells(p, 12).Value Then
            ws.Cells(4, 17).Value = GreatestTotalVolume
        
        ElseIf GreatestTotalVolume < ws.Cells(p, 12).Value Then
             ws.Cells(4, 17).Value = ws.Cells(p, 12).Value
             ws.Cells(4, 16).Value = ws.Cells(p, 9).Value
             
             GreatestTotalVolume = ws.Cells(p, 12).Value
        End If
    Next p
        
        For q = 3 To 5005
        If GreatestIncrease > ws.Cells(q, 11).Value Then
            ws.Cells(2, 17).Value = GreatestIncrease
        
        ElseIf GreatestIncrease < ws.Cells(q, 11).Value Then
             ws.Cells(2, 17).Value = ws.Cells(q, 11).Value
             ws.Cells(2, 16).Value = ws.Cells(q, 9).Value
             
             GreatestIncrease = ws.Cells(q, 11).Value
        End If
    Next q
         For r = 3 To 5005
        If GreatestDecrease < ws.Cells(r, 11).Value Then
            ws.Cells(3, 17).Value = GreatestDecrease
        
        ElseIf GreatestDecrease > ws.Cells(r, 11).Value Then
             ws.Cells(3, 17).Value = ws.Cells(r, 11).Value
             ws.Cells(3, 16).Value = ws.Cells(r, 9).Value
             
             GreatestDecrease = ws.Cells(r, 11).Value
        End If
    Next r
Next ws
End Sub
