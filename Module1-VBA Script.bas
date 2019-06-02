Attribute VB_Name = "Module1"
Sub stockshomework()

Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        
        Cells(1, 10).Value = "Ticker"
        Cells(1, 11).Value = "Yearly Change"
        Cells(1, 12).Value = "Percent Change"
        Cells(1, 13).Value = "Total Volume"
        
        
        Dim LastRow As Long
        
        LastRow = ActiveSheet.UsedRange.Rows.Count
        
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim ticker_name As String
        Dim percent_change As Double
        Dim volume As Double
        Dim i As Long
        Dim YCLastRow As Long
        Dim row As Double
        Dim column As Double
        volume = 0
        row = 2
        column = 1
             
        open_price = Cells(2, column + 2).Value
           
        For i = 2 To LastRow
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                ticker_name = Cells(i, column).Value
                Cells(row, column + 9).Value = ticker_name
                close_price = Cells(i, column + 5).Value
                yearly_change = close_price - open_price
                Cells(row, column + 10).Value = yearly_change
                If (open_price = 0 And close_price = 0) Then
                    percent_change = 0
                ElseIf (open_price = 0 And close_price <> 0) Then
                    percent_change = 1
                Else
                    percent_change = yearly_change / open_price
                    Cells(row, column + 11).Value = percent_change
                    Cells(row, column + 11).NumberFormat = "0.00%"
                End If
                
                volume = volume + Cells(i, column + 6)
                Cells(row, column + 12).Value = volume
                
                row = row + 1
                open_price = Cells(i + 1, column + 2)
                volume = 0
                
            Else
                volume = volume + Cells(i, column + 6).Value
            End If
            
        Next i
        
        YCLastRow = ws.Cells(Rows.Count, column + 10).End(xlUp).row
        
        For j = 2 To YCLastRow
            If (Cells(j, column + 10).Value > 0 Or Cells(j, column + 10).Value = 0) Then
                Cells(j, column + 10).Interior.ColorIndex = 10
            ElseIf Cells(j, column + 10).Value < 0 Then
                Cells(j, column + 10).Interior.ColorIndex = 3
            End If
            
            Next j
                 
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        For Z = 2 To YCLastRow
            If Cells(Z, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & YCLastRow)) Then
                Cells(2, 16).Value = Cells(Z, column + 9).Value
                Cells(2, 17).Value = Cells(Z, column + 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, column + 11).Value = Application.WorksheetFunction.Min(ws.Range("L2:L" & YCLastRow)) Then
                Cells(3, 16).Value = Cells(Z, column + 9).Value
                Cells(3, 17).Value = Cells(Z, column + 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, column + 12).Value = Application.WorksheetFunction.Max(ws.Range("M2:M" & YCLastRow)) Then
                Cells(4, 16).Value = Cells(Z, column + 9).Value
                Cells(4, 17).Value = Cells(Z, column + 12).Value
            End If
            
        Next Z
    
    
    
    Next ws



End Sub



