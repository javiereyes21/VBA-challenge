VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub vba_project()

    'dimensions
    Dim ws As Worksheet
    Dim total As Double
    Dim ri As Long
    Dim change As Double
    Dim ci As Integer
    Dim start As Long
    Dim RowCount As Long
    Dim percent_change As Double
    Dim days As Integer
    Dim daily_change As Single
    Dim avg_change As Double

    For Each ws In Worksheets
        ci = 0
        total = 0
        change = 0
        start = 2
        daily_change = 0
        
        ws.Range("j1").Value = "Ticker"
        ws.Range("k1").Value = "Quarterly Change"
        ws.Range("l1").Value = "% Change"
        ws.Range("m1").Value = "Total Stock Volume"
        ws.Range("q1").Value = "Value"
        ws.Range("r1").Value = "Ticker"
        ws.Range("p2").Value = "Greatest % Increase"
        ws.Range("p3").Value = "Greatest % Decrease"
        ws.Range("p4").Value = "Greatest Total Volume"
        
        'Getting row number of the last row with data
        RowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        For ri = 2 To RowCount
            If ws.Cells(ri + 1, 1).Value <> ws.Cells(ri, 1).Value Then
                'variable to store results
                total = total + ws.Cells(ri, 7).Value
                
                If total = 0 Then
                    ws.Range("j" & 2 + ci).Value = ws.Cells(ri, 1).Value
                    ws.Range("k" & 2 + ci).Value = 0
                    ws.Range("l" & 2 + ci).Value = "%" & 0
                    ws.Range("m" & 2 + ci).Value = 0
                Else
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To ri
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    
                    change = (ws.Cells(ri, 6).Value - ws.Cells(start, 3).Value)
                    percent_change = change / ws.Cells(start, 3).Value
                    
                    start = ri + 1
                    
                    ws.Range("j" & 2 + ci).Value = ws.Cells(ri, 1).Value
                    ws.Range("k" & 2 + ci).Value = change
                    ws.Range("k" & 2 + ci).NumberFormat = "0.00"
                    ws.Range("l" & 2 + ci).Value = percent_change
                    ws.Range("l" & 2 + ci).NumberFormat = "0.00%"
                    ws.Range("m" & 2 + ci).Value = total
                    
                    Select Case change
                        Case Is > 0
                            ws.Range("k" & 2 + ci).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("k" & 2 + ci).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("k" & 2 + ci).Interior.ColorIndex = 0
                    End Select
                End If
                
                total = 0
                change = 0
                ci = ci + 1
                days = 0
                daily_change = 0
            Else
                'if ticker is the same, add results
                total = total + ws.Cells(ri, 7).Value
            End If
        Next ri
        
        'MAX and MIN
        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("l2:l" & RowCount)) * 100
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("l2:l" & RowCount)) * 100
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("m2:m" & RowCount))
        
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("l2:l" & RowCount)), ws.Range("l2:l" & RowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("l2:l" & RowCount)), ws.Range("l2:l" & RowCount), 0)
        
        ws.Range("r2").Value = ws.Cells(increase_number + 1, 10).Value
        ws.Range("r3").Value = ws.Cells(decrease_number + 1, 10).Value
        ws.Range("r4").Value = WorksheetFunction.Max(ws.Range("m2:m" & RowCount))
        
        MsgBox ("Next ws")
    Next ws

End Sub

