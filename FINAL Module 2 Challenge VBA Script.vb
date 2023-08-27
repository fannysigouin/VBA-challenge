Sub Stock_Data()
    Dim ws As Worksheet
    Dim Ticker_Symbol As String
    Dim LastRow As Long
    Dim Start_Open_Price As Double
    Dim Open_Price_Captured As Boolean
    Dim End_Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Variant
    Dim Summary_Table_Row As Integer
    
    
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        Open_Price_Captured = False
            For i = 2 To LastRow
                If Open_Price_Captured = False Then
                    Start_Open_Price = ws.Cells(i, 3).Value
                    Open_Price_Captured = True
                End If
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    Ticker_Symbol = ws.Cells(i, 1).Value
                    ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol
                
                    End_Close_Price = ws.Cells(i, 6).Value
                    Yearly_Change = End_Close_Price - Start_Open_Price
                    ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
                        If Yearly_Change < 0 Then
                            ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                        Else
                            ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                        End If
                
                    Percent_Change = (End_Close_Price - Start_Open_Price) / Start_Open_Price
                    ws.Range("L" & Summary_Table_Row).Value = Percent_Change
                    ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    ws.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
                
                    Summary_Table_Row = Summary_Table_Row + 1
                
                    End_Close_Price = 0
                    Yearly_Change = 0
                    Percent_Change = 0
                    Total_Stock_Volume = 0
                    Open_Price_Captured = False
                Else
                    End_Close_Price = ws.Cells(i, 6).Value
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                End If
         Next i
    Next ws
    
    Dim Summary_Last_Row As Long
    Dim Max_Increase As Single
    Dim Max_Decrease As Single
    Dim Max_Total As Variant
    
    For Each ws In Worksheets
        Summary_Last_Row = ws.Cells(Rows.Count, 10).End(xlUp).Row
        Max_Increase = WorksheetFunction.Max(ws.Range("L:L"))
        Max_Decrease = WorksheetFunction.Min(ws.Range("L:L"))
        Max_Total = WorksheetFunction.Max(ws.Range("M:M"))
        For i = 2 To Summary_Last_Row
            For j = 12 To 12
                If ws.Cells(i, j).Value = Max_Increase Then
                    ws.Cells(2, 18).Value = Max_Increase * 100
                    ws.Cells(2, 17).Value = ws.Cells(i, j - 2).Value
                ElseIf ws.Cells(i, j).Value = Max_Decrease Then
                    ws.Cells(3, 18).Value = Max_Decrease * 100
                    ws.Cells(3, 17).Value = ws.Cells(i, j - 2).Value
                End If
            Next j
            For j = 13 To 13
                If ws.Cells(i, j).Value = Max_Total Then
                    ws.Cells(4, 18).Value = Max_Total
                    ws.Cells(4, 17).Value = ws.Cells(i, j - 3).Value
                End If
            Next j
        Next i
    Next ws
End Sub