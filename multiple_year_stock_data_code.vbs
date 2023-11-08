Sub multiple_year_stock_data()

    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Summary_Table_Row As Integer
    Dim Total_Stock As LongLong
    Dim Open_Price_Row As Long
    Dim Yearly_Change_Column As Integer
    Dim Percent_Change_Column As Integer
    Dim Total_Stock_Column As Integer
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume As LongLong

For Each ws In Worksheets

    ws.Cells.EntireColumn.AutoFit
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Cells.EntireColumn.AutoFit
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Open_Price = 0
    Close_Price = 0
    Yearly_Change = 0
    Total_Stock = 0
    Percent_Change = 0
    Summary_Table_Row = 2
    Open_Price_Row = 2
    Yearly_Change_Column = 10
    Percent_Change_Column = 11
    Total_Stock_Column = 12
    Greatest_Percent_Increase = 0
    Greatest_Percent_Decrease = 0
    Greatest_Total_Volume = 0
    

For i = 2 To LastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Ticker = ws.Cells(i, 1).Value
        Open_Price = ws.Cells(Open_Price_Row, 3).Value
        Closed_Price = ws.Cells(i, 6).Value
        Total_Stock = Total_Stock + ws.Cells(i, 7).Value
        Yearly_Change = Closed_Price - Open_Price
        Percent_Change = (Yearly_Change / Open_Price)
        Open_Price_Row = i + 1
            
            
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock
            
        ws.Range("K" & Summary_Table_Row).Style = "Percent"
            
        Summary_Table_Row = Summary_Table_Row + 1
            
        Open_Price = 0
        Closed_Price = 0
        Total_Stock = 0
        Percent_Change = 0
        Yearly_Change = 0
            
    Else
            
        Open_Price = Open_Price + ws.Cells(i, 3).Value
        Closed_Price = Closed_Price + ws.Cells(i, 6).Value
        Total_Stock = Total_Stock + ws.Cells(i, 7).Value
        Yearly_Change = Closed_Price - Open_Price
        Percent_Change = (Yearly_Change / Open_Price)
            
            
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock
            
            
    End If
          
Next i

          
For i = 2 To LastRow

    If ws.Cells(i, Yearly_Change_Column) > 0 Then
    ws.Cells(i, Yearly_Change_Column).Interior.ColorIndex = 4
        
    ElseIf ws.Cells(i, Yearly_Change_Column) < 0 Then
    ws.Cells(i, Yearly_Change_Column).Interior.ColorIndex = 3
        
    End If
        
  
Next i

For i = 2 To LastRow

    If ws.Cells(i, Percent_Change_Column) >= Greatest_Percent_Increase Then
    Greatest_Percent_Increase = ws.Cells(i, Percent_Change_Column).Value
    Ticker = ws.Cells(i, 9).Value
    ws.Range("P2").Value = Ticker
    ws.Range("Q2").Value = Greatest_Percent_Increase

    End If
    
Next i

For i = 2 To LastRow

    If ws.Cells(i, Percent_Change_Column) <= Greatest_Percent_Decrease Then
    Greatest_Percent_Decrease = ws.Cells(i, Percent_Change_Column).Value
    Ticker = ws.Cells(i, 9).Value
    ws.Range("P3").Value = Ticker
    ws.Range("Q3").Value = Greatest_Percent_Decrease
    
    End If
    
Next i

For i = 2 To LastRow

    If ws.Cells(i, Total_Stock_Column) >= Greatest_Total_Volume Then
    Greatest_Total_Volume = ws.Cells(i, Total_Stock_Column).Value
    Ticker = ws.Cells(i, 9).Value
    ws.Range("P4").Value = Ticker
    ws.Range("Q4").Value = Greatest_Total_Volume
    
    End If

Next i

    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
Next ws
    
End Sub
