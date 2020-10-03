Attribute VB_Name = "Module1"
Sub StockData():
    'Declare variables to assign ticker, total stock volume,yearly change in price, percentage change in yearly price, open price, close price, greatest % increase, greatest % decrease, greatest volume, number of rows for the main table, summary table row, starting row and the new starting row for the open price
    Dim Total_Stock_Volume As Double
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percentage_change As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim Row_Number As Long
    Dim Summary_Table_Row As Integer
    Dim Starting_Row As Double
    Dim New_Starting_Row As Double
    Dim i As Long
    Dim Second_Table_Row As Long
    Dim Third_Table_Row As Long
    Dim j As Long
    Dim New_Table_Ticker As String
    Dim Greatest_Inc As Double
    Dim Greatest_Dec As Double
    Dim Greatest_Vol As Double
    'Iterate all the worksheets in the document
    For Each ws In Worksheets
        'Set rows in the summary table and the starting row in the initial table to 2
        Summary_Table_Row = 2
        Starting_Row = 2
        'Set the initial value of the number of rows in the third table to 2
        'Third_Table_Row = 2
        'Retrieve the number of rows in each worksheet
        Row_Number = ws.Cells(Rows.Count, "A").End(xlUp).Row
        'open_price = ws.Cells(2, 3).Value
        'Iterate every row in each worksheet to retrieve their summary tickers, total stock volume, open price and close price
        For i = 2 To Row_Number
            'Check if the tickers are different
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'If they are different, set the first summary ticker
                Ticker = ws.Cells(i, 1).Value
                'Compute and update the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                'Set the close price
                close_price = ws.Cells(i, 6).Value
                'Check if the open price is equal to 0. This will help to bypass division by zero in the computation of the percentage change in price
                If open_price = 0 Then
                    For New_Starting_Row = Starting_Row To i
                        If ws.Cells(New_Starting_Row, 3).Value <> 0 Then
                            Starting_Row = New_Starting_Row
                            Exit For
                        End If
                    Next New_Starting_Row
                End If
                'Set the next open price
                open_price = Cells(Starting_Row, 3).Value
                'Compute the yearly change as the difference between the open price and the close price
                Yearly_Change = close_price - open_price
                'Compute the percentage change in price
                Percentage_change = Round(Yearly_Change / open_price, 2)
                'Update the starting row
                Starting_Row = i + 1
                'Print the ticker, total stock volume, close price, open price, yearly change and percentage change on the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                ws.Range("M" & Summary_Table_Row).Value = close_price
                ws.Range("N" & Summary_Table_Row).Value = open_price
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percentage_change
                'Color conditioning for the yearly change column
                If ws.Cells(Summary_Table_Row, 10).Value < 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                End If
                'Add one to the summary table
                Summary_Table_Row = Summary_Table_Row + 1
                'Reset the Total_Stock_Volume
                Total_Stock_Volume = 0
            Else
                'Add to the Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            End If
        Next i
        'Set the column headers for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("M1").Value = "Close Price"
        ws.Range("N1").Value = "Open Price"
        'ws.Range("O1").Value = "Worksheet Name"
        'Print out greatest % increase, greatest % decrease and greatest total volume on another table
        'Find the last row of the summary table
        Second_Table_Row = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        'Retrieve the location of the greatest % increase, greatest % decrease and the greatest total volume from the summary table
        Greatest_Inc = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Second_Table_Row)), ws.Range("K2:K" & Second_Table_Row), 0)
        Greatest_Dec = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Second_Table_Row)), ws.Range("K2:K" & Second_Table_Row), 0)
        Greatest_Vol = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Second_Table_Row)), ws.Range("L2:L" & Second_Table_Row))
        
        'Print the column header for the third table
        ws.Range("Q2").Value = "Greatest % Increase"
        ws.Range("Q3").Value = "Greatest % Decrease"
        ws.Range("Q4").Value = "Greatest Total Volume"
        ws.Range("R1").Value = "Ticker"
        ws.Range("S1").Value = "Value"
        
        'Retrieve the ticker with the greatest % increase, greatest % decrease and the greatest total volume from the summary table and print on the third table
        ws.Range("R2").Value = ws.Cells(Greatest_Inc + 1, 9).Value
        ws.Range("R3").Value = ws.Cells(Greatest_Dec + 1, 9).Value
        ws.Range("R4").Value = ws.Cells(Greatest_Vol + 1, 9).Value
        
        'Retrieve the value of the greatest % increase, greatest % decrease and the greatest total volume from the summary table and print on the third table
        ws.Range("S2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & Second_Table_Row)) * 100
        ws.Range("S3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & Second_Table_Row)) * 100
        ws.Range("S4").Value = WorksheetFunction.Max(ws.Range("L2:L" & Second_Table_Row))
        
    Next ws
End Sub
