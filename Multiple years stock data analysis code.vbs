Attribute VB_Name = "Module1"
Sub stock_ticker()

    ' Loop all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        ' Set up the table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Declare variables for ticker & total volume calculations
        Dim Ticker_name As String
        Dim Stock_Total As Double
        Dim RowCount As Long
        Dim Summary_Ticker_Row As Long
        Dim i As Long
        
        Stock_Total = 0
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        Summary_Ticker_Row = 2
        
        Dim StartRow As Long ' Variable to store the starting row of each ticker
        StartRow = 2 ' Initialize the starting row
        
        ' Iterate through rows
        For i = 2 To RowCount
            ' If the next ticker in the row is not the same, print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_name = ws.Cells(i, 1).Value
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                
                ' Print ticker name and stock total to the summary table
                ws.Cells(Summary_Ticker_Row, 9).Value = Ticker_name
                ws.Cells(Summary_Ticker_Row, 12).Value = Stock_Total
                
                ' Calculate yearly change and percentage change for each ticker
                Dim YearlyChange As Double
                Dim PercentageChange As String
                Dim StartPrice As Double
                Dim EndPrice As Double
                ' Dim roundedValue As Double
                
                StartPrice = ws.Cells(StartRow, 3).Value ' Opening price of the year
                EndPrice = ws.Cells(i, 6).Value ' Closing price of the year
                YearlyChange = EndPrice - StartPrice
                If StartPrice <> 0 Then
                    PercentageChange = YearlyChange / StartPrice
            
                Else
                    PercentageChange = 0
                End If
                
                ' Print yearly change and percentage change to the summary table
                ws.Cells(Summary_Ticker_Row, 10).Value = YearlyChange
                ws.Cells(Summary_Ticker_Row, 11).Value = PercentageChange
                
                ' Move to the next row in the summary table
                Summary_Ticker_Row = Summary_Ticker_Row + 1
                
                ' Reset variables for the new ticker
                Stock_Total = 0
                StartRow = i + 1 ' Update the starting row for the next ticker
            Else
                ' If ticker is still the same, add results
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            End If
        
        ' code for coloring
                If (ws.Cells(Summary_Ticker_Row, 10).Value) < 0 Then
                    ws.Cells(Summary_Ticker_Row, 10).Interior.ColorIndex = 3
                ElseIf (ws.Cells(Summary_Ticker_Row, 10).Value) > 0 Then
                    ws.Cells(Summary_Ticker_Row, 10).Interior.ColorIndex = 4
                End If
        
        Next i
        
        ' --------- Greatest  values ---------
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        Dim MaxIncrease As Double
        Dim MaxDecrease As Double
        Dim MaxTotalVolume As Double
        Dim MaxRow, MinRow, MaxVolumeRow As Double
        Dim Max_ticker, Min_ticker, MaxTotalVolume_ticker As String
        
        ' calculate and print out the greast % and value
        MaxIncrease = Application.WorksheetFunction.Max(ws.Columns("K")) ' calculate the max value
        MaxRow = Application.Match(MaxIncrease, ws.Columns("K"), 0)  ' calculate the max row number
        Max_ticker = ws.Cells(MaxRow, 9).Value ' ticker column
        ws.Cells(2, 17).Value = MaxIncrease ' print the value column
        ws.Cells(2, 16).Value = Max_ticker     ' print the ticker column

        MaxDecrease = Application.WorksheetFunction.Min(ws.Columns("K")) ' calculate the min value
        MinRow = Application.Match(MaxDecrease, ws.Columns("K"), 0)    ' calculate the min row number
        Min_ticker = ws.Cells(MinRow, 9).Value ' ticker column
        ws.Cells(3, 17).Value = MaxDecrease   ' print the value column
        ws.Cells(3, 16).Value = Min_ticker      ' print the ticker column
        
        MaxTotalVolume = Application.WorksheetFunction.Max(ws.Columns("L")) ' calculate the max total volume
        MaxVolumeRow = Application.Match(MaxTotalVolume, ws.Columns("L"), 0) ' calculate the row number
        MaxTotalVolume_ticker = ws.Cells(MaxVolumeRow, 9).Value ' ticker column
        ws.Cells(4, 17).Value = MaxTotalVolume ' print the value column
        ws.Cells(4, 16).Value = MaxTotalVolume_ticker     ' print the ticker column
    
    Next ws
    
End Sub
