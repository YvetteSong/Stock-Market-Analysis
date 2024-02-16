Attribute VB_Name = "Module1"

Sub stock_ticker()
    ' Set up the table headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    ' Declare variables for ticker & total volume calculations
    Dim Ticker_name As String
    Dim Stock_Total As Double
    Dim RowCount As Long
    Dim Summary_Ticker_Row As Long
    Dim i As Long
    
    Stock_Total = 0
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    Summary_Ticker_Row = 2
    
    ' Iterate through rows
    For i = 2 To RowCount
        ' If the next ticker in the row is not the same, print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker_name = Cells(i, 1).Value
            Stock_Total = Stock_Total + Cells(i, 7).Value
            
            ' Print ticker name and stock total to the summary table
            Cells(Summary_Ticker_Row, 9).Value = Ticker_name
            Cells(Summary_Ticker_Row, 12).Value = Stock_Total
            
            ' Calculate yearly change and percentage change for each row
            Dim yearly_change As Double
            Dim percentage_change As Double
            yearly_change = Cells(i, 6).Value - Cells(i, 3).Value
            If Cells(i, 3).Value <> 0 Then
                percentage_change = (yearly_change / Cells(i, 3).Value) * 100
            Else
                percentage_change = 0
            End If
            
            ' Print yearly change and percentage change to the summary table
            Cells(Summary_Ticker_Row, 10).Value = yearly_change
            Cells(Summary_Ticker_Row, 11).Value = percentage_change
            
            ' Move to the next row in the summary table
            Summary_Ticker_Row = Summary_Ticker_Row + 1
            
            ' Reset variables for the new ticker
            Stock_Total = 0
        Else
            ' If ticker is still the same, add results
            Stock_Total = Stock_Total + Cells(i, 7).Value
        End If
    Next i
    
    ' code for coloring
    If (Cells(i, 6).Value - Cells(i, 3).Value) < 0 Then
        Cells(Summary_Ticker_Row, 10).Interior.ColorIndex = 3
    ElseIf (Cells(i, 6).Value - Cells(i, 3).Value) > 0 Then
        Cells(Summary_Ticker_Row, 10).Interior.ColorIndex = 4
    
    End If
    
    
    
    ' --------- Greatest % decrease & "Greatest total volume ---------
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
End Sub
   


