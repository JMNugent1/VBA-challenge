Attribute VB_Name = "Module1"
Sub stonks()

'Set initial variables

Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock As Double
open_price = Cells(2, 3).Value
total_stock = 0

'Keep track of the location for each ticker symbol in the summary table
Dim summary_table_row As Integer
summary_table_row = 2

'Loop through all the stocks
For i = 2 To 759001
    
'Check if we are still within the same ticker name, if we are not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
'Set the ticker name
    ticker = Cells(i, 1).Value
    
'Set the closing price
    close_price = Cells(i, 6).Value
    
'Calculate the yearly change
    yearly_change = close_price - open_price
    
'Calculate the percent change
    percent_change = (close_price - open_price) / open_price
    
'Add to the total stock volume
    total_stock = total_stock + Cells(i, 7).Value
    
'Print the ticker name in the summary table
    Range("I" & summary_table_row).Value = ticker
    
'Print the yearly change in the summary table
    Range("J" & summary_table_row).Value = yearly_change
    
'Print the percent change in the summary table
    Range("K" & summary_table_row).Value = percent_change
    
'Print the total stock volume in the summary table
    Range("L" & summary_table_row).Value = total_stock
    
'Add one to the summary table row
    summary_table_row = summary_table_row + 1
    
'Reset the opening price
    open_price = Cells(i + 1, 3).Value

'Reset the total stock volume
    total_stock = 0
    
'If the cell immediately following a row is the same stock...
    Else
    
'Add to the Total Stock Volume
    total_stock = total_stock + Cells(i, 7).Value
    
    End If

Next i

End Sub
