Attribute VB_Name = "Module1"
Sub stockhomework()

' Set the initial holding variables
Dim ticker As String
Dim openprice As Double
openprice = 0
Dim closeprice As Double
closeprice = 0
Dim stockvolume As Integer
stockvolume = 0

' Location for each stock in the summary table
summary_ticker_row = 2

' loop function
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Function to see what names are there
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    ' Set my ticker name
    ticker = Cells(i, 1).Value
    
    ' Print in the summary table
    Range("H" & Summary_Table_Row).Value = ticker

    ' Print Total Volume in the summary table
    Range("I" & Summary_Table_Row).Value = ticker_total
    
    ' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    
    ' Reset the ticker_total
    ticker_total = 0
    
'If the next cell down is the same brand
Else

    ' Add to the ticker_total
    ticker_total = ticker_total + Cells(1, 3).Value
    
End If

End Sub
