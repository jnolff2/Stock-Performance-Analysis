VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockVolume()

Dim WS As Worksheet
    ' Loop through each worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
           
        ' Create variables for the tickers, yearly change, percent change, total volume, opening price, and closing price
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        
        ' Create a variable to hold the total volume
        TotalVolume = 0
        ' Identify the placement for the header titles in the summary table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        OpeningPrice = Cells(2, 3).Value
       
        ' Loop through each row
        For i = 2 To LastRow
        
        ' Make sure we are still in the same ticker data and if not then...
        If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
        Ticker = Cells(i, 1).Value
        ' Calculate the TotalVolume
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
        ' Calculate the YearlyChange
        ClosingPrice = Cells(i, 6).Value
        YearlyChange = ClosingPrice - OpeningPrice
        
        ' Calculate the PercentChange
        PercentChange = YearlyChange / OpeningPrice
        
        ' Print the ticker, total volume, yearly change, and percent change (formatted 0.00%) in the summary table
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("J" & Summary_Table_Row).Value = YearlyChange
        Range("K" & Summary_Table_Row).Value = PercentChange
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        Range("L" & Summary_Table_Row).Value = TotalVolume
            
            ' Color the cells green if they are positive or red if they are negative in the Yearly Change column of the summary table
            For j = 2 To Summary_Table_Row
                If (Cells(j, 10).Value >= 0) Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            Next j
            
        ' Add 1 to the summary table row to start the next ticker and reset the total volume to 0
        Summary_Table_Row = Summary_Table_Row + 1
        TotalVolume = 0
        
        Else
        ' Add the total volume for the current ticker
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
        End If
        
        Next i
    
    Next WS
    
End Sub



