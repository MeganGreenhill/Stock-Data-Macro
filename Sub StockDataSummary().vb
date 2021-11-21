Sub StockDataSummary()

  ' Set an initial variable for holding the ticker symbol
  Dim TickerSymbol As String

  ' Set an initial variable for holding the total stock volume for the year
  ' (This variable, and some others, were defined as a Variant data type to prevent 
  ' an overflow error that often occurs when using other data types with Excel for Mac OS.)
  Dim TotalStockVolume As Variant
  TotalStockVolume = 0

  ' Keep track of the location for each ticker in the Summary Table
  Dim Summary_Table_Row As Variant
  Summary_Table_Row = 2

  ' Print headers for Summary Table
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"

  ' Loop through all amounts
  For I = 2 To 797711

    ' If the cell is the first value for the ticker...
    If Cells(I, 1).Value <> Cells(I - 1, 1).Value Then
      
      ' Define Open Value
      Dim OpenValue As Variant
      OpenValue = Cells(I, 3).Value
    
    End If

    ' Check if we are still within the same ticker, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
      
      ' Define Close Value
      Dim CloseValue As Variant
      CloseValue = Cells(I, 6).Value
      
      ' Set the Ticker Symbol
      TickerSymbol = Cells(I, 1).Value

      ' Add to the Stock Volume Total
      TotalStockVolume = TotalStockVolume + Cells(I, 7).Value

      ' Calculate Yearly Change and Percent Change
      Dim YearlyChange As Variant
      Dim PercentChange As Variant
      YearlyChange = CloseValue - OpenValue
      ' Some open values at the start of the year are zero and thus result in a divide by zero error,
      ' so the following IF function was utilised to only calculate percentage change for nonzero values.
      ' Additionally, the FormatPercent function was used so that the percentage data format did not need
      ' to be applied manually via the worksheet.
      If OpenValue = 0 Then
      PercentChange = FormatPercent(0)
      Else
      PercentChange = FormatPercent(YearlyChange / OpenValue)
      End If

      ' Print the Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = TickerSymbol

      ' Print the Yearly Change in the Summary Table
      Range("J" & Summary_Table_Row).Value = YearlyChange

      ' Print the Percent Change to the Summary Table
      Range("K" & Summary_Table_Row).Value = PercentChange

      ' Print the Total Stock Volume Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = TotalStockVolume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Volume Total
      TotalStockVolume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Stock Volume Total
      TotalStockVolume = TotalStockVolume + Cells(I, 7).Value

    End If

  Next I

' Apply conditional formatting for Yearly Change in Summary Table
  For I = 2 To 3200

    If Cells(I, 10).Value >= 0 Then
    Cells(I, 10).Interior.ColorIndex = 4
    Else
    Cells(I, 10).Interior.ColorIndex = 3
    End If
    
  Next I

End Sub