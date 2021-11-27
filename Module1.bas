Attribute VB_Name = "Module1"
Sub ticker_hw()
  
  'Loop through all the worksheets
  For Each ws In Worksheets


  ' print headers tickers, yearly change, percent change, total ticker volume
  ws.Range("I1") = "Ticker"
  ws.Range("J1") = "Yearly Change"
  ws.Range("K1") = "Percent Change"
  ws.Range("L1") = "Total Ticker Volume"
  
  ' Set variable for holding the ticker name
  Dim Ticker_Name As String

  ' Set variable for holding the total volume per ticker
  Dim Total_Ticker_Volume As Double
  
  
  ' Set variable for holding the yearly change
  Dim YearlyChange As Double
  
  ' Set variable holding the percent change
  Dim PercentChange As Double
  
  ' Set variable for holding the tickers Open
  Dim TickerOpen As Double
  
  ' Set variable for holding the tickers Close
  Dim TickerClose As Double
  
  ' Find the last row within the sheet
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  Total_Ticker_Volume = 0

  ' Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2

  ' Loop through all tickers total ticker volume to last row
  For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Total ticker Volume
      Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value

      ' Print the ticker name in the summary table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Total Ticker volume to the summary table
      ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
      
      ' Reset the total ticker volume
      Total_Ticker_Volume = 0
    
    ' Loop through to tickers open & close to calculate the yearly change and the percentage change
    
    TickerClose = Cells(i, 6)
        If TickerOpen = 0 Then
            YearlyChange = 0
            PercentageChange = 0
            Else:
            ' Calculation to obtain Ticker Yearly Change
            YearlyChange = TickerClose - TickerOpen
            ' Calculation to obtain Ticker Percentage Change
            PercentageChange = (TickerClose - TickerOpen) / TickerOpen
            End If
            
        'Print the Summary table for the Ticker Yearly Change and Percentage Change
        ws.Range("J" & Summary_Table_Row).Value = YearlyChange
        ws.Range("K" & Summary_Table_Row).Value = PercentageChange
        'Set the format to %
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
           ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
            ElseIf Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
                TickerOpen = ws.Cells(i, 3)

    ' If the cell immediately following a row is the ticker
    Else

      ' Add to the Total ticker Volume
      Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value

    End If

  Next i
  
  ' Formatting the yearly change to show those ones in negative to red and normal as green
  ' Loop through all tickers total ticker volume to last row
  For i = 2 To lastrow
    'if it's less than 0 then format red
    If ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
    ' if it's greater 0 then its green
    ElseIf ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 4
    
  End If
    
    Next i
    
    Next ws
    
End Sub


