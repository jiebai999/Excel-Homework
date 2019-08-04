sub stock()
 ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws in Worksheets
     
        Dim ticker as string
        Dim Volume_Total As Double
        Volume_Total=0
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

     ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Range("I" & 1).Value = "ticker"
        Range("J" & 1).Value = "Total Stock Volume"

     'loop through all stock data
      for i=2 to LastRow
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         ' Set the ticker
           ticker = Cells(i, 1).Value
           Volume_Total = Volume_Total + Cells(i, 7).Value

            ' Print the ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = ticker

            ' Print the total volume to the Summary Table
            Range("J" & Summary_Table_Row).Value = Volume_Total

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the Brand Total
            Volume_Total = 0

            ' If the cell immediately following a row is the same brand...
            Else

             ' Add to the Brand Total
              Volume_Total = Volume_Total + Cells(i, 7).Value
            
         end if
      next i
    next ws


end sub