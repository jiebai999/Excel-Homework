sub Moderate()
 ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws in Worksheets
     
        Dim ticker as string
        'Dim Volume_Total As Double
        Dim Yearly_Change as Double
        Dim Percent_Change as Double
        Dim n as Integer
        n=0
        Volume_Total=0
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

     ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Range("I" & 1).Value = "ticker"
        Range("K" & 1).Value = "Yearly Change"
        Range("L" & 1).Value = "Percent Change"
        'Range("L" & 1).Value = "Total Stock Volume"

     'loop through all stock data
      for i=2 to LastRow
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

           Yearly_Change = Cells(i,6).Value - Cells(i-n,3).Value
           
           Percent_Change = Yearly_Change/Cells(i-n,3).Value

             ' Print Yearly_Change in the Summary Table
            Range("K" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Numberformat="0.000000000"
             ' Print Percent_Change in the Summary Table
            Range("L" & Summary_Table_Row).Value = Percent_Change
            Range("L" & Summary_Table_Row).Numberformat="0.00%"
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1

            'Reset the counter 
            n=0
            Else
              n=n+1
         end if
      next i

   
      ' Add Percentage style
      for i = 2 to Summary_Table_Row
         'ws.Cells(i, 12).Numberformat = "0.00%"
         if Cells(i,11).Value>=0 Then
           Cells(i, 11).Interior.ColorIndex = 4
         Else
           Cells(i, 11).Interior.ColorIndex = 3
         end If
      next i
    next ws


end sub