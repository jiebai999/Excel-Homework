sub Hard1()
 ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws in Worksheets
     
        Dim Greatest as Double
        Dim Greatest1 as Double
        Dim Greatest2 as Double
        Dim n, n1, n2 as Integer

     ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        Range("O" & 1).Value = "ticker"
        Range("P" & 1).Value = "Value"
        Cells(2,14).Value="Greatest % Increase"
        Cells(3,14).Value="Greatest % Decrease"
        Cells(4,14).Value="Greatest Total Volume"
    
      LastRow1 = ws.Cells(Rows.Count, 12).End(xlUp).Row
  '----------Greatest % Increase---------------
      Greatest = Cells(2,12).Value
      for i = 3 to LastRow1
           if Cells(i,12).Value>Greatest Then
             Greatest = Cells(i,12).Value
             n=i
           end If
      next i
      'Cells(2,15).Value=Cells(i,1).Value
      Cells(2,16).Value=Greatest
      Cells(2,16).Numberformat="0.00%"
      Cells(2,15).Value=Cells(n,9).Value

  '-----------Greatest % Decrease---------------
      Greatest1 = Cells(2,12).Value
      for i = 3 to LastRow1
           if Cells(i,12).Value<Greatest1 Then
             Greatest1 = Cells(i,12).Value
             n1=i
           end If
      next i
     ' Cells(3,15).Value=Cells(i,1).Value
      Cells(3,16).Value=Greatest1
      Cells(3,16).Numberformat="0.00%"
      Cells(3,15).Value=Cells(n1,9).Value

      Greatest2 = Cells(2,10).Value
      for i = 3 to LastRow1
           if Cells(i,10).Value>Greatest2 Then
             Greatest2 = Cells(i,10).Value
             n2=i
           end If
      next i
      'Cells(4,15).Value=Cells(i,1).Value
      Cells(4,16).Value=Greatest2
      Cells(4,15).Value=Cells(n2,9).Value


    next ws


end sub