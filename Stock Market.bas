Sub stock()

For Each ws In Worksheets

 'Set initial variables
  Dim ticker As String
  Dim opening_price As Double
  Dim closing_price As Double
  Dim yearly_difference As Double
  Dim yearly_change As Double
  
  'Set summary table names
  ws.cells(1, 9).value="Ticker"
  ws.cells(1, 10).value="Yearly Change"
  ws.cells(1, 11).value="Percent Change"
  ws.cells(1, 12).value="Total Stock Volume"
  
    
  'Set an initial variable for holding the total volume
  Dim total_volume As Double
  total_volume = 0
  

  ' Summary table index
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
   
    opening_price = ws.Cells(2, 3).Value

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

         
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
        ticker = ws.Cells(i, 1).Value
        closing_price = ws.Cells(i, 6).Value
        yearly_difference = closing_price - opening_price

        if opening_price=0 Then
        yearly_change=0 
        else
        yearly_change = yearly_difference / opening_price
        endif

        total_volume = total_volume + Cells(i, 7).Value
    
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("J" & Summary_Table_Row).Value = yearly_difference
        ws.Range("K" & Summary_Table_Row).Value = yearly_change
        ws.Range("L" & Summary_Table_Row).Value = total_volume
    
        Summary_Table_Row = Summary_Table_Row + 1
        total_volume = 0
        opening_price = ws.Cells(i + 1, 3).Value
   
      Else
        total_volume = total_volume + ws.Cells(i, 7).Value
      End If

    Next i
  
  'Formatting
   For i = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
    If ws.Cells(i, 10).Value >=0 Then
    ws.cells(i,10).Interior.ColorIndex=4
    Else
     ws.cells(i,10).Interior.ColorIndex=3
    Endif
  Next i

  ws.Columns("K").NumberFormat = "0.00%"

   Next ws
  End Sub