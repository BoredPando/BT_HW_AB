Sub TickerVol()

    ' Set an initial variable for holding the brand name
    Dim Ticker As String
    Cells(1, 9).Value = "Ticker"
    Cells(1, 12).Value = "Total Stock Volume"

    ' Set an initial variable for holding the total per credit card brand
    Dim Vol_Total As Double
    Vol_Total = 0

    ' Keep track of the location for each credit card brand in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through all credit card purchases
    For I = 2 To LastRow

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the Ticker
      Ticker = Cells(I, 1).Value
      
      ' Add to the Vol Total
      Vol_Total = Vol_Total + Cells(I, 7).Value

      ' Print the Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Vol to the Summary Table
      Range("L" & Summary_Table_Row).Value = Vol_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Vol Total
      Vol_Total = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Ticker Total
      Vol_Total = Vol_Total + Cells(I, 7).Value

    End If

    Next I

End Sub

Sub YearChange()

    ' Find start and end rows with unique Ticker
    Cells(1, 10).Value = "Yearly Change"
    Dim TickStartRow As Long
    Dim TickEndRow As Long
    Dim Summary_Table_Row1 As Integer
    Summary_Table_Row1 = 2
    LastRow1 = Cells(Rows.Count, 9).End(xlUp).Row
    Dim Change As Double
    Change = 0
    Dim LookUp As String
    
    If AutoFilterMode = False Then
    Cells.AutoFilter
    End If

    ' Loop through all credit card purchases
    For I = 2 To LastRow1

      LookUp = Cells(I, 9)

      Range("A:G").AutoFilter Field:=1, Criteria1:=LookUp

      'Find start and end row
      TickStartRow = Range("A:A").Find(what:=LookUp, after:=Cells(1, 1), LookAt:=xlWhole).Row
      TickEndRow = Range("A:A").Find(what:=LookUp, after:=Cells(1, 1), LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
      
      ' Print the Ticker in the Summary Table
      Change = Range("F" & TickEndRow).Value - Range("C" & TickStartRow).Value
      Range("J" & Summary_Table_Row1).Value = Change

      ' Add one to the summary table row
      Summary_Table_Row1 = Summary_Table_Row1 + 1
      
      Cells.AutoFilter
      
    Next I
    
    If AutoFilterMode = True Then
    Cells.AutoFilter
    End If

End Sub

Sub Format()

    Dim Summary_Table_Row2 As Integer
    Summary_Table_Row2 = 2
    LastRow2 = Cells(Rows.Count, 10).End(xlUp).Row

    For I = 2 To LastRow2

      If Cells(I, 10) >= 0 Then

        ' Set the Cell Colors to Green
        Cells(I, 10).Interior.ColorIndex = 4
        
        Else
        
        ' Set the Cell Colors to Red
        Cells(I, 10).Interior.ColorIndex = 3
      
      End If
      
    Next I

End Sub

Sub Percent()
    
    'Column Header
    Cells(1, 11).Value = "Percent Change"
    
    'Set dims
    Dim Summary_Table_Row3 As Integer
    Summary_Table_Row3 = 2
    LastRow3 = Cells(Rows.Count, 10).End(xlUp).Row

    ' Find start row with unique Ticker
    Dim TickStartRow As Double

    If AutoFilterMode = False Then
    Cells.AutoFilter
    End If

    Range("A:G").AutoFilter Field:=3, Criteria1:="<>0"

    ' Loop through all unique tickers
    For I = 2 To LastRow3

      'Find start row
      TickStartRow = Range("A:A").Find(what:=Cells(I, 9), after:=Range("A1"), LookAt:=xlWhole).Row
     
        If Cells(I, 10) = 0 Then
        Cells(I, 11) = 0
        
        Else
        Cells(I, 11) = Cells(I, 10) / Cells(TickStartRow, 3)
      
        End If
        
      Cells(I, 11).NumberFormat = "0.00%"

      ' Add one to the summary table row
      Summary_Table_Row1 = Summary_Table_Row1 + 1

    Next I

    Cells.AutoFilter

End Sub

Sub Max()

    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    LastRow4 = Cells(Rows.Count, 11).End(xlUp).Row
    Dim MaxRow As String
    Dim MaxPer As Double
    
        MaxPer = WorksheetFunction.Max(Range("K:K"))

        Cells(2, 17) = MaxPer
        MaxRow = Application.Match(Cells(2, 17).Value, Range("K:K"), 0)
        
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(2, 16) = Cells(MaxRow, 9)
    
    Dim MinPer As Double
        MinPer = WorksheetFunction.Min(Range("K:K"))

        Cells(3, 17) = MinPer
        MinRow = Application.Match(Cells(3, 17).Value, Range("K:K"), 0)
        
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(3, 16) = Cells(MinRow, 9)
    
    Dim MaxVol As Double
        MaxVol = WorksheetFunction.Max(Range("L:L"))
        
        Cells(4, 17) = MaxVol
        VolRow = Application.Match(Cells(4, 17).Value, Range("L:L"), 0)
        
        Cells(4, 17).NumberFormat = "0"
        Cells(4, 16) = Cells(VolRow, 9)

End Sub