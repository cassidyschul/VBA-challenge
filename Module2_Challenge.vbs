Sub Stocks()

Dim WS As Worksheet
    For Each WS In ThisWorkbook.Worksheets
   
    WS.Range("I1").EntireColumn.Insert
    WS.Cells(1, 9).Value = "Ticker"
    WS.Range("J1").EntireColumn.Insert
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Range("K1").EntireColumn.Insert
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Range("L1").EntireColumn.Insert
    WS.Cells(1, 12).Value = "Total Stock Volume"

  Dim Ticker As String

  Dim Yearly_Change As Double

  Dim Percent_Change As Double
  
  Dim Stock_Total As Double
  
  Stock_Total = 0

  Dim Open_Stock As Double
  Open_Stock = WS.Cells(2, 3).Value
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
      
  lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
      
    For i = 2 To lastrow
      
      If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        
        Ticker = WS.Cells(i, 1).Value
    
        Yearly_Change = (WS.Cells(i, 6).Value)- Open_Stock
            If Yearly_Change > 0 Then
                WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
          
        Percent_Change = ((WS.Cells(i, 6).Value- Open_Stock) / Open_Stock)
    
        Stock_Total = Stock_Total + WS.Cells(i, 7).Value
    
        WS.Range("I" & Summary_Table_Row).Value = Ticker
    
        WS.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
    
        WS.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        WS.Range("K" & Summary_Table_Row).Value = Percent_Change
        

        WS.Range("L" & Summary_Table_Row).Value = Stock_Total
    
        Summary_Table_Row = Summary_Table_Row + 1
          
        Open_Stock = WS.Cells(i + 1, 3).Value
    
        Stock_Total = 0
          
          
    
        Else
        
        Stock_Total = Stock_Total + WS.Cells(i, 7).Value
      
      

        End If

  Next i

  Dim max As Double
  Dim min As Double
  Dim Tinker_Max As String
  Dim Tinker_Min As String
  Dim Max_Volume As Double
  Dim Tinker_Max_Volume As String


  WS.Range("O2").Value = "Greatest % Increase"

  max = 0

  For i = 2 To lastrow

    If (WS.Cells(i, 11).Value) > max Then
      max = WS.Cells(i, 11).Value
      Tinker_Max = WS.Cells(i, 9).Value
    End If

  Next i

  WS.Range("Q2").Value = max
  WS.Range("Q2").NumberFormat = "0.00%"
  WS.Range("P2").Value = Tinker_Max

  WS.Range("O3").Value = "Greatest % Decrease"

  min = 0

  For i = 2 To lastrow

    If (WS.Cells(i, 11).Value) < min Then
      min = WS.Cells(i, 11).Value
        Tinker_Min = WS.Cells(i, 9).Value
    End If
      
  Next i

  WS.Range("Q3").Value = min
  WS.Range("Q3").NumberFormat = "0.00%"
  WS.Range("P3").Value = Tinker_Min

  WS.Range("O4").Value = "Greatest Total Volume"

  Max_Volume = 0

  For i = 2 To lastrow

    If (WS.Cells(i, 12).Value) > Max_Volume Then
      Max_Volume = WS.Cells(i, 12).Value
      Tinker_Max_Volume = WS.Cells(i, 9).Value
    End If
    
  Next i

  WS.Range("Q4").Value = Max_Volume
  WS.Range("P4").Value = Tinker_Max_Volume

  WS.Range("P1").Value = "Ticker"
  WS.Range("Q1").Value = "Value"
  
Next WS

End Sub
