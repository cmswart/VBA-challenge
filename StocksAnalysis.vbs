Attribute VB_Name = "Module1"

Sub Stocks()


Dim ticker As String
Dim Summary_Table_Row As Integer

Dim closeprice As Double
Dim openprice As Double
Dim volume As Double
Dim percentChange As Double
Dim change As Double
Dim greatestpercentincrease As Double
Dim greatestPercentIncreaseTicker As String
Dim greatestPercentDecreaseValue As Double
Dim greatestPercentDecreaseTicker As String
Dim greatesttotalvolumevalue As Double
Dim greatesttotalvolumeticker As String

'Dim Worksheet
Dim ws As Worksheet
Set ws = ActiveSheet

For Each ws In ActiveWorkbook.Worksheets

closeprice = 0
openprice = 0
percentChange = 0
change = 0

'get the row number of the last row with data
    RowCount = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    
'Set Range Names
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Format
ws.Range("K2:K" & RowCount).NumberFormat = "0.00%"
ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "#"

Summary_Table_Row = 2


  For i = 2 To RowCount
  
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
  
      ticker = ws.Cells(i, 1).Value
      ws.Range("I" & Summary_Table_Row).Value = ticker
      
    openprice = ws.Cells(i, 3).Value
    volume = 0
    volume = volume + ws.Cells(i, 7).Value

ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
volume = volume + ws.Cells(i, 7).Value

ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    volume = volume + ws.Cells(i, 7).Value

      closeprice = ws.Cells(i, 6).Value

    ws.Range("L" & Summary_Table_Row).Value = volume
    
    ws.Range("J" & Summary_Table_Row).Value = closeprice - openprice
    ws.Range("K" & Summary_Table_Row).Value = (closeprice - openprice) / openprice

      Summary_Table_Row = Summary_Table_Row + 1

End If
    Next i
    
For i = 2 To RowCount

If ws.Range("J" & i) > 0 Then
ws.Range("J" & i).Interior.Color = vbGreen
ElseIf ws.Range("J" & i) = 0 Then
ws.Range("J" & i).Interior.Color = vbYellow
ElseIf ws.Range("J" & i) < 0 Then
ws.Range("J" & i).Interior.Color = vbRed

End If

If ws.Range("k" & i) > 0 Then
ws.Range("k" & i).Interior.Color = vbGreen
ElseIf ws.Range("k" & i) = 0 Then
ws.Range("k" & i).Interior.Color = vbYellow
ElseIf ws.Range("J" & i) < 0 Then
ws.Range("k" & i).Interior.Color = vbRed

End If
Next i

SummaryRowCount = ActiveSheet.Cells(ActiveSheet.Rows.Count, 9).End(xlUp).Row

For i = 2 To SummaryRowCount

   If ws.Range("L" & i).Value > greatesttotalstockvolume Then
  greatesttotalstockvolume = ws.Range("L" & i).Value
  greatesttotalstockvolumeticker = ws.Range("I" & i).Value
  
  End If

   If ws.Range("K" & i).Value > greatestpercentincrease Then
  greatestpercentincrease = ws.Range("K" & i).Value
  greatestPercentIncreaseTicker = ws.Range("I" & i).Value

   End If
     
     If ws.Range("K" & i).Value < greatestPercentDecrease Then
  greatestPercentDecrease = ws.Range("K" & i).Value
  greatestPercentDecreaseTicker = ws.Range("I" & i).Value

   End If
   
     Next i
     
   ws.Range("Q4").Value = greatesttotalstockvolume
   ws.Range("P4").Value = greatesttotalstockvolumeticker
      ws.Range("Q2").Value = greatestpercentincrease
   ws.Range("P2").Value = greatestPercentIncreaseTicker
   ws.Range("Q3").Value = greatestPercentDecrease
   ws.Range("P3").Value = greatestPercentDecreaseTicker
  
ws.Range("L1:Q1").EntireColumn.AutoFit

Next ws

End Sub


