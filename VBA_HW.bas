Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information.

  
Sub Vba_Hw()


'Loop through all sheets
For Each ws In Worksheets

'Set column headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest % Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"



'Set variables
Dim ticker As String
Dim yearlyChange As Double
Dim percentageChange As Double
Dim volume As Double
Dim openPrice As Double
Dim closePrice As Double


Dim lastRow As Double
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

volume = 0

Dim summaryTableRow As Double
summaryTableRow = 2

Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Double

greatestIncrease = 0
greatestDecrease = 0
greatestVolume = 0

'Start
For i = 2 To lastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        ticker = ws.Cells(i, 1).Value
        volume = volume + ws.Cells(i, 7).Value

          ws.Range("I" & summaryTableRow).Value = ticker
          ws.Range("L" & summaryTableRow).Value = volume

volume = 0

        closePrice = ws.Cells(i, 6)
       
        If openPrice = 0 Then
            yearlyChange = 0
            percentChange = 0
        Else
        
            yearlyChange = closePrice - openPrice
            percentChange = (closePrice - openPrice) / openPrice
        End If

            ws.Range("J" & summaryTableRow).Value = yearlyChange
            ws.Range("K" & summaryTableRow).Value = percentChange
            ws.Range("K" & summaryTableRow).Style = "Percent"
            ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"

            summaryTableRow = summaryTableRow + 1

    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
         openPrice = ws.Cells(i, 3)

    Else: volume = volume + ws.Cells(i, 7).Value

    End If
    Next i


For r = 2 To lastRow

    If ws.Range("J" & r).Value > 0 Then
        ws.Range("J" & r).Interior.ColorIndex = 4

    ElseIf ws.Range("J" & r).Value < 0 Then
        ws.Range("J" & r).Interior.ColorIndex = 3
        
    End If

    Next r

For a = 2 To lastRow


    If ws.Cells(a, 11).Value > greatestIncrease Then
        greatestIncrease = ws.Cells(a, 11).Value
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = ws.Cells(a, 9).Value
    End If

    Next a

For b = 2 To lastRow
    
    If ws.Cells(b, 11).Value < greatestDecrease Then
        greatestDecrease = ws.Cells(b, 11).Value
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = ws.Cells(b, 9).Value
    End If
    
   Next b

For c = 2 To lastRow
    
    If ws.Cells(c, 12).Value > greatestVolume Then
        greatestVolume = ws.Cells(c, 12).Value
        ws.Range("Q4").Value = greatestVolume
        ws.Range("P4").Value = ws.Cells(c, 9).Value
    End If
  
    Next c
 
ws.Columns("A:Q").AutoFit
    
Next ws
End Sub





