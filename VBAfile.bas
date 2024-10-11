Attribute VB_Name = "Module1"
Sub Ticker()


Dim i As Long
Dim j As Long
Dim k As Long

Dim ws As Worksheet

For Each ws In Worksheets

Dim WorksheetName As String
Dim LastRow As Long
Dim Ticker As String

Dim Volume As Double


Volume = 0

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Quarterly Change"
ws.Range("L1").Value = "Percentage Change"
ws.Range("M1").Value = "Total Stock Volume"
ws.Range("N1").Value = "Open Price"
ws.Range("O1").Value = "Close Price"

Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

'Loop through all ticker symbols
For i = 2 To LastRow

'Check if we are still within the same ticker symbol, if it is not...
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

Ticker = ws.Cells(i, 1).Value

Volume = Volume + ws.Cells(i, 7).Value
 
 ' Print the ticker in the Summary Table
ws.Range("J" & Summary_Table_Row).Value = Ticker

' Print the Volume to the Summary Table
ws.Range("M" & Summary_Table_Row).Value = Volume

' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1


'Reset the volume
Volume = 0

Else

'Add to the volume
Volume = Volume + ws.Cells(i, 7).Value

 'Print the Volume to the Summary Table
ws.Range("M" & Summary_Table_Row).Value = Volume


End If

Next i

Dim TargetValue As String
Dim foundovRow As Long
Dim foundcvRow As Long
Dim OpenValue As Double
Dim CloseValue As Double
Dim QC As Double
Dim PC As Double


'Find the targetvalue and get the openvalue and closevalue
  Summary_Table_Row = 2

For j = 2 To LastRow

TargetValue = ws.Range("J" & Summary_Table_Row).Value
If ws.Cells(j, 1).Value = TargetValue And ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then

foundcvRow = j

CloseValue = ws.Cells(foundcvRow, 6).Value

ws.Range("O" & Summary_Table_Row).Value = CloseValue


'reset foundcvRow for the new targetvalue
foundcvRow = 0

' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

End If
Next j


  Summary_Table_Row = 2
  
For i = 2 To LastRow

TargetValue = ws.Range("J" & Summary_Table_Row).Value

If ws.Cells(i, 1).Value = TargetValue Then

foundovRow = i

OpenValue = ws.Cells(foundovRow, 3).Value

ws.Range("N" & Summary_Table_Row).Value = OpenValue


'reset foundcvRow for the new targetvalue
foundovRow = 0

' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
End If

Next i

Dim LR As Long

LR = ws.Cells(Rows.Count, 10).End(xlUp).Row

For k = 2 To LR

QC = ws.Cells(k, 15).Value - ws.Cells(k, 14).Value

ws.Cells(k, 11).Value = QC

'Check for division by zero to prevent overflow

If ws.Cells(k, 14).Value <> 0 Then

PC = QC / ws.Cells(k, 14).Value

Else
PC = 0
End If

ws.Cells(k, 12).Value = PC


Next k

For i = 2 To LR

ws.Cells(i, 12).NumberFormat = "0.00%"

Next i


Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As Double


ws.Range("Q2").Value = "Greatest % Increase"
ws.Range("Q3").Value = "Greatest % Decrease"
ws.Range("Q4").Value = "Greatest Total Volume"
ws.Range("R1").Value = "Ticker"
ws.Range("S1").Value = "Value"


greatestincrease = Application.WorksheetFunction.Max(ws.Range("L2:L" & LR))
ws.Range("S2").Value = greatestincrease
ws.Range("S2").NumberFormat = "0.00%"

greatestdecrease = Application.WorksheetFunction.Min(ws.Range("L2:L" & LR))
ws.Range("S3").Value = greatestdecrease
ws.Range("S3").NumberFormat = "0.00%"

greatestvolume = Application.WorksheetFunction.Max(ws.Range("M2:M" & LR))
ws.Range("S4").Value = greatestvolume


For i = 2 To LR

If ws.Cells(i, 12).Value = greatestincrease Then
ws.Range("R2").Value = ws.Cells(i, 10).Value

End If

If ws.Cells(i, 12).Value = greatestdecrease Then
ws.Range("R3").Value = ws.Cells(i, 10).Value

End If

If ws.Cells(i, 13).Value = greatestvolume Then
ws.Range("R4").Value = ws.Cells(i, 10).Value

End If

Next i

For i = 2 To LR

 If ws.Cells(i, 11).Value > 0 Then
 ws.Cells(i, 11).Interior.ColorIndex = 4
 
 ElseIf ws.Cells(i, 11).Value < 0 Then
 ws.Cells(i, 11).Interior.ColorIndex = 3
 
End If
If ws.Cells(i, 12).Value > 0 Then
 ws.Cells(i, 12).Interior.ColorIndex = 4
 
 ElseIf ws.Cells(i, 12).Value < 0 Then
 ws.Cells(i, 12).Interior.ColorIndex = 3
 
 End If
 

Next i

Next ws

End Sub

Sub reset_file(): 'Resets all sheets to pre-analysis state
    Dim i As Integer
    'Loop to cycle through all workbook sheets and delete columns I through Q - This also resets formating
    For i = 1 To Sheets.Count
        With Sheets(i)
            .Columns("I:T").Delete
        End With
    Next i
End Sub

