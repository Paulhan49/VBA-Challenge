Attribute VB_Name = "Module1"
Sub StockVolume()
'for all loops
For Each ws In Worksheets

'Define variables:

'Define Ticker and print header
Dim Ticker_Name As String
ws.Cells(1, "H").Value = "Ticker"

'Define Yearly_Change and print header
Dim Yearly_Change As Double
ws.Cells(1, "I").Value = "Yearly Change"

'Define Percent_Change and print header
 Dim Percent_Change As Double
 ws.Cells(1, "J").Value = "Percent Change"

'Define Stock_Total_Volume and print header
Dim Stock_Total_Volume As Double
Stock_Total_Volume = 0
ws.Cells(1, "K").Value = "Total_Stock_Volume"

'Define Summary_Row_Table
Dim Summary_Row_Table As Double
Summary_Row_Table = 2

'Define Open Price
Dim Open_Price As Double
Open_Price = ws.Cells(2, 3).Value


'Define Close Price
Dim Close_Price As Double

' Determine the last row
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Set Loop range
For i = 2 To lastRow

'if we are still within the same ticker symbol, if it is not then
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Set ticker name
Ticker_Name = ws.Cells(i, 1).Value
ws.Cells(Summary_Row_Table, 8).Value = Ticker_Name
                
' Set close price
 Close_Price = ws.Cells(i, 6).Value
                
'Set yearly change
Yearly_Change = Close_Price - Open_Price
ws.Range("I" & Summary_Row_Table).Value = Yearly_Change
                
' Set percent change
If (Open_Price = 0 And Close_Price = 0) Then
    Percent_Change = 0
ElseIf (Open_Price = 0 And Close_Price <> 0) Then
        Percent_Change = 1
Else
Percent_Change = (Close_Price - Open_Price) / Open_Price
ws.Range("J" & Summary_Row_Table).Value = Percent_Change
ws.Range("J" & Summary_Row_Table).NumberFormat = "0.00%"


End If

' Set Stock total volumn
Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
ws.Range("K" & Summary_Row_Table).Value = Total_Stock_Volume

' Add one to the summary table row
Summary_Row_Table = Summary_Row_Table + 1

' reset the open price
Open_Price = ws.Cells(i + 1, 3)

' reset the Total_Stock_Volume
Total_Stock_Volume = 0
            
'Else if ticker is the same
   Else
 Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
 
End If
 
 'if yearly change is positive, highlight green. if yearly change is negative, highlight red
If ws.Range("I" & Summary_Row_Table).Value > 0 Then
ws.Range("I" & Summary_Row_Table).Interior.ColorIndex = 4

ElseIf ws.Range("I" & Summary_Row_Table).Value <= 0 Then
ws.Range("I" & Summary_Row_Table).Interior.ColorIndex = 3

End If





   Next i
         
   
   Next ws
     End Sub



