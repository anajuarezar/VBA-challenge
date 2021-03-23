Attribute VB_Name = "Módulo2"
Sub Year_Change():

'Go through worksheets

For Each ws In Worksheets

'Define variables

Dim opening As Double
opening = ws.Cells(2, 3).Value
Dim closing As Double
Dim change_row As Integer
change_row = 2
Dim yearly_change As Double
Dim percentage_row As Integer
percentage_row = 2

'Compare the value of the cells to determine when it is a different ticker symbol

For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

  If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

'Define closing price and yearly change

  closing = ws.Cells(I, 6).Value
  yearly_change = closing - opening

'Adress the possible 0 in opening

    If opening = 0 Then
    change_percentage = 0

    Else
    change_percentage = yearly_change / opening
    End If

'Write it in the table

  ws.Range("J" & change_row).Value = yearly_change
  change_row = change_row + 1

  ws.Range("K" & percentage_row).Value = change_percentage
  
    'Convert them to percentages

  ws.Range("K" & percentage_row).NumberFormat = "0.00%"
  percentage_row = percentage_row + 1
  

  'Write the header

  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"

 
  End If


  Next I

Next ws

End Sub

