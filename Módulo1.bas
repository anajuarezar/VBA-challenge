Attribute VB_Name = "Módulo1"
Sub Ticker_Symbol():

'Go through the worksheets.

For Each ws In Worksheets

'Define variables

Dim ticker As String
Dim ticker_table As Integer
ticker_table = 2

'Loop through data for ticker symbol

For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
   
   ticker_name = ws.Cells(I, 1).Value

   'Write in the table

   ws.Range("I" & ticker_table).Value = ticker_name

   ticker_table = ticker_table + 1

   'Write the header

   ws.Cells(1, 9).Value = "Ticker"

   End If

Next I

Next ws

End Sub
