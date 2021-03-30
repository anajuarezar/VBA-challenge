Attribute VB_Name = "Módulo3"
Sub Total_volume():

'Go through the worksheets

For Each ws In Worksheets

'Define variables

Dim totalvolume As Double
totalvolume = 0
Dim totaltable As Integer
totaltable = 2

'Compare cells to know when the ticker symbol changes

For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

   'Define total volume

   totalvolume = totalvolume + ws.Cells(I, 7).Value

   'Write it in the table

   ws.Range("L" & totaltable).Value = totalvolume

   totaltable = totaltable + 1

   totalvolume = 0

   ' Write header

   ws.Cells(1, 12).Value = "Total Stock Volume"

   Else

   totalvolume = totalvolume + ws.Cells(I, 7).Value

   End If

   Next I

   Next ws

   End Sub



