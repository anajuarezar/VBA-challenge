Attribute VB_Name = "Módulo4"
Sub Color_Formating():

'Go through the worsheets

For Each ws In Worksheets

For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

If ws.Cells(I, 10).Value > 0 Then

ws.Cells(I, 10).Interior.ColorIndex = 4

ElseIf ws.Cells(I, 10).Value = "" Then

ws.Cells(I, 10).Interior.ColorIndex = 0

Else

ws.Cells(I, 10).Interior.ColorIndex = 3

End If

Next I

Next ws

End Sub
