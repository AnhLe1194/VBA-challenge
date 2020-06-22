Attribute VB_Name = "Module1"
Sub ticker():

Dim ws As Worksheet

For Each ws In Worksheets

Dim ticker As String
Dim vol As Double
Dim row As Long
Dim year_open As Double
Dim year_end As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim last_row As Long

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    

row = 2

year_open = ws.Cells(2, 3).Value
last_row = ws.Cells(Rows.Count, 1).End(xlUp).row

For i = 2 To last_row
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    ticker = ws.Cells(i, 1).Value
    year_end = ws.Cells(i, 6).Value
    yearly_change = year_end - year_open
    vol = vol + ws.Cells(i, 7).Value
    If year_open = 0 Then
    percent_change = 0
    Else
    percent_change = Round((yearly_change / year_open) * 100, 2)
    End If
    year_open = ws.Cells(i + 1, 3).Value

    ws.Cells(row, 9).Value = ticker
    ws.Cells(row, 10).Value = yearly_change
    ws.Cells(row, 11).Value = percent_change
    ws.Cells(row, 12).Value = vol

    
    If ws.Cells(row, 11).Value >= 0 Then
    
    ws.Cells(row, 11).Interior.ColorIndex = 4
    Else
    ws.Cells(row, 11).Interior.ColorIndex = 3
    
    End If
    
    vol = 0
    row = row + 1

    Else

    vol = vol + ws.Cells(i, 7).Value
    
    End If

Next i

Next ws

End Sub


