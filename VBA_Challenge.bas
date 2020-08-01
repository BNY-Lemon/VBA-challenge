Attribute VB_Name = "Module1"
Sub main()
Dim tickerColumn, summaryRow As Integer
Dim ticker As String
Dim volume, lRow As Long
Dim yearlyChange, percentChange, nextOpen, eoyClose As Double
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet


For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"

    tickerColumn = 1
    summaryRow = 2
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    nextOpen = Range("C2").Value
    eoyClose = 0
    
    For i = 2 To lRow
        If Cells(i + 1, tickerColumn).Value <> Cells(i, tickerColumn).Value Then
            ticker = Cells(i, tickerColumn).Value
            volume = volume + Cells(i, 7).Value
            eoyClose = Range("F" & i).Value
            yearlyChange = eoyClose - nextOpen
            If yearlyChange = 0 Then
                percentChange = 0
            ElseIf nextOpen = 0 Then
                nextOpen = 0.01
                percentChange = Application.WorksheetFunction.Round((yearlyChange / nextOpen), 4)
            Else
                percentChange = Application.WorksheetFunction.Round((yearlyChange / nextOpen), 4)
            End If
            Range("I" & summaryRow).Value = ticker
            Range("J" & summaryRow).Value = yearlyChange
            Range("K" & summaryRow).Value = percentChange
            Range("L" & summaryRow).Value = volume
            'reset the totals
            volume = 0
            'set summary to the next row
            summaryRow = summaryRow + 1
            'set next open
            nextOpen = Range("c" & i + 1).Value
        Else
            'we are in the same ticker, so add the value
            volume = volume + Cells(i, 7).Value
        End If
    Next i
    'Conditional Formatting
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveWindow.SmallScroll Down:=-99
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 1
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    
Next
starting_ws.Activate
End Sub



