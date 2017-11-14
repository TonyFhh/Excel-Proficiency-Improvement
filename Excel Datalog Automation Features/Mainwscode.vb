Private Sub Worksheet_Change(ByVal Target As Range)
'The following code only runs if values of cells from column B changes
    If Not Intersect(Target, Target.Worksheet.Range("B3:B" & Rows.Count)) Is Nothing Then
    Dim checkdup As Integer
    Dim cel As Range
    Dim sRow As Long

' Work compaitability with multiple entries, however countif checks against entire range that was changed, chronological
' appearances of entries are not correctly labelled for frequency.
For Each cel In Target 'for each row in range of changed cells...
    sRow = cel.Row
    checkdup = WorksheetFunction.CountIf(Range("B:B"), Range("B" & sRow).Value) 'Check frequency of changed row
    ' if frequency > x, change cell colour and font colour to highlight
    If checkdup = 2 Then
    Range("B" & sRow).Interior.Color = RGB(189, 214, 238)
    Range("B" & sRow).Font.ColorIndex = 1
    ElseIf checkdup >= 3 Then
    Range("B" & sRow).Interior.Color = RGB(192, 0, 0)
    Range("B" & sRow).Font.ColorIndex = 2
    Else
    Range("B" & sRow).Interior.ColorIndex = 0 'Restore to default if new entry with no prior occurance
    Range("B" & sRow).Font.ColorIndex = 1
    End If
    Next cel
    Call Update 'update frequency table


    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range) 'Auto add new ID entry on next row selection
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, "A").End(xlUp).Offset(1).Row 'Find last entry row then offset by 1 row down
'    Debug.Print lastrow
    If Target.Row = lastrow Then 'if selected cell is in new entry row
    Range("A3").AutoFill Destination:=Range("A3:A" & lastrow), Type:=xlFillSeries 'auto fill ID row
'    lastrow = Cells(Rows.Count, "A").End.Row
    End If
End Sub
