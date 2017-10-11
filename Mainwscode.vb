Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Target.Worksheet.Range("B3:B" & Rows.Count)) Is Nothing Then
    Dim checkdup As Integer
    Debug.Print Target.Value
    checkdup = WorksheetFunction.CountIf(Range("B:B"), Target.Value)
    Debug.Print checkdup
    If checkdup = 2 Then
    Target.Interior.Color = RGB(189, 214, 238)
    Target.Font.ColorIndex = 1
    ElseIf checkdup >= 3 Then
    Target.Interior.Color = RGB(192, 0, 0)
    Target.Font.ColorIndex = 2
    End If
    Call Update

    End If
End Sub