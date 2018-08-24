Sub CleanTnLSpaces()
'This macro will search all cells and trim leading and trailing spaces, ignoring cells that contains formulas

  Dim r As Range
  For Each r In ActiveSheet.UsedRange
    v = r.Value
    If v <> "" Then
      If Not r.HasFormula Then
        r.Value = Trim(v)
      End If
    End If
  Next r
End Sub
