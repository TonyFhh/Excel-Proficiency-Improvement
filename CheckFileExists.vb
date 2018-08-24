Option Explicit
Sub CheckSharedFilesRef()

Dim lastrowB As Long
Dim rList As Range, r As Range
Dim fullfilepath As String

lastrowB = Cells(Rows.Count, "B").End(xlUp).Row
Set rList = Range("B5:B" & lastrowB)

For Each r In rList
    fullfilepath = Range("B2").Value & "\" & r.Value & ".xlsx"
    
    If Dir(fullfilepath) <> "" Then
        r.Offset(0, 5).Value = FileDateTime(fullfilepath)
    Else
        r.Offset(0, 5).Value = "Not shared"
    End If
Next r

End Sub