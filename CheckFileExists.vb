Option Explicit
Sub CheckSharedFiles()
'Useful for eCompare, if table of reports to be compared is available, macro will iterate each file
'and check if the results are already shared in the shared drive, can be configured to possibly check
'for src/tgt as well

Dim lastrowB As Long
Dim rList As Range, r As Range
Dim fullfilepath As String

lastrowB = Cells(Rows.Count, "B").End(xlUp).Row
Set rList = Range("B2:B" & lastrowB) 'Define range from B2:B up till the column with value

For Each r In rList 'Iterate each row in B2:B

'A3 and A4 contain different information on full file path
fullfilepath = Range("J1").Value & "\" & r.Value & ".txt" 'Concatenates various strings together

'works even on shared drive
If Dir(fullfilepath) <> "" Then 'If file exists, based on cmd Dir command
    r.Offset(0, 1).Value = "Done on " & FileDateTime(fullfilepath) 'Paste value on column one to the right
Else
    r.Offset(0, 1).Value = "Not shared"
End If
Next r

End Sub
