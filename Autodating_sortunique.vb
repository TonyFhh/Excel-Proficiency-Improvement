Private Sub Worksheet_Change(ByVal Target As Range)
If Not Intersect(Target, Target.Worksheet.Range("B2:C" & ThisWorkbook.Worksheets(1).UsedRange.Rows.Count)) Is Nothing Then

'Used to sync with Vlookup in Sheet1, applying this to the worksheet where i will paste consolidation data
'Just paste in the stuff in B column onwards and this auto macro will add the date on A then sort descending and remove the duplicates
'Make sure A is formatted as Date else might not work (Paste too far down also weirdly not work

'TO DO: maybe can detect split files and do a sum...

    Application.ScreenUpdating = False
    currDT = Now() 'Get current date and time

    'For each cell that was changed, get row and assign value (this turns out to be Ax)
    For Each cel In Target
        sRow = cel.Row
        Cells(cel.Row, 1).Value = currDT
    Next cel

    'Sort Descending the range A:Z based on column A, accounting for headers
    Range("A:Z").Sort Key1:=Range("A1"), Order1:=xlDescending, Header:=xlGuess

    Application.EnableEvents = False 'This prevents worksheetchange sub from activating and causing infinite loop
    With Cells(1, 1).CurrentRegion
        .RemoveDuplicates Columns:=2, Header:=xlYes
    End With
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End If

End Sub
