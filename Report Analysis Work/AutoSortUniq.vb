Private Sub Worksheet_Change(ByVal Target As Range)
If Not Intersect(Target, Target.Worksheet.Range("A2:C" & Rows.Count)) Is Nothing Then

'Used to sync with Vlookup in Sheet1, applying this to the worksheet where i will paste consolidation data
'Just paste in the stuff in B column onwards and this auto macro will add the date on A then sort descending and remove the duplicates
'Make sure A is formatted as Date else might not work (Paste too far down also weirdly not work


    Application.ScreenUpdating = False
'    Sort Descending the range A:Z based on column A, accounting for headers
    Range("A2:Z999999").Sort Key1:=Range("B2"), Order1:=xlDescending, Header:=xlNo

    Application.EnableEvents = False 'This prevents worksheetchange sub from activating and causing infinite loop
    With Cells(1, 1).CurrentRegion
        .RemoveDuplicates Columns:=1, Header:=xlYes
    End With
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End If

End Sub
