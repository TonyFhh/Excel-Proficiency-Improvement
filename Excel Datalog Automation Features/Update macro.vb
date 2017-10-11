Option Explicit
Sub Update()
    Dim rList As Range
    Dim rData As Range
    Dim c As Range, r As Range
    Dim Freq As Integer
    Dim dRow As Integer
    Dim tRow As Integer
    Dim lastrowN As Long
    Dim lastrowB As Long
    
    lastrowB = Cells(Rows.Count, "B").End(xlUp).Row

Set rList = Range("B3:B" & lastrowB)
Set rData = Range("N:N")
dRow = 8
tRow = 3

Application.ScreenUpdating = False
'Clear cell contents of output row before recalculating
Range("N8:N" & Rows.Count).ClearContents
Range("O8:O" & Rows.Count).ClearContents
'Within output range N, scans each row along column B, determining frequency of occurance along B
'via Countif function. If freq >1 then search along N for occurance, if not found then input row
'value and freq in unused line along N. Repeat for new row until end
With rData
For Each r In rList
    If Not r = "" Then
    Freq = Application.WorksheetFunction.CountIf(rList, r.Value) 'Works correctly up to here
    'Now if duplicates are found for that row entry, search range (N:N) aka rData for existing entry,
    'if not found (c is Nothing) then input details
    If Freq > 1 Then
    Set c = .Find(what:=r.Value, LookIn:=xlValues, Lookat:=xlWhole, _
            SearchDirection:=xlNext)
'            Debug.Print
    If c Is Nothing Then
    .Cells(dRow, 1).Value = r.Value 'Note cells row column references are relative to with range
    .Cells(dRow, 2).Value = Freq
    dRow = dRow + 1
    End If
    End If
    End If
Next r
    
End With
'Sort column N data based on frequency and order of occcurance
lastrowN = Cells(Rows.Count, "N").End(xlUp).Row
Range("N8:O" & lastrowN).Sort key1:=Range("O8:O" & lastrowN), order1:=xlDescending, Header:=xlNo
Application.ScreenUpdating = True
    
End Sub