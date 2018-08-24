Option Explicit
Sub DraftSplit()

' Algorithm flow
' 1. Assuming the col "A" is always S/No, starting from A7 i'll find the last value of the row
' 2. Take a division by 900 and this will be number of copies needed
' 3. Create the number of copies as needed
' 4. Trim the range accordingly

' Other notes
' will be helpful if this macro is based on tool_do.xlsm (works if the wb is active screen when running)

' in VBA interface...
' ENSURE that in Tools -> References "Microsoft Scripting Reference" is checked ELSE THIS MACRO WILL NOT WORK!!

Dim lastRowA As Long
Dim lastSno As Long
Dim noSplit As Long

Dim wbName As String, wbExt As String
Dim splitName As String

Dim startRow As Long, endRow As Long
Dim wb As Workbook, sh As Worksheet


Application.ScreenUpdating = False

'Relabel the S/N of all the cells, in case the user forgots to, this will prevent uneeded errors
Range("A7").Value = 1
Range("A8").Value = 2
'Fill relative to B, since it's always has value "Onetime New"
Range("A7:A8").AutoFill Destination:=Range("A7:A" & Cells(Rows.Count, "B").End(xlUp).Row), Type:=xlFillDefault

ActiveWorkbook.Save


lastRowA = Cells(Rows.Count, "A").End(xlUp).Row
'lastCol6 = Cells(6, Columns.Count).End(xlToLeft).Column

Debug.Print "lastRowA is"; lastRowA;

'Check what is the last value in the S/No Row
lastSno = Cells(lastRowA, 1).Value
Debug.Print "lastsno is"; lastSno

noSplit = Application.WorksheetFunction.RoundUp(lastSno / 900, 0) 'Get no. of Splits needed, always rounded up
Debug.Print "NoSplit is"; noSplit

' If there is a need to split (S/No > 900)
If noSplit > 1 Then

    'FileSystemObject is a library from Microsoft Scripting Runtime
    ' We create an object of it here...
    ' then use it here to get FileName w/o ext, get ext and do file copy procedures
    Dim filesysobj As New Scripting.FileSystemObject
    wbName = filesysobj.GetBaseName(ActiveWorkbook.Name)
    wbExt = filesysobj.GetExtensionName(ActiveWorkbook.Name)
'    Debug.Print "wbName is"; wbName
'    Debug.Print "wbExt is"; wbExt
'    Debug.Print filesysobj.GetAbsolutePathName(".")
    
    'Loop across the no of splits needed
    Dim i As Integer
    For i = 1 To noSplit
        splitName = wbName & "_part" & i & "." & wbExt 'Splits will be named "<File>_part[x].ext"
'        Debug.Print "splitName is"; splitName
'        Debug.Print "ActiveWorkbookname is"; ActiveWorkbook.Name
'        Debug.Print ActiveWorkbook.Path
        filesysobj.CopyFile ActiveWorkbook.Path & "\" & ActiveWorkbook.Name, ActiveWorkbook.Path & "\" & splitName

        
        'Calculate the starting and ending rows of each split
        startRow = (i - 1) * 900 + 7
            
        If i = noSplit Then
            endRow = lastRowA
        Else
            endRow = i * 900 + 6
        End If
            
        'Then we modify the splitted excel files,
        'By deleting the rows above + below that of the start/end range
        Set wb = Workbooks.Open(ActiveWorkbook.Path & "\" & splitName)
        
        wb.Application.EnableEvents = False  'This appeared to got rid of the annoying msg boxes
        wb.Application.DisplayAlerts = False 'while this didn't work
        wb.Application.ScreenUpdating = False
        
        Set sh = wb.Sheets(1)
        
            
        'Special cases for modifying the first and last split
        
        ' We won't delete below if last split
        If i <> noSplit Then
'            Range(Cells(endRow + 1, 1), Cells(lastRowA, lastCol6)).Delete Shift:=xlUp
            Rows(endRow + 1 & ":" & lastRowA).Delete
        End If
        
        ' and likewise We won't delete above if first split
        If i <> 1 Then
'            Range("A7", Cells(startRow - 1, lastCol6)).Delete Shift:=xlUp 'Delete the range stated, shifting the cells up
            Rows(7 & ":" & startRow - 1).Delete Shift:=xlUp 'Delete the stated rows, shifting cells up
'            wb.Application.SendKeys "{ENTER}"
        End If

        'Repopulate the Serial No from 1 to 900
        Range("A7").Value = 1
        Range("A8").Value = 2
        Range("A7:A8").AutoFill Destination:=Range("A7:A" & Cells(Rows.Count, "A").End(xlUp).Row), Type:=xlFillDefault
'

'        Debug.Print "Before Close wb"
        wb.Application.DisplayAlerts = True
        wb.Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        wb.Close SaveChanges:=True
'        Debug.Print "wb Closed successfully"
    Next i
End If

Application.ScreenUpdating = True
End Sub