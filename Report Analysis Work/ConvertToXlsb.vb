Sub ConvertToXLSB()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    'Dim dialog As FileDialog:
    Set Dialog = Application.FileDialog(msoFileDialogFolderPicker)
    With Dialog
        .Title = "Select Folder"
        .AllowMultiSelect = False
        .Show
    End With

    If Dialog.SelectedItems.Count = 0 Then End
    Dim curFile As String, wb As Workbook, convDir As String: convDir = Dialog.SelectedItems(1)

    MkDir (convDir & "\converted")
    curFile = Dir(convDir & "\*.xl*")
    While Len(curFile) > 0
        Set wb = Workbooks.Open(convDir & "\" & curFile)
        wb.SaveAs convDir & "\converted\" & Left(curFile, InStrRev(curFile, ".")) & "xlsb", 50 '50 is the file format code in VBA, corresponding to xlsb format
       wb.Close
        curFile = Dir
    Wend
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
