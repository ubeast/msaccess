Public Function select_single_excel_file() As String
    Dim dir_location As String
    dir_location = CurDir()
    
    Dim f As Object 'FileDialog
    Set f = Application.FileDialog(3) 'msoFileDialogFilePicker
                                 ' 1 - msoFileDialogOpen
                                 ' 2 - msoFileDialogSaveAs
                                 ' 3 - msoFileDialogFolderPicker
                                 
    On Error GoTo Err_SomeName          ' Initialize error handling.
    
    With f
       .AllowMultiSelect = False
       .InitialFileName = dir_location
       .Filters.Clear
        .Filters.Add "Excel", "*.xlsx*"
       .Show
       
    End With
    
    select_single_excel_file = f.SelectedItems(1)
    
    Exit Function
    
Err_SomeName:
    If Err.Number = 5 Then Debug.Print "No file selected"
End Function
