Option Compare Database
Option Explicit

Function select_file()
    Dim f As Object 'FileDialog
    Set f = Application.FileDialog(3) 'msoFileDialogFilePicker
                                 ' 1 - msoFileDialogOpen
                                 ' 2 - msoFileDialogSaveAs
                                 ' 3 - msoFileDialogFolderPicker
    With f
       .AllowMultiSelect = False 'default
    '   .InitialFileName = "C:\Temp\"
    '   Specify filters
    '  .Filters.Clear
    '  .Filters.Add "All Files", "*.xlsx*"
       .Show
       
'       Dim varFile As Variant
'       For Each varFile In .SelectedItems
'          MsgBox Trim(varFile)
'       Next
    End With
    
    select_file = f.SelectedItems(1)
End Function

