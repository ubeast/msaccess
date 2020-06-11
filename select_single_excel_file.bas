Public Sub import_excel_file()

    Dim ingest_file_with_path As String
    ingest_file_with_path = select_single_excel_file()
    
    ' If user cancels the import exit this process
    If Len(ingest_file_with_path) < 5 Then Exit Sub
    
Continue:
    ' Get filename from "ingest_file_name" which includes the path
    Dim access_table_name, ingest_file As String
    ingest_file = get_filename(ingest_file_with_path)
    access_table_name = ingest_file
    
    ' Remove the extension from the filename
    access_table_name = Split(access_table_name, ".")(0)
    
    
Filename_Cleanup:
    ' Remove parenths "()"
    access_table_name = Replace(Replace(access_table_name, "(", ""), ")", "")
        
    ' Truncate characters beyond 50
    access_table_name = Left(access_table_name, 50)
        
    ' Replace spaces with dash "_", place in lowercase, and trim spaces
    access_table_name = Trim(LCase(Replace(access_table_name, " ", "_")))
        
Import_the_file:
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, access_table_name, ingest_file_with_path, True
    
Display_MsgBox:
    Dim msg As String
    msg = """" & _
          ingest_file & _
          """" & vbCr & vbCr & _
          " was successfully imported to " & _
          """" & vbCr & vbCr & _
          access_table_name & _
          """"
    MsgBox msg
    
View_the_file:
    DoCmd.OpenTable access_table_name, acViewNormal

End Sub
