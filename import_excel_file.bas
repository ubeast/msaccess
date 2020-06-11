Public Sub import_excel_file()
    Dim ingest_file_name As String
    ingest_file_name = select_single_excel_file()
    
    ' Get filename from "ingest_file_name" which includes the path
    Dim access_table_name As String
    access_table_name = get_filename(ingest_file_name)
    
    
    ' Remvoe the extension from the filenamne
    access_table_name = split(access_table_name, ".")(0)
    
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, access_table_name, ingest_file_name, True
    ' MsgBox """" & ingest_file_name & """" & " was successfully imported to " & """" & access_table_name & """"
    DoCmd.OpenTable access_table_name, acViewNormal
End Sub
