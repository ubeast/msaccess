Option Compare Database
Option Explicit

Sub import_excel_file(excel_filename As String, access_table_name As String)
    Dim file_extention As String
    file_extention = ".xlsx"
   
    ' If the ".xlsx" extension was not included in the "ingest_file_name" then add
    If Right(excel_filename, Len(file_extention)) <> file_extention Then _
        excel_filename = excel_filename + ".xlsx"
    
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, access_table_name, excel_filename, True
    MsgBox """" & excel_filename & """" & " was successfully imported to " & """" & access_table_name & """"

End Sub
