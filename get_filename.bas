Option Compare Database
Option Explicit

Public Function get_filename(ByVal path_filename As String) As String
    last_item = UBound(split(path_filename, "\"))
    get_filename = split(path_filename, "\")(last_item)
End Function
