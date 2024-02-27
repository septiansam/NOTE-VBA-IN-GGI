Attribute VB_Name = "Mdl_Function"
'
'
'

Public Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not Sheets(shtName) Is Nothing
    On Error GoTo 0
End Function
