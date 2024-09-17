Attribute VB_Name = "Mdl_SAM"
'+-------------------------------+
'        My Code SAM
'+-------------------------------+

Sub Add_Sheets_Preprocessing(ParamArray arr_sheet_names() As Variant) 'SEPAKET DENGAN FUNGSI WSX
    Dim i As Integer
    Dim sheet_name As String
    Dim new_sheet As Worksheet, ws As Worksheet
    
    For i = LBound(arr_sheet_names) To UBound(arr_sheet_names)
        sheet_name = CStr(arr_sheet_names(i))
        
        '..Periksa Apakah Sheets Sudah Ada. Jika Iya Maka Hapus Terlebih Dahulu"
        On Error Resume Next
        Set ws = Sheets(sheet_name)
        On Error GoTo 0
        
        If Not ws Is Nothing Then ws.Delete
        Set ws = Nothing
        
        Set new_sheet = Sheets.Add(after:=Sheets(Sheets.Count))
        new_sheet.Name = sheet_name
    Next i
End Sub

Sub Clear_Temporary_Sheets(ParamArray arr_sheet_names() As Variant) ' As String
    Application.DisplayAlerts = False
    
    Dim SH As Worksheet
    Dim i As Integer
    Dim sheetName As String
    Dim sheetNamesDict As Object
    
    ' Membuat dictionary untuk menyimpan nama sheet
    Set sheetNamesDict = CreateObject("Scripting.Dictionary")
    
    ' Memasukkan nama-nama sheet ke dalam dictionary
    For i = LBound(arr_sheet_names) To UBound(arr_sheet_names)
        sheetNamesDict(CStr(arr_sheet_names(i))) = True
    Next i
    
    ' Iterasi melalui sheet dari belakang
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        sheetName = ThisWorkbook.Sheets(i).Name
        ' Memeriksa apakah nama sheet ada dalam dictionary
        If sheetNamesDict.exists(sheetName) Then
            ThisWorkbook.Sheets(i).Delete
        End If
    Next i
    
    Application.DisplayAlerts = True
End Sub


Sub Delete_Sheets_Except_Assets()
    Application.DisplayAlerts = False
    Dim SH As Worksheet
    Dim SH_1 As String, SH_2 As String, SH_3 As String, SH_4 As String, SH_5 As String
    Dim SH_6 As String, SH_7 As String, SH_8 As String, SJ_9 As String, SH_10 As String
'~  >>>> -SHEET ASSET- <<<<  ~
    SH_1 = "HOME"
    SH_2 = "SetupDB"
    SH_3 = ""
    SH_4 = ""
    SH_5 = ""
    SH_6 = ""
    SH_7 = ""
    SH_8 = ""
    SH_9 = ""
    SH_10 = ""
'`````````````````````````````
'~  >>>> -DELETE SHEETS- <<<<  ~
    For Each SH In ThisWorkbook.Worksheets
        If SH.Name <> SH_1 And SH.Name <> SH_2 And SH.Name <> SH_3 And SH.Name <> SH_4 And SH.Name <> SH_5 And _
           SH.Name <> SH_6 And SH.Name <> SH_7 And SH.Name <> SH_8 And SH.Name <> SH_9 And SH.Name <> SH_10 Then
            
            SH.Delete
        
        End If
    Next SH
'`````````````````````````````
End Sub

Sub DeleteSheetsExcept(ParamArray sheetNames() As Variant)
    Dim ws As Worksheet
    Dim sheetName As Variant
    Dim keepSheets As Object ' Use dictionary to store sheets to keep
    Dim sheetExists As Boolean
    
    ' Create a dictionary for sheet names to keep
    Set keepSheets = CreateObject("Scripting.Dictionary")
    
    ' Add each sheet name to the dictionary
    For Each sheetName In sheetNames
        keepSheets(CStr(sheetName)) = True
    Next sheetName
    
'    ' Disable alerts to prevent confirmation dialogs when deleting sheets
'    Application.DisplayAlerts = False
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet name exists in the dictionary
        If Not keepSheets.exists(ws.Name) Then
            ws.Delete ' Delete sheet if it's not in the dictionary
        End If
    Next ws
    
'    ' Re-enable alerts after the process is complete
'    Application.DisplayAlerts = True
End Sub

Function wsx(sheet_names As String) As Boolean
    On Error Resume Next
        wsx = Not Sheets(sheet_names) Is Nothing
    On Error GoTo 0
End Function
