Attribute VB_Name = "Mdl_Main_Program"

Dim TWB As Workbook
Dim WB_Source As Workbook
Dim HOME As Worksheet, TMP1 As Worksheet, TMP2 As Worksheet, PESAN As Worksheet, SH As Worksheet
Dim path_src As String, lr As Long, lc As Long, jumlah_Request As Long, rowPaste As Long
Dim isi_pesan As String, arr_penerima_pesan As Variant, penerima As String


Sub Proses()
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
        Call Validation
        Call Delete_Sheets_Except_Assets
        Call Add_Sheets_Preprocessing("TMP1", "TMP2", "PESAN")
        Call Initialization
        Call Import_File
        Call Processing
        Call Clear_Temporary_Sheets("TMP1", "TMP2")
        Call End_Process
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
End Sub

Sub End_Process()
    HOME.Activate
    Range("A1").Select
End Sub

Sub Validation()
    Set TWB = ThisWorkbook
    Set HOME = TWB.Sheets("HOME")
    path_src = HOME.Range("H8").Value
    If Dir(path_src) = "" Then
        MsgBox "Source File, Doesn't Exists", vbInformation, "File Not Found"
        Exit Sub
    End If
End Sub

Sub Initialization()
    Set TMP1 = TWB.Sheets("TMP1")
    Set TMP2 = TWB.Sheets("TMP2")
    Set PESAN = TWB.Sheets("PESAN")
End Sub

Sub Import_File()
    Set WB_Source = Workbooks.Open(path_src)
    Windows(WB_Source.Name).Activate
    Set SH = WB_Source.Sheets(1): SH.AutoFilterMode = False
    Cells.Copy
    Windows(TWB.Name).Activate
    TMP1.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    WB_Source.Close False
End Sub

Sub Add_Sheets_Preprocessing(ParamArray arr_sheet_names() As Variant) 'SEPAKET DENGAN FUNGSI WSX
    Dim i As Integer
    Dim sheet_name As String
    Dim new_sheet As Worksheet
    
    For i = LBound(arr_sheet_names) To UBound(arr_sheet_names)
        sheet_name = CStr(arr_sheet_names(i))
        
        If wsx(sheet_name) Then Sheets(sheet_name).Delete
        Set new_sheet = Sheets.Add(AFTER:=Sheets(Sheets.Count))
        new_sheet.Name = sheet_name
    Next i
End Sub

Function Clear_Temporary_Sheets(ParamArray arr_sheet_names() As Variant) As String
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
End Function


Sub Delete_Sheets_Except_Assets()
    Application.DisplayAlerts = False
    Dim SH As Worksheet
    Dim SH_1 As String, SH_2 As String, SH_3 As String, SH_4 As String, SH_5 As String
    Dim SH_6 As String, SH_7 As String, SH_8 As String, SJ_9 As String, SH_10 As String
'~  >>>> -SHEET ASSET- <<<<  ~
    SH_1 = "HOME"
    SH_2 = ""
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

Function wsx(sheet_names As String) As Boolean
    On Error Resume Next
        wsx = Not Sheets(sheet_names) Is Nothing
    On Error GoTo 0
End Function
