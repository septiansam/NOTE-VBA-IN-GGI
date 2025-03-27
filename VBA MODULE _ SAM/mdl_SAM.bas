Attribute VB_Name = "mdl_SAM"

'+------------------------------------------------+
'        Modul: Mdl_SAM
' Deskripsi: Modul ini berisi berbagai prosedur
'           yang digunakan untuk manajemen sheet
'           dalam workbook Excel, seperti
'           menambah, menghapus, dan membersihkan
'           sheet sementara.
'+------------------------------------------------+

'+------------------------------------------------+
'        Subroutine: Add_Sheets
' Deskripsi: Prosedur ini digunakan untuk menambah
'           sheet baru ke workbook berdasarkan
'           array nama sheet yang diberikan.
'           Jika sheet dengan nama yang sama sudah
'           ada, sheet tersebut akan dihapus
'           terlebih dahulu.
' Parameter:
'   - arr_sheet_names: Array nama sheet yang
'                      akan ditambahkan.
'+------------------------------------------------+
Sub Add_Sheets(ParamArray arr_sheet_names() As Variant)
    Dim i As Integer
    Dim sheet_name As String
    Dim new_sheet As Worksheet, ws As Worksheet
    
    For i = LBound(arr_sheet_names) To UBound(arr_sheet_names)
        sheet_name = CStr(arr_sheet_names(i))
        
        '..Periksa apakah sheet sudah ada. Jika iya, hapus terlebih dahulu
        On Error Resume Next
        Set ws = Sheets(sheet_name)
        On Error GoTo 0
        
        If Not ws Is Nothing Then ws.Delete
        Set ws = Nothing
        
        ' Tambahkan sheet baru dengan nama yang diberikan
        Set new_sheet = Sheets.Add(after:=Sheets(Sheets.Count))
        new_sheet.Name = sheet_name
    Next i
End Sub

'+------------------------------------------------+
'        Subroutine: DeleteSheets_WithName
' Deskripsi: Prosedur ini digunakan untuk
'           menghapus sheet sementara yang
'           ada dalam array nama sheet yang
'           diberikan. Prosedur ini menggunakan
'           dictionary untuk menyimpan dan
'           memeriksa nama-nama sheet.
' Parameter:
'   - arr_sheet_names: Array nama sheet yang
'                      akan dihapus.
'+------------------------------------------------+
Sub DeleteSheets_WithName(ParamArray arr_sheet_names() As Variant)
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
    
'    Application.DisplayAlerts = True
End Sub

'+------------------------------------------------+
'        Subroutine: Delete_Sheets_Except_Assets
' Deskripsi: Prosedur ini digunakan untuk menghapus
'           semua sheet dalam workbook kecuali
'           sheet yang terdaftar sebagai sheet
'           aset penting. Sheet yang dikecualikan
'           didefinisikan dalam variabel SH_1
'           hingga SH_10.
'+------------------------------------------------+
Sub Delete_Sheets_Except_Assets()
    Application.DisplayAlerts = False
    Dim SH As Worksheet
    Dim SH_1 As String, SH_2 As String, SH_3 As String, SH_4 As String, SH_5 As String
    Dim SH_6 As String, SH_7 As String, SH_8 As String, SJ_9 As String, SH_10 As String
    
    ' Daftar nama sheet yang tidak akan dihapus
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
    
    ' Menghapus sheet kecuali sheet yang dikecualikan
    For Each SH In ThisWorkbook.Worksheets
        If SH.Name <> SH_1 And SH.Name <> SH_2 And SH.Name <> SH_3 And SH.Name <> SH_4 And SH.Name <> SH_5 And _
           SH.Name <> SH_6 And SH.Name <> SH_7 And SH.Name <> SH_8 And SH.Name <> SH_9 And SH.Name <> SH_10 Then
            
            SH.Delete
        
        End If
    Next SH
    
    Application.DisplayAlerts = True
End Sub

'+------------------------------------------------+
'        Subroutine: DeleteSheetsExcept
' Deskripsi: Prosedur ini menghapus semua sheet
'           kecuali sheet yang ada dalam parameter
'           array sheetNames. Sheet yang ingin
'           dipertahankan disimpan dalam dictionary
'           untuk pengecekan yang cepat.
' Parameter:
'   - sheetNames: Array nama sheet yang akan
'                 dipertahankan.
'+------------------------------------------------+
Sub DeleteSheetsExcept(ParamArray sheetNames() As Variant)
    Dim ws As Worksheet
    Dim sheetName As Variant
    Dim keepSheets As Object ' Gunakan dictionary untuk menyimpan sheet yang ingin dipertahankan
    Dim sheetExists As Boolean
    
    ' Membuat dictionary untuk sheet yang ingin dipertahankan
    Set keepSheets = CreateObject("Scripting.Dictionary")
    
    ' Memasukkan nama sheet ke dalam dictionary
    For Each sheetName In sheetNames
        keepSheets(CStr(sheetName)) = True
    Next sheetName
    
    ' Iterasi melalui setiap worksheet di workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Periksa apakah nama sheet ada di dictionary
        If Not keepSheets.exists(ws.Name) Then
            ws.Delete ' Hapus sheet jika tidak ada di dictionary
        End If
    Next ws
End Sub

'+------------------------------------------------+
'        Subroutine: ImportDataFile
' Deskripsi: Fungsi ini akan mencopy data
'            dari workbook lain ke dalam
'            workbook yang sedang aktif,
'            dari sheet ke 1 workbook lain
'            tersebut
' Parameter:
'   - WB_Dest : Lokasi Workbook yang akan menjadi destinasi file inport
'   - SH_Dest : Lokasi Sheet yang jadi acuan data import
'   - Rng_Dest: Lokasi File akan copyan akan di letakan.
'   - PathSource: Lokasi File SRC workbook yang akan di import.
'+------------------------------------------------+
Sub ImportDataFile(WB_Dest As Workbook, SH_Dest As Worksheet, Rng_Dest As Range, PathSource As String)
    Dim WB_SRC As Workbook
    Dim SH_SRC As Worksheet
    
    Set WB_SRC = Workbooks.Open(PathSource)
    Windows(WB_SRC.Name).Activate
    Set SH_SRC = WB_SRC.Worksheets(1)
    SH_SRC.Activate
    SH_SRC.AutoFilterMode = False
    SH_SRC.Cells.EntireColumn.Hidden = False
    SH_SRC.Cells.EntireRow.Hidden = False
    SH_SRC.Cells.Copy
    
    Windows(WB_Dest.Name).Activate
    SH_Dest.Activate
    Rng_Dest.PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SH_Dest.Cells.EntireColumn.AutoFit
    SH_Dest.Cells(1, 1).Select
    
    WB_SRC.Close False
    
    Set WB_SRC = Nothing
    Set SH_SRC = Nothing
End Sub

'+------------------------------------------------+
'        Subroutine: UnhideSheets
' Deskripsi: Fungsi ini memeriksa apakah sheet
'            dengan nama yang diberikan ada
'            di workbook atau tidak, dan
'            apakah sheet tersebut ter hide,
'            jika ter hide maka unhide.
' Parameter:
'   - sheetNames: Nama sheet yang ingin diperiksa.
'+------------------------------------------------+
Sub UnhideSheets(ParamArray sheetNames() As Variant)
    Dim ws As Worksheet
    Dim sheetName As Variant
    Dim keepSheets As Object ' Gunakan dictionary untuk menyimpan sheet yang ingin dipertahankan
    Dim sheetExists As Boolean
    
    ' Membuat dictionary untuk sheet yang ingin dipertahankan
    Set keepSheets = CreateObject("Scripting.Dictionary")
    
    ' Memasukkan nama sheet ke dalam dictionary
    For Each sheetName In sheetNames
        keepSheets(CStr(sheetName)) = True
    Next sheetName
    
    ' Iterasi melalui setiap worksheet di workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Periksa apakah nama sheet ada di dictionary
        If keepSheets.exists(ws.Name) Then
            ws.Visible = xlSheetVisible ' Unhide sheet jika ada di dictionary
        End If
    Next ws
    
End Sub

'+------------------------------------------------+
'        Subroutine: HideSheets
' Deskripsi: Fungsi ini memeriksa apakah sheet
'            dengan nama yang diberikan ada
'            di workbook atau tidak, dan
'            apakah sheet tersebut
'            visible (not hide), maka unhide.
' Parameter:
'   - sheetNames: Nama sheet yang ingin diperiksa.
'+------------------------------------------------+
Sub HideSheets(ParamArray sheetNames() As Variant)
    Dim ws As Worksheet
    Dim sheetName As Variant
    Dim keepSheets As Object ' Gunakan dictionary untuk menyimpan sheet yang ingin dipertahankan
    Dim sheetExists As Boolean
    
    ' Membuat dictionary untuk sheet yang ingin dipertahankan
    Set keepSheets = CreateObject("Scripting.Dictionary")
    
    ' Memasukkan nama sheet ke dalam dictionary
    For Each sheetName In sheetNames
        keepSheets(CStr(sheetName)) = True
    Next sheetName
    
    ' Iterasi melalui setiap worksheet di workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Periksa apakah nama sheet ada di dictionary
        If keepSheets.exists(ws.Name) Then
            ws.Visible = xlSheetHidden ' Hide sheet jika ada di dictionary
        End If
    Next ws
    
End Sub

'+------------------------------------------------+
'        Function: wsx
' Deskripsi: Fungsi ini memeriksa apakah sheet
'            dengan nama yang diberikan ada
'            di workbook atau tidak. Mengembalikan
'            nilai True jika sheet ada, dan False
'            jika tidak ada.
' Parameter:
'   - sheet_names: Nama sheet yang ingin diperiksa.
'+------------------------------------------------+
Function wsx(sheet_names As String) As Boolean
    On Error Resume Next
        wsx = Not Sheets(sheet_names) Is Nothing
    On Error GoTo 0
End Function

'+------------------------------------------------+
' Kode : vba code
' Deskripsi: Untuk Cek Sheet Ada Atau Tidak Menggunakan
' Parameter:
'   - Sebagai Contoh Nama Sheet yang akan di cek
'     adalah Sheet dengan nama TES
'+------------------------------------------------+
' Kode dibawah ini
' If Evaluate("isref('" & "TES" & "'!A1)") Then Sheets("TES").Delete

Sub MasukanKomen(WS As Worksheet, cellAddress As String, commentText As String)
    With WS.Range(cellAddress)
        ' Hapus komentar yang ada, jika ada
        If Not .Comment Is Nothing Then .Comment.Delete

        ' Tambahkan komentar baru
        .AddComment
        .Comment.Text Text:=commentText
    End With

    'CARA PAKE:
'    MasukanKomen SH1_CC3, "F6", "test"
End Sub
