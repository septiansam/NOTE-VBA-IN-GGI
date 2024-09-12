Attribute VB_Name = "btn_SelectFIle"
'____________________________________________________________________________________________________
'#SHORTCHUT   -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'----------------------------------------------------------------------------------------------------
'Sub OPEN_MainMenu()
''Keyboard Shortcut: Ctrl+Shift+Z
'    Fm_MenuUtama.Show
'End Sub
'
'Sub OPEN_About()
''Keyboard Shortcut: Ctrl+Shift+Z
'    fm_Help.Show
'End Sub


''____________________________________________________________________________________________________
''#BUTTON SHEET   -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
''----------------------------------------------------------------------------------------------------


Sub BUTTON_CariFile_Ledger()
    Dim FilePicker As FileDialog
    Dim myFile As String

    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx"
        .Title = "Select a File 'File Ledger'"
        .AllowMultiSelect = False

        ' Set lokasi default
        '.InitialFileName = ThisWorkbook.Worksheets("HOME").Range("E13").Value
'        .InitialFileName = ThisWorkbook.Worksheets("(HOME)").Range("E13").Value & "\FILE.xlsx"

        If .Show <> -1 Then Exit Sub
        myFile = .SelectedItems(1)
        ThisWorkbook.Worksheets("HOME").Range("E15").value = myFile
    End With
End Sub

Sub BUTTON_CariFile_ShipmentReport()
    Dim FilePicker As FileDialog
    Dim myFile As String

    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx"
        .Title = "Select a File 'File Shipment Report'"
        .AllowMultiSelect = False

        ' Set lokasi default
        '.InitialFileName = ThisWorkbook.Worksheets("HOME").Range("E13").Value
'        .InitialFileName = ThisWorkbook.Worksheets("(HOME)").Range("E13").Value & "\FILE.xlsx"

        If .Show <> -1 Then Exit Sub
        myFile = .SelectedItems(1)
        ThisWorkbook.Worksheets("HOME").Range("E16").value = myFile
    End With
End Sub
'
'
Sub BUTTON_CariFile_Report_WO_Buyer()
    Dim FilePicker As FileDialog
    Dim myFile As String

    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx"
        .Title = "Select a File 'File Report WO Buyer'"
        .AllowMultiSelect = False

        ' Set lokasi default
        '.InitialFileName = ThisWorkbook.Worksheets("HOME").Range("E13").Value
'        .InitialFileName = ThisWorkbook.Worksheets("(HOME)").Range("E13").Value & "\FILE.xlsx"

        If .Show <> -1 Then Exit Sub
        myFile = .SelectedItems(1)
        ThisWorkbook.Worksheets("HOME").Range("E17").value = myFile
    End With
End Sub

Sub BUTTON_CariFile_Inventory_Aging()
    Dim FilePicker As FileDialog
    Dim myFile As String

    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx"
        .Title = "Select a File 'File Inventory Aging'"
        .AllowMultiSelect = False

        ' Set lokasi default
'        .InitialFileName = "\\10.8.0.35\Bersama\santy\CLOSING bersama"

        If .Show <> -1 Then Exit Sub
        myFile = .SelectedItems(1)
        ThisWorkbook.Worksheets("HOME").Range("E18").value = myFile
    End With
End Sub
'
'

Sub BUTTON_CariFile_SalesLocal()
    Dim FilePicker As FileDialog
    Dim myFile As String

    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx"
        .Title = "Select a File 'File Sales Local'"
        .AllowMultiSelect = False

        ' Set lokasi default
'        .InitialFileName = "\\10.8.0.35\Bersama\santy\CLOSING bersama"

        If .Show <> -1 Then Exit Sub
        myFile = .SelectedItems(1)
        ThisWorkbook.Worksheets("HOME").Range("E19").value = myFile
    End With
End Sub


Sub BUTTON_CariFolder_TarikanJDE()
    Dim FldrPicker As FileDialog
    Dim myFolder As String

    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With FldrPicker
        .Title = "Select A Target Folder 'Tarikan JDE'"
        .AllowMultiSelect = False
'        .InitialFileName = ThisWorkbook.Worksheets("HOME").Range("E14").Value
        If .Show <> -1 Then Exit Sub
        myFolder = .SelectedItems(1)
        ThisWorkbook.Worksheets("HOME").Range("E14").value = myFolder
    End With
End Sub

'Sub BUTTON_CariFolder_HasilMacro()
'    Dim FldrPicker As FileDialog
'    Dim myFolder As String
'
'    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
'    With FldrPicker
'        .Title = "Select A Target Folder 'HASIL MACRO'"
'        .AllowMultiSelect = False
''        .InitialFileName = ThisWorkbook.Worksheets("HOME").Range("E14").Value
'        If .Show <> -1 Then Exit Sub
'        myFolder = .SelectedItems(1)
'        ThisWorkbook.Worksheets("HOME").Range("E20").Value = myFolder
'    End With
'End Sub

Sub BUTTON_CariFile_HasilMacro()
    Dim FilePicker As FileDialog
    Dim myFile As String

    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx"
        .Title = "Select a File 'File HASIL MAKRO'"
        .AllowMultiSelect = False

        ' Set lokasi default
'        .InitialFileName = "\\10.8.0.35\Bersama\santy\CLOSING bersama"

        If .Show <> -1 Then Exit Sub
        myFile = .SelectedItems(1)
        ThisWorkbook.Worksheets("HOME").Range("E20").value = myFile
    End With
End Sub


'
'
'Sub BUTTON_OpenFile_RemarkHistory()
'    On Error GoTo ErrorHandling
'        Workbooks.Open (ThisWorkbook.Worksheets("(HOME)").Range("E18").value)
'    End
'ErrorHandling:
'   MsgBox "Number: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error Handling"
'End Sub
'
'
'
''____________________________________________________________________________________________________
''#BUTTON FORM   -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
''----------------------------------------------------------------------------------------------------
'Private Sub btn_HapusDataPeriode_Click()
'    Dim Pesan As Integer
'    Pesan = MsgBox("Proses ini akan menghapus data Kondisi Perusahaan periode" & vbNewLine & _
'    lbl_NamaBulan.Caption & vbCrLf & _
'    "Lanjutkan ?", vbYesNo + vbQuestion, "Hapus Data")
'
'    If Pesan = vbYes Then
'        Call HapusPeriode
'        MsgBox "Done", vbInformation, "Hapus Data"
'    Else:
'        MsgBox "Cancel", vbInformation, "Hapus Data"
'    End If
'
'End Sub
'
'
'Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'    If Target.Column = 5 And Target.row = 14 Then ' Menggunakan nomor kolom (E) alih-alih huruf (E)
'        fm_Periode.Show
'        Cancel = True ' Menonaktifkan aksi standar double-click pada sel
'    End If
'
'    If Target.Column = 5 And Target.row = 15 Then ' Menggunakan nomor kolom (E) alih-alih huruf (E)
'        fm_Periode.Show
'        Cancel = True ' Menonaktifkan aksi standar double-click pada sel
'    End If
'End Sub
'








