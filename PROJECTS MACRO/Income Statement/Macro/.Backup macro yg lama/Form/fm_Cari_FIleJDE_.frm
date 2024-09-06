VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fm_Cari_FIleJDE_ 
   Caption         =   "START WIZARD (2)"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "fm_Cari_FIleJDE_.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fm_Cari_FIleJDE_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Txt_LokasiFile.Value = ThisWorkbook.Worksheets("(MENU)").Range("E13").Value
End Sub

Private Sub Lbl_CariFIle_Click()
    Dim FilePicker As FileDialog
    Dim myFile As String

    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)
    With FilePicker
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.csv"
        .Title = "Select a File 'File Budget'"
        .AllowMultiSelect = False
        
        .InitialFileName = ThisWorkbook.Worksheets("(MENU)").Range("E13").Value

        If .Show <> -1 Then Exit Sub
        myFile = .SelectedItems(1)
        ThisWorkbook.Worksheets("(MENU)").Range("E13").Value = myFile
        Txt_LokasiFile.Value = myFile
    End With
End Sub

Private Sub Btn_Next_Click()
    Dim Pesan As Integer
    Pesan = MsgBox("Macro akan memproses Laporan periode bulan: " & ThisWorkbook.Worksheets("(MENU)").Range("E12").Value & vbNewLine & _
    "Lolasi Tarikan JDE: " & Txt_LokasiFile.Value & vbCrLf & _
    "Lanjutkan ?", vbYesNo + vbQuestion, "Hapus Data")

    If Pesan = vbYes Then
'        Call HapusPeriode
        ThisWorkbook.Worksheets("(MENU)").Range("E13").Value = Txt_LokasiFile.Value
        MsgBox "Proses Selesai, Data tersimpan di Macro", vbInformation, "Execute"
    Else:
        MsgBox "Cancel", vbInformation, "Execute"
        End
    End If
End Sub


