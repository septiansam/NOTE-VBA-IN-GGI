Attribute VB_Name = "MDL_Any"
'____________________________________________________________________________________________________
'#RESET   -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'----------------------------------------------------------------------------------------------------
Sub DELETE_AllSheetOnlyAsset()
    Dim SH As Worksheet
    Dim SH_1 As String, SH_2 As String, SH_3 As String, SH_4 As String, SH_5 As String
    Dim SH_6 As String, SH_7 As String, SH_8 As String, SH_9 As String, SH_10 As String
'---SHEET ASSET:
    SH_1 = ("HOME")
    SH_2 = ("")
    SH_3 = ("")
    SH_4 = ("")
    SH_5 = ("")
    SH_6 = ("")
    SH_7 = ("")
    SH_8 = ("")
    SH_9 = ("")
    SH_10 = ("")
'    Application.DisplayAlerts = False ' Matikan peringatan
    For Each SH In ThisWorkbook.Sheets
        If SH.Name <> SH_1 And SH.Name <> SH_2 And SH.Name <> SH_3 And SH.Name <> SH_4 And SH.Name <> SH_5 And SH.Name <> SH_6 And SH.Name <> SH_7 And SH.Name <> SH_8 And SH.Name <> SH_9 And SH.Name <> SH_10 Then
            SH.Delete
        End If
    Next SH
'    Application.DisplayAlerts = True
End Sub

Function fn_HapusSheet(SH1 As String, SH2 As String, SH3 As String, SH4 As String, SH5 As String, SH6 As String, SH7 As String, SH8 As String, SH9 As String, SH10 As String) As String
    Dim SH As Worksheet
    Application.DisplayAlerts = False
    For Each SH In ThisWorkbook.Sheets
        If SH.Name = SH1 Or SH.Name = SH2 Or SH.Name = SH3 Or SH.Name = SH4 Or SH.Name = SH5 Or SH.Name = SH6 Or SH.Name = SH7 Or SH.Name = SH8 Or SH.Name = SH9 Or SH.Name = SH10 Then
            SH.Delete
        End If
    Next SH
'    Application.DisplayAlerts = True
    fn_HapusSheet = "Sheets deletion completed"
    
'    cara pake: MASUKAN NAMA SHEET YG AKAN DIHAPUS
'    Dim Ke_Fn As String
'    Ke_Fn = fn_HapusSheet("TES1", "TES2", "DI_WOBuyer", "RPA2", "CC2", "", "", "", "", "")
End Sub

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


'Sub CLEAR_FIELD()
'    Dim WB As Workbook
'    Dim SH_X As Worksheet
'    Set WB = ThisWorkbook
'
''---CLEAR Sheet CEISA
'    Set SH_X = WB.Worksheets("CEISA")
'    SH_X.Range("A6:Z1048576").ClearContents
'
''---CLEAR Summary
'    Set SH_X = WB.Worksheets("Summary")
'    SH_X.AutoFilterMode = False
'    SH_X.Range("A6:Z1048576").ClearContents
'
'End Sub
'
'Sub CLEAR_PATH()
'    Dim WB As Workbook
'    Dim SH_X As Worksheet
'    Set WB = ThisWorkbook
''---------------------------------------------
''---CLEAR Set_Cach
'    Set SH_X = WB1.Worksheets("Set_Cach")
'    SH_X.AutoFilterMode = False
'    SH_X.Range("C6:C15").ClearContents
'End Sub
'
'Sub HIDE_Sheets()
'On Error Resume Next
'    ThisWorkbook.Sheets("Summary").Visible = False
'On Error GoTo 0
'End Sub
'
'Sub UNHIDE_Sheets()
'On Error Resume Next
'    ThisWorkbook.Sheets("Backup_Done").Visible = True
'    ThisWorkbook.Sheets("Set_Def").Visible = True
'On Error GoTo 0
'End Sub








'____________________________________________________________________________________________________
'      -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'----------------------------------------------------------------------------------------------------
'Sub RESET_Macro()
'    Dim ws As Worksheet
'    For Each ws In ThisWorkbook.Sheets
'        If ws.Name <> "(M)" Then
'            ws.Delete
'        End If
'    Next ws
'End Sub

'Private Sub Lbl_ResetMacro_Click()
'Dim Pesan As Integer
'Pesan = MsgBox("Aksi ini akan menghapus semua field yg telah diimport sebelumnya" & vbCrLf & "Lanjutkan???", vbYesNo + vbQuestion, "Reset Macro")
'    If Pesan = vbYes Then
'        Call DELETE_AllSheetOnlyAsset

'        Call CLEAR_Macro
'        Call HIDE_Sheets
'        WB.Worksheets("(M)").Range("H8:M11").Select
'
'        fm_MenuUTama.Btn_ViewLast.Enabled = False
'        MsgBox "Done", vbInformation, "Reset Macro"
'        End
'    Else:
'        MsgBox "Cancel", vbInformation, "Reset Macro"
'    End If
'End Sub

