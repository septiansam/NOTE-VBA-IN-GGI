Attribute VB_Name = "mdl_Validation"
'---------------------------------------------------------------------------------------*
'PUBLIC VARIABLE ------------------------------------------------------------------------
    Dim WB1 As Workbook
    Dim SH1_HOME As Worksheet, SH1_DM1 As Worksheet
    Dim LR1_HOME As Long
    '`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.'
    Dim FilePath As String, FolderPath As String, LeftEntitas As String, FullPath As String
    Dim i As Long
    '`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.'
    
    
Function FN_Check_ExistFile(FilePath As String) As Boolean
'---------------------------------------------------------------------------------------*
    ' Periksa apakah file ada
    If Dir(FilePath) <> "" Then
        FN_Check_ExistFile = True
    Else
        FN_Check_ExistFile = False
    End If
End Function

Function FN_Check_ExistFolder(FolderPath As String) As Boolean
    ' Periksa apakah folder ada
    If Dir(FolderPath, vbDirectory) <> "" Then
        FN_Check_ExistFolder = True
    Else
        FN_Check_ExistFolder = False
    End If
End Function


'Sub Check_ExistingFileandFolder()
''---------------------------------------------------------------------------------------*
''[INISIALISASI]
'''''''''''''''''''''''''''''''''''
'    Set WB1 = ThisWorkbook
'    Set SH1_HOME = WB1.Worksheets("HOME")
'    LR1_HOME = SH1_HOME.Range("F" & Rows.Count).End(xlUp).row
'
'
''[PROSES]
'''''''''''''''''''''''''''''''''''
'    Dim RangePengecekan As Range
'    Set RangePengecekan = SH1_HOME.Range("F4:F" & LR1_HOME)
'
'    For Each cell In RangePengecekan
'        i = cell.row
'        If SH1_HOME.Range("F" & i).value <> "" Then
'            LeftEntitas = Left(SH1_HOME.Range("D" & i).value, 11)
'            If LeftEntitas = "LOKASI FILE" Then
'                SH1_HOME.Range("F" & i).Hyperlinks.Delete
'                FilePath = SH1_HOME.Range("F" & i).value
'                If FN_CheckExistensiFile(FilePath) Then
'                    SH1_HOME.Range("F" & i).Interior.Pattern = xlNone
'                Else
'                    SH1_HOME.Range("F" & i).Select
'                    SH1_HOME.Range("F" & i).Interior.Color = 255
'                    MsgBox "PROSES DIHENTIKAN!" & vbCrLf & "Ada entitas yang tidak ditemukan", vbCritical, "VALIDASI | Check_ExistingFileandFolder"
'                    Exit Sub
'                End If
'
'            ElseIf LeftEntitas = "LOKASI FOLD" Then
'                SH1_HOME.Range("F" & i).Hyperlinks.Delete
'                FolderPath = SH1_HOME.Range("F" & i).value
'                If FN_FolderExists(FolderPath) Then
'                    SH1_HOME.Range("F" & i).Interior.Pattern = xlNone
'                Else
'                    SH1_HOME.Range("F" & i).Select
'                    SH1_HOME.Range("F" & i).Interior.Color = 255
'                    MsgBox "PROSES DIHENTIKAN!" & vbCrLf & "Ada entitas yang tidak ditemukan", vbCritical, "VALIDASI | Check_ExistingFileandFolder"
'                    Exit Sub
'                End If
'            End If
'        End If
'    Next cell
'End Sub


Sub Check_ExistingFileandFolder2()
'---------------------------------------------------------------------------------------*
'[INISIALISASI]
''''''''''''''''''''''''''''''''''
    Set WB1 = ThisWorkbook
    Set SH1_HOME = WB1.Worksheets("HOME")
    LR1_HOME = SH1_HOME.Range("E" & Rows.Count).End(xlUp).Row

'[PROSES]
''''''''''''''''''''''''''''''''''
    Dim RangePengecekan As Range
    Dim RangeFind As Range
    
    Set RangePengecekan = SH1_HOME.Range("E12:E" & LR1_HOME)
    RangePengecekan.Hyperlinks.Delete
                
    For Each cell In RangePengecekan
        i = cell.Row
            LeftEntitas = Left(SH1_HOME.Range("C" & i).Value, 11)
            FullPath = SH1_HOME.Range("E" & i).Value
            
            If LeftEntitas = "LOKASI FILE" Then
                If FN_Check_ExistFile(FullPath) Then
                    SH1_HOME.Range("G" & i).ClearContents
                Else
                    SH1_HOME.Range("G" & i).Value = "NOT FOUND"
                End If
            ElseIf LeftEntitas = "LOKASI FOLD" Then
                If FN_Check_ExistFolder(FullPath) Then
                    SH1_HOME.Range("G" & i).ClearContents
                Else
                    SH1_HOME.Range("G" & i).Value = "NOT FOUND"
                End If
            End If
    Next cell
    
    Set RangeFind = SH1_HOME.Range("G12:G" & LR1_HOME)
    If Application.WorksheetFunction.CountA(RangeFind) > 0 Then
        MsgBox "Terdapat File Yang Tidak Ditemukan", vbExclamation, "FILE NOT FOUND"
        End
    End If
    
End Sub


'Sub CONTOH_Check_Existing_File()
'    Set WB1 = ThisWorkbook
'    Set SH1_HOME = WB1.Worksheets("HOME")
'    FilePath = SH1_HOME.Range("B10").value
'    If Not FN_CheckExistensiFile(FilePath) Then
'        MsgBox "PROSES DIHENTIKAN!" & vbNewLine & "File tidak ditemukan" & _
'        vbCrLf & FilePath, vbCritical, "Checking Existing File"
'        Exit Sub
'    End If
'End Sub
'
'
'Sub CONTOH_Check_Existing_Folder()
'    Set WB1 = ThisWorkbook
'    Set SH1_HOME = WB1.Worksheets("HOME")
'    FilePath = SH1_HOME.Range("B10").value
'    If Not FN_Check_ExistFile(FilePath) Then
'        MsgBox "PROSES DIHENTIKAN!" & vbNewLine & "File tidak ditemukan" & _
'        vbCrLf & FilePath, vbCritical, "Checking Existing File"
'        Exit Sub
'    End If
'End Sub

Sub Validate_Proses()
    Set WB1 = ThisWorkbook
    Set SH1_HOME = WB1.Sheets("HOME")
    If Not wsx("RPA1") Then
        SH1_HOME.Activate
        Cells(1, 1).Select
        MsgBox "Untuk Menjalankan Tombol PROSES" & vbCrLf & vbCrLf & _
                "Silahkan Klik Tombol RPA_1 Terlebih Dahulu...", _
                vbInformation, "INFORMATION"
        End
    End If
End Sub

