'_____________________________________________________________________________________________________
'## USER - MACRO_PROCESS Sales Update Sample di JDE   -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'## DEVELOPER := Septian Arif Maulana -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------

'=====================================================
'=> ------------VARIABLE INITIALIZATION------------ <=
'=====================================================
    Public WB1 As Workbook
    Public SH1_RefNoHP As Worksheet, SH1_Administrator As Worksheet
    Public SH1_Home As Worksheet
    Public SH1_INPUT_EMAIL As Worksheet
    Public CountFalse As Long
    Public PathMacroRpa As String
    Public Rng As Range, Cell As Range
    Public LR1 As Long, LC1 As Long
    Public LR1_InputEmail As Long
    Public i As Long, j As Long
    '-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
    Public WB2_RPA As Workbook
    Public SH2_InputanUser As Worksheet
    Public SH2_RefEmail As Worksheet
'+---------------------------------------------------+

Sub BTN_SEND_DATA()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Call HideSheets("ADMINISTRATOR")
    Call Validasi_MacroUser_Input_Email
    Call Send_Email_To_MacroRPA
    
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Sub Send_Email_To_MacroRPA()
    Set WB1 = ThisWorkbook
    Set SH1_Home = WB1.Worksheets("HOME")
    Set SH1_INPUT_EMAIL = WB1.Worksheets("INPUT EMAIL")
    
    PathMacroRpa = SH1_Administrator.Range("B2").value
    Set WB2_RPA = Workbooks.Open(PathMacroRpa)
    Windows(WB2_RPA.Name).Activate
    Set SH2_RefEmail = WB2_RPA.Worksheets("Ref_Mail")
    SH2_RefEmail.Activate
    
    ''' HAPUS EMAIL ADDRESS TO
    SH2_RefEmail.Range("A2:A" & Rows.Count).ClearContents
    
    ''' HAPUS EMAIL ADDRESS CC
    SH2_RefEmail.Range("D2:D" & Rows.Count).ClearContents
    
    ''' COPY TO EMAIL
    Windows(WB1.Name).Activate
    SH1_INPUT_EMAIL.Activate
    LR1 = SH1_INPUT_EMAIL.Range("A" & Rows.Count).End(xlUp).Row
    SH1_INPUT_EMAIL.Range("A6:A" & LR1).Copy
    
    Windows(WB2_RPA.Name).Activate
    SH2_RefEmail.Activate
    SH2_RefEmail.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    
    ''' COPY CC EMAIL -> CEK APAKAH CC NYA ADA
    Windows(WB1.Name).Activate
    SH1_INPUT_EMAIL.Activate
    LR1 = SH1_INPUT_EMAIL.Range("B" & Rows.Count).End(xlUp).Row
    If LR1 > 5 Then
        SH1_INPUT_EMAIL.Range("B6:B" & LR1).Copy
        Windows(WB2_RPA.Name).Activate
        SH2_RefEmail.Activate
        SH2_RefEmail.Range("D2").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End If
    
    Windows(WB2_RPA.Name).Activate
    SH2_RefEmail.Activate
    SH2_RefEmail.Cells.EntireColumn.AutoFit
    SH2_RefEmail.Cells(1, 1).Select
    
    Windows(WB1.Name).Activate
    SH1_Home.Activate
    SH1_Home.Cells(1, 1).Select
    
    WB2_RPA.Close True
    Set WB2_RPA = Nothing
    Set SH2_RefEmail = Nothing
    
End Sub

Sub Validasi_MacroUser_Input_Email()
    Set WB1 = ThisWorkbook
    Set SH1_Home = WB1.Worksheets("HOME")
    Set SH1_Administrator = WB1.Worksheets("ADMINISTRATOR")
    Set SH1_INPUT_EMAIL = WB1.Worksheets("INPUT EMAIL")
    
    PathMacroRpa = SH1_Administrator.Range("B2").value
    If Dir(PathMacroRpa) = "" Then
        SH1_Home.Activate
        SH1_Home.Range("A1").Select
        MsgBox "File Macro RPA tidak ditemukan", vbCritical, "NOT FOUND"
        End
    End If
    
    SH1_INPUT_EMAIL.Activate
    ''' VALIDASI TO EMAIL
    If SH1_INPUT_EMAIL.Range("A" & Rows.Count).End(xlUp).Row = 5 Then
        SH1_INPUT_EMAIL.Activate
        SH1_INPUT_EMAIL.Range("A6").Select
        MsgBox "TIDAK TERDAPAT INPUTAN USER ''TO EMAIL ADDRESS''" & vbNewLine & _
               "MOHON INPUT TERLEBIH DAHULU DATA ''TO EMAIL ADDRESS''", vbCritical, "DATA IS EMPTY"
        End
    End If
    
    ''' FORMULASI
    LR1_InputEmail = SH1_INPUT_EMAIL.Range("A:B").Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row

    '...VALIDASI TO EMAIL ADDRESS BY MACRO
    LR1 = SH1_INPUT_EMAIL.Range("A" & Rows.Count).End(xlUp).Row
    If LR1 > 5 Then
        SH1_INPUT_EMAIL.Range("F6:F" & LR1).FormulaR1C1 = _
            "=IF(AND(ISNUMBER(FIND(""."",RC[-5])),ISNUMBER(FIND(""@"",RC[-5]))),TRUE,FALSE)"
    End If
    
    '...VALIDASI CC EMAIL ADDRESS BY MACRO
    LR1 = SH1_INPUT_EMAIL.Range("B" & Rows.Count).End(xlUp).Row
    If LR1 > 5 Then
        SH1_INPUT_EMAIL.Range("G6:G" & LR1).FormulaR1C1 = _
            "=IF(RC[-5]="""",TRUE,IF(AND(ISNUMBER(FIND(""."",RC[-5])),ISNUMBER(FIND(""@"",RC[-5]))),TRUE,FALSE))"
    End If
    
    '...NOTE
    SH1_INPUT_EMAIL.Range("H6:H" & LR1_InputEmail).FormulaR1C1 = _
        "=IF(COUNTIF(R6C6:R1048576C7,FALSE)>0,""Silahkan check kembali Email address TO dan CC, pastikan alamat emial valid"","""")"
    
    '[*].. CEK APAKAH MASIH ADA YANG FALSE ATAU TIDAK
    CountFalse = WorksheetFunction.CountIf(SH1_INPUT_EMAIL.Range("F6:G1048576"), False)
    
    ''' JIKA MASIH ADA YANG FALSE
    If CountFalse > 0 Then
        SH1_INPUT_EMAIL.Range("A6").Select
        MsgBox "Hasil Validasi menyatakan" & vbNewLine & _
               "Ada nilai yang tidak sesuai, silahkan check inputan kembali", vbCritical, "VALIDASI INPUTAN USER"
        End
    End If
    
End Sub

Sub BTN_START_WIZARD()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Call Get_PathThisMacro
    Call Clear_StartWizard
    Call HideSheets("Ref_NoHP", "ADMINISTRATOR", "2. INPUT - WATo", "Results Validation")
    Call SelectedSheet_InputNoAju

    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Sub Validasi_InputanUser()
    Set WB1 = ThisWorkbook
    Set SH1_RefNoHP = WB1.Worksheets("Ref_NoHP")
    Set SH1_Home = WB1.Worksheets("HOME")
    Set SH1_InputNoAju = WB1.Worksheets("1. INPUT - NoAju")
    Set SH1_InputWATo = WB1.Worksheets("2. INPUT - WATo")
    Set SH1_ResultValidation = WB1.Worksheets("Results Validation")
    
    '[*].. UNHIDE SHEETS RESULTS VALIDATION
    SH1_ResultValidation.Visible = xlSheetVisible
    
    '[*].. VALIDASI DI INPUT NO AJU
    SH1_InputNoAju.Activate
    SH1_InputNoAju.Columns("E:G").Hidden = False
    LR1_NoAju = SH1_InputNoAju.Range("A" & Rows.Count).End(xlUp).Row
    For i = 6 To LR1_NoAju
        If IsNumeric(SH1_InputNoAju.Range("A" & i).value) And SH1_InputNoAju.Range("A" & i).value <> "" Then
            SH1_InputNoAju.Range("F" & i).value = True
        ElseIf SH1_InputNoAju.Range("A" & i).value = "" Then
            SH1_InputNoAju.Range("F" & i).value = False
        Else
            SH1_InputNoAju.Range("F" & i).value = False
        End If
    Next i
    SH1_InputNoAju.Range("G6:G" & LR1_NoAju).FormulaR1C1 = "=IF(RC[-1]<>"""",IF(RC[-1]=FALSE, ""Masukan No aju yg komplit 26 angka, Jika masih FALSE cobalah untuk Copas dari GCC dan pastikan format cell adalah 'TEXT' atau CEISA atau klik ulang tombol 'START WIZARD'"",""""),"""")"
    SH1_InputNoAju.Range("G6:G" & LR1_NoAju).Copy
    SH1_InputNoAju.Range("G6:G" & LR1_NoAju).PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    SH1_InputNoAju.Range("A1").Select
    
    '[*].. VALIDASI DI INPUT WA TO
    SH1_InputWATo.Activate
    SH1_InputWATo.Columns("E:G").Hidden = False
    LR1_WATo = SH1_InputWATo.Range("A" & Rows.Count).End(xlUp).Row
    For i = 6 To LR1_WATo
        If Left(SH1_InputWATo.Range("A" & i).value, 1) = 0 Then
            SH1_InputWATo.Range("F" & i).value = False
        ElseIf IsNumeric(SH1_InputWATo.Range("A" & i).value) And SH1_InputWATo.Range("A" & i).value <> "" Then
            SH1_InputWATo.Range("F" & i).value = True
        ElseIf SH1_InputWATo.Range("A" & i).value = "" Then
            SH1_InputWATo.Range("F" & i).value = False
        Else
            SH1_InputWATo.Range("F" & i).value = False
        End If
    Next i
    SH1_InputWATo.Range("G6:G" & LR1_WATo).FormulaR1C1 = "=IF(RC[-1]<>"""",IF(RC[-1]=FALSE, ""Pastikan angka awal adalah '62' sebagai pengganti angka '0', jika masih FALSE, cobalah klik ulang tombol 'START WIZARD'"",""""),"""")"
    SH1_InputWATo.Range("G6:G" & LR1_WATo).Copy
    SH1_InputWATo.Range("G6:G" & LR1_WATo).PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    SH1_InputWATo.Range("A1").Select
    
    '[*].. VALIDASI DI RESULTS VALIDATION
    SH1_ResultValidation.Activate
    SH1_ResultValidation.Columns("E:G").Hidden = False
    SH1_ResultValidation.Range("A6").value = "Checking Existing File Macro RPA"

    PathMacroRpa = fm_Admin.txt_LokasiMacroRPA.value

    If Dir(PathMacroRpa) <> "" Then
        SH1_ResultValidation.Range("F6").value = True
    Else
        SH1_ResultValidation.Range("F6").value = False
    End If
    SH1_ResultValidation.Range("G6").FormulaR1C1 = "=IF(RC[-1]=FALSE, ""SILAKAN HUBUNGI TEAM IT"","""")"
    
    SH1_ResultValidation.Range("A7").value = "Sheet ""1. INPUT- NoAju"""
    SH1_ResultValidation.Range("F7").FormulaR1C1 = "=IF(COUNTIF('1. INPUT - NoAju'!R6C6:R1048576C6,FALSE)>0,FALSE,TRUE)"
    SH1_ResultValidation.Range("G7").FormulaR1C1 = "=IF(RC[-1]=FALSE, ""Silakan check Sheet '1. INPUT- NoAju''"","""")"
    
    SH1_ResultValidation.Range("A8").value = "Sheet ""2. INPUT- WATo"""
    SH1_ResultValidation.Range("F8").FormulaR1C1 = "=IF(COUNTIF('2. INPUT - WATo'!R6C6:R1048576C6,FALSE)>0,FALSE,TRUE)"
    SH1_ResultValidation.Range("G8").FormulaR1C1 = "=IF(RC[-1]=FALSE, ""Silakan check Sheet '2. INPUT- WATo''"","""")"
    
    '[*].. PASTE VALUES
    SH1_ResultValidation.Range("F6:G8").Copy
    SH1_ResultValidation.Range("F6:G8").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    
    SH1_ResultValidation.Range("A1").Select
    
    '[*].. CEK APAKAH MASIH ADA YANG FALSE ATAU TIDAK
    CountFalse = WorksheetFunction.CountIf(SH1_ResultValidation.Range("F6:F1048576"), False)
    
    ''' JIKA MASIH ADA YANG FALSE
    If CountFalse > 0 Then
        SH1_ResultValidation.Range("A1").Select
        MsgBox "Hasil Validasi menyatakan" & vbNewLine & _
               "Ada nilai yang tidak sesuai, silahkan check inputan kembali", vbCritical, "VALIDASI INPUTAN USER"
        End
    End If
    
    ''' JIKA SUDAH TIDAK ADA YANG FALSE
'    Stop
'    fm_Antrian_Tambah.Show
'    fm_Antrian_Tambah.lbl_HasulValidasi.Caption = "TRUE"
'    fm_Antrian_Tambah.lbl_HasulValidasi.ForeColor = &HC000&
'    fm_Antrian_Tambah.fm_InputForm.Enabled = True
    
    ''' KIRIM DATA KE MACRO RPA
    Set WB2_RPA = Workbooks.Open(PathMacroRpa)
    Windows(WB2_RPA.Name).Activate
    Set SH2_InputanUser = WB2_RPA.Worksheets("INPUTAN_USER")
    SH2_InputanUser.Activate
    SH2_InputanUser.AutoFilterMode = False
    SH2_InputanUser.Cells.EntireColumn.Hidden = False
    
    SH2_InputanUser.Range("A2:XFD1048576").ClearContents
    
    ''' COPY NO AJU
    Windows(WB1.Name).Activate
    SH1_InputNoAju.Activate
    LR1 = SH1_InputNoAju.Range("A" & Rows.Count).End(xlUp).Row
    SH1_InputNoAju.Range("A6:A" & LR1).Copy
    Windows(WB2_RPA.Name).Activate
    SH2_InputanUser.Activate
    SH2_InputanUser.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    
    ''' COPY NO HP
    Windows(WB1.Name).Activate
    SH1_InputWATo.Activate
    LR1 = SH1_InputWATo.Range("A" & Rows.Count).End(xlUp).Row
    SH1_InputWATo.Range("A6:A" & LR1).Copy
    Windows(WB2_RPA.Name).Activate
    SH2_InputanUser.Activate
    SH2_InputanUser.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    
    SH2_InputanUser.Cells.EntireColumn.AutoFit
    SH2_InputanUser.Cells(1, 1).Select
    
    ''' COPY BU
    Windows(WB1.Name).Activate
    SH1_InputWATo.Activate
    SH2_InputanUser.Range("E2").value = SH1_Home.Range("E12").value
    
    WB2_RPA.Close True
    
    Windows(WB1.Name).Activate
    SH1_ResultValidation.Activate
    SH1_ResultValidation.Range("A1").Select
    
    MsgBox "Selamat, Hasil Validasi Dinyatakan Success" & vbNewLine & _
           "Data Telah Berhasil Terkirim" & vbNewLine & _
           "Silahkan Ajukan Tiket di GCC", vbInformation, "VALIDATION SUCCESS..."
    
End Sub

Sub Get_PathThisMacro()
    Set WB1 = ThisWorkbook
    Set SH1_RefNoHP = WB1.Worksheets("Ref_NoHP")
    Set SH1_Administrator = WB1.Worksheets("ADMINISTRATOR")
    
    SH1_Administrator.Range("B8").value = WB1.FullName
    SH1_Administrator.Range("B8").Hyperlinks.Delete
End Sub

Sub Clear_StartWizard()
    Set WB1 = ThisWorkbook
    Set SH1_Home = WB1.Worksheets("HOME")
    Set SH1_InputNoAju = WB1.Worksheets("1. INPUT - NoAju")
    Set SH1_InputWATo = WB1.Worksheets("2. INPUT - WATo")
    Set SH1_ResultValidation = WB1.Worksheets("Results Validation")
    
    '[*].. INPUT NO AJU - Clear Data dari kolom A6:XFD1048576, hide kolom E:G
    SH1_InputNoAju.Activate
    SH1_InputNoAju.AutoFilterMode = False
    SH1_InputNoAju.Cells.EntireColumn.Hidden = False
    SH1_InputNoAju.Range("A6:XFD1048576").ClearContents
    
    SH1_InputNoAju.Columns("B:D").Hidden = True
    SH1_InputNoAju.Columns("E:G").Hidden = True
    SH1_InputNoAju.Cells(1, 1).Select
    
    '[*].. INPUT WA TO - Clear Data dari kolom A6:XFD1048576, hide kolom E:G
    SH1_InputWATo.Activate
    SH1_InputWATo.AutoFilterMode = False
    SH1_InputWATo.Cells.EntireColumn.Hidden = False
    SH1_InputWATo.Range("A6:XFD1048576").ClearContents
    
    SH1_InputWATo.Columns("B:D").Hidden = True
    SH1_InputWATo.Columns("E:G").Hidden = True
    SH1_InputWATo.Cells(1, 1).Select
    
    '[*].. RESULT VALIDATION - Clear Data dari kolom A6:XFD1048576, hide kolom E:G
    SH1_ResultValidation.Activate
    SH1_ResultValidation.AutoFilterMode = False
    SH1_ResultValidation.Cells.EntireColumn.Hidden = False
    SH1_ResultValidation.Range("A6:XFD1048576").ClearContents
    
    SH1_ResultValidation.Columns("B:D").Hidden = True
    SH1_ResultValidation.Columns("E:G").Hidden = True
    SH1_ResultValidation.Cells(1, 1).Select
    
    SH1_Home.Activate
    SH1_Home.Cells(1, 1).Select
    
End Sub

Sub ClearData_InSheet(ByRef Sheet_Tujuan As Worksheet)
    Sheet_Tujuan.Activate
    Sheet_Tujuan.Range("A6:XFD1048576").ClearContents
    Sheet_Tujuan.Range("A6:A1048576").NumberFormat = "@"
    Sheet_Tujuan.Range("A6").Select
End Sub

Sub SelectedSheet_Home()
    Set WB1 = ThisWorkbook
    Set SH1_Home = WB1.Worksheets("HOME")
    SH1_Home.Activate: SH1_Home.Cells(1, 1).Select
End Sub

Sub SelectedSheet_InputNoAju()
    Set WB1 = ThisWorkbook
    Set SH1_InputNoAju = WB1.Worksheets("1. INPUT - NoAju")
    SH1_InputNoAju.Activate: SH1_InputNoAju.Cells(1, 1).Select
End Sub

Sub SelectedSheet_InputWATo()
    Set WB1 = ThisWorkbook
    Set SH1_InputWATo = WB1.Worksheets("2. INPUT - WATo")
    SH1_InputWATo.Activate: SH1_InputWATo.Cells(1, 1).Select
End Sub




















































































