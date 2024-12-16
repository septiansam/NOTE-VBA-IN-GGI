Sub BUTTON_Proses2()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Call Validasi_Source
    Call DeleteSheetsExcept("Ref_Mail", "LA", "HOME", "RPA1")
    Call Validasi_Data
    Call Processing_Proses_2

    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Sub Validasi_Data()

    Application.ScreenUpdating = False
    
    '[*]... VALIDASI DATA TARIKAN GCC
    Set WB1 = ThisWorkbook
    Set SH1_Home = WB1.Worksheets("HOME")
    Set SH1_RPA1 = WB1.Worksheets("RPA1")
    Set SH1_RPA_CHECK = Sheets.Add(after:=Sheets(Sheets.Count)): SH1_RPA_CHECK.Name = "RPA_CHECK"
    SH1_RPA_CHECK.Activate
    SH1_RPA_CHECK.Range("A1").Value = "STATUS"
    
    SH1_Home.Activate
    PathSource = SH1_Home.Range("F11").Value
    
    Set WB2 = Workbooks.Open(PathSource)
    Windows(WB2.Name).Activate
    Set SH2 = WB2.Worksheets(1)
    SH2.Activate
    SH2.AutoFilterMode = False
    SH2.Cells.EntireColumn.Hidden = False
    
    '[*]... JIKA TARIKAN GCC TIDAK ADA DATANYA
    If WorksheetFunction.CountA(SH2.Cells) = 0 Then
        WB2.Close False
        Set WB2 = Nothing
        Set SH2 = Nothing
        Windows(WB1.Name).Activate
        
        SH1_RPA_CHECK.Activate
        SH1_RPA_CHECK.Range("A2").Value = 0

        Set SH1_RPA_WA = Sheets.Add(after:=Sheets(Sheets.Count)): SH1_RPA_WA.Name = "RPA_WA"
        SH1_RPA_WA.Activate
        
        SH1_RPA_WA.Range("A1").Value = "TUJUAN"
        SH1_RPA_WA.Range("A2").FormulaR1C1 = "=Ref_Mail!R2C7"
        
        SH1_RPA_WA.Range("B1").Value = "PESAN"
        SH1_RPA_WA.Range("B2").FormulaR1C1 = "=IF(RPA_CHECK!RC[-1]=0,""PROSES RPA DIHENTIKAN: Macro Confirm Shipment Sample tidak berbayar - Tarikan GCC Kosong"")"
        
        SH1_RPA_WA.Range("A2:B2").Copy
        SH1_RPA_WA.Range("A2:B2").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        SH1_RPA_WA.Cells.EntireColumn.AutoFit
        SH1_RPA_WA.Cells(1, 1).Select
        
        SH1_Home.Activate
        SH1_Home.Cells(1, 1).Select

        Application.ScreenUpdating = True
        End
    End If
    
    '[*]... JIKA TARIKAN GCC PICK NUMBERNYA TIDAK ADA YANG BLANK
    Windows(WB2.Name).Activate
    SH2.Activate
    SH2.AutoFilterMode = False
    SH2.Cells.AutoFilter 39, "="
    If WorksheetFunction.CountA(SH2.Range("A:A").SpecialCells(xlCellTypeVisible)) = 1 Then
        WB2.Close False
        Set WB2 = Nothing
        Set SH2 = Nothing
        Windows(WB1.Name).Activate
        
        SH1_RPA_CHECK.Activate
        SH1_RPA_CHECK.Range("A2").Value = 0

        Set SH1_RPA_WA = Sheets.Add(after:=Sheets(Sheets.Count)): SH1_RPA_WA.Name = "RPA_WA"
        SH1_RPA_WA.Activate
        
        SH1_RPA_WA.Range("A1").Value = "TUJUAN"
        SH1_RPA_WA.Range("A2").FormulaR1C1 = "=Ref_Mail!R2C7"
        
        SH1_RPA_WA.Range("B1").Value = "PESAN"
        SH1_RPA_WA.Range("B2").FormulaR1C1 = "=IF(RPA_CHECK!RC[-1]=0,""PROSES RPA DIHENTIKAN: Macro Confirm Shipment Sample tidak berbayar - Tidak Terdapat Unique Key ID yang Kosong"")"
        
        SH1_RPA_WA.Range("A2:B2").Copy
        SH1_RPA_WA.Range("A2:B2").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        SH1_RPA_WA.Cells.EntireColumn.AutoFit
        SH1_RPA_WA.Cells(1, 1).Select
        
        SH1_Home.Activate
        SH1_Home.Cells(1, 1).Select

        Application.ScreenUpdating = True
        End
    End If
    
    Windows(WB1.Name).Activate
    
    SH1_RPA_CHECK.Activate
    SH1_RPA_CHECK.Range("A2").Value = 1
    
    WB2.Close False
    Set WB2 = Nothing
    Set SH2 = Nothing
    
    Windows(WB1.Name).Activate
    SH1_Home.Activate
    SH1_Home.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
End Sub