Sub Processing_Proses_3()
    Set WB1 = ThisWorkbook
    Set SH1_RefMail = WB1.Worksheets("Ref_Mail")
    Set SH1_Home = WB1.Worksheets("HOME")
    Set SH1_RPA2 = WB1.Worksheets("RPA2")
    Set SH1_RPA3 = WB1.Sheets.Add(after:=WB1.Sheets(WB1.Sheets.Count)): SH1_RPA3.Name = "RPA3"
    Set SH1_RESUME = WB1.Sheets.Add(after:=WB1.Sheets(WB1.Sheets.Count)): SH1_RESUME.Name = "RESUME"
    Set SH1_RPA4_Email = WB1.Sheets.Add(after:=WB1.Sheets(WB1.Sheets.Count)): SH1_RPA4_Email.Name = "RPA4_Email"
    Set SH1_OLAH1 = WB1.Sheets.Add(after:=WB1.Sheets(WB1.Sheets.Count)): SH1_OLAH1.Name = "OLAH1"

    '[*].. SETTING HEADER RPA3
    SH1_RPA3.Activate
    SH1_RPA3.Range("A1").Value = "PICK NUMBER"
    SH1_RPA3.Range("B1").Value = "ACTUAL SHIP DATE"
    SH1_RPA3.Range("C1").Value = "WO/LOCATION"
    SH1_RPA3.Range("D1").Value = "QTY SHIPPED"
    SH1_RPA3.Range("A1").EntireRow.Font.Bold = True
    SH1_RPA3.Cells.EntireColumn.AutoFit
    SH1_RPA3.Range("A1").Select
    
    '[*].. SETTING TEMPLATE RESUME1
    SH1_RESUME.Activate
    SH1_RESUME.Range("A1").Value = "Data Dibawah adalah data yg berhasil diproses ConfirmShipment Oleh RPA"
    SH1_RESUME.Range("A2").Value = "NO URUT"
    SH1_RESUME.Range("B2").Value = "PICK NUMBER"
    SH1_RESUME.Range("C2").Value = "NO AJU"
    SH1_RESUME.Range("A1").Font.Italic = True
    With SH1_RESUME.Range("A2:C2")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 176, 240)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    SH1_RESUME.Range("B:C").EntireColumn.AutoFit
    SH1_RESUME.Range("A1").Select
    
    '[*].. KE SHEETS RPA2, UNTUK ISI DATA RPA3
    '''''' DAPATKAN DATANYA, HANYA PICK NUMBER, DARI SHEETS RPA2, YANG DI FILTER DI KOLOM B, DENGAN KRITERIA FILTER "1"
    SH1_RPA2.Activate
    SH1_RPA2.AutoFilterMode = False
    SH1_RPA2.UsedRange.AutoFilter 2, 1
    
    '''''' JIKA ADA DATANYA, DAN JIKA DATA NYA TIDAK ADA -> UNTUK RPA3 DAN RESUME1
    If WorksheetFunction.CountA(SH1_RPA2.Range("B:B").SpecialCells(xlCellTypeVisible)) <> 1 Then
        SH1_OLAH1.Cells.ClearContents
        SH1_RPA2.Cells.SpecialCells(xlCellTypeVisible).Copy
        SH1_OLAH1.Activate
        SH1_OLAH1.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        LR1 = SH1_OLAH1.Range("A" & Rows.Count).End(xlUp).Row
        SH1_OLAH1.Range("A2:A" & LR1).Copy
        SH1_RPA3.Activate
        SH1_RPA3.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        SH1_OLAH1.Activate
        SH1_OLAH1.Range("B:B").Delete
        SH1_OLAH1.Range("A:A").Insert
        SH1_OLAH1.Range("A1").Value = "NO URUT"
        SH1_OLAH1.Range("A2:A" & LR1).FormulaR1C1 = "=ROW()-1"
        SH1_OLAH1.Range("A2:A" & LR1).Copy
        SH1_OLAH1.Range("A2:A" & LR1).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        SH1_OLAH1.Range("A2:C" & LR1).Copy
        SH1_RESUME.Activate
        SH1_RESUME.Range("A3").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        SH1_RESUME.Range("B:C").EntireColumn.AutoFit
        SH1_RESUME.Range("A1").Select
    Else
        SH1_RPA3.Activate
        SH1_RPA3.Range("A2").Value = "No Data Value"
        SH1_RPA3.Range("A1").Select
        SH1_RESUME.Activate
        SH1_RESUME.Range("A3").Value = "No Data Value"
        SH1_RESUME.Range("A1").Select
    End If
    
    '[*].. SETTING TEMPLATE RESUME2
    SH1_RESUME.Activate
    RowResume2 = SH1_RESUME.Range("A" & Rows.Count).End(xlUp).Row + 5
    SH1_RESUME.Range("A" & RowResume2).Value = "Data Dibawah ini tidak bisa diproses RPA, karena lebih dari 1 Row"
    SH1_RESUME.Range("A" & RowResume2 + 1).Value = "Kami infokan, kembalikan lagi ke user agar bisa diinput ulang"
    SH1_RESUME.Range("A" & RowResume2 & ":A" & RowResume2 + 1).Font.Italic = True
    
    SH1_RESUME.Range("A" & RowResume2 + 2).Value = "NO URUT"
    SH1_RESUME.Range("B" & RowResume2 + 2).Value = "PICK NUMBER"
    SH1_RESUME.Range("C" & RowResume2 + 2).Value = "NO AJU"
    With SH1_RESUME.Range("A" & RowResume2 + 2 & ":C" & RowResume2 + 2)
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 176, 240)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    SH1_RESUME.Range("B:C").EntireColumn.AutoFit
    SH1_RESUME.Range("A1").Select
    RowPaste = SH1_RESUME.Range("A" & Rows.Count).End(xlUp).Row + 1
    
    '[*].. KE SHEETS RPA2, UNTUK ISI DATA RESUME2
    '''''' DAPATKAN DATANYA, YANG DI FILTER DI KOLOM B, DENGAN KRITERIA FILTER "<>1"
    SH1_OLAH1.Cells.ClearContents
    SH1_RPA2.Activate
    SH1_RPA2.AutoFilterMode = False
    SH1_RPA2.UsedRange.AutoFilter 2, "<>1"
    
    If WorksheetFunction.CountA(SH1_RPA2.Range("A:A").SpecialCells(xlCellTypeVisible)) <> 1 Then
        SH1_RPA2.Cells.SpecialCells(xlCellTypeVisible).Copy
        SH1_OLAH1.Activate
        SH1_OLAH1.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        LR1 = SH1_OLAH1.Range("A" & Rows.Count).End(xlUp).Row
        SH1_OLAH1.Range("B:B").Delete
        SH1_OLAH1.Range("A:A").Insert
        SH1_OLAH1.Range("A1").Value = "NO URUT"
        SH1_OLAH1.Range("A2:A" & LR1).FormulaR1C1 = "=ROW()-1"
        SH1_OLAH1.Range("A2:A" & LR1).Copy
        SH1_OLAH1.Range("A2:A" & LR1).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        SH1_OLAH1.Range("A2:C" & LR1).Copy
        SH1_RESUME.Activate
        SH1_RESUME.Range("A" & RowPaste).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        SH1_RESUME.Range("B:C").EntireColumn.AutoFit
        SH1_RESUME.Range("A1").Select
    Else
        SH1_RESUME.Activate
        SH1_RESUME.Range("A" & RowPaste).Value = "No Data Value"
        SH1_RESUME.Range("B:C").EntireColumn.AutoFit
        SH1_RESUME.Range("A1").Select
    End If
    SH1_RPA2.Activate
    SH1_RPA2.AutoFilterMode = False
    SH1_RPA2.Range("A1").Select
    
    '[*].. EXPORT RESUME
    SH1_RESUME.Activate
    PathResume = SH1_Home.Range("E20").Value
    SH1_RESUME.Copy
    Set WB2_RESUME = ActiveWorkbook
    WB2_RESUME.SaveAs PathResume, xlOpenXMLWorkbook
    WB2_RESUME.Close False
    
    '[*].. SIAPKAN EMAIL
    SH1_RefMail.Activate
    SH1_RefMail.AutoFilterMode = False
    
    '''[##].. GET TO EMAIL
    SH1_OLAH1.Cells.ClearContents
    SH1_RefMail.Activate
    LR1 = SH1_RefMail.Range("E" & Rows.Count).End(xlUp).Row
    SH1_RefMail.Range("E2:E" & LR1).Copy
    SH1_OLAH1.Activate
    SH1_OLAH1.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SH1_OLAH1.Range("A:A").RemoveDuplicates 1, xlNo
    LR1 = SH1_OLAH1.Range("A" & Rows.Count).End(xlUp).Row
    For i = 1 To LR1
        Str_ToEmail = Str_ToEmail & SH1_OLAH1.Range("A" & i).Value & ","
    Next i
    Str_ToEmail = Left(Str_ToEmail, Len(Str_ToEmail) - 1)
    
    '''[##].. GET CC EMAIL
    SH1_OLAH1.Cells.ClearContents
    SH1_RefMail.Activate
    LR1 = SH1_RefMail.Range("H" & Rows.Count).End(xlUp).Row
    SH1_RefMail.Range("H2:H" & LR1).Copy
    SH1_OLAH1.Activate
    SH1_OLAH1.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SH1_OLAH1.Range("A:A").RemoveDuplicates 1, xlNo
    LR1 = SH1_OLAH1.Range("A" & Rows.Count).End(xlUp).Row
    For i = 1 To LR1
        Str_CCEmail = Str_CCEmail & SH1_OLAH1.Range("A" & i).Value & ","
    Next i
    Str_CCEmail = Left(Str_CCEmail, Len(Str_CCEmail) - 1)
    
    '[*].. SIMPAN DI SHEET RPA4_Email
    SH1_RPA4_Email.Activate
    SH1_RPA4_Email.Range("A1").Value = "TO"
    SH1_RPA4_Email.Range("A2").Value = Str_ToEmail
    
    SH1_RPA4_Email.Range("B1").Value = "CC"
    SH1_RPA4_Email.Range("B2").Value = Str_CCEmail
    
    SH1_RPA4_Email.Range("C1").Value = "SUBJECT"
    SH1_RPA4_Email.Range("C2").FormulaR1C1 = "=Ref_Mail!RC[-2]"
    SH1_RPA4_Email.Range("C2").Copy
    SH1_RPA4_Email.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    
    SH1_RPA4_Email.Range("D1").Value = "ATTACHMENT"
    SH1_RPA4_Email.Range("D2").FormulaR1C1 = "=HOME!R[18]C[1]"
    SH1_RPA4_Email.Range("D2").Copy
    SH1_RPA4_Email.Range("D2").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    
    SH1_RPA4_Email.Range("A1").EntireRow.Font.Bold = True
    SH1_RPA4_Email.Cells.EntireColumn.AutoFit
    SH1_RPA4_Email.Cells(1, 1).Select
    
    SH1_Home.Activate
    SH1_Home.Cells(1, 1).Select
    
    SH1_OLAH1.Delete
End Sub