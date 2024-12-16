Sub Main_Backup()

    '[*].. Validasi
    If Not wsx("RPA2") Then
        End
    End If
    
    '[*].. Pengolahan
    Set WB1 = ThisWorkbook
    Set SH1_History = WB1.Worksheets("History_WA")
    Set SH1_HOME = WB1.Worksheets("HOME")
    Set SH1_RPA2 = WB1.Worksheets("RPA2")
    
    SH1_RPA2.Activate
    For i = 2 To SH1_RPA2.Range("A" & Rows.Count).End(xlUp).Row
        SH1_RPA2.Range("D" & i).Value = Now()
    Next
    
    If Day(Date) <> 10 Then
        SH1_RPA2.Activate
        LR1 = SH1_RPA2.Range("A10000").End(xlUp).Row
        SH1_RPA2.Range("A2:D" & LR1).Copy
        
        SH1_History.Activate
        LR1_History = SH1_History.Range("A10000").End(xlUp).Offset(1).Row
        SH1_History.Range("A" & LR1_History).PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    Else
        SH1_History.Activate
        SH1_History.AutoFilterMode = False
        LR1 = SH1_History.Range("A" & Rows.Count).End(xlUp).Row
        
        LR1_History = SH1_History.Range("A10000").End(xlUp).Offset(1).Row
        If LR1 > 2 Then
            SH1_History.Range("E1") = "FILTER"
            SH1_History.Range("E2:E" & LR1).FormulaR1C1 = "=MONTH(RC[-1])"
            SH1_History.Range("A1:E" & LR1).AutoFilter 5, "<>" & Month(Date)
            SH1_History.Range("A1:E" & LR1).Offset(1).Delete
            SH1_History.AutoFilterMode = False
            SH1_History.Range("E:E").Delete
            SH1_History.Range("A1").Select
        End If
        SH1_RPA2.Activate
        LR1 = SH1_RPA2.Range("A" & Rows.Count).End(xlUp).Row
        
        SH1_RPA2.Range("A2:D" & LR1).Copy
        
        SH1_History.Activate
        LR1_History = SH1_History.Range("A" & Rows.Count).End(xlUp).Row + 1
        SH1_History.Range("A" & LR1_History).PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End If
    SH1_HOME.Activate
    SH1_HOME.Cells(1, 1).Select
End Sub