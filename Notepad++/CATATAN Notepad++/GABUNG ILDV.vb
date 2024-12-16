Sub CollectData()
Application.DisplayAlerts = False
    
    Set TWB = ThisWorkbook
    Set HOME = TWB.Sheets("HOME")
    
    N_PROCESS = HOME.Range("AF3")
    FILE_DATA = HOME.Range("E9") & Application.PathSeparator & HOME.Range("D9") & HOME.Range("F9")
    PATH_DATA = HOME.Range("E8") & Application.PathSeparator & HOME.Range("D8")
    For i = 1 To N_PROCESS
        Set WB_X = Nothing
        PATH_TARIKAN = ""
        PATH_TARIKAN = PATH_DATA & " " & i & HOME.Range("F8")
        If i = 1 Then
            Set WB_COLLECT = Workbooks.Open(PATH_TARIKAN): WB_COLLECT.Activate
            Set SH = WB_COLLECT.Sheets(1)
            ROW_PASTE = SH.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row + 1
        Else
            ROW_PASTE = SH.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row + 1
            Set WB_X = Workbooks.Open(PATH_TARIKAN)
            WB_X.Activate: Sheets(1).Select
            LR = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
            LC = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
            Range(Cells(8, 1), Cells(LR, LC)).Copy
            Windows(WB_COLLECT.Name).Activate
            SH.Activate
            Range("A" & ROW_PASTE).PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
            WB_X.Close False
        End If
    Next i
    Windows(WB_COLLECT.Name).Activate
    WB_COLLECT.SaveAs FILE_DATA, xlOpenXMLWorkbook
    WB_COLLECT.Close False
    
    HOME.Activate: Cells(1, 1).Select
    
Application.DisplayAlerts = True
End Sub
