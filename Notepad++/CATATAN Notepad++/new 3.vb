

'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] Olahan Untuk Performance WO Purchasing
'...................................................................................................................................

Sub WO_Purchasing()

    Path_File = HOME.Range("D" & 6) & Application.PathSeparator & HOME.Range("E" & 6) & HOME.Range("F" & 6)
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    '><><> ............................... <><><
    '      Import And Collect Data Monthly
    '><><> ``````````````````````````````` <><><
    Set WB_FILE = Workbooks.Open(Path_File): WB_FILE.Activate
    rPaste = 1
        WB_FILE.Activate
        MonthNames = Array( _
        "Januari", "January", "Februari", "February", "Maret", "March", "April", "Mei", "May", "Juni", "June", "Juli", "July", _
        "Agustus", "August", "September", "Oktober", "October", "November", "Desember", "December" _
        )
        MonthNumbers = Array(1, 1, 2, 2, 3, 3, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 10, 10, 11, 12, 12)
        
        For i = LBound(MonthNames) To UBound(MonthNames)
            WB_FILE.Activate
            If wsx(MonthNames(i)) Then
                Set SH = WB_FILE.Sheets(MonthNames(i))
                If SH.Visible = xlSheetVisible Then
                    SH.Activate
                    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
                    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
                    Range(Cells(1, 2), Cells(lr, lc)).Copy
                    Windows(TWB.Name).Activate: TMP1.Activate
                    Range("C" & rPaste).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
                    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
                    Range("A" & rPaste, "A" & lr) = MonthNumbers(i)
                    Range("B" & rPaste, "B" & lr) = MonthNames(i)
                    rPaste = Range("C" & Rows.Count).End(xlUp).Row + 1
                End If
            End If
            Set SH = Nothing
        Next i
    WB_FILE.Close False
    Set WB_FILE = Nothing
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    
    TMP1.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    
    Range("I1").FormulaR1C1 = "MANAGER"
    Range("I2:I" & lr).FormulaR1C1 = _
        "=IFERROR(INDEX(DB!C[6],MATCH('TMP1'!RC[-6],DB!C[5],0)),"""")"
        
    Range("J1").FormulaR1C1 = "PIC"
    Range("J2:J" & lr).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],DB!C[4]:C[5],1,FALSE),"""")"
    
    Range("K1").FormulaR1C1 = "TOTAL"
    Range("K2:K" & lr).FormulaR1C1 = "=RC[-7]"
    
    Range("L1").FormulaR1C1 = "FULL"
    Range("L2:L" & lr).FormulaR1C1 = "=RC[-7]"
    
    Range("M1").FormulaR1C1 = "PENDING"
    Range("M2:M" & lr).FormulaR1C1 = "=RC[-6]"
    
    Range("N1").FormulaR1C1 = "BULAN"
    Range("N2:N" & lr).FormulaR1C1 = "=RC[-12]"
    
    Range("O1").FormulaR1C1 = "NO BULAN"
    Range("O2:O" & lr).FormulaR1C1 = "=RC[-14]"
    
    With Range(Cells(1, 9), Cells(lr, 15))
        .NumberFormat = "General"
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
    
    TMP1.AutoFilterMode = False
    Set rng = Range(Cells(1, 9), Cells(lr, 15))
    rng.AutoFilter 1, "<>"
    rng.Resize(rng.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
    TMP2.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    
    '||----------------------------------||
        Call CreatePivot_WO_Purchasing
    '||----------------------------------||
    
    TMP3.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    With Range(Cells(1, 1), Cells(lr, lc))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
    Rows(1).Delete
    
    Cells.Copy
    shPurchasing.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    sumLoop = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column - 1
    SumPerformance = sumLoop / 2
    
    '{{--------- ISI PLAN, REALISASI ----------}}
    shAlert.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    Col = 5
    iterations = 0
    x = 1
    Do While iterations < sumLoop
        HeaderName = UCase(Cells(4, Col).Value)
        Select Case HeaderName
            Case "PLANN", "REALISASI"
                iterations = iterations + 1
                x = x + 1
                For i = 5 To lr
                    If Cells(i, Col) = "" And Cells(i, 3) = "WO Purchasing" Then
                        Cells(i, Col).FormulaR1C1 = _
                            "=IFERROR(VLOOKUP(RC1,'WO PURCHASING'!C1:C[-3]," & x & ",FALSE),"""")"
                    End If
                Next i
        End Select
        Col = Col + 1
    Loop
    
    '{{--------- PERFORMANCE ----------}}
    Col = 5
    iterations = 0
    Do While iterations < SumPerformance
        HeaderName = UCase(Cells(4, Col).Value)
        Select Case HeaderName
            Case "PERFORMANCE"
                iterations = iterations + 1
                For i = 5 To lr
                    If Cells(i, Col) = "" And Cells(i, 3) = "WO Purchasing" Then
                        Cells(i, Col).NumberFormat = "0.00%"
                        Cells(i, Col).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],"""")"
                    End If
                Next i
        End Select
        Col = Col + 1
    Loop
    
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    With Range(Cells(5, 5), Cells(lr, lc))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
        
    '><><> .................... <><><
    '        Cleancing Data
    '><><> .................... <><><
    TMP1.AutoFilterMode = False: TMP2.AutoFilterMode = False: TMP3.AutoFilterMode = False: TMP4.AutoFilterMode = False: TMP5.AutoFilterMode = False
    TMP1.Cells.ClearContents: TMP2.Cells.ClearContents: TMP3.Cells.ClearContents: TMP4.Cells.ClearContents: TMP5.Cells.ClearContents
    shPurchasing.Delete
    '><><> .................... <><><

    If TPL.Visible = xlSheetVisible Then TPL.Visible = xlSheetHidden
    HOME.Activate
    Cells(1, 1).Select
    '{{------------------ DONE -------------------}}
    
End Sub

'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] Create Pivot WO-Purchasing
'...................................................................................................................................
Sub CreatePivot_WO_Purchasing()
    TMP2.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set pv_Rng = Range(Cells(1, 1), Cells(lr, lc))
    Set pv_Cache = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=TMP3.Range("A1"), _
                        TableName:="pv_Purchasing")
                        
    TMP3.Activate
    
    With pv_Tb
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .ColumnGrand = False
        .ShowValuesRow = False
        .DisplayErrorString = False
        .RepeatAllLabels xlRepeatLabels
    End With
    
    With pv_Tb.PivotFields("MANAGER")
        .Caption = "MANAGER"
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("BULAN")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With pv_Tb.PivotFields("TOTAL")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
    End With
    
    With pv_Tb.PivotFields("FULL")
        .Orientation = xlDataField
        .Position = 2
        .Function = xlSum
    End With
    
    Set pv_Rng = Nothing
    Set pv_Cache = Nothing
    Set pv_Tb = Nothing
    
End Sub

