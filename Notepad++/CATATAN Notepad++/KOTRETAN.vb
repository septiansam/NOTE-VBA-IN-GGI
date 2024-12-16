

'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] Olahan Untuk Performance WO Preparation
'...................................................................................................................................

Sub WO_Preparation()
    
    Path_File = HOME.Range("D" & 5) & Application.PathSeparator & HOME.Range("E" & 5) & HOME.Range("F" & 5)
    
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
                    Range(Cells(1, 1), Cells(lr, lc)).Copy
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
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    
    '{{--------- Hapus Baris Yang Tidak Diperlukan dan Buat Huader Untuk Filter ----------}}
    lr = TMP1.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    For i = lr To 1 Step -1
        If Application.WorksheetFunction.CountA(Rows(i)) = 3 Then Rows(i).Delete
    Next i
    TMP1.Activate
    Rows(1).Insert
    Range("B1") = "FILTER"
    '---------------------------------------------------------------------------------------
    
    
    '><><> .........PRE-PROCESSING......... <><><
    '== >>>   LOOPING BERDASARKAN BULAN    <<< ==
    '== >>>    Urutkan Terlebih Dahulu     <<< ==
    '><><> ................................ <><><
    
    lr = TMP1.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = TMP1.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    With TMP1.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("A2:A" & lr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range(Cells(2, 1), Cells(lr, lc))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    
    Set rng = Range(Cells(2, 2), Cells(Cells(Rows.Count, 2).End(xlUp).Row, 2))
    Set Dict = CreateObject("Scripting.Dictionary")
    For Each Cell In rng
        If Not Dict.exists(Cell.Value) And Cell.Value <> "" Then
            Dict.Add Cell.Value, Nothing
        End If
    Next Cell
    Set rng = Nothing
    
    'SETTING TEMPLATE
    shAlert.Activate
    Col = 5
    Sub_Header = Array("PLANN", "REALISASI", "PERFORMANCE")
    For Each Key In Dict.Keys
        For i = LBound(Sub_Header) To UBound(Sub_Header)
            Cells(3, Col) = UCase(Key)
            Cells(4, Col) = UCase(Sub_Header(i))
            Col = Col + 1
        Next i
    Next Key
    Cells.EntireColumn.AutoFit
    
    TMP1.Activate
    i = 1
    For Each Key In Dict.Keys
        TMP1.Activate
        TMP1.AutoFilterMode = False
        TMP1.Cells.AutoFilter 2, Key
        Set FirstRange = TMP1.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Range("C1")
        Set LastRange = TMP1.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Range("C" & Rows.Count).End(xlUp)
        Set RangeFilter = Range(FirstRange, LastRange)
        TMP2.Activate
        Cells(1, i) = Key
        rPaste = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row + 1
        RangeFilter.Copy
        Cells(rPaste, i).PasteSpecial xlPasteAll: Application.CutCopyMode = False
        Set FirstRange = Nothing
        Set LastRange = Nothing
        Set RangeFilter = Nothing
        i = i + 1
    Next Key
    Cells.EntireColumn.AutoFit
    SumMonth = Range("SAM1").End(xlToLeft).Column

    TMP2.Activate
    col_PIC = Range("SAM1").End(xlToLeft).Column + 1
    col_Month = Range("SAM1").End(xlToLeft).Column + 2
    
    Cells(1, col_PIC) = "PIC"
    Cells(1, col_Month) = "BULAN"
    
    rPIC = 2
    For i = 1 To SumMonth
        StrMonth = Cells(1, i)
        SumPic = Cells(Rows.Count, i).End(xlUp).Row
        For j = rPIC To SumPic
            If Right(Cells(j, i), 4) <> ".PIC" And Cells(j, i) <> "" Then
                Cells(j, col_PIC) = Cells(j, i)
                Cells(j, col_Month) = StrMonth
            End If
        Next j
        rPIC = Cells(Rows.Count, i).End(xlUp).Row + 1
    Next i
    
    TMP2.Activate
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    col_Paste = lc + 1
    
    TMP1.Activate
    TMP1.AutoFilterMode = False
    lr = TMP1.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = TMP1.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set rng = TMP1.Range(TMP1.Cells(1, 3), TMP1.Cells(lr, lc))
    rPaste = 2
    For Each Key In Dict.Keys
        TMP1.Activate
        TMP1.AutoFilterMode = False
        TMP1.Cells.AutoFilter 2, Key
        rng.Offset(1).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        TMP2.Activate
        Cells(rPaste, col_Paste).PasteSpecial xlPasteAll: Application.CutCopyMode = False
        rPaste = Cells(Rows.Count, col_Paste).End(xlUp).Row + 1
    Next Key
    
    '======================================================
    '+++++----            PERHITUNGAN             ----+++++
    '-+_+-+_+-+_ [TOTAL, RELEASE, NOT RELEASE] _+-+_+-+_+-
    '======================================================
    TMP2.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    col_Total = lc + 1
    col_Release = lc + 2
    col_NotRelease = lc + 3
    

    Cells(1, col_Total).FormulaR1C1 = "PLANN"
    Range(Cells(2, col_Total), Cells(lr, col_Total)).FormulaR1C1 = "=IFERROR(RC[-15]+RC[-10]+RC[-5],"""")"
    Cells(1, col_Release).FormulaR1C1 = "REALISASI"
    Range(Cells(2, col_Release), Cells(lr, col_Release)).FormulaR1C1 = "=IFERROR(RC[-15]+RC[-10]+RC[-5],"""")"
    Cells(1, col_NotRelease).FormulaR1C1 = "NOT RELEASE"
    Range(Cells(2, col_NotRelease), Cells(lr, col_NotRelease)).FormulaR1C1 = "=IFERROR(RC[-14]+RC[-9]+RC[-4],"""")"
    
    With Range(Cells(1, col_Total), Cells(lr, col_NotRelease))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
    '-------------------------------------------------------------------------------------------------------------------------------
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    '><><> ............................... <><><
    '         OLAH DATA KE SHEETS TES3
    '><><> ``````````````````````````````` <><><
    Set rng = Nothing
    TMP2.Activate: TMP2.AutoFilterMode = False
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set rng = Range(Cells(1, col_PIC), Cells(lr, col_Month))
    Cells.AutoFilter col_PIC, "<>"
    rng.Resize(rng.Rows.Count).SpecialCells(xlCellTypeVisible).Copy
    TMP3.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    TMP1.AutoFilterMode = False: TMP2.AutoFilterMode = False
    
    Range("B:B").Insert xlToRight
    Range("B1") = "MANAGER"
    lr = Range("A" & Rows.Count).End(xlUp).Row
    
    Range("B1").FormulaR1C1 = "MANAGER"
    Range("B2:B" & lr).FormulaR1C1 = "=VLOOKUP(RC[-1],DB!C1:C2,2,FALSE)"
    Range("D1").FormulaR1C1 = "NO BULAN"
    Range("D2:D" & lr).FormulaR1C1 = _
        "=INDEX('TMP1'!C[-3],MATCH('TMP3'!RC[-1],'TMP1'!C[-2],0))"
    With Range(Cells(1, 2), Cells(lr, 4))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
    Set rng = Nothing
    TMP2.Activate: TMP2.AutoFilterMode = False
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set rng = Range(Cells(1, col_Total), Cells(lr, col_NotRelease))
    Cells.AutoFilter col_PIC, "<>"
    
    rng.SpecialCells(xlCellTypeVisible).Copy
    TMP3.Activate
    Range("E1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    TMP1.AutoFilterMode = False: TMP2.AutoFilterMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    
    '##===>>>...............<<<===##
    Call CreatePivot_WO_Preparation
    '##===>>>...............<<<===##
    
    TMP4.Activate
    
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    With Range(Cells(1, 1), Cells(lr, lc))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
    Rows(1).Delete
    
    Cells.Copy
    shPreparation.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select

    '{{--------- ISI PLAN, REALISASI, PERFORMANCE ----------}}
    shAlert.Activate
    lr = Range("C" & Rows.Count).End(xlUp).Row
    lc = Cells(4, Columns.Count).End(xlToLeft).Column
    x = 1
    For j = 5 To lc
        If Cells(4, j) <> "PERFORMANCE" Then
            x = x + 1
        End If
'        Range(Cells(5, j), Cells(lr, j)).NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Range(Cells(5, j), Cells(lr, j)).FormulaR1C1 = _
            "=IFERROR(IF(AND(RC3=""WO Preparation"",R4C<>""PERFORMANCE""),VLOOKUP(RC1,'WO PREPARATION'!C1:C[-3]," & x & ",FALSE),""""),"""")"
    Next j
    
    For j = 5 To lc
        If Cells(4, j) = "PERFORMANCE" Then
            Range(Cells(5, j), Cells(lr, j)).NumberFormat = "0.00%"
            Range(Cells(5, j), Cells(lr, j)).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],"""")"
        End If
    Next j
    
    With Range(Cells(5, 5), Cells(lr, lc))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
    
    With Range(Cells(3, 5), Cells(4, lc))
        .Font.Bold = True
        .Font.Size = 14
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .EntireColumn.AutoFit
    End With
    
    Range("A:A").Font.Bold = True
    Range("A4").Font.Size = 14
    
    
    '><><> .................... <><><
    '        Cleancing Data
    '><><> .................... <><><
    TMP1.Cells.ClearContents
    TMP2.Cells.ClearContents
    TMP3.Cells.ClearContents
    TMP4.Cells.ClearContents
    TMP5.Cells.ClearContents
    shPreparation.Delete
    '><><> .................... <><><

    If DB.Visible = xlSheetVisible Then DB.Visible = xlSheetHidden
    If TPL.Visible = xlSheetVisible Then TPL.Visible = xlSheetHidden
    HOME.Activate
    Cells(1, 1).Select
    '{{------------------ DONE -------------------}}
    
    
End Sub

'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] Create Pivot WO-Preparation
'...................................................................................................................................
Sub CreatePivot_WO_Preparation()
    TMP3.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set pv_Rng = Range(Cells(1, 1), Cells(lr, lc))
    Set pv_Cache = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=TMP4.Range("A1"), _
                        TableName:="pv_Preparation")
                        
    TMP4.Activate
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
    
    With pv_Tb.PivotFields("PLANN")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
    End With
    
    With pv_Tb.PivotFields("REALISASI")
        .Orientation = xlDataField
        .Position = 2
        .Function = xlSum
    End With
    
    Set pn_rng = Nothing
    Set pv_Cache = Nothing
    Set pv_Tb = Nothing
    
End Sub

