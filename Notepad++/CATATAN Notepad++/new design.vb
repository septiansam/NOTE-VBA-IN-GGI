
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] Olahan Untuk Performance WO Budgeting
'...................................................................................................................................

Sub WO_Budgeting()
    
    MonthNames = Array( _
        "Januari", "January", "Februari", "February", "Maret", "March", "April", "Mei", "May", "Juni", "June", "Juli", "July", _
        "Agustus", "August", "September", "Oktober", "October", "November", "Desember", "December" _
        )
    MonthNumbers = Array(1, 1, 2, 2, 3, 3, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 10, 10, 11, 12, 12)
    
    x = 1
    For i = LBound(MonthNames) To UBound(MonthNames)
        Path_File = HOME.Range("D" & 7)
        Nama_File = MonthNames(i)
        Path_File = Path_File & Application.PathSeparator & Nama_File & ".xlsx"
        If Dir(Path_File) <> "" Then
            Set WB_FILE = Workbooks.Open(Path_File): WB_FILE.Activate
            Set SH = WB_FILE.Sheets(1)
            SH.Activate
            SH.AutoFilterMode = False
            lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
            lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
            If x = 1 Then
                Set rng = Range(Cells(1, 1), Cells(lr, lc))
                rng.Copy
                Windows(TWB.Name).Activate: TMP1.Activate
                Range("A1") = "NO BULAN": Range("B1") = "BULAN"
                rPaste = 1
                Range("C" & rPaste).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
                lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
                Range("A" & rPaste + 1, "A" & lr) = MonthNumbers(i)
                Range("B" & rPaste + 1, "B" & lr) = MonthNames(i)
            Else
                Set rng = Range(Cells(2, 1), Cells(lr, lc))
                rng.Copy
                Windows(TWB.Name).Activate: TMP1.Activate
                rPaste = Range("C" & Rows.Count).End(xlUp).Row + 1
                Range("C" & rPaste).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
                lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
                Range("A" & rPaste, "A" & lr) = MonthNumbers(i)
                Range("B" & rPaste, "B" & lr) = MonthNames(i)
            End If
            
            Set SH = Nothing
            WB_FILE.Close False
            x = x + 1
        End If
    Next i
    Set WB_FILE = Nothing
    
    TMP1.Activate
    lr = Range("C" & Rows.Count).End(xlUp).Row
    lc = Cells(4, Columns.Count).End(xlToLeft).Column
    Range("A1:C" & lr).Copy TMP2.Range("A1")
    Range("L1:L" & lr).Copy TMP2.Range("D1")
    Range("N1:N" & lr).Copy TMP2.Range("E1")
    TMP2.Activate
    Cells.EntireColumn.AutoFit
    
    '*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
    Call CreatePivot_WO_Budgeting
    '*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
    
    TMP3.Activate
    lr = Range("C" & Rows.Count).End(xlUp).Row
    lc = Cells(4, Columns.Count).End(xlToLeft).Column
    With Range(Cells(1, 1), Cells(lr, lc))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
    Rows(1).Delete
    
    Cells.Copy
    shBudgeting.Activate
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
                    If Cells(i, Col) = "" And Cells(i, 3) = "WO Budgeting" Then
                        Cells(i, Col).FormulaR1C1 = _
                            "=IFERROR(VLOOKUP(RC1,'WO BUDGETING'!C1:C[-3]," & x & ",FALSE),"""")"
                        Cells(i, Col).NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_)"
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
    
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
    Call OLAH_DATA_BUDGED
    Call CreatePivot_Budgeting_Performance
    
'[*].. UNTUK MENDAPATKAN NILAI SETIAP PIC
'  ... --------- PERFORMANCE ----------
    Call Processing_Pivot_Performance
    
'[*].. UNTUK MEMBUAT SHEETS LAPORAN MONTHLY BUDGET
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    Call Create_Monthly_Budget
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
    
'[*].. UNTUK MEMBUAT SHEETS LAPORAN MONTHLY BUDGET DETAILS
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    Call Create_Monthly_Budget_Details
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_

    '><><> .................... <><><
    '        Cleancing Data
    '><><> .................... <><><
    TMP1.AutoFilterMode = False: TMP2.AutoFilterMode = False: TMP3.AutoFilterMode = False: TMP4.AutoFilterMode = False: TMP5.AutoFilterMode = False
    TMP1.Cells.Clear: TMP2.Cells.Clear: TMP3.Cells.Clear: TMP4.Cells.Clear: TMP5.Cells.Clear
    '><><> .................... <><><

    If TPL.Visible = xlSheetVisible Then TPL.Visible = xlSheetHidden
    HOME.Activate
    Cells(1, 1).Select
    '{{------------------ DONE -------------------}}
    

End Sub

'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] PROCESSING MONTHLY BUDGET
'.. [*] Pengolahan untuk mengisi membuat laporan monthly budged
'-----------------------------------------------------------------------------------------------------------------------------------
Sub Create_Monthly_Budget()
    
'.. [*] PROCESSING MONTHLY BUDGET
'==================================================
If wsx("Monthly Budget") Then Sheets("Monthly Budget").Delete
Set shMonth = Sheets.Add(after:=Sheets(Sheets.Count)): shMonth.Name = "Monthly Budget"
TMP2.Activate
Rows(1).Delete
Range("E1") = "FACTORY": Range("D1") = "BRANCH"
lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column

'.. [*] CREATE PIVOT FOR MONTHLY BUDGET (ONLY PERFORMANCE)
'==================================================
    TMP2.Sort.SortFields.Clear
    TMP2.Sort.SortFields.Add2 Key:=Range("A2:A" & lr) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With TMP2.Sort
        .SetRange Range(Cells(1, 1), Cells(lr, lc))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Set pv_Rng = Range(Cells(1, 1), Cells(lr, lc))
    Set pv_Cache = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=TMP5.Range("A1"), _
                        TableName:="pv_Budgeting_Monthly")
                        
    TMP5.Activate
    With pv_Tb
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .NullString = 0
'        .ColumnGrand = False
'        .ShowValuesRow = False
        .DisplayErrorString = False
        .RepeatAllLabels xlRepeatLabels
        .PivotCache.MissingItemsLimit = xlMissingItemsDefault
    End With
    
    With pv_Tb.PivotFields("PIC")
        .Caption = "PIC"
        .Orientation = xlRowField
        .Position = 1
'        .Subtotals = _
'            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("KATEGORI")
        .Caption = "KATEGORI"
        .Orientation = xlRowField
        .Position = 2
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("DEFINISI")
        .Caption = "DEFINISI"
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("FACTORY")
        .Caption = "FACTORY"
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("PIC-2")
        .Caption = "PIC-2"
        .Orientation = xlRowField
        .Position = 5
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("BULAN")
        .Orientation = xlColumnField
        .Position = 1
    End With

    '''ISI'''

    With pv_Tb.PivotFields("PERFORMANCE")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "0.00%"
    End With
    

    Set pv_Rng = Nothing
    Set pv_Cache = Nothing
    Set pv_Tb = Nothing
    Set rng = Nothing
    
'[*]..End Pivot
'==================================================


'[*]..Create Average
'-------------------
lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
Set rng = Range(Cells(1, 1), Cells(lr, lc))
With rng
    .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
End With
Set rng = Nothing
Rows(1).Delete
lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
Cells(1, lc + 1) = "AVERAGE"
Range(Cells(2, lc + 1), Cells(lr, lc + 1)).NumberFormat = "0.00%"
Range(Cells(2, lc + 1), Cells(lr, lc + 1)).FormulaR1C1 = "=AGGREGATE(1,6,RC[-" & lc + 1 - 6 & "]:RC[-1])"
lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
Set rng = Range(Cells(1, 1), Cells(lr, lc))
rng.Copy
shMonth.Activate
Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
Set rng = Nothing

'[*]..DESIGN MONTHLY BUDGET
'-------------------
shMonth.Activate
lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
Set rng = Range(Cells(1, 1), Cells(lr, lc))
Cells.Font.Name = "Trebuchet MS"
Cells.EntireRow.AutoFit

With Range(Cells(1, 1), Cells(1, lc))
    .Font.Bold = True
    .Font.Name = "Century Gothic"
    .Font.Size = 12
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColor = xlAutomatic
    .Interior.Color = RGB(31, 76, 81)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .EntireRow.RowHeight = .EntireRow.RowHeight + 3
End With

'[*]..Design Dasar
'(------------------------------------------------)
For i = 2 To lr
    If i Mod 2 = 0 Then
        With Range(Cells(i, 1), Cells(i, lc))
            .Interior.Pattern = xlSolid
            .Interior.PatternColor = xlAutomatic
            .Interior.Color = RGB(228, 240, 241)
        End With
    ElseIf i Mod 2 <> 0 Then
        With Range(Cells(i, 1), Cells(i, lc))
            .Interior.Pattern = xlSolid
            .Interior.PatternColor = xlAutomatic
            .Interior.Color = RGB(255, 255, 255)
        End With
    End If
Next i

'[*]..Design Sub Total
'(------------------------------------------------)
For i = 2 To lr - 1
    If Right(Cells(i, 1), 5) = "Total" Then
        With Range(Cells(i, 1), Cells(i, lc))
            .Font.Bold = True
            .Interior.Pattern = xlSolid
            .Interior.PatternColor = xlAutomatic
            .Interior.Color = RGB(0, 172, 168)
            .EntireRow.RowHeight = .EntireRow.RowHeight + 1
        End With
    End If
Next i

'[*]..Design Grand Total
'(------------------------------------------------)
With Range(Cells(lr, 1), Cells(lr, lc))
    .Font.Bold = True
    .Font.Size = 12
    .Font.Name = "Century Gothic"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColor = xlAutomatic
    .Interior.Color = RGB(31, 76, 81)
    .EntireRow.RowHeight = .EntireRow.RowHeight + 2
End With

Cells.EntireColumn.AutoFit
For Each Cell In rng.Columns
    Cell.ColumnWidth = Cell.ColumnWidth + 2
Next Cell

rng.Borders.LineStyle = xlContinuous
rng.Borders.Color = RGB(0, 108, 105)
rng.AutoFilter

Rows("1:2").Insert
With Range("A1")
    .Value = "MONTHLY BUDGET"
    .Font.Name = "Century Gothic"
    .Font.Size = 22
    .Font.Bold = True
    .VerticalAlignment = xlCenter
    .HorizontalAlignment = xlLeft
End With

Range("A:A").Insert
Range("A:A").ColumnWidth = 5
Rows(4).Select
ActiveWindow.FreezePanes = True
Cells(1, 1).Select

End Sub

'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] PROCESSING MONTHLY BUDGET DETAILS
'.. [*] Pengolahan untuk mengisi membuat laporan monthly budged details
'-----------------------------------------------------------------------------------------------------------------------------------
Sub Create_Monthly_Budget_Details()
    
    TMP4.AutoFilterMode = False: TMP5.AutoFilterMode = False
    TMP5.Cells.ClearContents
    If wsx("Monthly Budget Detail") Then Sheets("Monthly Budget Detail").Delete
    Set shMonthDetails = Sheets.Add(after:=Sheets(Sheets.Count)): shMonthDetails.Name = "Monthly Budget Detail"
    TMP4.Activate
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
'[*].. UNTUK MENGHAPUS KOLOM SELAIN BEST, CONTRACK, PERFORMANCE
'--------------------------------------------------------------
    For i = lc To 6 Step -1
        If Cells(2, i).Value = "Sum of PLAN BUDGET" Then
            Cells(2, i).Value = "PLAN BUDGET"
        ElseIf Cells(2, i).Value = "Sum of ACTUAL" Then
            Cells(2, i).Value = "ACTUAL"
        ElseIf Cells(2, i).Value = "Sum of BOBOT" Then
            Cells(2, i).Value = "BOBOT"
        ElseIf Cells(2, i).Value = "Sum of BEST" Then
            Cells(2, i).Value = "BEST"
        ElseIf Cells(2, i).Value = "Sum of ONTRACK" Then
            Cells(2, i).Value = "ONTRACK"
        ElseIf Cells(2, i).Value = "Sum of PERFORMANCE" Then
            Cells(2, i).Value = "PERFORMANCE"
        End If
    Next i
    
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    Range(Cells(1, 1), Cells(lr, lc)).Copy
    shMonthDetails.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit
    
'[*].. DESIGN
'--------------------------------------------------------------
    shMonthDetails.Activate
    
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set rng = Range(Cells(2, 1), Cells(lr, lc))
    Cells.Font.Name = "Trebuchet MS"
    Cells.EntireRow.AutoFit
    
    '[*]..Design Awal Header
    '(------------------------------------------------)
    With Range(Cells(1, 1), Cells(1, lc))
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Range(Cells(2, 1), Cells(2, lc))
        .Font.Bold = True
        .Font.Name = "Century Gothic"
        .Font.Size = 12
        .Font.Color = vbWhite
        .Interior.Pattern = xlSolid
        .Interior.PatternColor = xlAutomatic
        .Interior.Color = RGB(31, 76, 81)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .EntireRow.RowHeight = .EntireRow.RowHeight + 3
    End With
    
    '[*]..Design Dasar
    '(------------------------------------------------)
    For i = 3 To lr
        If i Mod 2 = 0 Then
            With Range(Cells(i, 1), Cells(i, lc))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(228, 240, 241)
            End With
        ElseIf i Mod 2 <> 0 Then
            With Range(Cells(i, 1), Cells(i, lc))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(255, 255, 255)
            End With
        End If
    Next i
    
    '[*]..Design Sub Total
    '(------------------------------------------------)
    For i = 2 To lr - 1
        If Right(Cells(i, 1), 5) = "Total" Then
            With Range(Cells(i, 1), Cells(i, lc))
                .Font.Bold = True
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(0, 137, 134)
                .EntireRow.RowHeight = .EntireRow.RowHeight + 1
            End With
        End If
    Next i
    
    Cells.EntireColumn.AutoFit
    Range("B:B").ColumnWidth = 28
    Range("B:B").ColumnWidth = 29
    Range("C:C").ColumnWidth = 30
    
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(0, 108, 105)
    rng.AutoFilter
    
    Rows("1:2").Insert
    With Range("A1")
        .Value = "MONTHLY BUDGET (Detail)"
        .Font.Name = "Century Gothic"
        .Font.Bold = True
        .Font.Size = 22
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
    
    Range("A:A").Insert
    Range("A:A").ColumnWidth = 5
    Range("G5").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.Zoom = 90
    Cells(1, 1).Select
    
    
End Sub



'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] PROCESSING PIVOT BUDGETING PERFORMANCE
'.. [*] Pengolahan untuk mengisi performa di sheets alert
'-----------------------------------------------------------------------------------------------------------------------------------
Sub Processing_Pivot_Performance()

    TMP3.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Range(Cells(1, 1), Cells(lr, lc)).Copy
    TMP4.Activate
    Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    Rows(1).Delete
    
    shAlert.Activate
    sumLoop = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    x = 11
    For i = 5 To sumLoop
        If Cells(4, i).Value = "PERFORMANCE" Then
            For j = 5 To lr
                If Cells(j, i) = "" And Cells(j, 3) = "WO Budgeting" Then
                    Cells(j, i).FormulaR1C1 = "=SUMIF('TMP4'!C1,ALERT!RC1,'TMP4'!C" & x & ")"
                End If
            Next j
            x = x + 6
        End If
    Next i
    
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    With Range(Cells(5, 5), Cells(lr, lc))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
    

End Sub
'...................................................................................................................................


'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] OLAH BUDGED
'...................................................................................................................................
Sub OLAH_DATA_BUDGED()
    TMP2.AutoFilterMode = False: TMP3.AutoFilterMode = False: TMP4.AutoFilterMode = False: TMP5.AutoFilterMode = False
    TMP2.Cells.ClearContents: TMP3.Cells.ClearContents: TMP4.Cells.ClearContents: TMP5.Cells.ClearContents
    
    TMP1.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    Range("A:A").Insert
    Range("D:D").Copy Range("A:A")
    Range("D1:E" & lr).Copy
    TMP2.Activate
    
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    Range("A1") = "PIC": Range("B1") = "KATEGORI": Range("C1") = "DEFINISI"
    Range("C2:C" & lr).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],DB!C5:C6,2,FALSE),"""")"
    TMP1.Range("G:G").Copy TMP2.Range("D1")
    Range("E1").FormulaR1C1 = "FACTORY2"
    Range("E2:E" & lr).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],DB!C8:C9,2,FALSE),"""")"
    TMP1.Range("I:I").Copy TMP2.Range("F1")
    TMP1.Range("C:C").Copy TMP2.Range("G1")
    TMP1.Range("B:B").Copy TMP2.Range("H1")
    Range("F1") = "PIC-2"
    Range("I1") = "CON"
    Range("I2:I" & lr).FormulaR1C1 = "=CONCATENATE(RC[-8],""_"",RC[-2])"
    TMP1.Range("M:M").Copy TMP2.Range("J1")
    TMP1.Range("O:O").Copy TMP2.Range("K1")
    Range("J1") = "PLAN BUDGET"
    Range("K1") = "ACTUAL"
    Range("L1") = "SUM PLANN"
    Range("M1") = "BOBOT"
    Range("N1") = "BEST"
    Range("O1") = "ONTRACK"
    Range("P1") = "PERFORMANCE"
    Range("J:L").NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_)"
    Range("M:M").NumberFormat = "0.00%"
    Range("N:O").NumberFormat = "_(* #,##0.00_);[Red]_(* (#,##0.00);_(* ""-""??_)"
    Range("P:P").NumberFormat = "0.00%"
    Rows(1).Insert
    Range("N1") = "80%"
    Range("O1") = "100%"

    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    Range(Cells(3, 12), Cells(lr, 12)).FormulaR1C1 = "=SUMIF(R3C9:R" & lr & "C9,RC[-3],R3C10:R" & lr & "C10)"
    Range(Cells(3, 13), Cells(lr, 13)).FormulaR1C1 = "=IFERROR(RC[-3]/RC[-1],"""")"
    Range(Cells(3, 14), Cells(lr, 14)).FormulaR1C1 = "=RC[-4]*R1C14"
    Range(Cells(3, 15), Cells(lr, 15)).FormulaR1C1 = "=RC[-5]*R1C15"
    Range(Cells(3, 16), Cells(lr, 16)).FormulaR1C1 = _
        "=IF(RC[-1]=0,0,IF(RC[-5]>RC[-1],0,RC[-1]/IF(RC[-5]>=RC[-2],RC[-5],RC[-2])*RC[-3]))"
    Range("N1:O1").NumberFormat = "0%"
    Range("P1").FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[1048575]C)"
    Cells.EntireColumn.AutoFit
    
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    With Range(Cells(1, 1), Cells(lr, lc))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
    
    
End Sub

'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] Create Pivot WO-Performance-Budgeting
'...................................................................................................................................
Sub CreatePivot_Budgeting_Performance()

    TMP2.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set pv_Rng = Range(Cells(2, 1), Cells(lr, lc))
    Set pv_Cache = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=TMP3.Range("A1"), _
                        TableName:="pv_Budgeting_Performance")
                        
    TMP3.Activate
    With pv_Tb
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .NullString = "-"
        .ColumnGrand = False
'        .ShowValuesRow = False
        .DisplayErrorString = False
        .RepeatAllLabels xlRepeatLabels
        .PivotCache.MissingItemsLimit = xlMissingItemsDefault
    End With
    
    With pv_Tb.PivotFields("PIC")
        .Caption = "PIC"
        .Orientation = xlRowField
        .Position = 1
'        .Subtotals = _
'            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("KATEGORI")
        .Caption = "KATEGORI"
        .Orientation = xlRowField
        .Position = 2
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("DEFINISI")
        .Caption = "DEFINISI"
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("FACTORY2")
        .Caption = "FACTORY2"
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("PIC-2")
        .Caption = "PIC-2"
        .Orientation = xlRowField
        .Position = 5
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("BULAN")
        .Orientation = xlColumnField
        .Position = 1
    End With

    '''ISI'''
    
    With pv_Tb.PivotFields("PLAN BUDGET")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_)"
    End With
    
    With pv_Tb.PivotFields("ACTUAL")
        .Orientation = xlDataField
        .Position = 2
        .Function = xlSum
        .NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_)"
    End With
    
    With pv_Tb.PivotFields("BOBOT")
        .Orientation = xlDataField
        .Position = 3
        .Function = xlSum
        .NumberFormat = "0.00%"
    End With
    
    With pv_Tb.PivotFields("BEST")
        .Orientation = xlDataField
        .Position = 4
        .Function = xlSum
        .NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_)"
    End With
    
    With pv_Tb.PivotFields("ONTRACK")
        .Orientation = xlDataField
        .Position = 5
        .Function = xlSum
        .NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_)"
    End With
    
    With pv_Tb.PivotFields("PERFORMANCE")
        .Orientation = xlDataField
        .Position = 6
        .Function = xlSum
        .NumberFormat = "0.00%"
    End With
    

    Set pv_Rng = Nothing
    Set pv_Cache = Nothing
    Set pv_Tb = Nothing


End Sub

'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------------------------------------
'.. [*] Create Pivot WO-Budgeting
'...................................................................................................................................
Sub CreatePivot_WO_Budgeting()
    TMP2.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set pv_Rng = Range(Cells(1, 1), Cells(lr, lc))
    Set pv_Cache = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=TMP3.Range("A1"), _
                        TableName:="pv_Budgeting")
                        
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
    
    With pv_Tb.PivotFields("LIMIT($)")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
    End With
    
    With pv_Tb.PivotFields("ACTUAL($)")
        .Orientation = xlDataField
        .Position = 2
        .Function = xlSum
    End With
    
    Set pv_Rng = Nothing
    Set pv_Cache = Nothing
    Set pv_Tb = Nothing
    
End Sub

