Sub ANALISA(Optional Periode As String)

Dim TWB As Workbook
Dim SH_HASIL As Worksheet
Dim OLAH As Worksheet, OLAH2 As Worksheet
Dim TEMP1 As Worksheet, TEMP2 As Worksheet
Dim LR As Long, LC As Long, LC2 As Long, i As Long, cell As Range
Dim WS As Worksheet

Application.DisplayAlerts = False

Set TWB = ThisWorkbook
Set SH_HASIL = TWB.Sheets("HASIL")
Set OLAH = TWB.Sheets("OLAH")

For Each WS In TWB.Worksheets
    If WS.Name <> "TOMBOL" And _
        WS.Name <> "DATA IM" And _
        WS.Name <> "DATA OV" And _
        WS.Name <> "Cari OV" And _
        WS.Name <> "Report WO Buyer" And _
        WS.Name <> "OLAH" And _
        WS.Name <> "HASIL" Then
    
        WS.Delete
        
    End If
Next WS

For i = 1 To 2
    If WorksheetExists("TEMP" & i) Then Sheets("TEMP" & i).Delete
    Sheets.Add(After:=TWB.Sheets(TWB.Sheets.Count)).Name = "TEMP" & i
Next i

Set TEMP1 = TWB.Sheets("TEMP1")
Set TEMP2 = TWB.Sheets("TEMP2")

OLAH.Activate
ActiveWindow.Zoom = 80
Range("K:Z").Delete
LR = OLAH.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = OLAH.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column
    
OLAH.Activate

Cells(1, LC + 1).Value = "KATEGORI"
With Range(Cells(2, LC + 1), Cells(LR, LC + 1))
        .FormulaR1C1 = _
        "=IFS(RC[-1]<=30,1,AND(RC[-1]>30,RC[-1]<=90),2,RC[-1]>90,3)"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

LC = OLAH.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

'[BUAT PIVOT KE SHEET TEMP1]................
Dim PV_RANGE As Range
Dim PV_TABLE As PivotTable
Dim PV_CACHE As PivotCache

Set PV_RANGE = OLAH.Range(Cells(1, 1), Cells(LR, LC))
Set PV_CACHE = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                    SourceData:=PV_RANGE)
Set PV_TABLE = PV_CACHE.CreatePivotTable _
                    (TableDestination:=TEMP1.Range("A1"), _
                    TableName:="PV_ANALISA")
TEMP1.Activate

'[INSERT FIELD FIELD NYA]...
With PV_TABLE.PivotFields("Buyer")
    .Caption = "Buyer"
    .Orientation = xlRowField
    .Position = 1
End With

With PV_TABLE.PivotFields("KATEGORI")
    .Caption = "Kategori    (Days)"
    .Orientation = xlRowField
    .Position = 2
End With

With PV_TABLE.PivotFields("G/L Cat")
    .Orientation = xlColumnField
    .Position = 1
End With

'[ISI].....
With PV_TABLE.PivotFields("Amount")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlSum
End With

'[SETTING]...
With PV_TABLE
    '.RowGrand = False
    .DisplayErrorString = False
    .NullString = 0
    .PageFieldOrder = 2
    .PreserveFormatting = True
    .PrintTitles = False
    .CompactRowIndent = 1
    .DisplayContextTooltips = True
    .ShowDrillIndicators = True
    .PrintDrillIndicators = False
    .AllowMultipleFilters = False
    .SortUsingCustomLists = True
    .FieldListSortAscending = False
    .ShowValuesRow = False
    .RowAxisLayout xlTabularRow
    .PivotCache.MissingItemsLimit = xlMissingItemsDefault
    .TableStyle2 = "PivotStyleDark7"
End With

LR = TEMP1.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = TEMP1.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

TEMP1.Range(Cells(1, 1), Cells(LR, LC)).Copy
TEMP2.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False

TEMP2.Activate
If Range("A1").Value = "Sum of Amount" Then Rows(1).Delete
Cells.EntireColumn.AutoFit
LR = TEMP2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = TEMP2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

'Memulai perulangan dari baris terakhir hingga baris pertama dengan langkah -1
For i = LR To 1 Step -1
    If Cells(i, "B").Value = 1 Then
        If Cells(i + 1, "B").Value <> 2 And Cells(i + 1, "B").Value <> 3 Then
            Rows(i + 1 & ":" & i + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i + 1, "B").Value = 2
            Cells(i + 2, "B").Value = 3
        ElseIf Cells(i + 1, "B").Value = 3 Then
            Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i + 1, "B").Value = 2
        End If
    ElseIf Cells(i, "B").Value = 2 Then
        If Cells(i - 1, "B").Value <> 1 And Cells(i + 1, "B").Value <> 3 Then
            Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i, "B").Value = 1
        ElseIf Cells(i - 1, "B").Value <> 1 And Cells(i + 1, "B").Value = 3 Then
            Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i, "B").Value = 1
            Cells(i, "A").Value = Cells(i + 1, "A").Value
            Cells(i + 1, "A").ClearContents
        ElseIf Cells(i - 1, "B").Value = 1 And Cells(i + 1, "B").Value <> 3 Then
            Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i + 1, "B").Value = 3
        End If
    ElseIf Cells(i, "B").Value = 3 Then
        If Cells(i - 1, "B").Value <> 1 And Cells(i - 1, "B").Value <> 2 Then
            Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i, "B").Value = 1
            Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i + 1, "B").Value = 2
            Cells(i, "A").Value = Cells(i + 2, "A").Value
            Cells(i + 2, "A").ClearContents
        ElseIf Cells(i - 1, "B").Value = 1 Then
            Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i, "B").Value = 2
        End If
    End If
Next i

'[ISI NILAI 0 PADA SEL YANG KOSONG DI G/L CAT]
'..................................................................
LR = TEMP2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = TEMP2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column
    
Dim RG_FILL As Range, RG_KATEGORI As Range
Set RG_FILL = Range(Cells(2, 3), Cells(LR, LC))

For Each cell In RG_FILL
    If cell.Value = vbNullString Then cell.Value = 0
Next cell

'Columns("C:G").NumberFormat = "_(* #,##0_);_((#,##0);_(* ""-""??_);_(@_)"
'Columns("C:G").NumberFormat = "_(* #,##0_);_((#,##0);_(    ""-""??_);_(@_)"
Columns("C:H").NumberFormat = "_(#,##0;_((#,##0);_(""-"";_(@_)"


'Set RG_KATEGORI = Range(Cells(1, 2), Cells(LR, 2))
'For Each cell In RG_KATEGORI
'    If cell.Value = 1 Then
'        cell.Value = "0 - 30"
'    ElseIf cell.Value = 2 Then
'        cell.Value = "31 - 90"
'    ElseIf cell.Value = 3 Then
'        cell.Value = ">90"
'    End If
'Next cell

Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter

Rows(1).HorizontalAlignment = xlCenter
Rows(1).VerticalAlignment = xlCenter
Range("B:B").VerticalAlignment = xlCenter

Rows(1).Insert
Cells.EntireColumn.AutoFit

Cells.Font.Name = "Verdana"
ActiveWindow.Zoom = 80

Range("A1:A2").Merge
Range("B1:B2").Merge
Range("B1:B2").WrapText = True

'Range("H1:H2").Merge
'Range("H1:H2").Value = "TOTAL"

Range(Cells(1, LC), Cells(2, LC)).Merge
Range(Cells(1, LC), Cells(2, LC)).Value = "TOTAL"


With Range(Cells(1, 3), Cells(1, LC - 1))
    .Merge
    .Value = "Amount($)"
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

With Range(Cells(1, 1), Cells(2, LC))
    .Font.Name = "Century Gothic"
    .Font.Color = vbWhite
    .Font.Bold = True
    .Font.Size = 12
    .Interior.Color = RGB(52, 98, 101)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .RowHeight = .RowHeight + 4
End With

LR = TEMP2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row
    
LC = TEMP2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

Range(Cells(3, 1), Cells(LR - 1, 1)).HorizontalAlignment = xlLeft

For i = 3 To LR
    If Cells(i, 1) <> "Grand Total" Then
        If Right(CStr(Cells(i, 1)), 5) Like "Total" Then
            
            With Range("A" & i)
                .Font.Italic = True
                .Font.Color = RGB(2, 112, 192)
                .Font.TintAndShade = 0
            End With
            
            With Range(Cells(i, 1), Cells(i, LC))
    '            .Font.Bold = True
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(228, 240, 241)
                .VerticalAlignment = xlCenter
                .RowHeight = .RowHeight + 2
            End With
            
        End If
    End If
    If Cells(i, 1) = "Grand Total" Then
        With Range(Cells(i, 1), Cells(i, LC))
            .Font.Bold = True
            .RowHeight = .RowHeight + 4
            .Interior.Pattern = xlSolid
            .Interior.PatternColor = xlAutomatic
            .Interior.Color = RGB(228, 240, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
Next i

With Range(Cells(1, 1), Cells(LR, LC))
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
End With

Dim col As Range
For Each col In Range(Cells(1, 1), Cells(LR, LC)).Columns
    col.EntireColumn.AutoFit
    col.ColumnWidth = col.ColumnWidth + 2
Next col

LR = TEMP2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row
    
LC = TEMP2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

Cells(LR + 4, 2) = 1
Cells(LR + 4, 2).Interior.Color = RGB(255, 255, 0)
Cells(LR + 5, 2) = 2
Cells(LR + 5, 2).Interior.Color = RGB(146, 208, 80)
Cells(LR + 6, 2) = 3
Cells(LR + 6, 2).Interior.Color = RGB(0, 176, 240)

If WorksheetExists("OLAH2") Then Sheets("OLAH2").Delete
Set OLAH2 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "OLAH2"
Call PIVOT_1

TEMP2.Activate

'Cells(LR + 2, 3) = "INAC"
'Cells(LR + 2, 4) = "INFA"
'Cells(LR + 2, 5) = "ININ"
'Cells(LR + 2, 6) = "INPA"
'Cells(LR + 2, 7) = "TOTAL"

'LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
LR = Range("A" & Rows.Count).End(xlUp).Row

Range(Cells(LR, 1), Cells(LR + 7, LC)).Font.Name = "Verdana"

With Range(Cells(LR + 3, 3), Cells(LR + 3, LC))
    .Font.Name = "Century Gothic"
    .Font.Color = vbWhite
    .Font.Bold = True
    .Font.Size = 12
    .Interior.Color = RGB(52, 98, 101)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .RowHeight = .RowHeight + 4
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

'Range(Cells(LR + 3, 3), Cells(LR + 5, 7)).FormulaR1C1 = _
'    "=SUMIF(R3C2:R" & LR & "C2,RC2,R3C:R" & LR & "C)"
'Range(Cells(LR + 6, 3), Cells(LR + 6, 7)).FormulaR1C1 = _
'    "=SUM(R[-3]C:R[-1]C)"
'Range(Cells(LR + 3, 7), Cells(LR + 6, 7)).FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
    
With Range(Cells(LR + 3, 3), Cells(LR + 6, LC))
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

Set RG_KATEGORI = Range(Cells(1, 2), Cells(LR + 6, 2))
For Each cell In RG_KATEGORI
    If cell.Value = 1 Then
        cell.Value = "0 - 30"
    ElseIf cell.Value = 2 Then
        cell.Value = "31 - 90"
    ElseIf cell.Value = 3 Then
        cell.Value = ">90"
    End If
Next cell
Range("B:B").ColumnWidth = 13
Range(Cells(LR + 7, 3), Cells(LR + 7, LC)).Interior.Pattern = xlSolid
Range(Cells(LR + 7, 3), Cells(LR + 7, LC)).Interior.PatternColor = xlAutomatic
Range(Cells(LR + 7, 3), Cells(LR + 7, LC)).Interior.Color = RGB(228, 240, 241)
Range("I:Z").NumberFormat = "0.0%"

'Stop
'LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
LR = Range("A" & Rows.Count).End(xlUp).Row

Range("C1", Cells(2, LC)).Copy
Cells(1, LC + 2).PasteSpecial xlPasteAll: Application.CutCopyMode = False
Cells(1, LC + 2).Value = "Percentage"
LC2 = Range("SAM1").End(xlToLeft).Column

'.FormulaR1C1 = "=IFERROR(IF(RC2<>"""",RC[-7]/R59C[-7],""""),0)"
With Range(Cells(3, LC + 2), Cells(LR - 1, LC2 - 1))
    '.FormulaR1C1 = "=IFERROR(RC[-" & LC + 2 - 3 & "]/R" & LR & "C[-" & LC + 2 - 3 & "],0)"
    .FormulaR1C1 = "=IFERROR(IF(RC2<>"""",RC[-" & LC + 2 - 3 & "]/R" & LR & "C[-" & LC + 2 - 3 & "],""""),0)"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats
End With

With Range(Cells(3, LC2), Cells(LR - 1, LC2))
    .FormulaR1C1 = _
        "=IFERROR(IF(RC[-" & LC2 - 3 & "]<>"""",AVERAGE(RC[-" & LC2 - (LC + 2) & "]:RC[-1]),""""),"""")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats
End With

Application.CutCopyMode = False

Range(Cells(LR + 3, 3), Cells(LR + 3, LC)).Copy
Cells(LR + 3, LC + 2).PasteSpecial xlPasteAll
Application.CutCopyMode = False

LR2 = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
With Range(Cells(LR + 4, LC + 2), Cells(LR2, LC2))
    .NumberFormat = "0%"
    .FormulaR1C1 = "=IFERROR(RC[-" & LC + 2 - 3 & "]/R" & LR & "C[-" & LC + 2 - 3 & "],0)"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats
End With

Range(Cells(LR + 7, LC + 2), Cells(LR + 7, LC2)).Interior.Pattern = xlSolid
Range(Cells(LR + 7, LC + 2), Cells(LR + 7, LC2)).Interior.PatternColor = xlAutomatic
Range(Cells(LR + 7, LC + 2), Cells(LR + 7, LC2)).Interior.Color = RGB(228, 240, 241)
Application.CutCopyMode = False
Cells(1, 1).Select

'Stop
'Cells(1, LC + 2) = "Percentage"
'Cells(1, LC + 6) = "TOTAL"
'Cells(2, LC + 2) = "INAC"
'Cells(2, LC + 3) = "INFA"
'Cells(2, LC + 4) = "ININ"
'Cells(2, LC + 5) = "INPA"
'
'With Range(Cells(1, LC + 2), Cells(1, LC + 5))
'    .Merge
'    .HorizontalAlignment = xlCenter
'End With
'
'Range(Cells(1, LC + 6), Cells(2, LC + 6)).Merge
'Range(Cells(1, LC + 6), Cells(2, LC + 6)).HorizontalAlignment = xlCenter
'
'With Range(Cells(1, LC + 2), Cells(2, LC + 6))
'    .Font.Name = "Century Gothic"
'    .Font.Color = vbWhite
'    .Font.Bold = True
'    .Font.Size = 12
'    .Interior.Color = RGB(52, 98, 101)
'    .Interior.Pattern = xlSolid
'    .Interior.PatternColorIndex = xlAutomatic
'    .RowHeight = .RowHeight + 4
'    .Borders.LineStyle = xlContinuous
'    .Borders.Weight = xlThin
'End With
'
'With Range(Cells(3, LC + 2), Cells(LR, LC + 2))
'    .FormulaR1C1 = "=IFERROR(IF(RC[-7]<>"""",RC[-6]/R" & LR & "C3,""""),0)"
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'End With
'With Range(Cells(3, LC + 3), Cells(LR, LC + 3))
'    .FormulaR1C1 = "=IFERROR(IF(RC[-8]<>"""",RC[-6]/R" & LR & "C4,""""),0)"
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'End With
'With Range(Cells(3, LC + 4), Cells(LR, LC + 4))
'    .FormulaR1C1 = "=IFERROR(IF(RC[-9]<>"""",RC[-6]/R" & LR & "C5,""""),0)"
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'End With
'With Range(Cells(3, LC + 5), Cells(LR, LC + 5))
'    .FormulaR1C1 = "=IFERROR(IF(RC[-10]<>"""",RC[-6]/R" & LR & "C6,""""),0)"
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'End With
'With Range(Cells(3, LC + 6), Cells(LR, LC + 6))
'    .FormulaR1C1 = "=IFERROR(IF(RC[-11]<>"""",RC[-6]/R" & LR & "C7,""""),0)"
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'End With
'
'Cells(LR + 2, LC + 2) = "INAC"
'Cells(LR + 2, LC + 3) = "INFA"
'Cells(LR + 2, LC + 4) = "ININ"
'Cells(LR + 2, LC + 5) = "INPA"
'Cells(LR + 2, LC + 6) = "TOTAL"
'
'With Range(Cells(LR + 2, LC + 2), Cells(LR + 2, LC + 6))
'    .Font.Name = "Century Gothic"
'    .Font.Color = vbWhite
'    .Font.Bold = True
'    .Font.Size = 12
'    .Interior.Color = RGB(52, 98, 101)
'    .Interior.Pattern = xlSolid
'    .Interior.PatternColorIndex = xlAutomatic
'    .RowHeight = .RowHeight + 4
'    .Borders.LineStyle = xlContinuous
'    .Borders.Weight = xlThin
'End With
'
'Range(Cells(LR + 2, LC + 2), Cells(LR + 6, LC + 6)).NumberFormat = "0%"
'
'Range(Cells(LR + 3, LC + 2), Cells(LR + 5, LC + 6)).FormulaR1C1 = _
'    "=IFERROR(RC[-6]/R" & LR & "C[-6],0)"
'
'Range(Cells(LR + 6, LC + 2), Cells(LR + 6, LC + 6)).FormulaR1C1 = _
'    "=SUM(R[-3]C:R[-1]C)"
'
'Range(Cells(LR + 6, LC + 2), Cells(LR + 6, LC + 6)).Interior.Pattern = xlSolid
'Range(Cells(LR + 6, LC + 2), Cells(LR + 6, LC + 6)).Interior.PatternColor = xlAutomatic
'Range(Cells(LR + 6, LC + 2), Cells(LR + 6, LC + 6)).Interior.Color = RGB(228, 240, 241)
'
'With Range(Cells(LR + 2, LC + 2), Cells(LR + 6, LC + 6))
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'End With

Columns(LC + 1).ColumnWidth = 1

For Each col In Range(Cells(1, LC + 2), Cells(LR + 6, LC2)).Columns
    col.EntireColumn.AutoFit
    col.ColumnWidth = col.ColumnWidth + 5
Next col

Cells(1, 1).Select

Rows(1).Insert
Range("A:A").Insert
Range("A:A").ColumnWidth = 3

Rows(1).Insert
With Range(Cells(2, 2), Cells(2, LC + 1))
    .Interior.Pattern = xlSolid
    .Interior.PatternColor = xlAutomatic
    .Interior.Color = RGB(228, 240, 241)
    .Merge
    .HorizontalAlignment = xlLeft
    .Value = Periode
    .Font.Name = "Calibri"
    .Font.Bold = False
    .Font.Italic = True
    .Font.Size = 14
End With

Range("D5").Select
ActiveWindow.FreezePanes = True

Range("A1").Select
ActiveWindow.Zoom = 85

TEMP2.Name = "Analisa"

For Each WS In TWB.Worksheets
    If WS.Name <> "TOMBOL" And _
        WS.Name <> "DATA IM" And _
        WS.Name <> "DATA OV" And _
        WS.Name <> "Cari OV" And _
        WS.Name <> "Report WO Buyer" And _
        WS.Name <> "HASIL" And _
        WS.Name <> "Analisa" Then
    
        WS.Delete
        
    End If
Next WS

TWB.Sheets("TOMBOL").Select
Range("A1").Select

End Sub
