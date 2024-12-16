'
'
'
'
'


Public TWB As Workbook
Public HOME As Worksheet
Public shTANGGAL As Worksheet, RgTanggal As Range, LRTanggal As Long
Public TEMP1 As Worksheet, TEMP2 As Worksheet, TEMP3 As Worksheet, TEMP4 As Worksheet
Public shRESUME As Worksheet, ROW_PASTE As Long
Public LR As Long, LC As Long, i As Long, j As Long, k As Long
Public Rng As Range, COUNT_LOOPING As Long, COUNT_TYPE As Long, STR_TYPE As String, STR_CODE As String, STR_STATUS As String
Public RgLine As Range, COUNT_LINE As Long
Public Rng_Data As Range

Public STR_WO_FACT As String, STR_WO As String, STR_FCT As String, STR_BUYER As String
Public QTY_ORDER As Long, VALUE_QTY As Long
Public COUNT_ROW As Long

Public EX_DATE As Date, Nomor As Long


Public PATH_EKSPOR As String, STR_EKSPOR As String, EKSPOR_PLAN As Worksheet, WB_EKSPOR_PLAN As Workbook
Public PATH_SEWING As String, STR_SEWING As String, UPLOAD_SEWING As Worksheet, WB_UPLOAD_SEWING As Workbook

Public RNG_COL As Range

Public WB_RESUME As Workbook, PATH_RESUME As String

Sub MAIN()
Application.DisplayAlerts = False

Set TWB = ThisWorkbook
Set HOME = TWB.Sheets("{HOME}")
For i = TWB.Sheets.Count To 2 Step -1
    Sheets(i).Delete
Next i

STR_EKSPOR = HOME.Range("H6")
PATH_EKSPOR = HOME.Range("I6") & Application.PathSeparator & STR_EKSPOR & ".xlsx"

STR_SEWING = HOME.Range("H7")
PATH_SEWING = HOME.Range("I7") & Application.PathSeparator & STR_SEWING & ".xlsx"

If Dir(PATH_EKSPOR) = "" Then
    Stop
End If

If Dir(PATH_SEWING) = "" Then
    Stop
End If

Set EKSPOR_PLAN = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "EKSPOR PLAN"
Set UPLOAD_SEWING = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "UPLOAD SEWING"
Set shTANGGAL = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TANGGAL"
Set TEMP1 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP1"
Set TEMP2 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP2"
Set TEMP3 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP3"
Set TEMP4 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP4"
Set shRESUME = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "RESUME"

shRESUME.Activate
Cells(1, 1) = "No"
Cells(1, 2) = "Buyer"
Cells(1, 3) = "Ex-Fact.Date"
Cells(1, 4) = "WO No"
Cells(1, 5) = "Line"
Cells(1, 6) = "Branch"
Cells(1, 7) = "Qty Order / Line"
Cells(1, 8) = "Ex Factory"
Cells(1, 9) = "Status Tarikan"

Set WB_EKSPOR_PLAN = Workbooks.Open(PATH_EKSPOR)
WB_EKSPOR_PLAN.Activate: Sheets(1).Select
Cells.Copy
EKSPOR_PLAN.Activate
EKSPOR_PLAN.Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
WB_EKSPOR_PLAN.Close (False)

Set WB_UPLOAD_SEWING = Workbooks.Open(PATH_SEWING)
WB_UPLOAD_SEWING.Activate: Sheets(1).Select
Cells.Copy
UPLOAD_SEWING.Activate
UPLOAD_SEWING.Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
WB_UPLOAD_SEWING.Close (False)

UPLOAD_SEWING.Activate
LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
Range("B:B").Copy
shTANGGAL.Activate
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
Range("A:A").RemoveDuplicates 1, xlYes

Range("A:A").Copy Range("B1")
Range("B:B").NumberFormat = "General"
LRTanggal = Range("B" & Rows.Count).End(xlUp).Row
Set RgTanggal = Range("B2:B" & LRTanggal)
Range("A2:A" & LRTanggal).Copy
shRESUME.Activate
Cells(1, 10).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
LC = Range("SAM1").End(xlToLeft).Column
Cells(1, LC + 1) = "Total Output"
Cells(1, LC + 2) = "Balance"
Cells.EntireColumn.AutoFit: Cells(1, 1).Select

EKSPOR_PLAN.Activate
If EKSPOR_PLAN.AutoFilterMode = True Then EKSPOR_PLAN.AutoFilterMode = False
LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
Range("A:A").Insert
Range("A1").FormulaR1C1 = "WO-FACTORY"
Range("A2:A" & LR).FormulaR1C1 = "=CONCATENATE(RC[19],""-"",RC[21])"
Range("A:A").Insert
Range("A1").FormulaR1C1 = "CON-WO-FACTORY"
Range("A2:A" & LR).FormulaR1C1 = "=COUNTIF(R1C2:RC[1],RC[1])"

With Range("A1:B" & LR)
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats
End With
Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select

Range("B:B").Copy Range("AA:AA")
Range("AB1").FormulaR1C1 = "BULAN"
With Range("AB2:AB" & LR)
    .FormulaR1C1 = "=IF(RC[-4]=""tarikan"",MONTH(RC[-19]),""xx"")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats
End With

Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
Range(Range("A1"), Range("A1").SpecialCells(xlLastCell)).AutoFilter 1, "1"
Range("B:B").Copy TEMP1.Range("A1")  '...WO-FACTORY
Range("U:U").Copy TEMP1.Range("B1")  '...WO No
Range("W:W").Copy TEMP1.Range("C1")  '...FACTORY
If EKSPOR_PLAN.AutoFilterMode = True Then EKSPOR_PLAN.AutoFilterMode = False


UPLOAD_SEWING.Activate
LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
Range("G:G").Insert
Range("G1").FormulaR1C1 = "WO-FACTORY"
Range("G2:G" & LR).FormulaR1C1 = "=CONCATENATE(RC[-1],""-"",RC[-3])"

Range("H:H").Insert
Range("H1").FormulaR1C1 = "CON-WO-FACTORY"
Range("H2:H" & LR).FormulaR1C1 = "=COUNTIF(R2C7:RC[-1],RC[-1])"

Range("I:I").Insert
Range("I1").FormulaR1C1 = "WO-FACTORY-LINE"
Range("I2:I" & LR).FormulaR1C1 = "=CONCATENATE(RC[-2],""-"",RC[-4])"

Range("A:A").Insert
Range("A1").FormulaR1C1 = "< EX"
'Range("A2:A" & LR).FormulaR1C1 = _
'    "=IF(COUNTIF('EKSPOR PLAN'!C[6],'UPLOAD SEWING'!RC[2])>0,""X"",""Y"")"

Range("AD1").FormulaR1C1 = "BULAN"
Range("AD2:AD" & LR).FormulaR1C1 = "=MONTH(RC[-27])"

Range("AE1").FormulaR1C1 = "CEK"
Range("AE2:AE" & LR).FormulaR1C1 = _
    "=IFERROR(VLOOKUP(RC[-24],'EKSPOR PLAN'!C[-10],1,0),""xx"")"
Range("AF1").FormulaR1C1 = "BULAN EKSPOR PLAN"
Range("AF2:AF" & LR).FormulaR1C1 = _
    "=IF(RC[-1]<>""xx"",VLOOKUP(RC[-24],'EKSPOR PLAN'!C[-5]:C[-4],2,0),""xx"")"

Range("A2:A" & LR).FormulaR1C1 = _
        "=IF(RC[31]<>""xx"",IF(RC[29]>RC[31],""X"",""Y""),""xx"")"
    
Range("A:A").Insert
Range("A1").FormulaR1C1 = "WO-FACTORY-LINE-< EX"
Range("A2:A" & LR).FormulaR1C1 = "=CONCATENATE(RC[10],""-"",RC[1])"

Range("A:A").Insert
Range("A1").FormulaR1C1 = "WO-FACTORY-LINE-< EX-Tanggal"
Range("A2:A" & LR).FormulaR1C1 = "=CONCATENATE(RC[1],""-"",RC[4])"

Range("L:L").Insert
Range("L1").FormulaR1C1 = "WO-FACTORY-CON"
Range("L2:L" & LR).FormulaR1C1 = "=CONCATENATE(RC[-2],""-"",RC[-1])"

With Range("A1:AZ" & LR)
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats
End With
Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select

If UPLOAD_SEWING.AutoFilterMode = True Then UPLOAD_SEWING.AutoFilterMode = False

TEMP1.Activate: Cells.EntireColumn.AutoFit
COUNT_LOOPING = Range("A" & Rows.Count).End(xlUp).Row - 1
Nomor = 0
For i = 1 To COUNT_LOOPING
'    If i = 1 Then
'        ROW_PASTE = 2
'    Else
'        shRESUME.Activate
'        ROW_PASTE = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
'    End If
    
    ' ??? PERTANYAAN WO - FACT = 181660-CVA
    
    Set Rng = Nothing
    Set RgLine = Nothing
    Set Rng_Data = Nothing
    
    TEMP2.Cells.Clear
    TEMP3.Cells.Clear
    TEMP4.Cells.Clear
    If EKSPOR_PLAN.AutoFilterMode = True Then EKSPOR_PLAN.AutoFilterMode = False
    If UPLOAD_SEWING.AutoFilterMode = True Then UPLOAD_SEWING.AutoFilterMode = False
    
    STR_WO_FACT = TEMP1.Range("A" & i + 1).Value
    STR_WO = TEMP1.Range("B" & i + 1).Value
    STR_FCT = TEMP1.Range("C" & i + 1).Value
    
    EKSPOR_PLAN.Activate
    Set Rng = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Rng.AutoFilter 2, STR_WO_FACT

    Rng.SpecialCells(xlCellTypeVisible).Copy
    TEMP2.Activate
    Range("A1").PasteSpecial xlPasteAll
    
    Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    If EKSPOR_PLAN.AutoFilterMode = True Then EKSPOR_PLAN.AutoFilterMode = False
    
    '[REMOVE DUPLICATE BERDASARKAN JENIS TARIKAN]'
    
'        Range(Range("A1"), Range("A1").SpecialCells(xlLastCell)).RemoveDuplicates Columns:=Array(2, 4, 5, 6, 7, 9, 16 _
'        , 17, 19, 20, 21, 22, 23, 24, 25, 26), Header:=xlYes
    
    Range(Range("A1"), Range("A1").SpecialCells(xlLastCell)).RemoveDuplicates Columns:=24, Header:=xlYes
    
    COUNT_TYPE = Application.WorksheetFunction.CountA(Range("B:B")) - 1
    If COUNT_TYPE > 1 Then
'        Stop
    End If
    '[DAPATKAN LINE]'
    UPLOAD_SEWING.Activate
    Range(Range("A1"), Range("A1").SpecialCells(xlLastCell)).AutoFilter 10, STR_WO_FACT
    Range(Range("A1"), Range("A1").SpecialCells(xlLastCell)).Copy
    
    TEMP3.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    If UPLOAD_SEWING.AutoFilterMode = True Then UPLOAD_SEWING.AutoFilterMode = False
    
    Range(Range("A1"), Range("A1").SpecialCells(xlLastCell)).RemoveDuplicates Columns:=8, Header:=xlYes
    
    '[URUTKAN LINE]
    TEMP3.Sort.SortFields.Clear
    TEMP3.Sort.SortFields.Add2 Key:=Range("H2:H" & Range("H" & Rows.Count).End(xlUp).Row), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With TEMP3.Sort
        .SetRange Range("A1").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Set RgLine = Range("H2:H" & Range("H1000").End(xlUp).Row)
    
    TEMP2.Activate
    For j = 1 To COUNT_TYPE
        Nomor = Nomor + 1
        TEMP4.Cells.Clear
        STR_BUYER = TEMP2.Range("D" & j + 1).Value
'            QTY_ORDER = TEMP2.Range("P" & j + 1).Value
        EX_DATE = TEMP2.Range("I" & j + 1).Value
        EX_DATE = Format(EX_DATE, "m/d/yyyy")
        STR_TYPE = TEMP2.Range("X" & j + 1).Value
        
        If STR_TYPE = "tarikan" Then
            STR_CODE = "Y"
            STR_STATUS = "Tarikan"
            QTY_ORDER = Application.WorksheetFunction. _
                SumIfs(EKSPOR_PLAN.Range("G:G"), _
                       EKSPOR_PLAN.Range("B:B"), STR_WO_FACT, _
                       EKSPOR_PLAN.Range("X:X"), STR_TYPE)
        Else
            STR_CODE = "X"
            STR_STATUS = ""
            QTY_ORDER = Application.WorksheetFunction. _
                SumIfs(EKSPOR_PLAN.Range("G:G"), _
                       EKSPOR_PLAN.Range("B:B"), STR_WO_FACT, _
                       EKSPOR_PLAN.Range("X:X"), STR_TYPE)
        End If
        
        TEMP4.Activate: Cells(1, 1).Select
        Range("A1") = "Buyer"
        Range("A2") = STR_BUYER
        
        Range("B1") = "Ex-Fact.Date"
        Range("B2") = EX_DATE
        
        Range("C1") = "WO No"
        Range("C2") = STR_WO
        
        Range("D1") = "LINE"
        Range("E1") = "Branch"
        Range("F1") = "QTY ORDER / LINE"
        Range("G1") = "EX FACTORY"
        Range("H1") = "status tarikan"
        Range("I1") = "LOOKUP"
        
        RgLine.Copy Range("D2")
        COUNT_ROW = Range("D1000").End(xlUp).Row
        
        VALUE_QTY = QTY_ORDER / (COUNT_ROW - 1)
        
'        If COUNT_TYPE = 1 Then
'            VALUE_QTY = QTY_ORDER
'        Else
'            VALUE_QTY = (QTY_ORDER / (COUNT_ROW - 1))
'        End If
        
        
        Range("E2:E" & COUNT_ROW).Value = STR_FCT
        Range("F2:F" & COUNT_ROW).Value = VALUE_QTY
        Range("G2:G" & COUNT_ROW).Value = EX_DATE
        Range("H2:H" & COUNT_ROW).Value = STR_STATUS
        
        Range("C" & COUNT_ROW + 1) = STR_WO & " Total"
        Range("F" & COUNT_ROW + 1).FormulaR1C1 = "=SUM(R2C:R[-1]C)"
        
        RgTanggal.Copy
        Range("J1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Application.CutCopyMode = False

        LC = Cells(1, Columns.Count).End(xlToLeft).Column
        Range("I2:I" & COUNT_ROW).FormulaR1C1 = "=""" & STR_WO_FACT & """&""-""&RC[-5]&""-""&""" & STR_CODE & """&""-"""
        
        For k = 1 To COUNT_ROW - 1
'                Range(Cells(1 + k, 10), Cells(1 + k, LC)).FormulaR1C1 = _
'                    "=IFERROR(VLOOKUP(R" & 1 + k & "C9&R1C,TEMP3!C1:C32,32,0),"""")"
            Range(Cells(1 + k, 10), Cells(1 + k, LC)).FormulaR1C1 = _
                "=SUMIF('UPLOAD SEWING'!C1,TEMP4!R" & 1 + k & "C9&TEMP4!R1C,'UPLOAD SEWING'!C32)"
                
            Range(Cells(1 + k, LC + 1), Cells(1 + k, LC + 1)).FormulaR1C1 = _
                "=SUBTOTAL(9,RC[-" & LC - 9 & "]:RC[-1])"
            Range(Cells(1 + k, LC + 2), Cells(1 + k, LC + 2)).FormulaR1C1 = _
                "=RC[-1]-RC[-" & LC - 4 & "]"
        Next k
        Range(Cells(COUNT_ROW + 1, LC + 1), Cells(COUNT_ROW + 1, LC + 2)).FormulaR1C1 = "=SUM(R2C:R[-1]C)"
        
        With Range(Cells(1, 1), Cells(COUNT_ROW + 1, LC + 2))
            .Copy
            .PasteSpecial xlPasteValuesAndNumberFormats
        End With
        
        Range("I:I").Delete
        
        LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        Range("A:A").Insert
        Range("A2:A" & LR).Value = Nomor
        
        LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
        Set Rng_Data = Range(Cells(2, 1), Cells(LR, LC))
        Rng_Data.Copy
        shRESUME.Activate
        ROW_PASTE = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
        Range("A" & ROW_PASTE).PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        Cells.EntireColumn.AutoFit
        Cells(1, 1).Select
    Next j
Next i

shRESUME.Activate

Call DesignResume

PATH_RESUME = HOME.Range("I8") & Application.PathSeparator & HOME.Range("H8") & ".xlsx"

Call SaveResume(shRESUME, PATH_RESUME)

HOME.Activate: Cells(1, 1).Select

If WorksheetExists("TANGGAL") Then Sheets("TANGGAL").Delete
If WorksheetExists("TEMP1") Then Sheets("TEMP1").Delete
If WorksheetExists("TEMP2") Then Sheets("TEMP2").Delete
If WorksheetExists("TEMP3") Then Sheets("TEMP3").Delete
If WorksheetExists("TEMP4") Then Sheets("TEMP4").Delete

Application.DisplayAlerts = True
End Sub

Sub DesignResume()

Set TWB = ThisWorkbook
Set shRESUME = TWB.Sheets("RESUME")

shRESUME.Activate
Cells.EntireColumn.AutoFit
Cells.EntireRow.AutoFit
ActiveWindow.Zoom = 90

LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

Set Rng_Data = Range(Cells(2, 10), Cells(LR, LC))
Rng_Data.NumberFormat = "#,##0;[Red]-#,##0;""-"";_(@_)"

For i = LC - 2 To 10 Step -1
    VALUE_QTY = Application.WorksheetFunction. _
                Sum(Range(Cells(2, i), Cells(LR, i)))
    If VALUE_QTY = 0 Then
        Columns(i).Delete
    End If
Next i

LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter
Cells.Font.Name = "Verdana"

With Range(Cells(1, 1), Cells(1, LC))
    .HorizontalAlignment = xlCenter
    .Font.Name = "Century Gothic"
    .Font.Color = vbWhite
    .Font.Bold = True
    .Font.Size = 13
    .Interior.Color = RGB(52, 98, 101)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .RowHeight = .RowHeight + 10
End With

Cells.EntireColumn.AutoFit: Cells(1, 1).Select
For i = 2 To LR
    If IsNumeric(Cells(i, 1).Value) Then
        If Cells(i, 1).Value Mod 2 = 0 Then
            With Range(Cells(i, 1), Cells(i, LC))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(255, 255, 255)
                .RowHeight = .RowHeight + 2
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(52, 98, 101)
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(52, 98, 101)
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
            End With
        ElseIf Cells(i, 1).Value Mod 2 <> 0 Then
            With Range(Cells(i, 1), Cells(i, LC))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(228, 240, 241)
                .RowHeight = .RowHeight + 2
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(52, 98, 101)
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(52, 98, 101)
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
            End With
        End If
    End If
Next i

LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

Set Range_Data = Range(Cells(1, 1), Cells(LR, LC))

For Each RNG_COL In Range_Data.Columns
    With RNG_COL
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = vbWhite
        .Borders(xlEdgeLeft).TintAndShade = 0
        .Borders(xlEdgeLeft).TintAndShade = 0
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = vbWhite
        .Borders(xlEdgeRight).TintAndShade = 0
        .Borders(xlEdgeRight).TintAndShade = 0
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = vbWhite
        .Borders(xlInsideVertical).TintAndShade = 0
        .Borders(xlInsideVertical).TintAndShade = 0
    End With
    RNG_COL.ColumnWidth = Range_Data.ColumnWidth + 1
Next RNG_COL

Range("A:A").Delete

Rows("1:2").Insert
Range("A:A").Insert: Range("A:A").ColumnWidth = 5
Rows(4).Select
ActiveWindow.FreezePanes = True

Cells(1, 1).Select

End Sub

Sub SaveResume(ByRef SheetResume As Worksheet, PathFile As String)
    SheetResume.Copy
    Set WB_RESUME = ActiveWorkbook
    Range(Range("B3"), Range("B3").SpecialCells(xlLastCell)).AutoFilter
    Range("A1").Select
    WB_RESUME.SaveAs PathFile, xlOpenXMLWorkbook
    WB_RESUME.Close (True)
End Sub


''''[ FUNGSI CEK SHEET ]''''
Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
        WorksheetExists = Not Sheets(shtName) Is Nothing
    On Error GoTo 0
End Function

Function wsx(sh_Name As Variant) As Boolean
    On Error Resume Next
        wsx = Not Sheets(sh_Name) Is Nothing
    On Error GoTo 0
End Function





