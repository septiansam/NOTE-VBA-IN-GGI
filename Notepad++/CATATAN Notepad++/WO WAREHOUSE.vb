'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'
'''''-----------------------------------------------------'''''

Dim TWB As Workbook, WB_TARIKAN As Workbook, WB_RESULTS As Workbook
Dim WB_OV As Workbook, OV As Worksheet
Dim WB_WO_SCHEDULE As Workbook, WO_SCHEDULE As Worksheet
Dim WB_WO_BUYER As Workbook, WO_BUYER As Worksheet
Dim WB_ITEM_MASTER_LIST As Workbook, ITEM_LIST As Worksheet
Dim wbResults As Workbook, pathResults As String

Dim HOME As Worksheet, TARIKAN As Worksheet
Dim TEMP1 As Worksheet, TEMP2 As Worksheet, TEMP3 As Worksheet, TEMP4 As Worksheet, SH_RESULTS As Worksheet
Dim pathOV As String, strOV As String
Dim OLAH_OV As Worksheet, CEK_OV As Worksheet
Dim firstCol As Long

Dim pathWO_SCH As String, strWO_SCH As String
Dim pathWO_BUYER As String, strWO_BUYER As String
Dim pathItemList As String, strItemList As String


Dim LR As Long, LC As Long
Dim rgFilter As Range
Dim rng As Range, col As Range

Public PATH_TARIKAN As String, PATH_RESULTS_EXCEL As String, PATH_RESULTS_PDF As String
Public PATH_SUMMARY_EXCEL As String, PATH_SUMMARY_PDF As String
Public i As Long, COL_PASTE As Long


'''''-----------------------------------------------------'''''
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'


Sub Main()

Application.DisplayAlerts = False

    Set TWB = ThisWorkbook
    Set HOME = TWB.Sheets("HOME")
    
    For i = TWB.Sheets.Count To 2 Step -1
        Sheets(i).Delete
    Next i
    
    strOV = HOME.Range("D13")
    pathOV = HOME.Range("E13") & Application.PathSeparator & strOV & ".xlsx"
    
    If Dir(pathOV) = "" Then
        MsgBox "File " & strOV & " Doesn't Exists", vbCritical, "File " & strOV & " Tidak Ditemukan"
        Exit Sub
    End If
    
    strWO_SCH = HOME.Range("D14")
    pathWO_SCH = HOME.Range("E14") & Application.PathSeparator & strWO_SCH & ".xlsx"
    
    If Dir(pathWO_SCH) = "" Then
        MsgBox "File " & strWO_SCH & " Doesn't Exists", vbCritical, "File " & strWO_SCH & " Tidak Ditemukan"
        Exit Sub
    End If
    
    strWO_BUYER = HOME.Range("D16")
    pathWO_BUYER = HOME.Range("E16") & Application.PathSeparator & strWO_BUYER & ".xlsx"
    
    If Dir(pathWO_BUYER) = "" Then
        MsgBox "File " & strWO_BUYER & " Doesn't Exists", vbCritical, "File " & strWO_BUYER & " Tidak Ditemukan"
        Exit Sub
    End If
    
    strItemList = HOME.Range("D15")
    pathItemList = HOME.Range("E15") & Application.PathSeparator & strItemList & ".csv"
    
    If Dir(pathItemList) = "" Then
        MsgBox "File " & strItemList & " Doesn't Exists", vbCritical, "File " & strItemList & " Tidak Ditemukan"
        Exit Sub
    End If
    
    pathResults = HOME.Range("E17") & Application.PathSeparator & _
                    HOME.Range("D17") & ".xlsx"

    
'[AMBIL FILE OV LEDGER]
'_______________________________________________________________________________.
    If WorksheetExists("OV") Then Sheets("OV").Delete
    Set OV = Sheets.Add(AFTER:=Sheets(Sheets.Count)): ActiveSheet.Name = "OV"
    
    Set WB_OV = Workbooks.Open(pathOV)
    WB_OV.Activate: Sheets(1).Select
    
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    Range(Cells(1, 1), Cells(LR, LC)).Copy
    OV.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    WB_OV.Close (False)
'-------------------------------------------------------------------------------
    
'[AMBIL FILE WO SCHEDULE]
'_______________________________________________________________________________.
    If WorksheetExists("WO SCHEDULE") Then Sheets("WO SCHEDULE").Delete
    Set WO_SCHEDULE = Sheets.Add(AFTER:=Sheets(Sheets.Count)): ActiveSheet.Name = "WO SCHEDULE"
    
    Set WB_WO_SCHEDULE = Workbooks.Open(pathWO_SCH)
    WB_WO_SCHEDULE.Activate: Sheets(1).Select
    
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    Range(Cells(1, 1), Cells(LR, LC)).Copy
    WO_SCHEDULE.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    WB_WO_SCHEDULE.Close (False)
'-------------------------------------------------------------------------------

'[AMBIL FILE WO BUYER]
'_______________________________________________________________________________.
    If WorksheetExists("WO BUYER") Then Sheets("WO BUYER").Delete
    Set WO_BUYER = Sheets.Add(AFTER:=Sheets(Sheets.Count)): ActiveSheet.Name = "WO BUYER"
    
    Set WB_WO_BUYER = Workbooks.Open(pathWO_BUYER)
    WB_WO_BUYER.Activate: Sheets(1).Select
    
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    Range(Cells(1, 1), Cells(LR, LC)).Copy
    WO_BUYER.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    WB_WO_BUYER.Close (False)
'-------------------------------------------------------------------------------

'[AMBIL FILE ITEM MASTER LIST]
'_______________________________________________________________________________.
    If WorksheetExists("ITEM LIST") Then Sheets("ITEM LIST").Delete
    Set ITEM_LIST = Sheets.Add(AFTER:=Sheets(Sheets.Count)): ActiveSheet.Name = "ITEM LIST"
    
    Set WB_ITEM_MASTER_LIST = Workbooks.Open(pathItemList)
    WB_ITEM_MASTER_LIST.Activate: Sheets(1).Select
    
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    Range(Cells(1, 1), Cells(LR, LC)).Copy
    ITEM_LIST.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    WB_ITEM_MASTER_LIST.Close (False)
'-------------------------------------------------------------------------------
    
'[PRE-PROCESSING]
'_______________________________________________________________________________.
    If WorksheetExists("OLAH OV") Then Sheets("OLAH OV").Delete
    If WorksheetExists("CEK OV") Then Sheets("CEK OV").Delete
    Set OLAH_OV = Sheets.Add(AFTER:=Sheets(Sheets.Count)): ActiveSheet.Name = "OLAH OV"
    Set CEK_OV = Sheets.Add(AFTER:=Sheets(Sheets.Count)): ActiveSheet.Name = "CEK OV"

    If WorksheetExists("TEMP1") Then Sheets("TEMP1").Delete
    If WorksheetExists("TEMP2") Then Sheets("TEMP2").Delete
    If WorksheetExists("TEMP3") Then Sheets("TEMP3").Delete
    Set TEMP1 = Sheets.Add(AFTER:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP1"
    Set TEMP2 = Sheets.Add(AFTER:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP2"
    Set TEMP3 = Sheets.Add(AFTER:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP3"
    
    WO_SCHEDULE.Activate
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Range("R1") = "XXX"
    
    Range("R2:R" & LR).FormulaR1C1 = _
        "=IFERROR(INDEX('WO BUYER'!C[-17],MATCH('WO SCHEDULE'!RC[-15],'WO BUYER'!C[-15],0)),""XXX"")"
    Range("S1") = "XXX-OR No"
    Range("S2:S" & LR).FormulaR1C1 = "=CONCATENATE(RC[-1],""-"",RC[-15])"
    With Range("R1:S" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
    Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    
    OV.Activate
    If OV.AutoFilterMode = True Then OV.AutoFilterMode = False
    Set rgFilter = Range("A6", Range("A6").SpecialCells(xlLastCell))
    
    rgFilter.AutoFilter Field:=2, Criteria1:="<>"
    rgFilter.AutoFilter Field:=10, Criteria1:="OV"
    
    Range("A1").Select
    
    Range("A1", Range("A1").SpecialCells(xlLastCell)).Copy
    OLAH_OV.Activate
    Range("A1").PasteSpecial (xlPasteAll)
    Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    OV.AutoFilterMode = False

    Range("C:G").Delete xlLeft
    Range("K:K").Delete xlLeft
    Range("N:N").Delete xlLeft
    Range("P:P").Delete xlLeft
    
    Columns("I:I").TextToColumns Destination:=Range("R1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
    Range("W6").FormulaR1C1 = "or infa/inin"
    Range("W7:W" & LR).FormulaR1C1 = _
        "=IF(OR(AND(RC[-20]=1201,RC[-6]=""infa""),AND(RC[-20]=1201,RC[-6]=""inin""),AND(RC[-20]=1204,RC[-6]=""infa""),AND(RC[-20]=1204,RC[-6]=""inin"")),RC[-2],RC[-5])"
    
    Range("X6").FormulaR1C1 = "cari or"
    Range("X7:X" & LR).FormulaR1C1 = _
        "=IF(AND(NOT(RC[-7]=""infa""),NOT(RC[-7]=""inin"")),RIGHT(RC[-15],8),RC[-1])"
    
    Range("Y6").FormulaR1C1 = "con 1"
    Range("Y7:Y" & LR).FormulaR1C1 = _
        "=IF(RC[-12]>0,CONCATENATE(RC[-24],""-"",RC[-22],""-"",RC[-16],""-"",RC[-12],""-"",RC[-9]),"""")"
    
    Range("Z6").FormulaR1C1 = "cf con 1"
    Range("Z7:Z" & LR).FormulaR1C1 = "=COUNTIF(R2C14:RC[-1],RC[-1])"
    
    Range("AA6").FormulaR1C1 = "con 2"
    Range("AA7:AA" & LR).FormulaR1C1 = _
        "=IF(RC[-14]<0,CONCATENATE(RC[-26],""-"",RC[-24],""-"",RC[-18],""-"",RC[-14]*-1,""-"",RC[-11]),"""")"
    
    Range("AB6").FormulaR1C1 = "cf con 2"
    Range("AB7:AB" & LR).FormulaR1C1 = "=COUNTIF(R2C16:RC[-1],RC[-1])"
    
    Range("AC6").FormulaR1C1 = "con 3"
    Range("AC7:AC" & LR).FormulaR1C1 = _
        "=IF(RC[-16]>0,CONCATENATE(RC[-4],""-"",RC[-3]),CONCATENATE(RC[-2],""-"",RC[-1]))"
    
    Range("M6:M" & LR).Copy
    TEMP1.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    Range("M6:M" & LR).Copy
    CEK_OV.Range("A6").PasteSpecial xlPasteValuesAndNumberFormats
    
    Range("AC6:AC" & LR).Copy
    TEMP1.Range("B1").PasteSpecial xlPasteValuesAndNumberFormats
    Range("AC6:AC" & LR).Copy
    CEK_OV.Range("B6").PasteSpecial xlPasteValuesAndNumberFormats
    
    Application.CutCopyMode = False
    
    TEMP1.Activate
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    Range("C1").FormulaR1C1 = "COUNTIF"
    With Range("C2:C" & LR)
        .FormulaR1C1 = "=COUNTIF(R1C2:RC[-1],RC[-1])"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
    Application.CutCopyMode = False
    
    TEMP1.AutoFilterMode = False
    Range("A1:C" & LR).AutoFilter 3, "1"
    Range("B1:B" & LR).SpecialCells(xlCellTypeVisible).Copy
    
    CEK_OV.Activate
    Range("C6").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit
    Cells(1, 1).Select
    
    TEMP1.Activate
    TEMP1.AutoFilterMode = False
    TEMP1.Cells.CLEAR: Cells(1, 1).Select
    
    CEK_OV.Activate
    
    Range("C6").Value = "Row Labels"
    Range("D6").Value = "Sum of QTY.PRIMARY"
    LR = Range("C" & Rows.Count).End(xlUp).Row
    
    With Range("D7:D" & LR)
        .FormulaR1C1 = "=SUMIF(C[-2],RC[-1],C[-3])"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
    With Range("E7:E" & LR)
        .FormulaR1C1 = "=IF(RC[-1]=0,""list out"",""list in"")"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
    
    Application.CutCopyMode = False
    Range("A:B").Delete xlLeft
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select

    OLAH_OV.Activate
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    Range("A6").Value = "ITEM NUMBER"
    
    Range("AD6").FormulaR1C1 = "vl (vlookup cek ov)"
    Range("AD7:AD" & LR).FormulaR1C1 = "=VLOOKUP(RC[-1],'CEK OV'!C[-29]:C[-27],3,0)"
    
    Range("AE6").FormulaR1C1 = "con 4"
    Range("AE7:AE" & LR).FormulaR1C1 = _
        "=IF(RC[-1]=""list in"",CONCATENATE(RC[-28],""-"",RC[-7]),"""")"
        
    Range("AF6").FormulaR1C1 = "start sewing (xlookup wo schedule)"
    Range("AF7:AF" & LR).NumberFormat = "M/D/YYYY"
    Range("AF7:AF" & LR).FormulaR1C1 = _
        "=IFERROR(INDEX('WO SCHEDULE'!C[-22],MATCH('OLAH OV'!RC[-1],'WO SCHEDULE'!C[-13],0)),""xx"")"
        
    Range("AG6").FormulaR1C1 = "count day"
    Range("AG7:AG" & LR).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-26],""xx"")"
    
    Range("AH6").FormulaR1C1 = "KATEGORI"
    Range("AH7:AH" & LR).FormulaR1C1 = _
        "=IF(RC[-1]=""xx"","""",IF(RC[-1]<=14,""0-2 week"",IF(RC[-7]<=28,""3-4 week"",IF(RC[-1]<=42,""5-6 week"",IF(RC[-1]<=56,""7-8 week"",IF(RC[-1]<=70,""9-10 week"",""> 10 week""))))))"
    
    Range("AI6").FormulaR1C1 = "x"
    Range("AI7:AI" & LR).FormulaR1C1 = _
        "=IF(RC[-1]=""0-2 week"",1,IF(RC[-1]=""3-4 week"",2,IF(RC[-1]=""5-6 week"",3,IF(RC[-1]=""7-8 week"",4,IF(RC[-1]=""9-10 week"",5,IF(RC[-1]=""> 10 week"",6,""""))))))"
    
    Range("AJ6").FormulaR1C1 = "search text"
    Range("AJ7:AJ" & LR).FormulaR1C1 = _
        "=IFERROR(INDEX('ITEM LIST'!C[-34],MATCH('OLAH OV'!RC[-35],'ITEM LIST'!C[-31],0)),""xxx"")"

    With Range(Cells(1, 1), Cells(LR, "AJ"))
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats
    End With
    Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    
    '___________________________________________
    '[KELOMPOKKAN BERDASARKAN KATEGORI TERTENTU]
    '```````````````````````````````````````````
    'LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Range("AK6").Value = "ITEM"
    Range("AL6").Value = "SORT"
    
    For i = 7 To LR
        On Error Resume Next
        If Cells(i, 17) Like "*INFA*" Then
            Cells(i, 37).Value = "FABRIC"
            Cells(i, 38).Value = "1"
        ElseIf Cells(i, 17) Like "*ININ*" Then
            Cells(i, 37).Value = "FABRIC"
            Cells(i, 38).Value = "1"
        ElseIf Cells(i, 36) Like "*ELASTIC*" Then
            Cells(i, 37).Value = "ELASTIC"
            Cells(i, 38).Value = "2"
        ElseIf Cells(i, 36) Like "*CARTON*" Then
            Cells(i, 37).Value = "CARTON"
            Cells(i, 38).Value = "4"
        ElseIf Cells(i, 36) Like "*BOX*" And _
                Not Cells(i, 36) Like "*CARTON*" Then
            Cells(i, 37).Value = "BOX"
            Cells(i, 38).Value = "5"
        ElseIf Cells(i, 36) Like "*THREAD*" Then
            Cells(i, 37).Value = "THREAD"
            Cells(i, 38).Value = "3"
        Else
            Cells(i, 37).Value = "OTHER"
            Cells(i, 38).Value = "6"
        End If
        On Error GoTo 0
    Next i
    
'[LAST PRE-PROCESSING]
'-------------------------------------------------------------------------------
    
'[GET DATA FROM PIVOT]
'_______________________________________________________________________________.
    OLAH_OV.Activate
    If OLAH_OV.AutoFilterMode = True Then OLAH_OV.AutoFilterMode = False
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Range(Cells(6, 1), Cells(LR, LC)).AutoFilter 34, "<>"
    
    'ITEM
    Range("A:A").Copy TEMP1.Range("A1")
    
    'BU
    Range("C:C").Copy TEMP1.Range("B1")
    
    'EXT
    Range("N:N").Copy TEMP1.Range("C1")
    
    'G/L CAT
    Range("Q:Q").Copy TEMP1.Range("D1")
    
    'KATEGORI
    Range("AH:AH").Copy TEMP1.Range("E1")
    
    'x
    Range("AI:AI").Copy TEMP1.Range("F1")
    
    'ITEM
    Range("AK:AK").Copy TEMP1.Range("G1")
    
    'SORT
    Range("AL:AL").Copy TEMP1.Range("H1")
    
    
    If OLAH_OV.AutoFilterMode = True Then OLAH_OV.AutoFilterMode = False
    
'-------------------------------------------------------------------------------

    TEMP1.Activate
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column


'[BUAT PIVOT]
'-------------------------------------------------------------------------------

    Dim pvRng As Range
    Dim pvTbl As PivotTable
    Dim pvChc As PivotCache
    
    Set pvRng = Range(Cells(6, 1), Cells(LR, LC))
    Set pvChc = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pvRng)
    Set pvTbl = pvChc.CreatePivotTable _
                        (TableDestination:=TEMP2.Range("A1"), _
                        TableName:="pvReport")
                        
    '[BARIS]
    TEMP2.Activate
    With pvTbl.PivotFields("BUSINESS")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pvTbl.PivotFields("SORT")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pvTbl.PivotFields("ITEM")
        .Orientation = xlRowField
        .Position = 3
    End With
    With pvTbl.PivotFields("x")
        .Orientation = xlRowField
        .Position = 4
    End With
    With pvTbl.PivotFields("KATEGORI")
        .Orientation = xlRowField
        .Position = 5
    End With

    
    '[KOLOM]
    With pvTbl.PivotFields("G/L Cat")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    '[ISI]
    With pvTbl.PivotFields("EXTENDED AMT")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End With
    
    With pvTbl
        .DisplayErrorString = False
        .RepeatAllLabels xlRepeatLabels
        .PivotFields("SORT").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("ITEM NUMBER").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("BUSINESS").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("EXTENDED AMT").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("G/L Cat").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("KATEGORI").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("x").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("ITEM").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    
        .RowAxisLayout xlTabularRow
    End With
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    Range(Cells(1, 1), Cells(LR, LC)).Copy
    
    
'[CREATE TEMPLATE AND PERCENTAGE]
'-------------------------------------------------------------------------------

    TEMP3.Activate
    ActiveWindow.Zoom = 90
    
    Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False: Cells(1, 1).Select
    Range("B:B").Delete
    Range("C:C").Delete
    
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Cells.Font.Name = "Verdana"
    Cells.VerticalAlignment = xlCenter
    Cells.HorizontalAlignment = xlCenter
    Range("B3:C" & LR - 1).HorizontalAlignment = xlLeft
    
    Rows(1).ClearContents
    Range("A1:A2").Merge
    Range("A1:A2").Value = "BUSINESS"
    Range("B1:B2").Merge
    Range("B1:B2").Value = "ITEM"
    Range("C1:C2").Merge
    Range("C1:C2").Value = "KATEGORI"

    Range(Cells(1, 4), Cells(1, LC - 1)).Merge
    Range(Cells(1, 4), Cells(1, LC - 1)).Value = "G/L Cat"
    Range(Cells(1, LC), Cells(2, LC)).Merge
    Range(Cells(1, LC), Cells(2, LC)).Value = "Grand Total"
    
    Range(Cells(LR, 1), Cells(LR, 3)).Merge
    With Range(Cells(1, 1), Cells(LR, LC))
        .Borders.LineStyle = xlContinuous
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
    End With
    
    'Judul
    With Range(Cells(1, 1), Cells(2, LC))
        .Font.Name = "Century Gothic"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Pattern = xlSolid
        .Interior.PatternColor = xlAutomatic
        .Interior.Color = RGB(44, 83, 86)
        .RowHeight = .RowHeight + 2
    End With
    
    'Grand Total
    With Range(Cells(LR, 1), Cells(LR, LC))
        .Font.Name = "Century Gothic"
        .Font.Bold = True
        .Font.Color = vbWhite
        .Font.Size = 12
        .Interior.Pattern = xlSolid
        .Interior.PatternColor = xlAutomatic
        .Interior.Color = RGB(44, 83, 86)
        .RowHeight = .RowHeight + 2
    End With
    
    Cells.EntireColumn.AutoFit

    For i = 3 To LR - 1
        If i Mod 2 = 0 Then
            With Range(Cells(i, 1), Cells(i, LC))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(255, 255, 255)
                .RowHeight = .RowHeight + 2
            End With
        End If
        If i Mod 2 <> 0 Then
            With Range(Cells(i, 1), Cells(i, LC))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(228, 240, 241)
                .RowHeight = .RowHeight + 2
            End With
        End If
    Next i
    
    For i = LR - 1 To 4 Step -1
        If Cells(i, 1) <> Cells(i - 1, 1) Then
            Rows(i).Insert
            With Range(Cells(i, 1), Cells(i, LC))
                .Borders.LineStyle = xlNone
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(90, 164, 170)
                .RowHeight = 7
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).TintAndShade = 0
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeRight).Weight = xlThin
            End With
        End If
    Next i
    Rows(3).Insert
    With Range(Cells(3, 1), Cells(i, LC))
        .Borders.LineStyle = xlNone
        .Interior.Pattern = xlSolid
        .Interior.PatternColor = xlAutomatic
        .Interior.Color = RGB(90, 164, 170)
        .RowHeight = 7
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).TintAndShade = 0
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).TintAndShade = 0
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).TintAndShade = 0
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    
    
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Cells(1, 1).Select
    
    Range(Cells(1, 1), Cells(LR, LC)).Copy
    Cells(1, LC + 2).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    firstCol = LC + 5
    
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Range(Cells(1, LC), Cells(2, LC)).Value = "%"
    Range(Cells(3, firstCol), Cells(LR, LC)).ClearContents
    Range(Cells(3, firstCol), Cells(LR, LC)).NumberFormat = "0.00%"
    
    With Range(Cells(3, firstCol), Cells(LR - 1, LC))
        .FormulaR1C1 = _
            "=IF(RC[-" & (firstCol - 4) & "]<>"""",RC[-" & (firstCol - 4) & "]/R" & LR & "C[-" & (firstCol - 4) & "],"""")"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
    With Range(Cells(LR, firstCol), Cells(LR, LC))
        .FormulaR1C1 = _
            "=IFERROR(IF(RC[-" & (firstCol - 4) & "]<>"""",R" & LR & "C[-" & (firstCol - 4) & "]/R" & LR & "C[-" & (firstCol - 4) & "],""""),"""")"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
    Cells.EntireColumn.AutoFit
    
    Set rng = Range(Cells(1, 1), Cells(LR, LC))
    For Each col In rng.Columns
        col.EntireColumn.AutoFit
        col.ColumnWidth = col.ColumnWidth + 2
    Next col
    
    Rows("1:2").Insert
    Range("A:A").Insert: Range("A:A").ColumnWidth = 5
    Range("B5").Activate
    ActiveWindow.FreezePanes = True
    
    Cells(1, 1).Select
    
'[SAVE RESULTS]
'-------------------------------------------------------------------------------
    '''[ SAVE TO EXCEL ]'''
    
    TEMP3.Name = "RESULTS"
    TEMP3.Copy
    Set wbResults = ActiveWorkbook
    wbResults.SaveAs pathResults, xlOpenXMLWorkbook
    wbResults.Close (True)
    
    '''[ SAVE TO PDF }'''
    TEMP3.Activate
    
    PATH_RESULTS_PDF = HOME.Range("E18") & Application.PathSeparator & HOME.Range("D18")

    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Set RNG_RESULTS = Range(Cells(1, 1), Cells(LR, LC))

    '''[ SAVE TO PDF }'''
    TEMP3.Activate
    With TEMP3.PageSetup
        .PrintArea = RNG_RESULTS.Address
'        .Orientation = xlLandscape
        .Orientation = xlPortrait
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With

    TEMP3.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PATH_RESULTS_PDF


'[END]
'-------------------------------------------------------------------------------
    TWB.Activate
    For i = Sheets.Count - 1 To 2 Step -1
        If Sheets(i).Name <> "OLAH OV" Then
            Sheets(i).Delete
        End If
    Next i
    
'    If WorksheetExists("OV") Then Sheets("OV").Delete
'    If WorksheetExists("WO SCHEDULE") Then Sheets("WO SCHEDULE").Delete
'    If WorksheetExists("ITEM LIST") Then Sheets("ITEM LIST").Delete
'    If WorksheetExists("OLAH OV") Then Sheets("OLAH OV").Delete
'    If WorksheetExists("CEK OV") Then Sheets("CEK OV").Delete
'    If WorksheetExists("TEMP1") Then Sheets("TEMP1").Delete
'    If WorksheetExists("TEMP2") Then Sheets("TEMP2").Delete
    
    HOME.Activate
    Cells(1, 1).Select


Application.DisplayAlerts = True

End Sub

Sub IsiTanggal()

Dim TWB As Workbook, HOME As Worksheet
Dim tanggalSebelumnya As Date
Dim tanggalHariIni As Date
Dim tanggalKemarin As Date
Dim shTanggal As Worksheet

Application.DisplayAlerts = False

    Set TWB = ThisWorkbook
    Set HOME = TWB.Sheets("HOME")
    
    For i = Sheets.Count To 2 Step -1
        Sheets(i).Delete
    Next i
    
    If WorksheetExists("PERIODE") Then Sheets("PERIODE").Delete
    Set shTanggal = Sheets.Add(AFTER:=HOME): ActiveSheet.Name = "PERIODE"
    
    ' Tentukan tanggal hari ini
    tanggalHariIni = Date
    tanggalKemarin = DateAdd("d", -1, tanggalHariIni)
    
    ' Hitung tanggal satu bulan sebelumnya (selalu tanggal 1)
    If Month(tanggalHariIni) = 1 Then
        ' Jika bulan saat ini adalah Januari, maka tahun diubah menjadi tahun sebelumnya
        tanggalSebelumnya = DateSerial(Year(tanggalHariIni) - 1, 12, 1)
    Else
        ' Jika bulan saat ini bukan Januari, maka tanggal 1 dari bulan sebelumnya
        tanggalSebelumnya = DateSerial(Year(tanggalHariIni), Month(tanggalHariIni) - 1, 1)
    End If
    
    ' Isi data ke dalam sel A2 dan B2
    shTanggal.Activate
    Range("A:B").NumberFormat = "m/d/yyyy"
    Range("A1") = "TANGGAL AWAL"
    Range("B1") = "TANGGAL AKHIR"
    Range("A2").Value = tanggalSebelumnya
    Range("B2").Value = tanggalKemarin
    Cells.EntireColumn.AutoFit
    HOME.Activate: Cells(1, 1).Select
    
Application.DisplayAlerts = True
End Sub

Sub CLEAR()
    Application.DisplayAlerts = False
    Set TWB = ThisWorkbook
    For i = TWB.Sheets.Count To 2 Step -1
        Sheets(i).Delete
    Next i
    Application.DisplayAlerts = True
End Sub

''''[ FUNGSI CEK SHEET ]''''
Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
        WorksheetExists = Not Sheets(shtName) Is Nothing
    On Error GoTo 0
End Function





