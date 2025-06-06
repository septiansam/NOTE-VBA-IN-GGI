'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'
'''''-----------------------------------------------------'''''

Public TWB As Workbook, WB_TARIKAN As Workbook, WB_RESULTS As Workbook
Public SH_HOME As Worksheet, SH_TARIKAN As Worksheet
Public SH_TEMP1 As Worksheet, SH_TEMP2 As Worksheet, SH_TEMP3 As Worksheet, SH_TEMP4 As Worksheet, SH_RESULTS As Worksheet
Public PATH_TARIKAN As String, PATH_RESULTS_EXCEL As String, PATH_RESULTS_PDF As String
Public PATH_SUMMARY_EXCEL As String, PATH_SUMMARY_PDF As String
Public i As Long, COL_PASTE As Long

'''''-----------------------------------------------------'''''
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'

'''' [ FIELD YANG AKAN DI BANDINGKAN DENGAN PLAN CUTTING ] ''''

' (1). Worksheet.Release
' (2). Trimcard.Release
' (3). Sample.Release
' (4). Pilot.Run
' (5). Machine.Setting.Release
' (6). Mika.Release
' (7). Layout.Range.Release

Public ARR_COMPARISON As Variant
Public STR_COMPARISON As String
Public COL_COMPARISON As Long
Public LR_DATA As Long, LC_DATA As Long
Public IsFound As Boolean
Public SUM_DATA_COMPARISON As Long
Public ROW_HEIGHT As Long
Public COL_WIDTH As Long
Public RNG_RESULTS As Range
Public COL As Range

'*)... TAMBAHAN FEB-05-2024
''_ SPLIT FACTORY
Public SH_FCT As Worksheet
Public SH_SPLIT As Worksheet
Public Z As Long

Public LR_RESULTS As Long
Public LC_RESULTS As Long
Public ROW_PASTE_RESULT As Long, COL_PASTE_RESULT As Long
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'
'''''-----------------------------------------------------'''''

Sub PROSES()
    
    Application.DisplayAlerts = False
    
    Set TWB = ThisWorkbook
    Set SH_HOME = TWB.Sheets("HOME")

    For i = TWB.Sheets.Count To 2 Step -1
        Sheets(i).Delete
    Next i
    
    PATH_TARIKAN = SH_HOME.Range("E13") & Application.PathSeparator & SH_HOME.Range("D13") & ".xlsx"
    If Dir(PATH_TARIKAN) = "" Then
        Call MsgBox("File " & SH_HOME.Range("D13") & " Doesn't Exosts", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    If WorksheetExists("TARIKAN GCC") Then Sheets("TARIKAN GCC").Delete
    If WorksheetExists("TEMP1") Then Sheets("TEMP1").Delete
    If WorksheetExists("TEMP2") Then Sheets("TEMP2").Delete
    If WorksheetExists("TEMP3") Then Sheets("TEMP3").Delete
    If WorksheetExists("TEMP4") Then Sheets("TEMP4").Delete
    If WorksheetExists("RESULTS") Then Sheets("RESULTS").Delete
    
    Set SH_TARIKAN = Sheets.Add(After:=SH_HOME): ActiveSheet.Name = "TARIKAN GCC"
    Set SH_TEMP1 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP1"
    Set SH_TEMP2 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP2"
    Set SH_TEMP3 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP3"
    Set SH_TEMP4 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP4"
    Set SH_RESULTS = Sheets.Add(After:=SH_HOME): ActiveSheet.Name = "RESULTS": ActiveWindow.Zoom = 70
    
    Set WB_TARIKAN = Workbooks.Open(PATH_TARIKAN)
    WB_TARIKAN.Activate: Sheets(1).Select: Cells.Copy SH_TARIKAN.Range("A1")
    WB_TARIKAN.Close False
    
    Call PROSES_SUMMARY
    
    SH_TEMP1.Cells.CLEAR
    SH_TEMP2.Cells.CLEAR
    SH_TEMP3.Cells.CLEAR
    SH_TEMP4.Cells.CLEAR
    
    SH_TARIKAN.Activate: Cells.Copy SH_TEMP1.Range("A1")
    
    'SH_TEMP1.Activate: Cells.Copy SH_TEMP2.Range("A1")
    
    '*)... TAMBAHAN FEB-05-2024
    
    If WorksheetExists("FACTORY") Then Sheets("FACTORY").Delete
    If WorksheetExists("SPLIT FCT") Then Sheets("SPLIT FCT").Delete
    Set SH_SPLIT = Sheets.Add(After:=SH_HOME): ActiveSheet.Name = "SPLIT FCT"
    Set SH_FCT = Sheets.Add(After:=SH_SPLIT): ActiveSheet.Name = "FACTORY"

    SH_TEMP1.Activate
    LR_DATA = SH_TEMP1.UsedRange.Row + SH_TEMP1.UsedRange.Rows.Count - 1
    Range("B2:B" & LR_DATA).Copy SH_SPLIT.Range("A2")
    SH_SPLIT.Activate
    Range("A1").Value = "FCT"
    Range("A:A").RemoveDuplicates 1, xlYes
    LR_DATA = SH_SPLIT.UsedRange.Row + SH_SPLIT.UsedRange.Rows.Count - 1
    
    Z = 1
    
    '..[CLN,CJL,MJ1,MJ2]____________________________________________________________________________________
    If SH_SPLIT.AutoFilterMode = True Then SH_SPLIT.AutoFilterMode = False
    Range("A1").AutoFilter Field:=1, Criteria1:=Array("CLN", "CJL", "MJ1", "MJ2"), Operator:=xlFilterValues
    If Range("A" & Rows.Count).End(xlUp).Value <> "FCT" Then
        SH_SPLIT.UsedRange.Offset(1).Copy SH_FCT.Cells(1, Z)
        SH_SPLIT.UsedRange.Offset(1).Delete
        Z = Z + 1
    End If
    
    '..[KLB,CHW]____________________________________________________________________________________________
    If SH_SPLIT.AutoFilterMode = True Then SH_SPLIT.AutoFilterMode = False
    Range("A1").AutoFilter Field:=1, Criteria1:=Array("KLB", "CHW"), Operator:=xlFilterValues
    If Range("A" & Rows.Count).End(xlUp).Value <> "FCT" Then
        SH_SPLIT.UsedRange.Offset(1).Copy SH_FCT.Cells(1, Z)
        SH_SPLIT.UsedRange.Offset(1).Delete
        Z = Z + 1
    End If
    
    '..[CNJ GROUP]__________________________________________________________________________________________
    If SH_SPLIT.AutoFilterMode = True Then SH_SPLIT.AutoFilterMode = False
    SH_SPLIT.UsedRange.Offset(1).Copy SH_FCT.Cells(1, Z)
    
    Dim STR_SH_FCT As String, SH_HASIL As Worksheet
    Dim ARR_FCT() As Variant
    Dim SUM_SPLIT As Long, i_SPLIT As Integer
    Dim CELL As Range, DATA_RANGE As Range
    Dim NAME_SHEET As String
    
    SUM_SPLIT = SH_FCT.UsedRange.Columns.Count
    
    For i_SPLIT = 1 To SUM_SPLIT
        If i_SPLIT = 1 Then
            NAME_SHEET = "CLN, MJ1, MJ2"
        ElseIf i_SPLIT = 2 Then
            NAME_SHEET = "KLB, CHW"
        ElseIf i_SPLIT = 3 Then
            NAME_SHEET = "CBA, CNJ2, CVA, CVA2"
        End If
        
        SH_TEMP2.Cells.CLEAR
        SH_FCT.Activate
        Set DATA_RANGE = SH_FCT.Range(SH_FCT.Cells(1, i_SPLIT), _
                                        SH_FCT.Cells(SH_FCT.Cells(Rows.Count, i_SPLIT).End(xlUp).Row, i_SPLIT))
        DATA_RANGE.Select
        SH_FCT.Sort.SortFields.CLEAR
        SH_FCT.Sort.SortFields.Add2 Key:=DATA_RANGE _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortTextAsNumbers
        With SH_FCT.Sort
            .SetRange DATA_RANGE
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        STR_SH_FCT = ""
        For Each CELL In DATA_RANGE
            STR_SH_FCT = STR_SH_FCT & CELL.Value & ", "
        Next CELL
        STR_SH_FCT = Left(STR_SH_FCT, Len(STR_SH_FCT) - 2)
        
        ReDim ARR_FCT(1 To DATA_RANGE.Rows.Count)
        For i = 1 To DATA_RANGE.Rows.Count
            ARR_FCT(i) = SH_FCT.Cells(i, i_SPLIT).Value2
        Next i
        
        If WorksheetExists(NAME_SHEET) Then Sheets(NAME_SHEET).Delete
        Set SH_HASIL = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = NAME_SHEET
        
        'SH_TEMP1.Activate: Cells.Copy SH_TEMP2.Range("A1")
        SH_TEMP1.Activate
        If SH_TEMP1.AutoFilterMode = True Then SH_TEMP1.AutoFilterMode = False
        
        '.... CUTTING FACTORY ADA DI KOLOM B
        SH_TEMP1.UsedRange.AutoFilter Field:=2, Criteria1:=ARR_FCT, Operator:=xlFilterValues
        SH_TEMP1.UsedRange.Copy SH_TEMP2.Range("A1")
        SH_TEMP2.Activate
        Range("C:C,G:G").Delete Shift:=xlToLeft
        
        ARR_COMPARISON = Array("Worksheet.Release", _
                                "Trimcard.Release", _
                                "Sample.Release", _
                                "Pilot.Run", _
                                "Machine.Setting.Release", _
                                "Mika.Release", _
                                "Layout.Range.Release", _
                                "PPM")
        
        For i = LBound(ARR_COMPARISON) To UBound(ARR_COMPARISON)
            SH_TEMP3.Cells.CLEAR
            SH_TEMP4.Cells.CLEAR
            
            STR_COMPARISON = ARR_COMPARISON(i)
            If STR_COMPARISON = "" Then
            
            End If
            SH_TEMP2.Activate
            IsFound = Not IsEmpty(Rows(1).Find(STR_COMPARISON, , , xlPart))
            If IsFound = True Then
                COL_COMPARISON = Rows(1).Find(STR_COMPARISON, , , xlPart).Column
                SUM_DATA_COMPARISON = Application.WorksheetFunction.CountA(Columns(COL_COMPARISON))
    
                'If SUM_DATA_COMPARISON <> 1 Then
                
                '''[ PROSES ]'''
                SH_TEMP2.Activate
                Range("A:E").Copy SH_TEMP3.Cells(1, 1)
                Columns(COL_COMPARISON).Copy SH_TEMP3.Cells(1, 6)
                Columns(COL_COMPARISON + 1).Copy SH_TEMP3.Cells(1, 7)
                Application.CutCopyMode = False
                
                SH_TEMP3.Activate
                
                LR_DATA = Cells(Rows.Count, 1).End(xlUp).Row
                
                Range("H2:H" & LR_DATA).FormulaR1C1 = "=TODAY()"
                Range("I2:I" & LR_DATA).FormulaR1C1 = "=RC[-4]-RC[-1]"
                With Range("H2:I" & LR_DATA)
                    .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
                End With
                Range("H:H").Delete Shift:=xlToLeft
                Range("H1") = "Diff Days"
                
                SH_TEMP3.Sort.SortFields.CLEAR
                SH_TEMP3.Sort.SortFields.Add2 Key:=Range("H2:H" & LR_DATA) _
                    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("TEMP3").Sort
                    .SetRange Range("A1:H" & LR_DATA)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
    
                Cells(1, 1).Select
                Cells.EntireColumn.AutoFit
                
                If SH_TEMP3.AutoFilterMode = True Then SH_TEMP3.AutoFilterMode = False
                Range("A1").AutoFilter Field:=6, Criteria1:="="
                
                If Range("A" & Rows.Count).End(xlUp).Value <> "No" Then
                    SH_TEMP3.UsedRange.Copy SH_TEMP4.Range("a1")
                    SH_TEMP4.Activate: Cells.EntireColumn.AutoFit
                    Range("E:E").Delete Shift:=xlToLeft: Cells(1, 1).Select
                    LR_DATA = Cells(Rows.Count, 1).End(xlUp).Row
                    Range("A2:A" & LR_DATA).CLEAR
                    
                    ''[ BUAT NOMOR ]''
                    Range("A2") = "1": Range("A2").DataSeries xlColumns, xlLinear, , 1, LR_DATA - 1
                    Call Yellow_Highlight
                    
                    Rows(1).Insert
                    With Range("A1")
                        .Value = "Plan Cutting Vs " & STR_COMPARISON
                        .Font.Bold = True
                        .Font.Name = "Century Gothic"
                        .Font.Size = 16
                    End With
                    
                    Range("A2:G2").Font.Bold = True
                    Range("A2:Z9999").Font.Name = "Verdana"
                    Range("A1:G1").Merge
                    Range("A1:G1").Font.Color = vbWhite
                    Range("A2:G2").Font.Color = vbWhite
                    
                    
                    Range("A1:G1").Interior.Color = RGB(79, 146, 151)
                    Range("A2:G2").Interior.Color = RGB(52, 98, 101)
                    Rows(1).Insert: Range("A1:G1").Interior.Color = RGB(228, 240, 241)
                    Rows(3).Insert: Range("A3:G3").Interior.Color = RGB(228, 240, 241)
                    
                    LR_DATA = Cells(Rows.Count, 1).End(xlUp).Row
                    LC_DATA = SH_TEMP4.Cells(4, Columns.Count).End(xlToLeft).Column
                            
                    Set RNG_RESULTS = Range(Cells(1, 1), Cells(LR_DATA, LC_DATA))
'                    With RNG_RESULTS
'                        .Borders.LineStyle = xlContinuous
'                        .HorizontalAlignment = xlCenter
'                        .VerticalAlignment = xlCenter
'                        .Borders(xlEdgeLeft).Weight = xlMedium
'                        .Borders(xlEdgeTop).Weight = xlMedium
'                        .Borders(xlEdgeBottom).Weight = xlMedium
'                        .Borders(xlEdgeRight).Weight = xlMedium
'                    End With
'                    With Range("A1:G1, A3:G3")
'                        .Borders.LineStyle = xlNone
'                        .Borders(xlEdgeTop).LineStyle = xlContinuous
'                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
'                        .Borders(xlEdgeRight).LineStyle = xlContinuous
'                        .Borders(xlEdgeRight).Weight = xlMedium
'                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
'                        .Borders(xlEdgeLeft).Weight = xlMedium
'                    End With

                    Cells.Borders.LineStyle = xlNone
                    With Range("A2:G2, A4:G4")
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                    Dim RNG_FILL As Range
                    Dim RNG_ROW As Range
                    Set RNG_FILL = Range(Cells(5, 1), Cells(LR_DATA, LC_DATA))
                    
                    For Each RNG_ROW In RNG_FILL.Rows
                        If RNG_ROW.Row Mod 2 = 0 Then
                            RNG_ROW.Font.Name = "Verdana"
                            With RNG_ROW.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = RGB(255, 255, 255)
                            End With
                            RNG_ROW.HorizontalAlignment = xlCenter
                            RNG_ROW.VerticalAlignment = xlCenter
                        Else
                            RNG_ROW.Font.Name = "Verdana"
                            With RNG_ROW.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = RGB(228, 240, 241)
                            End With
                            RNG_ROW.HorizontalAlignment = xlCenter
                            RNG_ROW.VerticalAlignment = xlCenter
                        End If
                    Next RNG_ROW
                    Call Yellow_Highlight
                    SH_HASIL.Activate
                    COL_PASTE = SH_HASIL.Cells(4, Columns.Count).End(xlToLeft).Column
                    If COL_PASTE <> 1 Then
                        Columns(COL_PASTE + 1).ColumnWidth = 10
                        COL_PASTE = COL_PASTE + 2
                    End If
                    
                    RNG_RESULTS.Copy
                    Cells(1, COL_PASTE).PasteSpecial xlPasteAll: Application.CutCopyMode = False
                    
                    Set RNG_RESULTS = Selection
                    ActiveWindow.Zoom = 85
                    RNG_RESULTS.EntireColumn.AutoFit
                    
                    For Each COL In RNG_RESULTS.Columns
                        COL_WIDTH = COL.ColumnWidth
                        COL.ColumnWidth = COL_WIDTH + 3
                    Next COL
                End If
                If SH_TEMP3.AutoFilterMode = True Then SH_TEMP3.AutoFilterMode = False
                '''[ AKHIR PROSES ]'''
                    
            End If
    
        Next i
        SH_HASIL.Activate
    
        Set RNG_RESULTS = SH_HASIL.UsedRange
                    
        LR_DATA = RNG_RESULTS.Rows.Count + RNG_RESULTS.Row - 1
        LC_DATA = RNG_RESULTS.Columns.Count + RNG_RESULTS.Column - 1
        
        Rows(1).RowHeight = 3
        Rows(2).RowHeight = 30
        Rows(3).RowHeight = 3
        Rows(4).RowHeight = 20
        Rows("5:" & LR_DATA).RowHeight = 17
        
        '------------------------------'
        '_-_-_-_-_-[ HEADER ]-_-_-_-_-_'
        '------------------------------'
        Rows("1:2").Insert
        With Range("A1")
            .Value = "WO Production Sewing" & " - " & NAME_SHEET
            .Font.Bold = True
            .Font.Name = "Century Gothic"
            .Font.Size = 30
            .Font.Color = RGB(79, 146, 151)
        End With
        
        SH_HASIL.Activate
        With Range(Cells(1, 1), Cells(1, LC_DATA))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ROW_HEIGHT = Rows(1).RowHeight
        Rows(1).RowHeight = ROW_HEIGHT + 3
        Rows("2:2").Insert
        Rows("2:2").RowHeight = 15
        Cells(1, 1).Select
        
        '*)...[SET PAGE AREA]
        SH_HASIL.Activate
        Set RNG_RESULTS = SH_HASIL.UsedRange
        SH_HASIL.Activate
        With SH_HASIL.PageSetup
            .PrintArea = RNG_RESULTS.Address
            .Orientation = xlLandscape
            .CenterHorizontally = True
            .Zoom = False
            .FitToPagesTall = 1
            .FitToPagesWide = 1
        End With
        Rows("1:2").Insert xlDown
        Range("A:A").Insert xlRight
        Cells(1, 1).Select
        
        'On Error Resume Next
        Rows("10:10").Select
        ActiveWindow.FreezePanes = True
        'On Error GoTo 0
        Cells(1, 1).Select

    Next i_SPLIT
    
    
    '*)...[SET PAGE AREA SUMMARY]
    SH_RESULTS.Activate
    Set RNG_RESULTS = SH_RESULTS.UsedRange
    With SH_RESULTS.PageSetup
        .PrintArea = RNG_RESULTS.Address
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    Cells(1, 1).Select
    
    If WorksheetExists("SPLIT FCT") Then Sheets("SPLIT FCT").Delete
    If WorksheetExists("FACTORY") Then Sheets("FACTORY").Delete
    If WorksheetExists("TARIKAN GCC") Then Sheets("TARIKAN GCC").Delete
    If WorksheetExists("TEMP1") Then Sheets("TEMP1").Delete
    If WorksheetExists("TEMP2") Then Sheets("TEMP2").Delete
    If WorksheetExists("TEMP3") Then Sheets("TEMP3").Delete
    If WorksheetExists("TEMP4") Then Sheets("TEMP4").Delete

    '-------------------------------------'
    '_-_-_-_-_-[ SAVE RESULTS ]-_-_-_-_-_'
    '_-_-_-_-_-_[ EXCEL & PDF ]_-_-_-_-_-_'
    '-------------------------------------'
    Dim ARR_SH_SAVE() As Variant
    Dim SUM_SH_SAVE As Integer
    
    SUM_SH_SAVE = TWB.Sheets.Count - 1
    ReDim ARR_SH_SAVE(2 To SUM_SH_SAVE)
    For i = 2 To SUM_SH_SAVE
        ARR_SH_SAVE(i) = Sheets(i + 1).Name
    Next i
    
    PATH_RESULTS_EXCEL = SH_HOME.Range("E14") & Application.PathSeparator & SH_HOME.Range("D14") & ".xlsx"
    PATH_SUMMARY_EXCEL = SH_HOME.Range("E15") & Application.PathSeparator & SH_HOME.Range("D15") & ".xlsx"
    
    PATH_RESULTS_PDF = SH_HOME.Range("E16") & Application.PathSeparator & SH_HOME.Range("D16")
    PATH_SUMMARY_PDF = SH_HOME.Range("E17") & Application.PathSeparator & SH_HOME.Range("D17")
    
    '''[SAVE TO EXCEL RESULTS ALL}'''
    Sheets(ARR_SH_SAVE).Copy
    Set WB_RESULTS = ActiveWorkbook
    WB_RESULTS.Activate: Sheets(1).Select
    ActiveWindow.Zoom = 85
    Cells(1, 1).Select
    WB_RESULTS.SaveAs PATH_RESULTS_EXCEL, xlOpenXMLStrictWorkbook
    WB_RESULTS.Close True
    
    '''[SAVE TO EXCEL SUMMARY}'''
    SH_RESULTS.Copy
    Set WB_RESULTS = ActiveWorkbook
    WB_RESULTS.Activate: Sheets(1).Select: Sheets(1).Name = "SUMMARY"
    ActiveWindow.Zoom = 85
    Cells(1, 1).Select
    WB_RESULTS.SaveAs PATH_SUMMARY_EXCEL, xlOpenXMLStrictWorkbook
    WB_RESULTS.Close True
    
    '''[SAVE TO PDF SPLIT]'''
    Sheets(ARR_SH_SAVE).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        PATH_RESULTS_PDF _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
    
    SH_HOME.Activate
    Cells(1, 1).Select
    
    '''[SAVE TO PDF SUMMARY]'''
    SH_RESULTS.Select
    SH_RESULTS.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PATH_SUMMARY_PDF
    
    SH_HOME.Activate
    Cells(1, 1).Select
    
    TWB.Save
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

Sub PIVOT_1()

Dim TWB As Workbook
Dim SH_HASIL As Worksheet
Dim OLAH As Worksheet, OLAH2 As Worksheet
Dim TEMP1 As Worksheet, TEMP2 As Worksheet
Dim LR As Long, LC As Long, i As Long, cell As Range
Dim WS As Worksheet, RNG As Range

Dim PV_RANGE As Range
Dim PV_TABLE As PivotTable
Dim PV_CACHE As PivotCache

Application.DisplayAlerts = False

Set TWB = ThisWorkbook
Set OLAH = TWB.Sheets("OLAH")
Set OLAH2 = TWB.Sheets("OLAH2")
Set TEMP1 = TWB.Sheets("TEMP1")
Set TEMP2 = TWB.Sheets("TEMP2")

OLAH.Activate
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

Set PV_RANGE = OLAH.Range(Cells(1, 1), Cells(LR, LC))

Set PV_CACHE = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                    SourceData:=PV_RANGE)
Set PV_TABLE = PV_CACHE.CreatePivotTable _
                    (TableDestination:=OLAH2.Range("A1"), _
                    TableName:="PV_1")
OLAH2.Activate

'[INSERT FIELD FIELD NYA]...
With PV_TABLE.PivotFields("KATEGORI")
'    .Caption = "Kategori    (Days)"
    .Orientation = xlRowField
    .Position = 1
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

LR = OLAH2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = OLAH2.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

OLAH2.Range(Cells(1, 1), Cells(LR, LC)).Copy
OLAH2.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False

Columns("B:H").NumberFormat = "_(#,##0;_((#,##0);_(""-"";_(@_)"
Range("A:A").Delete
Rows(1).Delete
LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
Cells(1, LC).Value = "TOTAL"

Set RNG = Range(Cells(1, 1), Cells(LR, LC))

TEMP2.Activate
LR = Range("C" & Rows.Count).End(xlUp).Row

RNG.Copy
Range("C" & LR + 3).PasteSpecial xlPasteAll: Application.CutCopyMode = False

End Sub
