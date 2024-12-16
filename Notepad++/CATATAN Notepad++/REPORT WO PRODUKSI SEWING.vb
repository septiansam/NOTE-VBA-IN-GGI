'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'
'''''-----------------------------------------------------'''''

Dim TWB As Workbook, WB_TARIKAN As Workbook, WB_RESULTS As Workbook
Dim SH_HOME As Worksheet, SH_TARIKAN As Worksheet
Dim SH_TEMP1 As Worksheet, SH_TEMP2 As Worksheet, SH_TEMP3 As Worksheet, SH_TEMP4 As Worksheet, SH_RESULTS As Worksheet
Dim PATH_TARIKAN As String, PATH_RESULTS_EXCEL As String, PATH_RESULTS_PDF As String
Dim i As Long, COL_PASTE As Long

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

Dim ARR_COMPARISON As Variant
Dim STR_COMPARISON As String
Dim COL_COMPARISON As Long
Dim LR_DATA As Long, LC_DATA As Long
Dim IsFound As Boolean
Dim SUM_DATA_COMPARISON As Long
Dim ROW_HEIGHT As Long
Dim COL_WIDTH As Long
Dim RNG_RESULTS As Range
Dim COL As Range

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
    
    If WorksheetExists("Tarikan GCC") Then Sheets("Tarikan GCC").Delete
    If WorksheetExists("TEMP1") Then Sheets("TEMP1").Delete
    If WorksheetExists("TEMP2") Then Sheets("TEMP2").Delete
    If WorksheetExists("TEMP3") Then Sheets("TEMP3").Delete
    If WorksheetExists("TEMP4") Then Sheets("TEMP4").Delete
    If WorksheetExists("RESULTS") Then Sheets("RESULTS").Delete
    
    Set SH_TARIKAN = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "Tarikan GCC"
    Set SH_TEMP1 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP1"
    Set SH_TEMP2 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP2"
    Set SH_TEMP3 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP3"
    Set SH_TEMP4 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP4"
    Set SH_RESULTS = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "RESULTS": ActiveWindow.Zoom = 70
    
    Set WB_TARIKAN = Workbooks.Open(PATH_TARIKAN)
    WB_TARIKAN.Activate: Sheets(1).Select: Cells.Copy SH_TEMP1.Range("A1")
    WB_TARIKAN.Close False
    
    SH_TEMP1.Activate: Cells.Copy SH_TEMP2.Range("A1")
    
    SH_TEMP2.Activate
    Range("B:B,C:C,G:G").Delete Shift:=xlToLeft
    
    ARR_COMPARISON = Array("Worksheet.Release", _
                            "Trimcard.Release", _
                            "Sample.Release", _
                            "Pilot.Run", _
                            "Machine.Setting.Release", _
                            "Mika.Release", _
                            "Layout.Range.Release")
    

    
    For i = LBound(ARR_COMPARISON) To UBound(ARR_COMPARISON)
        SH_TEMP3.Cells.CLEAR
        SH_TEMP4.Cells.CLEAR
        
        STR_COMPARISON = ARR_COMPARISON(i)
        SH_TEMP2.Activate
        IsFound = Not IsEmpty(Rows(1).Find(STR_COMPARISON, , , xlPart))
        If IsFound = True Then
            COL_COMPARISON = Rows(1).Find(STR_COMPARISON, , , xlPart).Column
            SUM_DATA_COMPARISON = Application.WorksheetFunction.CountA(Columns(COL_COMPARISON))

            If SUM_DATA_COMPARISON <> 1 Then
            
                '''[ PROSES ]'''
                SH_TEMP2.Activate
                Range("A:D").Copy SH_TEMP3.Cells(1, 1)
                Columns(COL_COMPARISON).Copy SH_TEMP3.Cells(1, 5)
                Columns(COL_COMPARISON + 1).Copy SH_TEMP3.Cells(1, 6)
                Application.CutCopyMode = False
                
                SH_TEMP3.Activate
                
                LR_DATA = Cells(Rows.Count, 1).End(xlUp).Row
                
                Range("G2:G" & LR_DATA).FormulaR1C1 = "=TODAY()"
                Range("H2:H" & LR_DATA).FormulaR1C1 = "=RC[-4]-RC[-1]"
                With Range("G2:H" & LR_DATA)
                    .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
                End With
                Range("G:G").Delete Shift:=xlToLeft
                Range("G1") = "Diff Days"
                
                SH_TEMP3.Sort.SortFields.CLEAR
                SH_TEMP3.Sort.SortFields.Add2 Key:=Range("G2:G" & LR_DATA) _
                    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("TEMP3").Sort
                    .SetRange Range("A1:G" & LR_DATA)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
    
                Cells(1, 1).Select
                Cells.EntireColumn.AutoFit
                
                If SH_TEMP3.AutoFilterMode = True Then SH_TEMP3.AutoFilterMode = False
                Range("A1").AutoFilter Field:=5, Criteria1:="="
                
                If Range("A" & Rows.Count).End(xlUp).Value <> "No" Then
                    SH_TEMP3.UsedRange.Copy SH_TEMP4.Range("a1")
                    SH_TEMP4.Activate: Cells.EntireColumn.AutoFit
                    Range("E:E").Delete Shift:=xlToLeft: Cells(1, 1).Select
                    LR_DATA = Cells(Rows.Count, 1).End(xlUp).Row
                    Range("A2:A" & LR_DATA).CLEAR
                    
                    ''[ BUAT NOMOR ]''
                    Range("A2") = "1": Range("A2").DataSeries xlColumns, xlLinear, , 1, LR_DATA - 1
                    
                    Rows(1).Insert
                    With Range("A1")
                        .Value = "Plan Cutting Vs " & STR_COMPARISON
                        .Font.Bold = True
                        .Font.Name = "Calibri Light"
                        .Font.Size = 16
                    End With
                    
                    Range("A2:F2").Font.Bold = True
                    
                    Range("A1:F1").Merge
                    Range("A1:F1").Interior.ColorIndex = 2
                    Range("A2:F2").Interior.Color = RGB(198, 198, 198)
                    Rows(1).Insert: Range("A1:F1").Interior.Color = RGB(126, 126, 126)
                    Rows(3).Insert: Range("A3:F3").Interior.Color = RGB(126, 126, 126)
                    
                    LR_DATA = Cells(Rows.Count, 1).End(xlUp).Row
                    LC_DATA = SH_TEMP4.Cells(4, Columns.Count).End(xlToLeft).Column
                            
                    Set RNG_RESULTS = Range(Cells(1, 1), Cells(LR_DATA, LC_DATA))
    
                    With RNG_RESULTS
                        .Borders.LineStyle = xlContinuous
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Borders(xlEdgeLeft).Weight = xlMedium
                        .Borders(xlEdgeTop).Weight = xlMedium
                        .Borders(xlEdgeBottom).Weight = xlMedium
                        .Borders(xlEdgeRight).Weight = xlMedium
                    End With
                    With Range("A1:F1, A3:F3")
                        .Borders.LineStyle = xlNone
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).Weight = xlMedium
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).Weight = xlMedium
                    End With
                End If
                If SH_TEMP3.AutoFilterMode = True Then SH_TEMP3.AutoFilterMode = False
                
                SH_RESULTS.Activate
                COL_PASTE = SH_RESULTS.Cells(4, Columns.Count).End(xlToLeft).Column
                            
                If COL_PASTE <> 1 Then
                    Columns(COL_PASTE + 1).ColumnWidth = 10
                    COL_PASTE = COL_PASTE + 2
                End If
                
                RNG_RESULTS.Copy
                Cells(1, COL_PASTE).PasteSpecial xlPasteAll: Application.CutCopyMode = False
                
                Set RNG_RESULTS = Selection
                RNG_RESULTS.EntireColumn.AutoFit
                
                For Each COL In RNG_RESULTS.Columns
                    COL_WIDTH = COL.ColumnWidth
                    COL.ColumnWidth = COL_WIDTH + 3
                Next COL
                
                
                '''[ AKHIR PROSES ]'''
            End If
        End If

    Next i
    
    SH_RESULTS.Activate: Cells(1, 1).Select
    On Error Resume Next
    SH_RESULTS.Tab.Color = 15773696
    On Error GoTo 0
    
    Set RNG_RESULTS = SH_RESULTS.UsedRange
    
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
        .Value = "WO Production Sewing"
        .Font.Bold = True
        .Font.Name = "Calibri Light"
        .Font.Size = 30
    End With
    
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
    
    If WorksheetExists("Tarikan GCC") Then Sheets("Tarikan GCC").Delete
    If WorksheetExists("TEMP1") Then Sheets("TEMP1").Delete
    If WorksheetExists("TEMP2") Then Sheets("TEMP2").Delete
    If WorksheetExists("TEMP3") Then Sheets("TEMP3").Delete
    If WorksheetExists("TEMP4") Then Sheets("TEMP4").Delete

    '-------------------------------------'
    '_-_-_-_-_-[ SAVE RESULTS ]-_-_-_-_-_'
    '_-_-_-_-_-_[ EXCEL & PDF ]_-_-_-_-_-_'
    '-------------------------------------'
    
    PATH_RESULTS_EXCEL = SH_HOME.Range("E14") & Application.PathSeparator & SH_HOME.Range("D14") & ".xlsx"
    PATH_RESULTS_PDF = SH_HOME.Range("E15") & Application.PathSeparator & SH_HOME.Range("D15")
    
    Set RNG_RESULTS = SH_RESULTS.UsedRange
    
    '''[ SAVE TO EXCEL ]'''
    RNG_RESULTS.Select
    SH_RESULTS.Activate
    SH_RESULTS.Copy
    Set WB_RESULTS = ActiveWorkbook
    WB_RESULTS.Activate: Sheets(1).Select
    Rows("1:2").Insert xlDown
    Range("A:A").Insert xlRight
    ActiveWindow.Zoom = 70
    Cells(1, 1).Select
    WB_RESULTS.SaveAs PATH_RESULTS_EXCEL, xlOpenXMLStrictWorkbook
    WB_RESULTS.Close True
    
    '''[ SAVE TO PDF }'''
    SH_RESULTS.Activate
    With SH_RESULTS.PageSetup
        .PrintArea = RNG_RESULTS.Address
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    
    SH_RESULTS.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PATH_RESULTS_PDF
    
    SH_RESULTS.Activate
    Rows("1:2").Insert xlDown
    Range("A:A").Insert xlRight
    Cells(1, 1).Select
    
    
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
