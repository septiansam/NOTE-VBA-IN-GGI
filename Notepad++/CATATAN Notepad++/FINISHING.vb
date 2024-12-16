Option Explicit
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'
'''''-----------------------------------------------------'''''

Dim TWB As Workbook, WB_TARIKAN As Workbook, WB_RESULTS As Workbook
Dim HOME As Worksheet, TARIKAN As Worksheet
Dim TEMP1 As Worksheet, TEMP2 As Worksheets, RESULTS As Worksheet
Dim COL_SEWING As Long, COL_DUEDAYS As Long
Dim RG_CEK As Range, CEK_VALUE As Long
Dim PATH_TARIKAN As String, PATH_RESULTS_EXCEL As String, PATH_RESULTS_PDF As String
Dim RNG As Range, RNG_BORDER As Range, CELL As Range
Dim LR_TARIKAN As Long, LC_TARIKAN As Long, LR As Long, LC As Long, FR As Long, FC As Long
Dim COL_REF As Long
Dim i As Long, j As Long, x As Long, COL_PASTE As Long
Dim RNG_PERIODE As Range, PERIODE_AWAL As Date, PERIODE_AKHIR As Date, TITLE As String
Dim BULAN_AWAL As String, BULAN_AKHIR As String
Dim RNG_ROW As Range, RNG_COLUMN As Range, RNG_RESULTS As Range

'''''-----------------------------------------------------'''''
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'

Sub PROSES()
    Application.DisplayAlerts = False
    
    Set TWB = ThisWorkbook
    Set HOME = TWB.Sheets("HOME")

    For i = TWB.Sheets.Count To 2 Step -1
        Sheets(i).Delete
    Next i
    
    PATH_TARIKAN = HOME.Range("E13") & Application.PathSeparator & HOME.Range("D13") & ".xlsx"
    If Dir(PATH_TARIKAN) = "" Then
        Call MsgBox("File " & HOME.Range("D13") & " Doesn't Exosts", vbCritical + vbOKOnly)
        Exit Sub
    End If

    Set TARIKAN = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "Tarikan GCC"
    
    Set TEMP1 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP1"

    Set RESULTS = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "RESULTS"

    Set WB_TARIKAN = Workbooks.Open(PATH_TARIKAN)
    WB_TARIKAN.Activate: Sheets(1).Select: Cells.Copy TARIKAN.Range("A1")
    WB_TARIKAN.Close False
    
    TARIKAN.Activate
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Range(Cells(1, 1), Cells(LR, LC)).Copy
    
    TEMP1.Activate
    Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    Cells.UnMerge
    Cells.Font.Name = "Verdana"
    Rows(1).Delete
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    COL_SEWING = Cells.Find(What:="Sewing.Factory", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Column
    COL_DUEDAYS = Rows("1:1").Find(What:="Due.Days", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Column
    Set RNG = Range(Cells(1, 1), Cells(LR, LC))

    TEMP1.Sort.SortFields.CLEAR
    TEMP1.Sort.SortFields.Add2 Key:=Range(Cells(2, COL_SEWING), Cells(LR, COL_SEWING)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    TEMP1.Sort.SortFields.Add2 Key:=Range(Cells(2, COL_DUEDAYS), Cells(LR, COL_DUEDAYS)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With TEMP1.Sort
        .SetRange RNG
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Set RNG = Range(Cells(1, 1), Cells(LR, LC))
    RNG.Copy
    RESULTS.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False

    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    With Range(Cells(1, 1), Cells(1, LC))
        .HorizontalAlignment = xlCenter
        .Font.Name = "Century Gothic"
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 13
        .Interior.Color = RGB(31, 59, 61)
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .RowHeight = .RowHeight + 10
    End With
    
    Cells.VerticalAlignment = xlCenter
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    
    Range(Cells(2, COL_DUEDAYS), Cells(LR, COL_DUEDAYS)).HorizontalAlignment = xlCenter
    
    For i = 2 To LR
        If i Mod 2 = 0 Then
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
        ElseIf i Mod 2 <> 0 Then
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
    Next i
    
    Set RG_CEK = Range(Cells(2, COL_DUEDAYS), Cells(LR, COL_DUEDAYS))
    For Each CELL In RG_CEK
        If CELL.Value <> "" And IsNumeric(CELL.Value) Then
            CEK_VALUE = CLng(CELL.Value)
            If CEK_VALUE < 0 Then
                i = CELL.Row
                With Range(Cells(i, 1), Cells(i, LC))
                    .Interior.Pattern = xlSolid
                    .Interior.PatternColor = xlAutomatic
                    .Interior.Color = vbYellow
                    CELL.Font.Color = vbRed
                End With
            End If
        End If
    Next CELL
    
    For i = LR To 2 Step -1
        If Cells(i, 1) <> Cells(i - 1, 1) Then
            Rows(i).Insert
            With Range(Cells(i, 1), Cells(i, LC))
                .Borders.LineStyle = xlNone
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(90, 164, 170)
                .RowHeight = 5
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeTop).Color = RGB(52, 98, 101)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeBottom).Color = RGB(52, 98, 101)
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).TintAndShade = 0
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeLeft).Color = RGB(52, 98, 101)
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeRight).Weight = xlThin
                .Borders(xlEdgeRight).Color = RGB(52, 98, 101)
            End With
        End If
    Next i
    
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Set RNG = Range(Cells(1, 1), Cells(LR, LC))
    For Each CELL In RNG.Columns
        CELL.ColumnWidth = CELL.ColumnWidth + 1
    Next CELL
    
    Rows("1:3").Insert
    With Range(Cells(2, 1), Cells(2, LC))
        .Merge
        .Value = "WO FINISHING"
        .Font.Name = "Century Gothic"
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .EntireRow.AutoFit
        .RowHeight = .RowHeight + 20
        .Interior.Color = RGB(79, 146, 151)
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
    End With
    
    With Range(Cells(3, 1), Cells(3, LC))
        .Merge
        .RowHeight = 25
        .Interior.Color = RGB(228, 240, 241)
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
    End With
    
    Range("A:A").Insert
    Range("A:A").ColumnWidth = 5
    
    Rows("5:5").Select
    ActiveWindow.FreezePanes = True
    Cells(1, 1).Select

    '[SAVE FILE EXCEL & PDF]....
    RESULTS.Activate
    
    PATH_RESULTS_EXCEL = HOME.Range("E14") & Application.PathSeparator & HOME.Range("D14") & ".xlsx"
    PATH_RESULTS_PDF = HOME.Range("E15") & Application.PathSeparator & HOME.Range("D15")

    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column


    Set RNG_RESULTS = Range(Cells(1, 1), Cells(LR, LC))
    
    '''[ SAVE TO EXCEL ]'''
    RESULTS.Copy
    Set WB_RESULTS = ActiveWorkbook
    Cells(1, 1).Select
    WB_RESULTS.SaveAs PATH_RESULTS_EXCEL, xlOpenXMLStrictWorkbook
    WB_RESULTS.Close True

    '''[ SAVE TO PDF }'''
    RESULTS.Activate
    With RESULTS.PageSetup
        .PrintArea = RNG_RESULTS.Address
'        .Orientation = xlLandscape
        .Orientation = xlPortrait
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With

    RESULTS.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PATH_RESULTS_PDF
    With RESULTS.Tab
        .Color = 15773696
        .TintAndShade = 0
    End With
    With TARIKAN.Tab
        .Color = 65280
        .TintAndShade = 0
    End With
    HOME.Activate
    Cells(1, 1).Select
    If WorksheetExists("TEMP1") Then Sheets("TEMP1").Delete

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

Sub ColorCheck()
    x = ActiveCell.Interior.ColorIndex
    Debug.Print x
End Sub


