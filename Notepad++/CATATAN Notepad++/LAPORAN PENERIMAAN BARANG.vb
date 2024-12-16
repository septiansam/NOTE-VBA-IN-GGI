Option Explicit
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'
'''''-----------------------------------------------------'''''

Dim TWB As Workbook, WB_TARIKAN As Workbook, WB_RESULTS As Workbook
Dim SH_HOME As Worksheet, SH_TARIKAN As Worksheet, SH_HELP As Worksheet
Dim SH_TEMP1 As Worksheet, SH_TEMP2 As Worksheet, SH_RESULTS As Worksheet
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
    Set SH_HOME = TWB.Sheets("HOME")

    For i = TWB.Sheets.Count To 2 Step -1
        Sheets(i).Delete
    Next i
    
    PATH_TARIKAN = SH_HOME.Range("E13") & Application.PathSeparator & SH_HOME.Range("D13") & ".xlsx"
    If Dir(PATH_TARIKAN) = "" Then
        Call MsgBox("File " & SH_HOME.Range("D13") & " Doesn't Exosts", vbCritical + vbOKOnly)
        Exit Sub
    End If

    Set SH_TARIKAN = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "Tarikan GCC"
    Set SH_TEMP1 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP1"
    Set SH_TEMP2 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP2"
    
    Set WB_TARIKAN = Workbooks.Open(PATH_TARIKAN)
    WB_TARIKAN.Activate: Sheets(1).Select: Cells.Copy SH_TARIKAN.Range("A1")
    WB_TARIKAN.Close False
    
    SH_TARIKAN.Activate
    SH_TARIKAN.UsedRange.Copy SH_TEMP1.Range("a1")
    SH_TEMP1.Activate
    
    With SH_TEMP1.UsedRange
        LR = .Row + .Rows.Count - 1
        LC = .Column + .Columns.Count - 1
    End With
    
    Set RNG = SH_TEMP1.UsedRange.Find(What:="Cummulative Percentage", LookAt:=xlPart)
    If Not RNG Is Nothing Then
        FR = RNG.Row
        FR = FR + 1
        COL_REF = RNG.Column
    Else
        Stop
    End If
    
    ' ISI SEL KOSONG
    ' SUPPLIER
    Range(Cells(FR, 1), Cells(LR, 1)).SpecialCells(xlCellTypeBlanks) _
        .FormulaR1C1 = "=IF(RC[5]<>"""",R[-1]C,"""")"
    
    ' ORDER TYPE
    Range("C9:C" & LR).SpecialCells(xlCellTypeBlanks) _
        .FormulaR1C1 = "=R[-1]C"
        
    ' PO NUMBER
    Range(Cells(FR, 5), Cells(LR, 5)).SpecialCells(xlCellTypeBlanks) _
        .FormulaR1C1 = "=IF(RC[1]<>"""",R[-1]C,"""")"
        
    For Each CELL In Range("B7:B" & LR)
        If CELL.Offset(0, -1).Value = "Business Unit" Then
            CELL.Value = CELL.Offset(0, 1)
        End If
    Next CELL
    
    Range("B7:B" & LR).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    
    SH_TEMP1.UsedRange.Copy
    Range("A1").PasteSpecial xlPasteValues: Application.CutCopyMode = False
    
    If SH_TEMP1.AutoFilterMode = True Then SH_TEMP1.AutoFilterMode = False
    Rows("1:9").Delete
    With SH_TEMP1.UsedRange
        LR = .Row + .Rows.Count - 1
        LC = .Column + .Columns.Count - 1
    End With
    Range(Cells(1, 1), Cells(LR, LC)).AutoFilter Field:=COL_REF, Criteria1:="="
    If Range("a" & Rows.Count).End(xlUp).Value <> "Supplier" Then
        SH_TEMP1.UsedRange.Offset(1).Delete
    End If
    Cells(1, 1).Select
    If SH_TEMP1.AutoFilterMode = True Then SH_TEMP1.AutoFilterMode = False
    
    With SH_TEMP1.UsedRange
        LR = .Row + .Rows.Count - 1
        LC = .Column + .Columns.Count - 1
    End With
    
    Range("T:T").Insert
    Range("T2:T" & LR).FormulaR1C1 = "=RC[-1]-RC[-2]"
    
    Range("B1") = "Branch": Range("C1") = "Order Type": Range("T1") = "Balance": Range("V1") = "Persentase PO/Kedatangan"
    Range("V2:V" & LR).FormulaR1C1 = "=RC[-1]&RC[1]"
    SH_TEMP1.UsedRange.Copy
    Range("A1").PasteSpecial xlPasteValues: Application.CutCopyMode = False

    Range("D:D,J:J,L:M,O:P,U:U,W:W").Delete
    Range("SAM1").Copy
    Range("O2:O" & LR).PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A:A").Insert: Cells(1, 1) = "Uniq-Value"
    Range("A2:A" & LR).FormulaR1C1 = _
        "=RC[1] & ""-"" & RC[2] & ""-"" & RC[3] & ""-"" & RC[4] & ""-"" & RC[5] & ""-"" & RC[6] & ""-"" & RC[7] & ""-"" & RC[8] & ""-"" & RC[9] & ""-"" & RC[10] & ""-"" & RC[11] & ""-"" & RC[12] & ""-"" & RC[13] & ""-"" & RC[14] & ""-"" & RC[15]"
    SH_TEMP1.UsedRange.Copy
    Range("A1").PasteSpecial xlPasteValues: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    
    '''[SET DATA YANG DI PERLUKAN]...
    Range("A1:A" & LR).Copy SH_TEMP2.Range("A1")
    SH_TEMP2.Activate
    Range("B1") = "Suplier"
    Range("C1") = "Branch"
    Range("D1") = "PO Number"
    Range("E1") = "Item Number"
    Range("F1") = "Receipt Number"
    Range("G1") = "Description"
    Range("H1") = "UOM"
    Range("I1") = "Request Date"
    Range("J1") = "Receipt Date"
    Range("K1") = "Quantity Ordered"
    Range("L1") = "Quantity Received"
    Range("M1") = "Balance PO"
    Range("N1") = "Persentase PO/kedatangan"
    Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    
    With SH_TEMP2.UsedRange
        LR = .Row + .Rows.Count - 1
        LC = .Column + .Columns.Count - 1
    End With
    
    Range("B2:B" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],TEMP1!C[-1]:C,2,0),"""")"
    With Range("B2:B" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .Replace "0", "", xlWhole
    End With
    
    Range("C2:C" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-2],TEMP1!C[-2]:C,3,0),"""")"
    With Range("C2:C" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .Replace "0", "", xlWhole
    End With
    
    Range("D2:D" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],TEMP1!C[-3]:C[1],5,0),"""")"
    With Range("D2:D" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .Replace "0", "", xlWhole
    End With
    
    Range("E2:E" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-4],TEMP1!C[-4]:C[4],9,0),"""")"
    With Range("E2:E" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .Replace "0", "", xlWhole
    End With
    
    Range("F2:F" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-5],TEMP1!C[-5]:C[1],7,0),"""")"
    With Range("F2:F" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .Replace "0", "", xlWhole
    End With
    
    Range("G2:G" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-6],TEMP1!C[-6]:C[3],10,0),"""")"
    With Range("G2:G" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .Replace "0", "", xlWhole
    End With
    
    Range("H2:H" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-7],TEMP1!C[-7]:C[4],12,0),"""")"
    With Range("H2:H" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .Replace "0", "", xlWhole
    End With
        
    Range("I2:I" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-8],TEMP1!C[-8]:C[-3],6,0),"""")"
    Range("J2:J" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-9],TEMP1!C[-9]:C[-2],8,0),"""")"
    With Range("I2:J" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .NumberFormat = "M/D/YYYY"
    End With
    
    Range("K2:K" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-10],TEMP1!C[-10]:C[2],13,0),"""")"
    Range("L2:L" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-11],TEMP1!C[-11]:C[2],14,0),"""")"
    Range("M2:M" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-12],TEMP1!C[-12]:C[2],15,0),"""")"
    With Range("K2:M" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .Style = "Comma"
        .NumberFormat = "#,##0.00;[Red]-#,##0.00;#,##0"
    End With
    
    For Each CELL In Range("K2:M" & LR)
        If CELL.Value <> 0 And InStr(1, CELL.Value, ".") = 0 And InStr(1, CELL.Text, ".00") > 0 Then
            CELL.NumberFormat = "#,##0;[Red]-#,##0"
        End If
        If InStr(1, CELL.Value, ".") <> 0 And Len(CELL.Text) = Len(CELL.Value) + 1 And Left(CELL.Text, 1) = "0" Then
            CELL.NumberFormat = "#,##0.0;[Red]-#,##0.0"
        End If
    Next CELL

    Range("N2:N" & LR).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-13],TEMP1!C[-13]:C[2],16,0),"""")"
    With Range("N2:N" & LR)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
        .NumberFormat = "#,##0.00%;[Red]-#,##0.00%;#,##0%"
    End With
    
    For Each CELL In Range("N2:N" & LR)
        If CELL.Value = -1 Or CELL.Value = 1 Or InStr(1, CELL.Text, ".00") > 0 Then
            CELL.NumberFormat = "#,##0%;[Red]-#,##0%"
        End If
    Next CELL
    
    With Range("A2:A" & LR)
        .ClearContents
        .FormulaR1C1 = "=RC[1]&RC[2]&RC[3]"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    Cells(1, 1).Select
    
    For i = LR To 2 Step -1
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            Range(Cells(i + 1, 2), Cells(i + 1, 4)).ClearContents
        End If
    Next i
    Range("A:A").CLEAR
    Range("A1") = "No."
    
    i = 2
    j = 1
    
    Do Until i = LR + 1
        If Cells(i, 2) <> vbNullString Then
            Cells(i, 1) = j
            j = j + 1
        End If
        i = i + 1
    Loop
    
    '..[Border And Style Report]..
    Set RNG = SH_TEMP2.UsedRange
    RNG.Borders.LineStyle = xlNone
    
    With RNG
        LR = .Row + .Rows.Count - 1
        LC = .Column + .Columns.Count - 1
    End With
    
    With RNG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Rows(1).Font.Bold = True
        .Rows(1).Font.Size = 12
        .Rows(1).Font.Name = "Arial"
        .Rows(1).Borders.LineStyle = xlContinuous
        .Rows(1).Interior.ColorIndex = 15
        .EntireColumn.AutoFit
    End With
    
    i = 2
    Do Until i > LR
        If Cells(i, 1) <> vbNullString Then
            x = i
            Do Until Cells(x + 1, 1) <> vbNullString Or x = LR
                x = x + 1
            Loop
            Set RNG_BORDER = Range(Cells(i, 1), Cells(x, LC))
            With RNG_BORDER
                .Borders.LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
            i = x + 1
        Else
            i = i + 1
        End If
    Loop
    
    For Each RNG_ROW In RNG.Rows
        x = RNG_ROW.RowHeight
        RNG_ROW.RowHeight = x + 2
    Next RNG_ROW
    
    For Each RNG_COLUMN In RNG.Columns
        x = RNG_COLUMN.ColumnWidth
        RNG_COLUMN.ColumnWidth = x + 1
        RNG_COLUMN.Borders(xlEdgeTop).LineStyle = xlContinuous
        RNG_COLUMN.Borders(xlEdgeBottom).LineStyle = xlContinuous
        RNG_COLUMN.Borders(xlEdgeLeft).LineStyle = xlContinuous
        RNG_COLUMN.Borders(xlEdgeRight).LineStyle = xlContinuous
    Next RNG_COLUMN
    
    RNG.Rows(2).Insert
    With RNG.Rows(2)
        .Borders.LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Interior.ColorIndex = 16
        .RowHeight = 3
    End With
    
    Rows("1:2").Insert
    Range("A:A").Insert
    Rows("5:5").Select
    ActiveWindow.FreezePanes = True
    Cells(1, 1).Select
    
    SH_TEMP2.Name = "RESULTS"
    Set SH_RESULTS = TWB.Sheets("RESULTS")
    Cells(1, 1).Select
    
    With SH_RESULTS.Tab
        .Color = RGB(15, 19, 19)
        .TintAndShade = 0
    End With
    
    '[SAVE FILE EXCEL & PDF]....
    PATH_RESULTS_EXCEL = SH_HOME.Range("E14") & Application.PathSeparator & SH_HOME.Range("D14") & ".xlsx"
    PATH_RESULTS_PDF = SH_HOME.Range("E15") & Application.PathSeparator & SH_HOME.Range("D15")
    
    Set RNG_RESULTS = SH_RESULTS.UsedRange
    
    '''[ SAVE TO EXCEL ]'''
    SH_RESULTS.Activate
    SH_RESULTS.Copy
    Set WB_RESULTS = ActiveWorkbook
    Cells(1, 1).Select
    WB_RESULTS.SaveAs PATH_RESULTS_EXCEL, xlOpenXMLStrictWorkbook
    WB_RESULTS.Close True
    
    '''[ SAVE TO PDF }'''
    SH_RESULTS.Activate
    With SH_RESULTS.PageSetup
        .PrintArea = RNG_RESULTS.Address
'        .Orientation = xlLandscape
        .Orientation = xlPortrait
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    
    SH_RESULTS.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PATH_RESULTS_PDF
    
    SH_RESULTS.Activate
    
    SH_HOME.Activate
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
