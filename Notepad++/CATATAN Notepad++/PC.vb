'Option Explicit

Dim TWB As ThisWorkbook, DB As Worksheet, HOME As Worksheet, INV As Worksheet, FSI As Worksheet, LOAD As Worksheet
Dim sh_Input As Worksheet
Dim TEMP As Worksheet, RNG As Range, CELL As Range
Dim lr As Long, lc As Long, firstRowData As Long, lastRowData As Long, i As Long, j As Long, val As Long
Dim wb_INV As Workbook, wb_Src As Workbook
Dim FullPathFile As Variant
Dim FileDialog As FileDialog
Dim randomColor As Long


Sub MAIN()

Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

    Set TWB = ThisWorkbook
    Set HOME = TWB.Sheets("HOME")
    If Sheets("DB").Visible = False Then Sheets("DB").Visible = True
    Set DB = TWB.Sheets("DB")
    
    For i = Sheets.Count To 3 Step -1
        Sheets(i).Delete
    Next i
    
    If WorksheetExists("TEMP") Then Sheets("TEMP").Delete

    Set INV = Sheets.Add(AFTER:=Sheets(Sheets.Count)): INV.Name = "INV"
    Set FSI = Sheets.Add(AFTER:=Sheets(Sheets.Count)): FSI.Name = "FSI"
    Set LOAD = Sheets.Add(AFTER:=Sheets(Sheets.Count)): LOAD.Name = "LOAD"
    Set TEMP = Sheets.Add(AFTER:=Sheets(Sheets.Count)): TEMP.Name = "TEMP"
    
    '::::: AMBIL FILE INVOICE :::::'
        FullPathFile = HOME.Range("D7")
        Set wb_INV = Workbooks.Open(FullPathFile): wb_INV.Activate: Sheets("PACKING LIST").Select: ActiveSheet.AutoFilterMode = False: Cells.EntireColumn.Hidden = False: Cells.EntireRow.Hidden = False
        lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
        lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
        Range(Cells(1, 1), Cells(lr, lc)).Copy
        INV.Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
        wb_INV.Close False
        FullPathFile = ""


    '::::: AMBIL FILE FINAL SI :::::'
        FullPathFile = HOME.Range("D8")
        Set wb_Src = Workbooks.Open(FullPathFile): wb_Src.Activate: Sheets(1).Select: ActiveSheet.AutoFilterMode = False: Cells.EntireColumn.Hidden = False: Cells.EntireRow.Hidden = False
        lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
        lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
        Range(Cells(1, 1), Cells(lr, lc)).Copy FSI.Range("A1")
        wb_Src.Close False
        Set wb_Src = Nothing
        FullPathFile = ""

    '::::: AMBIL FILE LOADINGAN :::::'
        
        FullPathFile = HOME.Range("D9")
        Set wb_Src = Workbooks.Open(FullPathFile): wb_Src.Activate: Sheets(1).Select: ActiveSheet.AutoFilterMode = False: Cells.EntireColumn.Hidden = False: Cells.EntireRow.Hidden = False
        lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
        lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
        Range(Cells(1, 1), Cells(lr, lc)).Copy LOAD.Range("A1")
        wb_Src.Close False
        Set wb_Src = Nothing
        FullPathFile = ""

    Call Data_Processing
    
    If Sheets("DB").Visible = True Then Sheets("DB").Visible = False
    
    Randomize
    randomColor = RGB(Int(Rnd() * 256), Int(Rnd() * 256), Int(Rnd() * 256))
    With TEMP.Tab
        .Color = randomColor
        .TintAndShade = 0
    End With
    
    HOME.Activate
    Cells(1, 1).Select

Application.DisplayAlerts = True
Application.AskToUpdateLinks = True

End Sub

Sub Data_Processing()

Set TWB = ThisWorkbook

Set DB = TWB.Sheets("DB")
Set HOME = TWB.Sheets("HOME")
Set INV = TWB.Sheets("INV")
Set FSI = TWB.Sheets("FSI")
Set LOAD = TWB.Sheets("LOAD")
Set TEMP = TWB.Sheets("TEMP")

DB.Activate
Range("A1").CurrentRegion.Copy
TEMP.Activate
Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select

'.....................................
''''OLAH DATA INV, DAN CLEAN DATA''''
'`````````````````````````````````````
INV.Activate
Cells.EntireColumn.AutoFit
INV.AutoFilterMode = False
firstRowData = Range("A:A").Find("CARTON" & "*" & "NUMBER", , xlFormulas, xlPart, xlByRows, xlNext).Row
Rows("1:" & firstRowData - 1).Delete

firstRowData = Range("A:A").Find("GRAND TOTAL", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
lastRowData = Range("A" & Rows.Count).End(xlUp).Row

Rows("" & firstRowData & "" & ":" & "" & lastRowData & "").Delete

lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column

Range(Cells(1, 1), Cells(lr, lc)).AutoFilter 4, "="
If Application.WorksheetFunction.CountA(Range("B:B")) = 1 Then
    Stop
End If

Range(INV.AutoFilter.Range.Offset(2).SpecialCells(xlCellTypeVisible).Cells(1, 1), _
    INV.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Cells(Rows.Count, lc).End(xlUp)).Delete xlUp

INV.AutoFilterMode = False

lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column

Range(Cells(1, 1), Cells(lr, lc)).AutoFilter 5, "=", xlOr, "PO NO"

Range(INV.AutoFilter.Range.Offset(2).SpecialCells(xlCellTypeVisible).Cells(1, 1), _
    INV.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Cells(Rows.Count, lc).End(xlUp)).Delete xlUp
    
INV.AutoFilterMode = False

lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row

INV.Range("A:A").TextToColumns Destination:=TEMP.Range("Q1"), DataType:=xlDelimited, _
        Other:=True, OtherChar:="-"
        
INV.Activate

Range("Q3:R" & lr).Copy
TEMP.Activate
Range("B3").PasteSpecial xlPasteValuesAndNumberFormats

INV.Activate
Range("D3:D" & lr).Copy TEMP.Range("D3")
Range("E3:E" & lr).Copy TEMP.Range("E3")
Range("L3:L" & lr).Copy TEMP.Range("J3")
Range("N3:O" & lr).Copy TEMP.Range("M3")

TEMP.Activate
lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
Range("A3:A" & lr).Value = "DIVISION"
Range("F3:F" & lr).Value = "BRIEF"
Range("R3:R" & lr) = "x"
Range("T3:T" & lr) = "x"

Range("H3:H" & lr).FormulaR1C1 = _
    "=IFERROR(IF(RIGHT(RC[-4],1)=""A"",""S"",IF(RIGHT(RC[-4],1)=""B"",""M"",IF(RIGHT(RC[-4],1)=""C"",""L"",IF(RIGHT(RC[-4],1)=""D"",""XL"",IF(RIGHT(RC[-4],1)=""E"",""XXL"",""""))))),"""")"
Range("K3:K" & lr).FormulaR1C1 = "=RC[-8]-RC[-9]+1"
Range("I3:I" & lr).FormulaR1C1 = "=IFERROR(RC[1]/RC[2],"""")"

Range("O3:O" & lr).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-4],"""")"
Range("O3:O" & lr).NumberFormat = "0.00"

Range("P3:P" & lr).FormulaR1C1 = "=IFERROR(RC[-3]/RC[-5],"""")"
Range("P3:P" & lr).NumberFormat = "0.00"

'.....................................
'''''    OLAH DATA LOADINGAN     '''''
'`````````````````````````````````````
LOAD.Activate
lr = Range("B" & Rows.Count).End(xlUp).Row
Range("A:A").Insert
With Range("A3:A" & lr)
    .FormulaR1C1 = "=CONCATENATE(RC[1],""-"",RC[2])"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With
With Range("AH3:AH" & lr)
    .FormulaR1C1 = "=CONCATENATE(RC[-2],"" "",RC[-3])"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With
Cells.EntireColumn.AutoFit: Cells(1, 1).Select

TEMP.Activate
lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row

'COLOR
With Range("G3:G" & lr)
    .FormulaR1C1 = _
        "=IFERROR(VLOOKUP(CONCATENATE(RC[-2],""-"",RC[-3]),LOAD!C1:C22,22,FALSE),"""")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    .EntireColumn.AutoFit
End With

'L
Range("Q3:Q" & lr).FormulaR1C1 = _
    "=IFERROR(VLOOKUP(CONCATENATE(RC[-12],""-"",RC[-13]),LOAD!C1:C27,23,FALSE),"""")"
'W
Range("S3:S" & lr).FormulaR1C1 = _
    "=IFERROR(VLOOKUP(CONCATENATE(RC[-14],""-"",RC[-15]),LOAD!C1:C27,25,FALSE),"""")"
'H
Range("U3:U" & lr).FormulaR1C1 = _
    "=IFERROR(VLOOKUP(CONCATENATE(RC[-16],""-"",RC[-17]),LOAD!C1:C27,27,FALSE),"""")"

'CTN DIMS (CM)
With Range("Q3:U" & lr)
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    .EntireColumn.AutoFit
End With

'CBM / CTN
Range("V3:V" & lr).FormulaR1C1 = "=IFERROR((RC[-5]*RC[-3]*RC[-1])/1000000,"""")"
Range("V3:V" & lr).EntireColumn.AutoFit
Range("V3:V" & lr).NumberFormat = "0.0000"

'TTL VOL (CBM)
Range("L3:L" & lr).FormulaR1C1 = "=IFERROR(RC[-1]*RC[10],"""")"
Range("L3:L" & lr).NumberFormat = "0.00"

'FABRIC CONTENT
With Range("Y3:Y" & lr)
    .FormulaR1C1 = _
    "=IFERROR(VLOOKUP(CONCATENATE(RC[-20],""-"",RC[-21]),LOAD!C1:C34,34,FALSE),"""")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    .EntireColumn.AutoFit
End With

'.....................................
'''''    OLAH DATA FINAL SI     '''''
'`````````````````````````````````````
FSI.Activate
AW = Cells.Find("DESCRIPTION OF PACKAGES AND GOODS").Row
Cells.UnMerge
ActiveWindow.Zoom = 65
Cells.EntireColumn.AutoFit
LASTCOLUMN = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column

Range(Cells(1, 1), Cells(AW, LASTCOLUMN)).Delete

DATATES3 = Cells(Rows.Count, 4).End(xlUp).Row

For i = 3 To DATATES3
    If Left(Cells(i, 1), 4) = "SEAL" Then
        Cells(i - 1, 20) = Cells(i - 1, 1)
        Cells(i + 1, 21) = Cells(i + 1, 1)
    End If
    
    MYSTRING = Cells(i, 3).Text
    MYSTRING = Left(MYSTRING, Len(MYSTRING) - 1)
    
    Cells(i, 3) = MYSTRING
Next i

Dim AB As String, AC As String
AB = Range("T3:T" & DATATES3).End(xlDown).Text
AC = Range("U3:U" & DATATES3).End(xlDown).Text

Range("T3") = AB
Range("U3") = AC

Range("T3:T" & DATATES3).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
Range("U3:U" & DATATES3).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"

Columns("T:U").Copy
Cells(1, 20).PasteSpecial xlPasteValues: Application.CutCopyMode = False
Range("S3:S" & DATATES3).Formula = "=C3&""_""&B3"
Range("S3:S" & DATATES3).Copy
Range("S3:S" & DATATES3).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False

TEMP.Select
DATA_TES2 = Cells(Rows.Count, 1).End(xlUp).Row
Range("Z3:Z" & DATA_TES2).Formula = "=D3&""_""&E3"
Range("W3:W" & DATA_TES2).FormulaR1C1 = _
        "=IFERROR(INDEX(FSI!C20,MATCH(TEMP!R3C26,FSI!C19,0)),""DATA TIDAK DITEMUKAN"")"
Range("X3:X" & DATA_TES2).FormulaR1C1 = _
        "=IFERROR(INDEX(FSI!C21,MATCH(TEMP!R3C26,FSI!C19,0)),""DATA TIDAK DITEMUKAN"")"

Columns("W:X").Copy
Cells(1, 23).PasteSpecial xlPasteValues: Application.CutCopyMode = False
Range("Z:Z").Clear

Range("D" & DATA_TES2 + 1).Value = "Total"

Range(Cells(DATA_TES2 + 1, 10), Cells(DATA_TES2 + 1, 14)).Formula = "=SUM(J3:J" & DATA_TES2 & ")"

Rows(DATA_TES2 + 1).Font.Bold = True

Cells.EntireColumn.AutoFit
lr = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
lc = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
With Range(Cells(3, 1), Cells(lr, lc))
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlHairline
End With
Set RNG = Range(Cells(1, 1), Cells(lr, lc))
For Each CELL In RNG.Columns
    CELL.ColumnWidth = CELL.ColumnWidth + 3
Next CELL
Cells(1, 1).Select

TEMP.Name = "Packing List"
INV.Delete
FSI.Delete
LOAD.Delete

End Sub


Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
        WorksheetExists = Not Sheets(shtName) Is Nothing
    On Error GoTo 0
End Function


