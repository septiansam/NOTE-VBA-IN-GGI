Option Explicit

Public Periode As String

Sub ALERTGUDANG_2()

'Call ALERTGUDANG_1

Application.DisplayAlerts = False
Dim TWB As Workbook, WS1 As Worksheet, i As Integer, j As Integer, lastRow As Integer, DirFile As String, FILEPATHWO As String, FILEWO As String, HASIL As Worksheet, IM As Worksheet, OV As Worksheet
Dim CARI_OV As Worksheet, LAST As Long
Dim LR As Long, LC As Long
Dim RNG_HASIL As Range, Baris As Range

Set TWB = ThisWorkbook
Set WS1 = TWB.Sheets("TOMBOL")
Set IM = TWB.Sheets("DATA IM")
Set CARI_OV = TWB.Sheets("CARI OV")
Set HASIL = TWB.Sheets("HASIL")

Dim WS As Worksheet

For Each WS In TWB.Worksheets
    If WS.Name <> "TOMBOL" And _
        WS.Name <> "DATA IM" And _
        WS.Name <> "DATA OV" And _
        WS.Name <> "Cari OV" And _
        WS.Name <> "Report WO Buyer" And _
        WS.Name <> "HASIL" Then
    
        WS.Delete
        
    End If
Next WS

'[*]Septian... 26-02-2024
IM.Activate
lastRow = IM.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

'[ISI YG KOSONG UNTUK G/L CAT]...
For i = 9 To lastRow
    If Cells(i, 25) = "" And Cells(i - 1, 25) <> "" Then
        Cells(i, 25) = Cells(i - 1, 25)
    End If
Next i

HASIL.Range("A3:Z100000").Clear

If Evaluate("isref('" & "DATA OV" & "'!A1)") Then Sheets("DATA OV").Delete
Sheets.Add(After:=IM).Name = "DATA OV"
Set OV = TWB.Sheets("DATA OV")

If Evaluate("isref('" & "TES2" & "'!A1)") Then Sheets("TES2").Delete
If Evaluate("isref('" & "TES1" & "'!A1)") Then Sheets("TES1").Delete
Sheets.Add(After:=IM).Name = "TES1"

'BUKA DATA OV
FILEPATHWO = WS1.Range("G8").Value & Application.PathSeparator & WS1.Range("G7").Value & ".xlsx"
FILEWO = WS1.Range("G7").Value & ".xlsx"
DirFile = FILEPATHWO
If Dir(DirFile) = "" Then
        TWB.Activate: MsgBox "File " & FILEWO & " doesn't exist", vbCritical
    Exit Sub
Else
    Workbooks.Open FILEPATHWO
End If
Sheets(1).Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
Cells.Copy
OV.Cells(1, 1).PasteSpecial xlPasteValues: Application.CutCopyMode = False
Workbooks(FILEWO).Close False

IM.Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
lastRow = IM.Cells(Rows.Count, 1).End(xlUp).Row

'BUKA REPORT WO BUYER
FILEPATHWO = WS1.Range("G11").Value & Application.PathSeparator & WS1.Range("G10").Value & ".xlsx"
FILEWO = WS1.Range("G10").Value & ".xlsx"
DirFile = FILEPATHWO
If Dir(DirFile) = "" Then
        TWB.Activate: MsgBox "File " & FILEWO & " doesn't exist", vbCritical
    Exit Sub
Else
    Workbooks.Open FILEPATHWO
End If
Sheets(1).Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
Cells.Copy
TWB.Sheets("TES1").Cells(1, 1).PasteSpecial xlPasteValues: Application.CutCopyMode = False
Workbooks(FILEWO).Close False

Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TES2"

IM.Select
LR = ActiveSheet.Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'Rows(7).AutoFilter
'Range("$A$6:$AZ$" & LR).AutoFilter Field:=10, Criteria1:="IM"
'Range("$A$6:$Z$100000").SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets("TES2").Cells(1, 1)
Cells.AutoFilter
Cells.AutoFilter 10, "IM"
Range("$A$1:$Z$" & LR).SpecialCells(xlCellTypeVisible).Copy
Sheets("TES2").Select
'Sheets("TES2").Cells(1, 1)
Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False

Sheets("TES2").Select
'Rows(2).Delete

Cells(1, 3) = "OLAH"
LAST = Sheets("TES2").Cells(Rows.Count, 1).End(xlUp).Row

'Cells(1, 3) = "BUYER"

'Range("C2:C" & LAST).Formula = "=A2&""_""&N2"

'[NEW SAM 15-03-2024]..............................
Cells(1, 3) = "ITEM-BU-LOCATION-LOT NUM-UC"
Range("C2:C" & LAST).Formula = "=A2&""-""&H2&""-""&M2&""-""&N2&""-""&R2"

'[DONE].......

Range("D2:D" & LAST).FormulaR1C1 = _
        "=IFERROR(IF(RC[5]<>"""",INDEX('TES1'!C[9],MATCH('TES2'!RC[5],'TES1'!C[-1],0)),""""),"""")"

Columns("C:D").Copy
Cells(1, 3).PasteSpecial xlPasteValues: Application.CutCopyMode = False

Range("C2:C" & LAST).Copy HASIL.Cells(3, 3)
HASIL.Select
Range("C:C").RemoveDuplicates Columns:=1, Header:=xlYes

lastRow = HASIL.Cells(Rows.Count, 3).End(xlUp).Row
Range("B3:B" & lastRow).Formula = "=IF(INDEX('TES2'!D:D,MATCH(HASIL!C3,'TES2'!C:C,0)) = 0,"""",INDEX('TES2'!D:D,MATCH(HASIL!C3,'TES2'!C:C,0)))"
Range("D3:D" & lastRow).Formula = "=IF(INDEX('TES2'!B:B,MATCH(HASIL!C3,'TES2'!C:C,0)) = 0,"""",INDEX('TES2'!B:B,MATCH(HASIL!C3,'TES2'!C:C,0)))"
Range("E3:E" & lastRow).Formula = "=IF(INDEX('TES2'!E:E,MATCH(HASIL!C3,'TES2'!C:C,0)) = 0,"""",INDEX('TES2'!E:E,MATCH(HASIL!C3,'TES2'!C:C,0)))"
Range("F3:F" & lastRow).Formula = "=IF(INDEX('TES2'!N:N,MATCH(HASIL!C3,'TES2'!C:C,0)) = 0,"""",INDEX('TES2'!N:N,MATCH(HASIL!C3,'TES2'!C:C,0)))"
Range("H3:H" & lastRow).Formula = "=IF(INDEX('TES2'!K:K,MATCH(HASIL!C3,'TES2'!C:C,0)) = 0,"""",INDEX('TES2'!K:K,MATCH(HASIL!C3,'TES2'!C:C,0)))"

'TAMBAHKAN GL/CAT SEPTIAN
With Range("P3:P" & lastRow)
    .Formula = "=INDEX('TES2'!Y:Y,MATCH(HASIL!C3,'TES2'!C:C,0))"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

'TAMBAHKAN AMOUNT
With Range("Q3:Q" & lastRow)
    .FormulaR1C1 = _
        "=IFERROR(SUMIF('TES2'!C[-14],HASIL!RC[-14],'TES2'!C[4]),""-"")"
        '"=IFERROR(IF(INDEX('TES2'!C[4],MATCH(HASIL!RC[-14],'TES2'!C[-14],0)) = 0,""-"",INDEX('TES2'!C[4],MATCH(HASIL!RC[-14],'TES2'!C[-14],0))),""-"")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

Cells.Copy: Cells(1, 1).PasteSpecial xlPasteValues: Application.CutCopyMode = False

Cells(3, 1) = 1
Range("A3").DataSeries Rowcol:=xlColumns, Step:=1, Stop:=lastRow

Sheets("TES2").Cells.Clear

Sheets("DATA OV").Select
'Rows(7).AutoFilter
LR = ActiveSheet.Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'Range("$A$7:$AZ$100000").AutoFilter Field:=10, Criteria1:="OV"
'Range("$A$6:$Z$100000").SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets("TES2").Cells(1, 1)

Cells.AutoFilter
Cells.AutoFilter 10, "OV"
Range("$A$1:$Z$" & LR).SpecialCells(xlCellTypeVisible).Copy
'Sheets("TES2").Cells(1, 1)
Sheets("TES2").Select
Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False

Sheets("TES2").Select
'Rows(2).Delete

'Cells(1, 3) = "OLAH"
Sheets("TES2").Cells(Rows.Count, 1).End(xlUp).Select
LAST = Selection.Row
'Range("C2:C" & LAST).Formula = "=A2&""_""&N2"

'[NEW SAM 15-03-2024]..............................
Cells(1, 3) = "ITEM-BU-LOCATION-LOT NUM-UC"
Range("C2:C" & LAST).Formula = "=A2&""-""&H2&""-""&M2&""-""&N2&""-""&R2"

'[DONE].......

Columns("C").Copy

Cells(1, 3).PasteSpecial xlPasteValues: Application.CutCopyMode = False

HASIL.Select
Range("G3:G" & lastRow).Formula = "=IFNA(VLOOKUP(C3,'TES2'!C:L,9,0),""Not found"")"

Columns("G:H").NumberFormat = "M/D/YYYY"
Columns("G").Copy: Cells(1, 7).PasteSpecial xlPasteValues: Application.CutCopyMode = False
Range("I3:I" & lastRow).Formula = "=IFERROR(H3-G3,0)"

Cells.HorizontalAlignment = xlCenter

If Evaluate("isref('" & "TES2" & "'!A1)") Then Sheets("TES2").Delete
If Evaluate("isref('" & "TES1" & "'!A1)") Then Sheets("TES1").Delete

IM.Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

Sheets("DATA OV").Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

HASIL.Select

lastRow = HASIL.Cells(Rows.Count, 3).End(xlUp).Row
For i = 3 To lastRow
    Cells(i, 3) = Left(Cells(i, 3), 6)
Next i

Range("D:D").Insert
Range("Q2:Q" & lastRow).Copy Range("D2")
Range("Q:Q").Delete

Range("H:H").Insert
Range("R2:R" & lastRow).Copy Range("H2")
Range("R:R").Delete

Cells(1, 1).Select

HASIL.Select
LR = HASIL.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

    
For i = LR To 3 Step -1
    If Cells(i, 2) = "" Then
        Range("a" & i).EntireRow.Delete xlUp
    End If
Next i

LR = HASIL.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row
LC = HASIL.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column
    
Set RNG_HASIL = Range(Cells(2, 1), Cells(LR, LC))

RNG_HASIL.Select
'[TAMBAHAN JIKA Diff Day <0, JADIKAN 0]
'.....................................................
For i = 3 To LR
    If IsNumeric(Cells(i, 11)) Then
        If Cells(i, 11).Value < 0 Then Cells(i, 11) = 0
    End If
Next i

Rows(2).Clear
Cells(2, 1) = "No"
Cells(2, 2) = "Buyer"
Cells(2, 3) = "Item"
Cells(2, 4) = "G/L Cat"
Cells(2, 5) = "Description"
Cells(2, 6) = "Description 2"
Cells(2, 7) = "Lot Serial"
Cells(2, 8) = "Amount"
Cells(2, 9) = "Receipt Material Date"
Cells(2, 10) = "Issue Material Date"
Cells(2, 11) = "Diff Day"

ActiveWorkbook.Worksheets("HASIL").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("HASIL").Sort.SortFields.Add2 Key:=Range("K3:K" & LR) _
    , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("HASIL").Sort
    .SetRange Range("A2:K" & LR)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Range("A3:A" & LR).ClearContents
Range("A3") = "1"
Range("A3").DataSeries Rowcol:=xlColumns, Step:=1, Stop:=LR - 2
Range("G:G").Delete
HASIL.Cells.Interior.Color = xlNone

LR = HASIL.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row
LC = HASIL.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column
    
Cells.EntireColumn.AutoFit

'[BUAT DULU SHEET OLAH, COPYKAN DATANYA]
'[UNTUK PEMBUATAN SHEETS ANALISA DI AKHIR]
'[23-02-2024]............................................

If WorksheetExists("OLAH") Then Sheets("OLAH").Delete
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "OLAH"

HASIL.Activate

Range(Cells(2, 1), Cells(LR, LC)).Copy
Sheets("OLAH").Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False

Set RNG_HASIL = Range(Cells(2, 1), Cells(LR, LC))

With RNG_HASIL
    .Rows(1).Font.Bold = True
    .Rows(1).Font.Name = "Century Gothic"
    .Rows(1).Font.Color = vbWhite
    .Rows(1).Interior.Pattern = xlSolid
    .Rows(1).Interior.PatternColorIndex = xlAutomatic
    .Rows(1).Interior.Color = RGB(52, 98, 101)
    .Rows(1).VerticalAlignment = xlCenter
    For Each Baris In .Rows
    
        If Baris.Row > 2 And Baris.Row Mod 2 = 0 Then
            Baris.Font.Name = "Verdana"
            With Baris
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.Color = RGB(255, 255, 255)
                .VerticalAlignment = xlCenter
            End With
        ElseIf Baris.Row > 2 And Baris.Row Mod 2 <> 0 Then
            Baris.Font.Name = "Verdana"
            With Baris
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.Color = RGB(228, 240, 241)
                .VerticalAlignment = xlCenter
            End With
        End If
    
    Next Baris
End With
With RNG_HASIL
    .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
End With

Cells.EntireColumn.AutoFit
Cells.EntireRow.AutoFit

Cells(1, 1).Select

Rows("1:3").Insert
With Range("a2")
    .Value = "Receipt Material Date(OV) Vs Issue Material Date(IM)"
'    .Font.Bold = True
    .Font.Name = "Century Gothic"
    .Font.Size = 36
    .Font.Color = vbWhite
End With

LR = HASIL.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = HASIL.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

Set RNG_HASIL = Range(Cells(5, 1), Cells(LR, LC))

RNG_HASIL.HorizontalAlignment = xlCenter
Rows(5).RowHeight = Rows(5).RowHeight + 5
Range(Cells(6, 1), Cells(LR, LC)).RowHeight = Rows(6).RowHeight + 2

With Range(Cells(2, 1), Cells(3, LC))
    .Merge
    .Interior.Color = RGB(79, 146, 151)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .RowHeight = .RowHeight + 5
End With

With Range(Cells(4, 1), Cells(4, LC))
    .Merge
    .Interior.Color = RGB(228, 240, 241)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .RowHeight = 20
End With
Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter

Range("B6:B" & LR).HorizontalAlignment = xlLeft
'Range("E6:F" & LR).HorizontalAlignment = xlLeft
'Range("H6:I" & LR).HorizontalAlignment = xlLeft
'Range("J6:J" & LR).HorizontalAlignment = xlRight

LC = HASIL.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

If HASIL.AutoFilterMode = True Then HASIL.AutoFilterMode = False
'Range("A5", Cells(5, LC)).AutoFilter

Range("A:A").Insert
Range("A:A").ColumnWidth = 3

If ActiveWindow.FreezePanes = True Then ActiveWindow.FreezePanes = False

Rows("6:6").Select
ActiveWindow.FreezePanes = True

Periode = "Periode : " & WS1.Range("C17")

With Range("b4:k4")
    .Value = Periode
    .Font.Name = "Calibri"
    .Font.Bold = False
    .Font.Italic = True
    .HorizontalAlignment = xlRight
    .Font.Size = 14
End With
ActiveWindow.Zoom = 85
Range("A1").Select

Application.DisplayAlerts = False

Call ANALISA(Periode)

WS1.Select

Dim PATH_HASIL As String, WB_HASIL As Workbook
PATH_HASIL = WS1.Range("G14") & Application.PathSeparator & WS1.Range("G13") & ".xlsx"

Sheets(Array("HASIL", "Analisa")).Copy
Set WB_HASIL = ActiveWorkbook
WB_HASIL.Activate
Sheets(1).Select: ActiveWindow.Zoom = 85
Sheets(2).Select: ActiveWindow.Zoom = 85
Sheets(1).Name = "Details"
Sheets(2).Name = "Analisa"
Sheets(1).Select: Range("A1").Select
WB_HASIL.SaveAs PATH_HASIL, xlOpenXMLWorkbook
WB_HASIL.Close (True)

TWB.Activate
WS1.Activate
Cells(1, 1).Select

TWB.Save

Application.DisplayAlerts = True

End Sub

