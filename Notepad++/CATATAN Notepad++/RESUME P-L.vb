'
'
'
'

Public TWB As Workbook
Public WB_COST As Workbook
Public WB_KETERANGAN As Workbook
Public PATH_COST As String, PATH_KETERANGAN As String
Public PATH_HASIL As String
Public LR As Long, LC As Long, i As Long
Public TEMP1 As Worksheet, TEMP2 As Worksheet, TEMP3 As Worksheet, TEMP4 As Worksheet, TEMP5 As Worksheet

Sub MAIN()

Application.DisplayAlerts = False

Set TWB = ThisWorkbook
TWB.Activate
HOME.Activate

PATH_COST = HOME.Range("D7") & Application.PathSeparator & _
            HOME.Range("C7") & _
            HOME.Range("E7")

PATH_KETERANGAN = HOME.Range("D8") & Application.PathSeparator & _
            HOME.Range("C8") & _
            HOME.Range("E8")
            
PATH_HASIL = HOME.Range("D9") & Application.PathSeparator & _
            HOME.Range("C9") & _
            HOME.Range("E9")

If Dir(PATH_COST) = "" Or Dir(PATH_KETERANGAN) = "" Then
    MsgBox "File Tarikan Doesn't Exists", vbCritical, "File Tarikan Tidak Ditemukan"
    Exit Sub
End If

For i = TWB.Sheets.Count To 2 Step -1
    Sheets(i).Delete
Next i

For i = 1 To 3
    If WorksheetExists("TEMP" & i) Then Sheets("TEMP" & i).Delete
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TEMP" & i
Next i

Set TEMP1 = TWB.Sheets("TEMP1")
Set TEMP2 = TWB.Sheets("TEMP2")
Set TEMP3 = TWB.Sheets("TEMP3")

Dim SRC_FOUND As Boolean

SRC_FOUND = False
Set WB_COST = Workbooks.Open(PATH_COST)
WB_COST.Activate
If WorksheetExists("Resume per Buyer") Then
Sheets("Resume per Buyer").Select
If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
Cells.Copy: TEMP1.Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
SRC_FOUND = True
End If
WB_COST.Close (False)

If SRC_FOUND = False Then Stop

SRC_FOUND = False
Set WB_KETERANGAN = Workbooks.Open(PATH_KETERANGAN)
WB_KETERANGAN.Activate
Sheets(1).Select
Cells.Copy: TEMP2.Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
SRC_FOUND = True
WB_KETERANGAN.Close (False)

TWB.Activate

TEMP1.Activate

If TEMP1.AutoFilterMode = True Then TEMP1.AutoFilterMode = False
Rows("4:4").AutoFilter
ActiveSheet.Range("$A$4:$P$100000").AutoFilter Field:=1, Criteria1:=RGB(198, _
    224, 180), Operator:=xlFilterCellColor
If Range("B" & Rows.Count).End(xlUp).Value <> "Buyer" Then
    Range("A4").CurrentRegion.Offset(1).Delete
End If
TEMP1.ShowAllData

ActiveSheet.Range("$A$4:$P$100000").AutoFilter Field:=1, Criteria1:=RGB(0, 176 _
    , 240), Operator:=xlFilterCellColor
If Range("A" & Rows.Count).End(xlUp).Value <> "Factory" Then
    Range("A4").CurrentRegion.Offset(1).Delete
End If
If TEMP1.AutoFilterMode = True Then TEMP1.AutoFilterMode = False

With TEMP1.UsedRange
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With
Range("A1").Select

Rows("1:3").Delete: Cells(1, 1).Select

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

With TEMP1.Range(Cells(1, 1), Cells(LR, LC))
    .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

Range("A1").Select

Range("D:D").Insert
Range("D1") = "Item"
With Range(Cells(2, 4), Cells(LR, 4))
    .FormulaR1C1 = _
        "=IFERROR(INDEX(TEMP2!C[1],MATCH(TEMP1!RC[1],TEMP2!C[-1],0)),RC[1])"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
End With
Range("A1").Select
Range("E:E").Delete

TEMP1.Activate
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
    
Dim PV_RANGE As Range
Dim PV_TABLE As PivotTable
Dim PV_CACHE As PivotCache

Set PV_RANGE = TEMP1.Range(Cells(1, 1), Cells(LR, LC))
Set PV_CACHE = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                    SourceData:=PV_RANGE)
Set PV_TABLE = PV_CACHE.CreatePivotTable _
                    (TableDestination:=TEMP3.Range("A1"), _
                    TableName:="PV_REPORT")
TEMP3.Activate

'[INSERT FIELD FIELD NYA]...
With PV_TABLE.PivotFields("Buyer")
    .Caption = "Buyer"
    .Orientation = xlRowField
    .Position = 1
End With

With PV_TABLE.PivotFields("Factory")
    .Caption = "Factory"
    .Orientation = xlRowField
    .Position = 2
End With

With PV_TABLE.PivotFields("Item")
    .Caption = "Product"
    .Orientation = xlRowField
    .Position = 3
End With

'[ISI].....
With PV_TABLE.PivotFields("Sum of QTY")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlSum
    .Caption = "Quantity"
    .NumberFormat = "#,##0_);[Red](#,##0);_(* ""-""??_)" '"_(* #,##0_);[Red]_(* (#,##0);_(* ""-""??_);_(@_)"
End With

With PV_TABLE.PivotFields("Total Sales")
    .Orientation = xlDataField
    .Position = 2
    .Function = xlSum
    .Caption = "Sales"
    .NumberFormat = "#,##0_);[Red](#,##0);_(* ""-""??_)" '"_(* #,##0_);[Red]_(* (#,##0);_(* ""-""??_);_(@_)"
End With

With PV_TABLE.PivotFields("Sum of Profit.Lost(USD)")
    .Orientation = xlDataField
    .Position = 3
    .Function = xlSum
    .Caption = "Profit/Loss (USD)"
    .NumberFormat = "#,##0_);[Red](#,##0);_(* ""-""??_)" '"_(* #,##0_);[Red]_(* (#,##0);_(* ""-""??_);_(@_)"
End With

'[SETTING]...
With PV_TABLE
    .RowGrand = False
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
'    .PivotFields("Factory").Subtotals = Array _
'                (False, False, False, False, False, False, False, False, False, False, False, False)
    .TableStyle2 = "PivotStyleDark7"
End With

LR = TEMP3.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = TEMP3.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

With Range(Cells(1, 1), Cells(LR, LC))
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

ActiveWindow.Zoom = 80

Cells(1, LC + 1) = "Profit/Loss (%)"
Cells(1, LC + 2) = "Profit/Loss (PCS)"

Range("G:G").NumberFormat = "0.00%;[Red]-0.00%"
Range("H:H").NumberFormat = "#,##0.00_);[Red](#,##0.00);_(* ""-""??_)"

With Range(Cells(2, 7), Cells(LR, 7))
    .FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],""-"")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

With Range(Cells(2, 8), Cells(LR, 8))
    .FormulaR1C1 = _
        "=IFERROR(IF(RIGHT(RC[-6],5)<>""Total"",ROUNDUP(RC[-2]/RC[-4],2),""-""),""-"")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

LR = TEMP3.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

For i = LR To 2 Step -1
    If IsNumeric(Cells(i, 5)) Then
        If Cells(i, 5) = 0 Then
            Cells(i, 5).EntireRow.Delete xlUp
        End If
    End If
Next i

LR = TEMP3.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = TEMP3.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column
    
Cells.Font.Name = "Verdana"
Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter

For i = 1 To LR
    If Cells(i, 1) = "Buyer" Then
        With Range(Cells(i, 1), Cells(i, LC))
            .Font.Name = "Century Gothic"
            .Font.Color = vbWhite
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(52, 98, 101)
            .Interior.Pattern = xlSolid
            .Interior.PatternColorIndex = xlAutomatic
            .RowHeight = .RowHeight + 4
        End With
        With Range(Cells(i, 1), Cells(i, LC))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
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
    If Cells(i, 1) <> "Grand Total" Then
        If Right(CStr(Cells(i, 1)), 5) Like "Total" Then
            
            With Range("A" & i)
                .Font.Italic = True
                .Font.Color = RGB(2, 112, 192)
                .Font.TintAndShade = 0
'                .Font.Bold = True
            End With
            
            With Range(Cells(i, 1), Cells(i, LC))
    '            .Font.Bold = True
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(228, 240, 241)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .RowHeight = .RowHeight + 2
            End With
            
        End If
    End If
    If Right(CStr(Cells(i, 2)), 5) Like "Total" Then
        With Range(Cells(i, 2), Cells(i, LC))
'            .Font.Bold = True
            .Interior.Pattern = xlSolid
            .Interior.PatternColor = xlAutomatic
            .Interior.Color = RGB(228, 240, 241)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = .RowHeight + 2
        End With
    End If
    
'    If Cells(i, 1) <> "Buyer" Then
'
'        If Cells(i, 1).Row Mod 2 = 0 Then
'            With Range(Cells(i, 1), Cells(i, LC))
'                .Interior.Pattern = xlSolid
'                .Interior.PatternColor = xlAutomatic
'                .Interior.Color = RGB(255, 255, 255)
'                .RowHeight = .RowHeight + 2
'            End With
'            With Range(Cells(i, 2), Cells(i, LC))
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'            End With
'            Cells(i, LC).HorizontalAlignment = xlRight
'        Else
'            With Range(Cells(i, 1), Cells(i, LC))
'                .Interior.Pattern = xlSolid
'                .Interior.PatternColor = xlAutomatic
'                .Interior.Color = RGB(228, 240, 241)
'                .RowHeight = .RowHeight + 2
'            End With
'            With Range(Cells(i, 2), Cells(i, LC))
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'            End With
'            Cells(i, LC).HorizontalAlignment = xlRight
'        End If
'    End If
    
Next i

Cells.EntireColumn.AutoFit

Range(Cells(2, 1), Cells(LR - 1, 1)).HorizontalAlignment = xlLeft

With Range(Cells(1, 1), Cells(LR, LC))
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
End With

Dim Title As String
Dim STR_DATE As String

STR_DATE = Application.WorksheetFunction.Text(Date, "[$-id-ID]mmmm" & "'" & "yy")
Title = "Resume Profit Loss " & STR_DATE

Rows("1:3").Insert

LR = TEMP3.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = TEMP3.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column
    
Range("a2") = Title
With Range(Cells(2, 1), Cells(2, LC))
    .Merge
    .Font.Name = "Century Gothic"
    .Font.Color = vbWhite
    .Font.Size = 36
    .Interior.Color = RGB(79, 146, 151)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .RowHeight = .RowHeight + 5
End With

Cells.EntireColumn.AutoFit
With Range(Cells(3, 1), Cells(3, LC))
    .Merge
    .Interior.Color = RGB(228, 240, 241)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .RowHeight = 20
End With

Range("a:a").Insert
Range("a:a").ColumnWidth = 3

LR = TEMP3.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = TEMP3.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

With Range(Cells(1, 4), Cells(LR, LC))
    .RowHeight = .RowHeight + 1
End With

Dim RNG_PDF As Range
Set RNG_PDF = Range(Cells(1, 2), Cells(LR, LC))
With ActiveSheet.PageSetup
     .Orientation = xlPortrait
     .CenterHorizontally = True
     .LeftMargin = Application.InchesToPoints(0.5)
     .RightMargin = Application.InchesToPoints(0.5)
     .PrintArea = RNG_PDF.Address
     .FitToPagesTall = 1
     .FitToPagesWide = 1
     .Zoom = False
End With

Rows("5:5").Select
ActiveWindow.FreezePanes = True
Range("A1").Select
TEMP3.ExportAsFixedFormat xlTypePDF, PATH_HASIL, xlQualityStandard, False, False, , , False

On Error Resume Next
TEMP1.Delete
TEMP2.Delete
On Error GoTo 0

TEMP3.Activate
TEMP3.Name = "RESULTS"
With TEMP3.Tab
    .Color = RGB(79, 146, 151)
    .TintAndShade = 0
End With

HOME.Activate: Cells(1, 1).Select


Application.DisplayAlerts = True

End Sub

'
'
'
'
Public Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not Sheets(shtName) Is Nothing
    On Error GoTo 0
End Function

