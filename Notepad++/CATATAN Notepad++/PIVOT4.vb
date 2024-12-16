Sub CreatePivot()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Dim twb As Workbook, shAbout As Worksheet, shHome As Worksheet, i As Long, lr As Long, lc As Long, r As Long

Dim shMutasi1 As Worksheet, shMutasi2 As Worksheet, shMesin1 As Worksheet, shMesin2 As Worksheet

Set twb = ThisWorkbook: Set shAbout = twb.Sheets("ABOUT"): Set shHome = twb.Sheets("HOME")
Set shMutasi1 = twb.Sheets("P1.MUTASI")
Set shMutasi2 = twb.Sheets("P2.MUTASI")
Set shMesin1 = twb.Sheets("P1.MESIN")
Set shMesin2 = twb.Sheets("P2.MESIN")

If Evaluate("isref('" & "PIVOT" & "'!A1)") Then Sheets("PIVOT").Delete
If Evaluate("isref('" & "RESUME" & "'!A1)") Then Sheets("RESUME").Delete

Dim PivotSheet As Worksheet, shResume As Worksheet
Sheets.Add(After:=Sheets(Sheets.Count)).name = "PIVOT"
Set PivotSheet = twb.Sheets("PIVOT")

Sheets.Add(After:=Sheets(Sheets.Count)).name = "RESUME"
Set shResume = twb.Sheets("RESUME")

Dim DataSrc As Worksheet
Dim PivotTable As PivotTable
Dim PivotCache As PivotCache
Dim PivotRange As Range

Set DataSrc = twb.Sheets("OLAH")
PivotSheet.Activate
For Each PivotTable In PivotSheet.PivotTables
        PivotTable.TableRange2.Clear
Next PivotTable

Set PivotRange = DataSrc.UsedRange
Set PivotCache = twb.PivotCaches.CREATE(SourceType:=xlDatabase, _
                    SourceData:=PivotRange)
Set PivotTable = PivotCache.CreatePivotTable _
                    (TableDestination:=PivotSheet.Range("A1"), _
                    TableName:="PV_MutasiBarang")
                    
'''' INSERT FIELD FIELD NYA ''''

With PivotTable.PivotFields("BUSINESS UNIT")
    .Caption = "BUSINESS UNIT"
    .Orientation = xlRowField
    .Position = 1
End With

With PivotTable.PivotFields("CATEGORY")
    .Caption = "CATEGORY"
    .Orientation = xlRowField
    .Position = 2
End With

Dim fieldName As String
Dim targetField As PivotField
' Nama field yang dicari
fieldName = "KD.BRG"

' Coba gunakan "KD.BRG"
On Error Resume Next
Set targetField = PivotTable.PivotFields(fieldName)
On Error GoTo 0
' Jika "KD.BRG" tidak ditemukan, coba "KD. BRG"
If targetField Is Nothing Then
    fieldName = "KD. BRG"
    On Error Resume Next
    Set targetField = PivotTable.PivotFields(fieldName)
    On Error GoTo 0
End If
' Jika field ditemukan, atur propertinya
If Not targetField Is Nothing Then
    With targetField
        .Caption = fieldName
        .Orientation = xlRowField
        .Position = 3
    End With
End If

' BARIS
'With PivotTable.PivotFields("KD.BRG")
'    .Caption = "KD.BRG"
'    .Orientation = xlRowField
'    .Position = 1
'End With

With PivotTable.PivotFields("SAT")
    .Caption = "SAT"
    .Orientation = xlRowField
    .Position = 4
End With

'KOLOM
With PivotTable.PivotFields("SOURCE")
    .Orientation = xlColumnField
    .Position = 1
End With

'ISI
With PivotTable.PivotFields("SALDO AKHIR")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlSum
    .Caption = "Sum of SALDO AKHIR"
End With

' SETTING
With PivotTable
    .ColumnGrand = False
    .RowGrand = False
    .DisplayErrorString = False
    .NullString = 0
    .PageFieldOrder = 2
    .PreserveFormatting = True
    .PrintTitles = False
    .RepeatItemsOnEachPrintedPage = True
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
    .RepeatAllLabels xlRepeatLabels
End With

With PivotSheet.PivotTables("PV_MutasiBarang")
    .PivotSelect "", xlDataAndLabel, True
    ' RINGKAS ISI ARRAY FALSE SUBTOTAL
    Dim FieldsArray(1 To 12) As Boolean
    For i = 1 To 12
        FieldsArray(i) = False
    Next i
    .PivotFields("BUSINESS UNIT").Subtotals = FieldsArray
    .PivotFields("CATEGORY").Subtotals = FieldsArray
    .PivotFields(fieldName).Subtotals = FieldsArray
    .PivotFields("SAT").Subtotals = FieldsArray
    .PivotFields("SALDO AKHIR").Subtotals = FieldsArray
    .PivotFields("SOURCE").Subtotals = FieldsArray
    .InGridDropZones = True
End With

PivotSheet.PivotTables("PV_MutasiBarang").PivotSelect "", xlDataAndLabel, True

With Selection
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

lr = PivotSheet.Cells.Find(what:="*" _
    , lookat:=xlPart _
    , LookIn:=xlFormulas _
    , Searchorder:=xlByRows _
    , searchdirection:=xlPrevious).Row

Cells(2, 7) = "CEK": Cells(2, 8) = "SELISIH"
Range(Cells(3, 7), Cells(lr, 7)).FormulaR1C1 = "=RC[-1]=RC[-2]"
Range(Cells(3, 8), Cells(lr, 8)).FormulaR1C1 = "=RC[-3]-RC[-2]"
Cells.Select
With Selection
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

PivotSheet.Activate
Sheets(3).UsedRange.Clear
Cells.Copy
Sheets(3).Select
Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
With Cells
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .EntireColumn.AutoFit
End With
Rows(1).Delete

PivotSheet.Activate
Dim xyz As Long, colSelisih As Integer
colSelisih = PivotSheet.Cells.Find("SELISIH", , , xlPart).Column
If PivotSheet.AutoFilterMode Then Selection.AutoFilter
Range("A2").AutoFilter colSelisih, "<>0"

xyz = Application.WorksheetFunction.CountA(Range("A:A").SpecialCells(xlCellTypeVisible))
shResume.UsedRange.Clear
If xyz > 2 Then
    PivotSheet.UsedRange.SpecialCells(xlCellTypeVisible).Copy
    shResume.Activate
    Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Range("A1").Select
    Rows(1).Delete
    Cells.EntireColumn.AutoFit
End If
If PivotSheet.AutoFilterMode Then PivotSheet.AutoFilterMode = False

End Sub